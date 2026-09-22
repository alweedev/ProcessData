"""Casos de uso de estruturas de aprovação: validação de CPF/aprovador,
detecção de colunas da base de carga e cálculo de impacto/remoção."""

import re
from typing import Any

import pandas as pd

from backend.shared.cpf_utils import format_cpf_for_output, is_valid_cpf, limpar_cpf_raw, raw_cpf_digits
from backend.shared.text_utils import upper_no_accents


def _digits_matrix(df: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    """CPF (11 dígitos) de cada célula das colunas dadas; "" onde a célula está vazia.

    Cada valor distinto é normalizado uma única vez: as bases repetem muito o mesmo login e
    `limpar_cpf_raw` (regex) é o custo dominante.
    """
    cache: dict[str, str] = {}
    out: dict[str, pd.Series] = {}
    for col in columns:
        raw = df[col].astype(str).str.strip()
        for value in raw.unique():
            if value and value not in cache:
                cache[value] = limpar_cpf_raw(value)
        out[col] = raw.map(cache).fillna("")
    return pd.DataFrame(out, index=df.index)


def _valor_parece_cpf_viajante(df: pd.DataFrame, cols: dict[str, Any]) -> bool:
    """True se, nas linhas AprovacaoPor=VIAJANTE, a maioria dos valores de `Valor` tiver 11 dígitos.

    Detecta o padrão documentado pela Argo (Valor(VIAJANTE) = LOGIN, que neste sistema é o CPF
    formatado `XXXXXXXXX-XX`) quando não há coluna de CPF dedicada. Nunca usado sem confirmação
    explícita do operador — só sinaliza que dá para propor a confirmação.
    """
    valor_col = cols.get("valor")
    por_col = cols.get("aprovacao_por")
    if not valor_col or not por_col:
        return False
    is_traveler = df[por_col].astype(str).str.strip().str.upper() == "VIAJANTE"
    if not is_traveler.any():
        return False
    valores = df[valor_col].astype(str).str.strip()[is_traveler]
    raw_lens = valores.apply(lambda v: len(raw_cpf_digits(v)))
    validos = int(raw_lens.isin([10, 11]).sum())
    return validos > 0 and validos / len(raw_lens) >= 0.8


def _orphan_rows_mask(
    df_base: pd.DataFrame,
    cpfs: set[str],
    cols: dict[str, Any],
    target_ids: set[str],
    remove_second_level: bool,
) -> pd.Series:
    """Linhas das estruturas-alvo que terminam SEM nenhum aprovador (vetorizado)."""
    approver_cols: list[str] = cols["approver_cols"]
    login_segundo_col = cols.get("login_segundo")
    wanted = list(cpfs)
    in_target = df_base[cols["aprovacao_id"]].astype(str).str.strip().isin(target_ids)
    raw = df_base[approver_cols].astype(str).apply(lambda col: col.str.strip())
    remaining = ((raw != "") & ~_digits_matrix(df_base, approver_cols).isin(wanted)).any(axis=1)
    has_second = pd.Series(False, index=df_base.index)
    if login_segundo_col and login_segundo_col in df_base.columns:
        second_raw = df_base[login_segundo_col].astype(str).str.strip()
        removed = pd.Series(False, index=df_base.index)
        if remove_second_level:
            removed = _digits_matrix(df_base, [login_segundo_col])[login_segundo_col].isin(wanted)
        has_second = (second_raw != "") & ~removed
    return in_target & ~remaining & ~has_second


def _get_or_create_structure(
    store: dict[str, dict[str, Any]],
    row: pd.Series,
    cols: dict[str, Any],
) -> dict[str, Any] | None:
    """Obtém (ou cria) o registro agregado por AprovacaoId."""

    aprov_id_col = cols.get("aprovacao_id")
    if not aprov_id_col:
        return None

    aprov_id = str(row.get(aprov_id_col, "")).strip()
    if not aprov_id:
        return None

    if aprov_id in store:
        return store[aprov_id]

    aprovacao_por_col = cols.get("aprovacao_por")
    aprovacao_col = cols.get("aprovacao")
    tipo_col = cols.get("tipo")
    valor_col = cols.get("valor")
    desc_ccusto_col = cols.get("desc_ccusto")
    cod_ccusto_col = cols.get("cod_ccusto")
    traveler_col = cols.get("traveler_name_col")

    aprovacao_por_val = str(row.get(aprovacao_por_col, "")).strip() if aprovacao_por_col else ""
    aprovacao_val = str(row.get(aprovacao_col, "")).strip() if aprovacao_col else ""
    tipo_val = str(row.get(tipo_col, "")).strip() if tipo_col else ""
    valor_val = str(row.get(valor_col, "")).strip() if valor_col else ""

    traveler_name_raw = str(row.get(traveler_col, "")).strip() if traveler_col else ""
    cod_cc = str(row.get(cod_ccusto_col, "")).strip() if cod_ccusto_col else ""
    desc_cc = str(row.get(desc_ccusto_col, "")).strip() if desc_ccusto_col else ""
    if cod_cc and desc_cc:
        cost_center = f"{cod_cc} - {desc_cc}"
    else:
        cost_center = cod_cc or desc_cc or ""

    aprov_por_upper = aprovacao_por_val.upper()
    if aprov_por_upper == "VIAJANTE":
        traveler_out: str | None = traveler_name_raw or None
        cost_center_out: str | None = None
    elif aprov_por_upper == "CCEMPRESA":
        traveler_out = None
        cost_center_out = cost_center or None
    else:
        traveler_out = traveler_name_raw or None
        cost_center_out = cost_center or None

    record: dict[str, Any] = {
        "aprovacao_id": aprov_id,
        "aprovacao_por": aprovacao_por_val or None,
        "aprovacao": aprovacao_val or None,
        "tipo": tipo_val or None,
        "valor": valor_val or None,
        "traveler_name": traveler_out,
        "cost_center": cost_center_out,
        "positions": [],  # preenchido posteriormente
        "in_second_level": False,
        "occurrences_count": 0,
    }

    store[aprov_id] = record
    return record


def _check_structures_without_approvers(
    df_base: pd.DataFrame,
    cpfs: set[str],
    cols: dict[str, Any],
    target_ids: set[str],
    remove_second_level: bool,
) -> list[dict[str, Any]]:
    """Verifica quais estruturas ficarão sem aprovadores após a remoção dos CPFs.

    Retorna lista de estruturas que ficarão vazias (sem nenhum aprovador).
    """
    approver_cols: list[str] = cols.get("approver_cols") or []
    aprov_id_col = cols.get("aprovacao_id")
    login_segundo_col = cols.get("login_segundo")

    if not aprov_id_col or not approver_cols:
        return []

    structures_without_approvers: list[dict[str, Any]] = []

    # O laço só monta o registro das linhas que de fato ficam vazias (o filtro é vetorizado).
    orphan_mask = _orphan_rows_mask(df_base, cpfs, cols, target_ids, remove_second_level)
    for _idx, row in df_base[orphan_mask].iterrows():
        aprov_id = str(row.get(aprov_id_col, "")).strip()
        if not aprov_id or aprov_id not in target_ids:
            continue

        # Contar aprovadores atuais (excluindo o CPF que será removido)
        remaining_approvers: list[str] = []
        for col in approver_cols:
            raw_login = str(row.get(col, "")).strip()
            if not raw_login:
                continue
            # Se for o CPF que será removido, não conta
            if limpar_cpf_raw(raw_login) in cpfs:
                continue
            remaining_approvers.append(raw_login)

        # Verificar segundo nível (se não estiver sendo removido)
        has_second_level = False
        if login_segundo_col and login_segundo_col in df_base.columns:
            raw_second = str(row.get(login_segundo_col, "")).strip()
            if raw_second:
                # Se remove_second_level=True e o segundo nível é o CPF, não conta
                if remove_second_level and limpar_cpf_raw(raw_second) in cpfs:
                    has_second_level = False
                else:
                    has_second_level = True

        # Se não sobrar nenhum aprovador, adiciona à lista de alertas
        if len(remaining_approvers) == 0 and not has_second_level:
            aprovacao_por_col = cols.get("aprovacao_por")
            valor_col = cols.get("valor")
            desc_ccusto_col = cols.get("desc_ccusto")
            cod_ccusto_col = cols.get("cod_ccusto")
            traveler_col = cols.get("traveler_name_col")

            aprovacao_por_val = str(row.get(aprovacao_por_col, "")).strip() if aprovacao_por_col else ""
            valor_val = str(row.get(valor_col, "")).strip() if valor_col else ""

            # Contexto baseado em AprovacaoPor
            contexto = ""
            if aprovacao_por_val.upper() == "VIAJANTE":
                traveler_name = str(row.get(traveler_col, "")).strip() if traveler_col else ""
                contexto = traveler_name or valor_val
            elif aprovacao_por_val.upper() == "CCEMPRESA":
                cod_cc = str(row.get(cod_ccusto_col, "")).strip() if cod_ccusto_col else ""
                desc_cc = str(row.get(desc_ccusto_col, "")).strip() if desc_ccusto_col else ""
                if cod_cc and desc_cc:
                    contexto = f"{cod_cc} - {desc_cc}"
                else:
                    contexto = cod_cc or desc_cc or valor_val
            else:
                contexto = valor_val

            structures_without_approvers.append(
                {
                    "aprovacaoId": aprov_id,
                    "aprovacaoPor": aprovacao_por_val,
                    "valor": valor_val,
                    "contexto": contexto,
                }
            )

    # Remover duplicados por aprovacaoId
    seen: set[str] = set()
    unique_structures: list[dict[str, Any]] = []
    for s in structures_without_approvers:
        if s["aprovacaoId"] not in seen:
            seen.add(s["aprovacaoId"])
            unique_structures.append(s)

    return unique_structures


def _check_structures_with_duplicate(
    df_base: pd.DataFrame,
    new_cpf_digits: str,
    cols: dict[str, Any],
    target_ids: set[str],
) -> list[dict[str, Any]]:
    """Verifica estruturas onde o novo CPF já é aprovador (main ou segundo nível).

    Substituir um CPF por outro que já aprova a mesma estrutura criaria uma
    duplicidade (o mesmo login em duas posições) — reportado para o usuário
    decidir antes de exportar.
    """
    approver_cols: list[str] = cols.get("approver_cols") or []
    aprov_id_col = cols.get("aprovacao_id")
    login_segundo_col = cols.get("login_segundo")

    if not aprov_id_col:
        return []

    duplicates: list[dict[str, Any]] = []
    for _idx, row in df_base.iterrows():
        aprov_id = str(row.get(aprov_id_col, "")).strip()
        if not aprov_id or aprov_id not in target_ids:
            continue

        already_present = False
        for col in approver_cols:
            raw_login = str(row.get(col, "")).strip()
            if raw_login and limpar_cpf_raw(raw_login) == new_cpf_digits:
                already_present = True
                break
        if not already_present and login_segundo_col and login_segundo_col in df_base.columns:
            raw_second = str(row.get(login_segundo_col, "")).strip()
            if raw_second and limpar_cpf_raw(raw_second) == new_cpf_digits:
                already_present = True

        if not already_present:
            continue

        aprovacao_por_col = cols.get("aprovacao_por")
        valor_col = cols.get("valor")
        desc_ccusto_col = cols.get("desc_ccusto")
        cod_ccusto_col = cols.get("cod_ccusto")
        traveler_col = cols.get("traveler_name_col")

        aprovacao_por_val = str(row.get(aprovacao_por_col, "")).strip() if aprovacao_por_col else ""
        valor_val = str(row.get(valor_col, "")).strip() if valor_col else ""

        contexto = ""
        if aprovacao_por_val.upper() == "VIAJANTE":
            traveler_name = str(row.get(traveler_col, "")).strip() if traveler_col else ""
            contexto = traveler_name or valor_val
        elif aprovacao_por_val.upper() == "CCEMPRESA":
            cod_cc = str(row.get(cod_ccusto_col, "")).strip() if cod_ccusto_col else ""
            desc_cc = str(row.get(desc_ccusto_col, "")).strip() if desc_ccusto_col else ""
            contexto = f"{cod_cc} - {desc_cc}" if cod_cc and desc_cc else (cod_cc or desc_cc or valor_val)
        else:
            contexto = valor_val

        duplicates.append(
            {
                "aprovacaoId": aprov_id,
                "aprovacaoPor": aprovacao_por_val,
                "valor": valor_val,
                "contexto": contexto,
            }
        )

    seen: set[str] = set()
    unique_duplicates: list[dict[str, Any]] = []
    for d in duplicates:
        if d["aprovacaoId"] not in seen:
            seen.add(d["aprovacaoId"])
            unique_duplicates.append(d)

    return unique_duplicates


class ApprovalService:
    @staticmethod
    def normalize_cpf_input(raw_cpf: str | None) -> tuple[str, str]:
        """Normaliza o CPF de entrada.

        Retorna (cpf_digits, cpf_formatado) ou lança ValueError em caso de CPF inválido.
        """

        if not raw_cpf:
            raise ValueError("Informe um CPF para o aprovador.")

        digits = limpar_cpf_raw(raw_cpf)
        if len(digits) != 11:
            raise ValueError("CPF inválido. Informe 11 dígitos.")
        if not is_valid_cpf(digits):
            raise ValueError("CPF inválido (dígito verificador).")

        formatted = format_cpf_for_output(digits)
        return digits, formatted

    @staticmethod
    def load_users_and_find_approver(users_path: str, cpf_digits: str) -> tuple[pd.DataFrame, str]:
        """Carrega base de usuários e retorna o nome completo do aprovador.

        Lança ValueError se CPF não existir na base.
        """

        try:
            df_users = pd.read_excel(users_path, dtype=str).fillna("")
        except Exception as exc:  # pragma: no cover - erro de IO
            raise ValueError(f"Falha ao ler base de usuários: {exc}") from exc

        # Mapa de colunas normalizadas (case-insensitive, sem acentos/separadores)
        norm_cols: dict[str, str] = {}
        for col in df_users.columns:
            key = upper_no_accents(str(col)).replace(" ", "").replace("-", "").replace("_", "")
            norm_cols.setdefault(key, col)

        cpf_col = norm_cols.get("CPF")
        if not cpf_col:
            raise ValueError("Base de usuários não contém coluna 'CPF'.")

        nome_completo_col = norm_cols.get("NOMECOMPLETO")
        nome_col = norm_cols.get("NOME")
        sobrenome_col = norm_cols.get("SOBRENOME")
        if not nome_completo_col and not nome_col:
            raise ValueError("Base de usuários não contém coluna de nome ('NomeCompleto' ou 'Nome').")

        status_col = next((v for k, v in norm_cols.items() if "STATUS" in k), None)

        df_users = df_users.copy()
        df_users["CPFdigits"] = df_users[cpf_col].apply(limpar_cpf_raw)
        matches = df_users[df_users["CPFdigits"] == cpf_digits]
        if matches.empty:
            raise ValueError("CPF não encontrado na base de usuários.")

        row = matches.iloc[0]

        if status_col:
            status_val = upper_no_accents(str(row.get(status_col, ""))).strip()
            if status_val != "ATIVO":
                raise ValueError(f"Aprovador não está ATIVO na base de usuários (Status: '{status_val or 'vazio'}').")

        nome_completo = ""
        if nome_completo_col:
            nome_completo = str(row.get(nome_completo_col, "")).strip()
        if not nome_completo:
            primeiro = str(row.get(nome_col, "")).strip() if nome_col else ""
            sobrenome = str(row.get(sobrenome_col, "")).strip() if sobrenome_col else ""
            nome_completo = f"{primeiro} {sobrenome}".strip()

        return df_users, nome_completo

    @staticmethod
    def valor_parece_cpf_viajante(df: pd.DataFrame, cols: dict[str, Any]) -> bool:
        return _valor_parece_cpf_viajante(df, cols)

    @staticmethod
    def detect_approval_columns(df: pd.DataFrame) -> dict[str, Any]:
        """Detecta colunas relevantes da base de carga de aprovação.

        Usa nomes esperados, mas de forma case-insensitive.
        Detecta dinamicamente LoginAprovador_1..100.
        """

        col_map: dict[str, str] = {}
        for col in df.columns:
            key = upper_no_accents(str(col)).replace(" ", "").replace("-", "").replace("_", "")
            col_map[key] = col

        def pick(*candidates: str) -> str | None:
            for cand in candidates:
                key = upper_no_accents(cand).replace(" ", "").replace("-", "").replace("_", "")
                if key in col_map:
                    return col_map[key]
            return None

        aprovacao_id = pick("AprovacaoId")
        aprovacao_por = pick("AprovacaoPor")
        aprovacao = pick("Aprovacao")
        tipo = pick("Tipo")
        valor = pick("Valor")
        desc_ccusto = pick("DescricaoCCusto", "DescricaoCentroDeCusto", "DescricaoCCustoEmpresa")
        cod_ccusto = pick("CodigoCCusto", "CodigoCentroDeCusto", "CodigoCCustoEmpresa")
        login_segundo = pick("LoginAprovador_SEGUNDO_NIVEL", "LoginAprovadorSegundoNivel")
        segundo_master = pick("SegundoNivelMaster")
        traveler_name_col = pick("NomeViajante", "NomeCompletoViajante", "NomeCompleto")
        traveler_cpf_col = pick("CPFViajante", "CPFDoViajante", "CPF")

        approver_cols: list[str] = []
        for col in df.columns:
            m = re.match(r"(?i)^LoginAprovador_(\d+)$", str(col))
            if m:
                approver_cols.append(col)

        def _slot_num(c: str) -> int:
            match = re.search(r"(\d+)$", str(c))
            return int(match.group(1)) if match else 0

        approver_cols.sort(key=_slot_num)

        return {
            "aprovacao_id": aprovacao_id,
            "aprovacao_por": aprovacao_por,
            "aprovacao": aprovacao,
            "tipo": tipo,
            "valor": valor,
            "desc_ccusto": desc_ccusto,
            "cod_ccusto": cod_ccusto,
            "login_segundo": login_segundo,
            "segundo_master": segundo_master,
            "traveler_name_col": traveler_name_col,
            "traveler_cpf_col": traveler_cpf_col,
            "approver_cols": approver_cols,
        }

    @staticmethod
    def build_preview_for_cpf(
        df_base: pd.DataFrame,
        cpf_digits: str,
        cols: dict[str, Any],
        check_empty: bool = False,
        remove_second_level: bool = False,
    ) -> dict[str, Any]:
        """Gera estruturas afetadas e estatísticas de preview para um CPF."""

        structures: dict[str, dict[str, Any]] = {}

        approver_cols: list[str] = cols.get("approver_cols") or []
        aprov_id_col = cols.get("aprovacao_id")
        if not aprov_id_col or not approver_cols:
            return {
                "structures": [],
                "total_structures": 0,
                "total_occurrences": 0,
                "affected_ids": [],
                "structures_without_approvers": [],
            }

        id_vars: list[str] = []
        for key in [
            "aprovacao_id",
            "aprovacao_por",
            "aprovacao",
            "tipo",
            "valor",
            "desc_ccusto",
            "cod_ccusto",
            "traveler_name_col",
            "login_segundo",
        ]:
            colname = cols.get(key)
            if colname and colname in df_base.columns and colname not in id_vars:
                id_vars.append(colname)

        melted = df_base.melt(
            id_vars=id_vars,
            value_vars=approver_cols,
            var_name="slot_col",
            value_name="login",
        ).fillna("")

        melted["CPFdigits"] = melted["login"].apply(limpar_cpf_raw)
        matches_main = melted[melted["CPFdigits"] == cpf_digits]

        for _, row in matches_main.iterrows():
            rec = _get_or_create_structure(structures, row, cols)
            if not rec:
                continue
            slot_col = str(row.get("slot_col", ""))
            m = re.search(r"(\d+)$", slot_col)
            if m:
                pos = int(m.group(1))
                if pos not in rec["positions"]:
                    rec["positions"].append(pos)
                    rec["occurrences_count"] += 1

        # SEGUNDO_NIVEL
        login_segundo_col = cols.get("login_segundo")
        if login_segundo_col and login_segundo_col in df_base.columns:
            for _, row in df_base.iterrows():
                raw_login = str(row.get(login_segundo_col, "")).strip()
                if not raw_login:
                    continue
                if limpar_cpf_raw(raw_login) != cpf_digits:
                    continue
                rec = _get_or_create_structure(structures, row, cols)
                if not rec:
                    continue
                if not rec["in_second_level"]:
                    rec["in_second_level"] = True
                rec["occurrences_count"] += 1

        affected_ids: set[str] = set(structures.keys())
        total_occurrences = int(sum(rec.get("occurrences_count", 0) for rec in structures.values()))

        ordered_structs = sorted(structures.values(), key=lambda r: str(r.get("aprovacao_id") or ""))

        # Verificar estruturas que ficarão sem aprovadores
        structures_without_approvers: list[dict[str, Any]] = []
        if check_empty and affected_ids:
            structures_without_approvers = _check_structures_without_approvers(
                df_base=df_base,
                cpfs={cpf_digits},
                cols=cols,
                target_ids=affected_ids,
                remove_second_level=remove_second_level,
            )

        return {
            "structures": ordered_structs,
            "total_structures": len(affected_ids),
            "total_occurrences": total_occurrences,
            "affected_ids": sorted(affected_ids),
            "structures_without_approvers": structures_without_approvers,
        }

    @staticmethod
    def remove_cpf_and_compact(
        df_base: pd.DataFrame,
        cpf_digits: str,
        cols: dict[str, Any],
        target_ids: set[str],
        remove_second_level: bool,
    ) -> tuple[pd.DataFrame, dict[str, Any]]:
        """Remove todas as ocorrências do CPF e compacta aprovadores 1..100."""

        return ApprovalService.remove_cpfs_and_compact(df_base, {cpf_digits}, cols, target_ids, remove_second_level)

    @staticmethod
    def remove_cpfs_and_compact(
        df_base: pd.DataFrame,
        cpfs: set[str],
        cols: dict[str, Any],
        target_ids: set[str],
        remove_second_level: bool,
    ) -> tuple[pd.DataFrame, dict[str, Any]]:
        """Remove TODOS os `cpfs` de LoginAprovador_1..100 (e, se pedido, do 2º nível) e compacta.

        Uma passada só: só as linhas das estruturas-alvo em que algum CPF aparece passam pelo laço.
        Não altera `df_base`.
        """
        approver_cols: list[str] = cols.get("approver_cols") or []
        aprov_id_col = cols.get("aprovacao_id")
        if not aprov_id_col or not approver_cols or not cpfs:
            return df_base, {
                "structures_updated": 0,
                "occurrences_removed": 0,
                "changed_indices": set(),
                "promotions": 0,
            }

        df_out = df_base.copy()
        login_segundo_col = cols.get("login_segundo")
        has_segundo = bool(login_segundo_col and login_segundo_col in df_out.columns)
        segundo_master_col = cols.get("segundo_master")
        wanted = list(cpfs)

        in_target = df_out[aprov_id_col].astype(str).str.strip().isin(target_ids)
        has_main = _digits_matrix(df_out, approver_cols).isin(wanted).any(axis=1)
        if has_segundo:
            second_col = str(login_segundo_col)
            has_second = _digits_matrix(df_out, [second_col])[second_col].isin(wanted)
        else:
            has_second = pd.Series(False, index=df_out.index)
        candidates = df_out.index[(in_target & (has_main | has_second)).to_numpy()]

        structures_updated: set[str] = set()
        changed_indices: set[Any] = set()
        occurrences_removed = 0
        promotions = 0

        for idx in candidates:
            row = df_out.loc[idx]
            aprov_id = str(row.get(aprov_id_col, "")).strip()
            changed = False

            if bool(has_main.at[idx]):
                kept: list[str] = []
                for col in approver_cols:
                    value = str(row.get(col, ""))
                    digits = limpar_cpf_raw(value)
                    if digits and digits in cpfs:
                        occurrences_removed += 1
                        changed = True
                        continue
                    if value.strip():
                        kept.append(value)
                for pos, col in enumerate(approver_cols):
                    df_out.at[idx, col] = kept[pos] if pos < len(kept) else ""

            if remove_second_level and has_segundo and bool(has_second.at[idx]):
                df_out.at[idx, login_segundo_col] = ""
                occurrences_removed += 1
                changed = True

            # Se o 1º nível esvaziou e sobrou um 2º nível que NÃO está sendo removido, ele sobe.
            if has_segundo:
                remaining_main = [
                    str(df_out.at[idx, col]).strip() for col in approver_cols if str(df_out.at[idx, col]).strip()
                ]
                current_second = str(df_out.at[idx, login_segundo_col]).strip()
                if not remaining_main and current_second and limpar_cpf_raw(current_second) not in cpfs:
                    df_out.at[idx, approver_cols[0]] = current_second
                    df_out.at[idx, login_segundo_col] = ""
                    promotions += 1
                    changed = True

            if changed:
                structures_updated.add(aprov_id)
                changed_indices.add(idx)

        if segundo_master_col and segundo_master_col in df_out.columns:
            df_out[segundo_master_col] = df_out[segundo_master_col].astype(str).fillna("")

        return df_out, {
            "structures_updated": len(structures_updated),
            "occurrences_removed": int(occurrences_removed),
            "promotions": int(promotions),
            "changed_indices": changed_indices,
        }

    @staticmethod
    def structures_left_without_approvers(
        df_base: pd.DataFrame,
        cpfs: set[str],
        cols: dict[str, Any],
        target_ids: set[str],
        remove_second_level: bool,
    ) -> list[dict[str, Any]]:
        """Estruturas de `target_ids` que ficam SEM nenhum aprovador depois de remover todos os `cpfs`."""
        if not cpfs or not target_ids:
            return []
        return _check_structures_without_approvers(df_base, set(cpfs), cols, set(target_ids), remove_second_level)

    @staticmethod
    def find_traveler_structures(df_base: pd.DataFrame, cpfs: set[str], cols: dict[str, Any]) -> dict[str, set[str]]:
        """`{cpf: {AprovacaoId}}` das estruturas AprovacaoPor=VIAJANTE cujo CPF do viajante é o do usuário."""
        found: dict[str, set[str]] = {cpf: set() for cpf in cpfs}
        aprov_id_col = cols.get("aprovacao_id")
        por_col = cols.get("aprovacao_por")
        if not aprov_id_col or not por_col or not cpfs:
            return found
        is_traveler = df_base[por_col].astype(str).str.strip().str.upper() == "VIAJANTE"
        if not is_traveler.any():
            return found
        cpf_col = cols.get("traveler_cpf_col")
        if not cpf_col:
            raise ValueError("Base de estruturas não contém a coluna CPF do viajante.")
        digits = _digits_matrix(df_base, [cpf_col])[cpf_col]
        ids = df_base[aprov_id_col].astype(str).str.strip()
        hit = is_traveler & digits.isin(list(cpfs)) & (ids != "")
        for idx in df_base.index[hit.to_numpy()]:
            found[digits.at[idx]].add(ids.at[idx])
        return found

    @staticmethod
    def structures_with_foreign_rows(
        df_base: pd.DataFrame, ids: set[str], cpfs: set[str], cols: dict[str, Any]
    ) -> set[str]:
        """Dos `ids`, os que têm ao menos uma linha que NÃO é VIAJANTE de um dos `cpfs`.

        `delete_structures` leva todas as linhas de um AprovacaoId; se o id reúne linhas de outro viajante,
        de CCEMPRESA ou sem CPF, excluir a estrutura inteira apagaria o que ninguém escolheu.
        """
        aprov_id_col = cols.get("aprovacao_id")
        por_col = cols.get("aprovacao_por")
        cpf_col = cols.get("traveler_cpf_col")
        if not ids or not aprov_id_col or not por_col or not cpf_col:
            return set()
        row_ids = df_base[aprov_id_col].astype(str).str.strip()
        in_ids = row_ids.isin(ids)
        is_traveler = df_base[por_col].astype(str).str.strip().str.upper() == "VIAJANTE"
        digits = _digits_matrix(df_base, [cpf_col])[cpf_col]
        own = is_traveler & digits.isin(list(cpfs))
        return set(row_ids[in_ids & ~own])

    @staticmethod
    def find_approver_structures(
        df_base: pd.DataFrame, cpfs: set[str], cols: dict[str, Any]
    ) -> dict[str, dict[str, dict[str, Any]]]:
        """`{cpf: {aprovacaoId: {"posicoes": [1, 3], "segundoNivel": False}}}` — onde cada CPF é aprovador."""
        result: dict[str, dict[str, dict[str, Any]]] = {cpf: {} for cpf in cpfs}
        approver_cols: list[str] = cols.get("approver_cols") or []
        aprov_id_col = cols.get("aprovacao_id")
        if not aprov_id_col or not approver_cols or not cpfs:
            return result
        wanted = list(cpfs)
        ids = df_base[aprov_id_col].astype(str).str.strip()
        digits = _digits_matrix(df_base, approver_cols)
        for col in approver_cols:
            hit = digits[col].isin(wanted) & (ids != "")
            if not hit.any():
                continue
            match = re.search(r"(\d+)$", str(col))
            pos = int(match.group(1)) if match else 0
            for idx in df_base.index[hit.to_numpy()]:
                entry = result[digits.at[idx, col]].setdefault(ids.at[idx], {"posicoes": [], "segundoNivel": False})
                if pos not in entry["posicoes"]:
                    entry["posicoes"].append(pos)
        second_col = cols.get("login_segundo")
        if second_col and second_col in df_base.columns:
            second = _digits_matrix(df_base, [second_col])[second_col]
            hit = second.isin(wanted) & (ids != "")
            for idx in df_base.index[hit.to_numpy()]:
                entry = result[second.at[idx]].setdefault(ids.at[idx], {"posicoes": [], "segundoNivel": False})
                entry["segundoNivel"] = True
        for by_id in result.values():
            for entry in by_id.values():
                entry["posicoes"].sort()
        return result

    @staticmethod
    def delete_structures(df_base: pd.DataFrame, ids: set[str], cols: dict[str, Any]) -> pd.DataFrame:
        """Linhas das estruturas em `ids` com Operacao=DELETE. Não altera `df_base`."""
        aprov_id_col = cols.get("aprovacao_id")
        if not aprov_id_col or not ids:
            return df_base.iloc[0:0].copy()
        rows = df_base[df_base[aprov_id_col].astype(str).str.strip().isin(ids)].copy()
        rows["Operacao"] = "DELETE"
        return rows

    @staticmethod
    def check_new_approver_duplicates(
        df_base: pd.DataFrame,
        new_cpf_digits: str,
        cols: dict[str, Any],
        target_ids: set[str],
    ) -> list[dict[str, Any]]:
        """Estruturas (dentro de target_ids) onde o novo CPF já é aprovador."""

        if not target_ids:
            return []
        return _check_structures_with_duplicate(df_base, new_cpf_digits, cols, target_ids)

    @staticmethod
    def replace_cpf(
        df_base: pd.DataFrame,
        old_cpf_digits: str,
        new_cpf_digits: str,
        cols: dict[str, Any],
        target_ids: set[str],
        replace_second_level: bool,
    ) -> tuple[pd.DataFrame, dict[str, Any]]:
        """Substitui todas as ocorrências do CPF antigo pelo novo, na mesma posição/nível.

        Ao contrário de `remove_cpf_and_compact`, não compacta nem promove
        segundo nível: a posição do aprovador é preservada, só o login muda.
        """

        approver_cols: list[str] = cols.get("approver_cols") or []
        aprov_id_col = cols.get("aprovacao_id")
        if not aprov_id_col or not approver_cols:
            return df_base, {
                "structures_updated": 0,
                "occurrences_replaced": 0,
                "changed_indices": set(),
            }

        df_out = df_base.copy()
        login_segundo_col = cols.get("login_segundo")
        has_segundo = bool(login_segundo_col and login_segundo_col in df_out.columns)
        segundo_master_col = cols.get("segundo_master")
        new_login = format_cpf_for_output(new_cpf_digits)

        structures_updated: set[str] = set()
        changed_indices: set[Any] = set()
        occurrences_replaced = 0

        for idx, row in df_out.iterrows():
            aprov_id = str(row.get(aprov_id_col, "")).strip()
            if not aprov_id or aprov_id not in target_ids:
                continue

            changed = False

            for col in approver_cols:
                raw_login = str(row.get(col, "")).strip()
                if raw_login and limpar_cpf_raw(raw_login) == old_cpf_digits:
                    df_out.at[idx, col] = new_login
                    occurrences_replaced += 1
                    changed = True

            if replace_second_level and has_segundo:
                raw_second = str(row.get(login_segundo_col, "")).strip()
                if raw_second and limpar_cpf_raw(raw_second) == old_cpf_digits:
                    df_out.at[idx, login_segundo_col] = new_login
                    occurrences_replaced += 1
                    changed = True

            if changed:
                structures_updated.add(aprov_id)
                changed_indices.add(idx)

        if segundo_master_col and segundo_master_col in df_out.columns:
            df_out[segundo_master_col] = df_out[segundo_master_col].astype(str).fillna("")

        stats = {
            "structures_updated": len(structures_updated),
            "occurrences_replaced": int(occurrences_replaced),
            "changed_indices": changed_indices,
        }
        return df_out, stats
