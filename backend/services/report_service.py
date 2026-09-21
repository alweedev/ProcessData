import pandas as pd

from backend.domain.rules import REQUIRED_FICHA_FIELDS

# "Sem preenchimento": todos os obrigatórios em branco, com a possível exceção do nome (sem nome a linha nem chega
# aqui: uma linha sem nome, CPF e e-mail é tratada como vazia e descartada).
_WITHOUT_REQUIRED = frozenset(
    f"Campo obrigatório em branco: {label}" for field, label in REQUIRED_FICHA_FIELDS.items() if field != "NomeCompleto"
)


class ReportService:
    @staticmethod
    def build_quality_report(df: pd.DataFrame, errors: dict, file_errors: dict | None = None) -> dict:
        """``file_errors``: ``{nome do arquivo: mensagem}`` dos arquivos que não puderam ser lidos enquanto
        outros foram. Entram no erro geral: sem isso o arquivo sumiria do resultado sem o usuário saber."""
        errors = errors or {}
        # O pipeline deixa em df.attrs o que não cabe nas colunas do arquivo de carga; com um
        # DataFrame montado à mão (sem attrs) o relatório se vira só com o que há nas colunas.
        attrs = df.attrs if df is not None else {}
        row_labels = attrs.get("row_labels", {})
        row_names = attrs.get("row_names", {})
        line_errors = {row_labels.get(k, str(k)): v for k, v in errors.items() if isinstance(k, int)}
        # O mesmo, um item por linha com problema e cada problema separado (o texto de `line_errors` os junta com "; "):
        # é o que a tela agrupa por tipo de problema, com o passageiro de cada linha.
        line_details = []
        for k, message in errors.items():
            if not isinstance(k, int):
                continue
            problems = message.split("; ")
            line_details.append(
                {
                    "label": row_labels.get(k, str(k)),
                    "nome": row_names.get(k, ""),
                    "erros": problems,
                    # Linha sem nenhum obrigatório (só o nome): a tela diz isso numa frase, em vez de listar todos.
                    "sem_preenchimento": _WITHOUT_REQUIRED.issubset(problems),
                }
            )
        general_errors = "; ".join(
            part for part in [errors.get("__geral__"), *(f"{n}: {m}" for n, m in (file_errors or {}).items())] if part
        )

        total_rows = int(df.shape[0]) if df is not None else 0
        invalid_rows = len(line_errors)
        valid_rows = max(0, total_rows - invalid_rows)

        if "duplicated_rows" in attrs:
            # O pipeline remove as repetidas antes de devolver o df: só ele sabe quantas eram.
            duplicated_rows = int(attrs["duplicated_rows"])
        else:
            duplicated_rows = 0
            if df is not None and not df.empty and "Login" in df.columns and "NomeCompleto" in df.columns:
                duplicated_rows = int(df.duplicated(subset=["Login", "NomeCompleto"]).sum())

        # Em branco por campo obrigatório, com o nome que o usuário vê na ficha.
        if "required_blank" in attrs:
            required_blank = dict(attrs["required_blank"])
        else:
            required_blank = {}
            if df is not None and not df.empty:
                for campo, rotulo in REQUIRED_FICHA_FIELDS.items():
                    if campo in df.columns:
                        required_blank[rotulo] = int(df[campo].astype(str).str.strip().eq("").sum())

        return {
            "total_rows": total_rows,
            "valid_rows": valid_rows,
            "invalid_rows": invalid_rows,
            "duplicated_rows": duplicated_rows,
            "general_errors": general_errors or "",
            "line_errors": line_errors,
            "line_details": line_details,
            "required_blank": required_blank,
        }
