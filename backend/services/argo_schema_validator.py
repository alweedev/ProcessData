"""Valida a base de estruturas exportada contra o schema oficial da carga de aprovação da Argo.

Regras de `Formulario Requisitos - Carga Aprovacao.pdf`: campos obrigatórios, valores aceitos e
tamanhos máximos. Roda só antes do download (nunca bloqueia a análise) — se falhar aqui é sinal de
bug na geração, não de dado ruim do cliente.
"""

import re

import pandas as pd

_OPERACAO_VALIDA = {"INSERT", "UPDATE", "DELETE"}
_APROVACAO_POR_VALIDA = {
    "VIAJANTE", "COMUNIDADE", "MOTIVO", "CCCLIENTE", "CCEMPRESA", "CLIENTE", "EMPRESA", "PROJETO",
}
_TIPO_VALIDO = {"U", "S", "P", "N"}
_LOGIN_MAX = 12


class ArgoSchemaValidator:
    @staticmethod
    def validar(df: pd.DataFrame) -> list[str]:
        erros: list[str] = []
        if df.empty:
            return erros

        if "Operacao" not in df.columns or (df["Operacao"].astype(str).str.strip() == "").any():
            erros.append("Operacao: há linha(s) sem valor (obrigatório: INSERT, UPDATE ou DELETE).")
        elif not df["Operacao"].astype(str).str.strip().str.upper().isin(_OPERACAO_VALIDA).all():
            erros.append("Operacao: há valor fora de INSERT, UPDATE ou DELETE.")

        if "AprovacaoPor" in df.columns:
            preenchidos = df["AprovacaoPor"].astype(str).str.strip()
            invalidos = preenchidos[(preenchidos != "") & ~preenchidos.str.upper().isin(_APROVACAO_POR_VALIDA)]
            if not invalidos.empty:
                erros.append("AprovacaoPor: há valor fora da lista aceita pela Argo.")

        if "Tipo" in df.columns:
            preenchidos = df["Tipo"].astype(str).str.strip()
            invalidos = preenchidos[(preenchidos != "") & ~preenchidos.str.upper().isin(_TIPO_VALIDO)]
            if not invalidos.empty:
                erros.append("Tipo: há valor fora de U, S, P ou N.")

        for col in df.columns:
            if re.match(r"(?i)^LoginAprovador_\d+$", str(col)):
                grandes = df[col].astype(str).str.len() > _LOGIN_MAX
                if grandes.any():
                    erros.append(f"{col}: há login com mais de {_LOGIN_MAX} caracteres.")

        if "Status" in df.columns and (df["Status"].astype(str).str.strip() != "").any():
            erros.append("Status: deve vir vazio na carga.")

        return erros
