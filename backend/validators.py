# backend/validators.py
import logging
from typing import List, Dict
from .utils import limpar_cpf_raw, format_cpf_for_output, upper_no_accents

MODEL_COLS = [
    "Operacao","UserId","Login","CodigoCCustoCliente","DescricaoCCustoCliente",
    "NomeEmpresa","CodigoCCustoEmpresa","DescricaoCCustoEmpresa","EmpresaCCustoParaUsuario",
    "NroMatricula","Nome","SobreNome","NomeCompleto","Email","Telefone","Cargo","Departamento","Nivel",
    "Endereco","Cidade","Estado","CEP","Solicitante","Vip","ViajanteMasterNacional",
    "ViajanteMasterInternacional","SolicitanteMaster","MasterAdiantamento","MasterReembolso","Terceiro",
    "CodigoIntegracao","Status","Observacoes"
]

logger = logging.getLogger("robo_backend")

def validar_linha(reg):
    """
    reg: dict com campos extraidos
    retorna lista de mensagens de validação (vazia se ok)
    """
    msgs = []
    
    # Solicitante obrigatório (deve ser 'S' ou 'N')
    solicitante = str(reg.get("Solicitante", "")).strip().upper()
    if solicitante not in ("S", "N"):
        msgs.append("Solicitante obrigatório (deve ser S ou N)")
    
    # CPF se existir
    cpf_raw = reg.get("CPF", "") or reg.get("Login", "")
    digits = limpar_cpf_raw(cpf_raw)
    if digits and len(digits) != 11:
        msgs.append("CPF deve ter 11 dígitos")
    elif not digits and "CPF" in reg and reg["CPF"]:  # CPF vazio ou inválido
        logger.warning("CPF ausente ou inválido para registro: %s", reg.get("NomeCompleto", "desconhecido"))

    # Email simples (opcional)
    email = reg.get("Email","").strip()
    if email and ("@" not in email or "." not in email.split("@")[-1]):
        msgs.append("Email inválido")
    elif not email and "Email" in reg:  # Email vazio mas esperado
        logger.warning("Email ausente para registro: %s", reg.get("NomeCompleto", "desconhecido"))

    # Nome completo
    nomec = reg.get("NomeCompleto","").strip()
    if not nomec:
        msgs.append("NomeCompleto vazio")

    # Nivel: deve ser OPERACIONAL, GERENCIA, DIRETORIA ou vazio
    nivel = upper_no_accents(reg.get("Nivel",""))
    if nivel and nivel not in ("OPERACIONAL","GERENCIA","DIRETORIA"):
        # se tiver outro texto, tentar mapear
        if "OPER" in nivel:
            reg["Nivel"] = "OPERACIONAL"
        elif "GER" in nivel:
            reg["Nivel"] = "GERENCIA"
        elif "DIR" in nivel:
            reg["Nivel"] = "DIRETORIA"
        else:
            reg["Nivel"] = ""
            msgs.append("Nivel inválido, ajustado para vazio")

    return msgs

def validar_dataframe_for_output(df):
    """
    df: pandas DataFrame final
    retorna lista de mensagens gerais
    """
    msgs = []
    # verificar colunas obrigatórias
    req = ["Operacao","Login","NomeCompleto","Nome","SobreNome","CodigoIntegracao","EmpresaCCustoParaUsuario"]
    for c in req:
        if c not in df.columns:
            msgs.append(f"Coluna obrigatória ausente: {c}")
    return msgs