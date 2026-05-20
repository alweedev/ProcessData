# backend/utils.py
import re
import unicodedata

def normalize_text(s):
    if s is None:
        return ""
    return str(s).strip()

def upper_no_accents(s):
    s = normalize_text(s).upper()
    s = unicodedata.normalize("NFKD", s).encode("ASCII", "ignore").decode("utf-8")
    return s

def limpar_cpf_raw(cpf):
    """retorna apenas digitos do CPF"""
    if cpf is None:
        return ""
    s = str(cpf)
    return re.sub(r"\D", "", s)

def format_cpf_for_output(cpf_digits):
    """formata com traço antes dos 2 ultimos digitos: 12345678901 -> 123456789-01"""
    if not cpf_digits:
        return ""
    s = re.sub(r"\D", "", str(cpf_digits))
    if len(s) == 11:
        return f"{s[:-2]}-{s[-2:]}"
    return s

def validar_extensao_arquivo(filename: str, allowed_extensions: set = None) -> tuple[bool, str]:
    """Valida se a extensão do arquivo é permitida.
    
    Args:
        filename: Nome do arquivo com extensão
        allowed_extensions: Conjunto de extensões permitidas (ex: {'.xlsx', '.xls'})
                           Se None, usa padrão {'.xlsx', '.xls', '.xltx'}
    
    Returns:
        (bool, str): (é_válido, mensagem_erro)
        
    Examples:
        >>> validar_extensao_arquivo('dados.xlsx')
        (True, '')
        
        >>> validar_extensao_arquivo('script.txt')
        (False, 'Extensão não permitida. Aceitos: .xlsx, .xls, .xltx')
    """
    if allowed_extensions is None:
        allowed_extensions = {'.xlsx', '.xls', '.xltx'}
    
    if not filename:
        return False, "Nenhum arquivo fornecido"
    
    _, ext = filename.rsplit('.', 1) if '.' in filename else ('', '')
    ext = f".{ext.lower()}" if ext else ""
    
    if not ext:
        return False, "Arquivo sem extensão"
    
    if ext not in allowed_extensions:
        exts_str = ", ".join(sorted(allowed_extensions))
        return False, f"Extensão não permitida. Aceitos: {exts_str}"
    
    return True, ""
