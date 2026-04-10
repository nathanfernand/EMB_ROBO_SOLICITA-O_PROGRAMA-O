# =============================================================================
# config.py — Constantes globais, caminhos e mensagens de STATUS
# =============================================================================

# ---------------------------------------------------------------------------
# Caminhos dos recursos
# ---------------------------------------------------------------------------
SAP_GUI_PATH: str    = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplgpad.exe"
EXCEL_PATH: str      = r"C:\base\ZLOLMM025.xlsx"
TXT_EXPORT_PATH: str = r"C:\base\ZDPQPL126.txt"

SAP_SERVER_NAME: str = "02- EBP - SAP Corp (FI/CO)"

# ---------------------------------------------------------------------------
# Mensagens de STATUS gravadas na planilha
# ---------------------------------------------------------------------------
class Status:
    SOLICITADA         = "ORDEM SOLICITADA!"
    SOLICITADA_MTS     = "ORDEM SOLICITADA VIA ZLOLMM027"
    ORDEM_NAO_SINC     = "ORDEM NÃO SINCRONIZADA NO MES, ACIONAR PPCP"
    OPERACAO_INCORRETA = "OPERAÇÃO INCORRETA P2S, ACIONAR PPCP"
    ORDEM_CONGELADA    = "ORDEM CONGELADA, ACIONAR PPCP"
    TXT_INCOMPLETO     = "ARQUIVO TXT INCOMPLETO, ACIONAR PPCP"
    TXT_NAO_ENCONTRADO = "ARQUIVO TXT NÃO ENCONTRADO, ACIONAR PPCP"
    ERRO_LEITURA_TXT   = "ERRO NA LEITURA DO TXT, ACIONAR PPCP"
    FALHA_POPUP        = "FALHA AO CONFIRMAR POPUP, ACIONAR PPCP"
    FALHA_MTS          = "FALHA NA SOLICITAÇÃO ORDEM ZTMS, ACIONAR PPCP"


# =============================================================================
# DICIONÁRIO DE TIPOS DE CARRO
# =============================================================================
# Regra de roteamento:
#   • Código PRESENTE neste dicionário → processa via ZDPQPL126_MONTADOR
#   • Código AUSENTE  neste dicionário → processa via ZDPQPL126_FABRICANTE
# =============================================================================
TIPO_CARRO: dict[str, str] = {
    "I1K": "Montador",
    "I1M": "Montador",
    "I1O": "Montador",
    "I1Q": "Montador",
    "I1S": "Montador",
    "I2K": "Montador",
    "I2O": "Montador",
    "I2Q": "Montador",
    "I2S": "Montador",
    "I2Y": "Montador",
    "I3J": "Montador",
    "I3Q": "Montador",
    "I3S": "Montador",
    "I4S": "Montador",
    "I4Y": "Montador",
    "IA2": "Montador",
    "IAR": "Montador",
    "IAZ": "Montador",
    "IB2": "Montador",
    "IBJ": "Montador",
    "I4J": "Montador",
    "I5Q": "Montador",
    "IA3": "Montador",
    "IA4": "Montador",
    "IAK": "Montador",
    "IAL": "Montador",
    "IAM": "Montador",
    "IAN": "Montador",
    "IAP": "Montador",
    "IAT": "Montador",
    "IAU": "Montador",
    "IAV": "Montador",
    "IAW": "Montador",
    "ICJ": "Montador",
    "IDC": "Montador",
    "IL1": "Montador",
    "IL2": "Montador",
    "IT1": "Montador",
    "IT2": "Montador",
    "IAX": "Montador",
    "IBK": "Montador",
    "IW1": "Montador",
    "I2G": "Montador",
    "I1V": "Montador",
    "I2J": "Montador",
    "I1J": "Montador",
    "I3K": "Montador",
    "I4K": "Montador",
    "I1Y": "Montador",
    "I3Y": "Montador",
}

# =============================================================================
# DICIONÁRIO DE CARROS — ZLOLMM027
# =============================================================================
# Códigos de carro processados diretamente via ZLOLMM027 (antes de FABRICANTE).
# =============================================================================
CARRO_ZLOLMM027: dict[str, str] = {
    "MMU": "ZLOLMM027",
    "I1D": "ZLOLMM027",
    "MMX": "ZLOLMM027",
    "MMW": "ZLOLMM027",
    "MMY": "ZLOLMM027",
    "MMZ": "ZLOLMM027",
    "Z1R": "ZLOLMM027",
    "ZAR": "ZLOLMM027",
    "KVC": "ZLOLMM027",
    "KVA": "ZLOLMM027",
    "JAK": "ZLOLMM027",
    "JAD": "ZLOLMM027",
    "JAB": "ZLOLMM027",
    "JAJ": "ZLOLMM027",
    "JAI": "ZLOLMM027",
    "JAG": "ZLOLMM027",
    "JAH": "ZLOLMM027",
    "JAF": "ZLOLMM027",
    "JAE": "ZLOLMM027",
    "JAC": "ZLOLMM027",
    "SO9": "ZLOLMM027",
    "SO6": "ZLOLMM027",
    "SO5": "ZLOLMM027",
    "SO4": "ZLOLMM027",
    "SO3": "ZLOLMM027",
    "SO2": "ZLOLMM027",
    "SO1": "ZLOLMM027",
    "KS1": "ZLOLMM027",
    "KS2": "ZLOLMM027",
    "KS4": "ZLOLMM027",
    "KS5": "ZLOLMM027",
    "KS3": "ZLOLMM027",
    "KS6": "ZLOLMM027",
    "KS7": "ZLOLMM027",
    "KS8": "ZLOLMM027",
    "KC2": "ZLOLMM027",
    "TCA": "ZLOLMM027",
    "KC5": "ZLOLMM027",
    "IAD": "ZLOLMM027",
    "IEB": "ZLOLMM027", 
    "ILA": "ZLOLMM027",
}
