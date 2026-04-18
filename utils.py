"""
Funções auxiliares
"""

from datetime import datetime
from config import CORES_CURSOS, HORA_INICIO_MIN, HORA_FIM_MAX

def validar_horario(hora_str: str) -> bool:
    """Valida se o horário está no formato correto"""
    try:
        datetime.strptime(hora_str, "%H:%M")
        return True
    except ValueError:
        return False

def validar_intervalo_horario(hora_inicio: str, hora_fim: str) -> tuple[bool, str]:
    """
    Valida se o intervalo de horário é válido
    Retorna (válido, mensagem_erro)
    """
    if not validar_horario(hora_inicio):
        return False, "Hora de início inválida. Use formato HH:MM"

    if not validar_horario(hora_fim):
        return False, "Hora de fim inválida. Use formato HH:MM"

    h_ini = datetime.strptime(hora_inicio, "%H:%M")
    h_fim = datetime.strptime(hora_fim, "%H:%M")
    h_min = datetime.strptime(HORA_INICIO_MIN, "%H:%M")
    h_max = datetime.strptime(HORA_FIM_MAX, "%H:%M")

    if h_ini < h_min or h_fim > h_max:
        return False, f"Horário deve estar entre {HORA_INICIO_MIN} e {HORA_FIM_MAX}"

    if h_fim <= h_ini:
        return False, "Hora de fim deve ser maior que hora de início"

    return True, ""

def obter_cor_curso(curso: str) -> str:
    """Retorna a cor de fundo do curso"""
    return formatar_cor_hex(CORES_CURSOS.get(curso.upper(), "#FFFFFF"))

def formatar_cor_hex(cor: str) -> str:
    """Garante que a cor hexadecimal comece com '#' e tenha 6 dígitos (ou use fallback)."""
    if not cor:
        return "#ffffff"
    cor = cor.strip()
    if not cor.startswith('#'):
        cor = f"#{cor}"
    
    # Comprimentos válidos (incluindo o '#'): 4 (#RGB), 7 (#RRGGBB)
    if len(cor) not in [4, 7]:
        return "#ffffff"
    
    # Verifica se os caracteres são hexadecimais válidos
    try:
        int(cor[1:], 16)
        return cor
    except ValueError:
        return "#ffffff"

def hex_to_rgb(hex_color: str) -> tuple:
    """Converte cor hexadecimal para RGB, suportando formatos curtos e longos."""
    cor = formatar_cor_hex(hex_color).lstrip('#')
    
    try:
        if len(cor) == 3:
            # Formato curto: F00 -> FF0000
            cor = "".join([c*2 for c in cor])
        
        return tuple(int(cor[i:i+2], 16) for i in (0, 2, 4))
    except (ValueError, IndexError):
        return (255, 255, 255) # Fallback para branco

def texto_contraste(hex_color: str) -> str:
    """Retorna cor de texto (preto ou branco) com melhor contraste"""
    r, g, b = hex_to_rgb(hex_color)
    # Fórmula de luminância (W3C)
    luminancia = (0.299 * r + 0.587 * g + 0.114 * b) / 255
    return "#000000" if luminancia > 0.5 else "#FFFFFF"

def remover_acentos(texto: str) -> str:
    """Remove acentos e caracteres especiais para comparação robusta."""
    import unicodedata
    if not texto: return ""
    return "".join(
        c for c in unicodedata.normalize('NFD', texto)
        if unicodedata.category(c) != 'Mn'
    ).upper().strip()
