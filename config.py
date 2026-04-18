"""
Configurações do Sistema de Agendamento de Laboratórios
"""

import os

import sys

# Determina onde salvar os dados
def get_data_path():
    # Se estiver rodando como executável PyInstaller temporário (_MEIPASS existe)
    if hasattr(sys, '_MEIPASS'):
        return "C:\\SCHEDULE_LABS"
    
    # Se estiver rodando instalado (onde os arquivos .py estão)
    # Pegamos a pasta onde o config.py (ou main.py) está localizado
    return os.path.dirname(os.path.abspath(__file__))

PASTA_SISTEMA = get_data_path()

# Arquivo JSON que armazena todos os dados
DATABASE_PATH = os.path.join(PASTA_SISTEMA, "schedule_labs.json")

# Configurações de horário
HORA_INICIO_MIN = "07:30"
HORA_FIM_MAX = "23:15"

# Laboratórios disponíveis
LABORATORIOS = [f"{i:02d}" for i in range(1, 15)]

# Dias da semana
DIAS_SEMANA = [
    "SEGUNDA-FEIRA",
    "TERÇA-FEIRA",
    "QUARTA-FEIRA",
    "QUINTA-FEIRA",
    "SEXTA-FEIRA",
    "SÁBADO"
]

# Faculdades (ordem alfabética)
FACULDADES = []

# Mapeamento Faculdade -> Cursos
CURSOS_POR_FACULDADE = {}

# Cores por curso (background em hexadecimal)
CORES_CURSOS = {}

# Pesos para observações
PESOS_OBSERVACAO = ["Não mostrar", "Baixo", "Média", "Importante"]
