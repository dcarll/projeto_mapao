"""
Modelos de dados
"""

from dataclasses import dataclass
from typing import Optional

@dataclass(unsafe_hash=True)
class Aula:
    """Representa uma aula agendada"""
    id: Optional[int]
    laboratorio: str
    dia_semana: str
    hora_inicio: str
    hora_fim: str
    disciplina: str
    turma: str
    professor: str
    qtde_alunos: int
    faculdade: str
    curso: str
    cor_fundo: str
    observacoes: str = ""
    peso_observacao: str = "Baixo"
    is_eventual: bool = False
    data_eventual: Optional[str] = None

@dataclass(unsafe_hash=True)
class Laboratorio:
    """Representa um laboratório"""
    id: Optional[int]
    nome: str
    descricao: str
    observacao: str
    qtd_micros: int = 0
    planta_path: Optional[str] = None
    inventario: Optional[dict] = None # Dados técnicos do lab

    def __str__(self):
        return self.nome
