# 🏨 Sistema de Gerenciamento de Laboratórios (Mapão de Aulas)

![Status](https://img.shields.io/badge/Status-Finalizado-success?style=for-the-badge)
![Python](https://img.shields.io/badge/Python-3.10+-blue?style=for-the-badge&logo=python)
![Interface](https://img.shields.io/badge/Interface-Tkinter_Premium-orange?style=for-the-badge)

Um sistema completo e elegante para gestão de agendamentos de laboratórios, projetado para oferecer visibilidade total e controle preciso sobre o uso de salas, horários e recursos acadêmicos.

---

## ✨ Principais Funcionalidades

*   **📅 Grade Inteligente**: Visualização completa em tempo real de todos os laboratórios ou dias específicos. 
*   **🔍 Busca Dinâmica**: Encontre professores, disciplinas ou turmas instantaneamente através da barra de pesquisa em tempo real.
*   **⚡ Clique Rápido (Context Menu)**: Adicione, edite, mova ou exclua aulas clicando apenas com o botão direito diretamente na grade.
*   **🎯 Efeito Cross-Hover**: Navegação facilitada com destaque automático de linha (horário) e coluna (laboratório) ao passar o mouse.
*   **⚠️ Sensor de Conflitos**: O sistema impede automaticamente agendamentos duplicados no mesmo laboratório e horário.
*   **🌗 Aula Eventual**: Suporte para agendamentos pontuais, reposições e eventos sem interferir na grade fixa.
*   **📊 Importação CSV**: Migre seus dados facilmente importando planilhas inteiras de aulas em segundos.
*   **🎨 Gestão de Cores**: Personalize as cores de cada curso para facilitar a identificação visual rápida.

---

## 🛠️ Tecnologias Utilizadas

*   **Linguagem**: [Python](https://www.python.org/)
*   **Interface Gráfica**: Tkinter com temas personalizados e componentes modernos.
*   **Banco de Dados**: JSON (Persistência leve e rápida).
*   **Bibliotecas**:
    *   `tksheet`: Para tabelas dinâmicas.
    *   `tkcalendar`: Para seleção intuitiva de datas.
    *   `fpdf2`: Para geração de relatórios.
    *   `openpyxl`: Para manipulação de arquivos Excel.

---

## 🚀 Instalação e Uso

### Pré-requisitos
1. Certifique-se de ter o **Python 3.10 ou superior** instalado.
2. No ato da instalação do Python, marque a opção **"Add Python to PATH"**.

### Instalar dependências
Abra o terminal na pasta do projeto e execute:
```bash
pip install tksheet tkcalendar fpdf2 openpyxl
```

### Rodando o programa
Basta executar o arquivo principal:
```bash
python main.py
```

---

## 📖 manuais
Para um guia detalhado de como gerenciar cada função, consulte nosso [MANUAL DO USUÁRIO](./MANUAL.md).

---

## 👤 Autor
Desenvolvido para facilitar a logística acadêmica e otimizar o tempo de gestores e professores.📦

---
*Este projeto é parte do Agendamento de Aulas Final Edition.*
