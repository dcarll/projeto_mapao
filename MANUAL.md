# 📄 Manual do Usuário e Guia de Instalação - Sistema de Agendamento

Este manual foi desenvolvido para ajudar você a instalar e operar o sistema de agendamento de laboratórios de forma simples e eficiente.

---

## 🚀 1. Preparação (Instalação)

### Passo 1: Instalar o Python
O sistema necessita do Python para funcionar.
1.  Acesse o site oficial: [python.org/downloads](https://www.python.org/downloads/)
2.  Clique no botão **Download Python** (Versão 3.10 ou superior recomendada).
3.  **MUITO IMPORTANTE:** No instalador, marque a caixa **"Add Python to PATH"** antes de clicar em "Install Now".

### Passo 2: Instalar as Bibliotecas
Abra o **Prompt de Comando (CMD)** do Windows e execute o comando abaixo:
```bash
pip install tksheet tkcalendar fpdf2 openpyxl
```
Aguarde o término da instalação de todos os pacotes.

---

## 🖥️ 2. Como Iniciar o Sistema
1.  Navegue até a pasta do projeto.
2.  Localize e abra o arquivo `main.py` (ou execute `python main.py` no terminal).
3.  A janela principal azul marinho deverá aparecer.

---

## 🧭 3. Navegação na Janela Principal

A tela principal exibe o mapa de ocupação dos laboratórios.

*   **Filtros de Visão (Topo):**
    *   **Laboratório:** Selecione um lab específico ou visualize todos simultaneamente.
    *   **Dia da Semana:** Alterna a visualização entre os dias da semana.
    *   **Turno:** Filtro rápido para Manhã, Tarde ou Noite.
    *   **Pesquisa Dinâmica:** Digite nomes de professores ou disciplinas para filtrá-los em tempo real na grade.
*   **Aulas Próximas:** O sistema destaca automaticamente em **Laranja** o horário atual e a aula que está para começar.
*   **Efeito de Foco (Hover):** Ao mover o mouse sobre a grade, uma linha e coluna de destaque aparecerão, indicando exatamente o cruzamento de horário e laboratório.

---

## 🖱️ 4. Menus de Contexto (Clique Direito)

Utilize o botão direito do mouse para acessar funções rápidas:

### Em Espaços Vazios:
*   **Adicionar Aula Aqui:** Abre o formulário de cadastro já pré-preenchido com o laboratório e dia corretos daquela posição.

### Em Aulas Existentes:
1.  **Editar Aula:** Altera dados de uma reserva existente.
2.  **Alterar LAB:** Move a aula para outro laboratório. O sistema fará uma checagem automática de conflitos.
3.  **Excluir Aula:** Remove o agendamento permanentemente.
4.  **Informações:** Exibe todos os detalhes da aula em uma janela de design premium. Você pode selecionar e copiar textos desta janela.

---

## 📝 5. Formulários de Aula

### Detalhes do Cadastro:
*   **Aula Eventual (Toggle):** Ative este interruptor para agendamentos únicos (reposições ou eventos). Desative para aulas fixas que se repetem semanalmente.
*   **Seleção de Curso/Faculdade:** As listas são inteligentes; ao mudar a faculdade, os cursos correspondentes são filtrados.
*   **Observações e Ícones:** Defina observações com pesos (Baixo, Médio, Importante). Observações importantes geram um ícone colorido no "card" da aula.

---

## ⚙️ 6. Gestão e Configurações
Acesse o menu de **Cadastros** para:
*   **Gerenciar Cursos:** Defina as cores de cada curso para facilitar a identificação visual na grade.
*   **Importar Dados (CSV):** Carregue centenas de aulas de uma vez através de um arquivo de planilha.

---

## ❓ 7. Solução de Problemas

*   **Mensagem de Conflito:** O sistema impede que duas aulas ocupem o mesmo lugar e hora. Verifique se o laboratório já não possui uma aula fixa antes de tentar agendar.
*   **A grade não carrega dados:** Verifique se o arquivo `schedule_labs.json` está na pasta. Ele armazena todos os seus dados.
*   **Erro de Cor Inválida:** Se encontrar uma aula com cor estranha, o sistema usará branco por padrão para evitar travamentos.

---
> **Sugestão:** Mantenha os nomes de professores e disciplinas padronizados para que a busca funcione da melhor forma possível!
