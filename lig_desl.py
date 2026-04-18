import tkinter as tk
from tkinter import ttk
import os
from database import Database

# Paleta dark global (Sincronizada com gui.py)
_D = {
    "bg":      "#1e2540",   # fundo principal
    "card":    "#252d4a",   # fundo de cards
    "input":   "#2e3760",   # fundo dos campos
    "border":  "#3d4f82",   # borda
    "primary": "#3d5af1",   # azul primário
    "success": "#10b981",   # verde
    "danger":  "#ef4444",   # vermelho
    "warning": "#f59e0b",   # âmbar
    "fg":      "#e8ecf8",   # texto principal
    "fg2":     "#8f9dc7",   # texto secundário
    "hdr":     "#161d36",   # cabeçalho dark
}

# Estados possíveis e suas cores (Texto, BG_Fundo)
STATE_THEME = {
    "Ligado":    {"text": "Ligado",     "fg": "#ffffff", "bg": "#10b981"},
    "Desligado": {"text": "Desligado",  "fg": "#ffffff", "bg": "#ef4444"},
    "Finalizado": {"text": "Finalizado", "fg": "#ffffff", "bg": "#4b5563"}, # cinza escuro
}

STATE_ORDER = ["Desligado", "Ligado", "Finalizado"]


class LabControlApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Controle de Laboratórios - Status")
        self.geometry("600x800")
        self.configure(bg=_D["bg"])

        # Ajuste para ícone no Windows (Taskbar)
        try:
            import ctypes
            myappid = 'schedulelabs.labstatus.1.0'
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        except Exception:
            pass

        # Configura ícone se existir
        try:
            from PIL import Image, ImageTk
            script_dir = os.path.dirname(os.path.abspath(__file__))
            icon_path = os.path.join(script_dir, "icone", "ico-raio.png")
            
            if os.path.exists(icon_path):
                img_pil = Image.open(icon_path)
                # Redimensiona para icon standard se necessario (ex: 32x32)
                img_pil = img_pil.resize((32, 32), Image.Resampling.LANCZOS)
                icon_img = ImageTk.PhotoImage(img_pil)
                self.iconphoto(True, icon_img)
                self._icon_ref = icon_img # Manter referência
            else:
                print(f"Icone nao encontrado em: {icon_path}")
        except Exception as e:
            print(f"Erro ao carregar icone: {e}")
            # Fallback para PhotoImage basico se PIL falhar
            try:
                icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icone", "ico-raio.png")
                if os.path.exists(icon_path):
                    img = tk.PhotoImage(file=icon_path)
                    self.iconphoto(True, img)
                    self._icon_ref = img
            except:
                pass

        # Lista de laboratórios
        self.labs = [
            "Lab 01", "Lab 02", "Lab 03", "Lab 04",
            "Lab 05", "Lab 06", "Lab 07", "Lab 08",
            "Lab 09", "Lab 10", "Lab 11", "Lab 12", "Lab 13", "Lab 14"
        ]

        self.state_labels = {}   # lab_name -> tk.Label (estado)
        self.state_index = {}    # lab_name -> índice atual em STATE_ORDER
        self.action_widgets = {} # lab_name -> combobox da ação

        self.setup_styles()
        self.db = Database()
        self.create_widgets()
        self.refresh_from_db() # Inicia loop de sincronização

    def setup_styles(self):
        style = ttk.Style()
        if 'clam' in style.theme_names():
            style.theme_use('clam')

        # Estilo para Combobox
        style.configure("TCombobox", 
                        fieldbackground=_D["input"], 
                        background=_D["input"],
                        foreground=_D["fg"],
                        arrowcolor=_D["fg2"],
                        padding=5)
        
        style.map("TCombobox",
                  fieldbackground=[("readonly", _D["input"])],
                  foreground=[("readonly", _D["fg"])])

        # Estilo para botões se necessário (não usados no momento mas bom ter)
        style.configure("TButton", font=("Segoe UI", 9, "bold"), padding=5)

    def create_widgets(self):
        # Container principal com padding
        main_container = tk.Frame(self, bg=_D["bg"], padx=20, pady=20)
        main_container.pack(fill=tk.BOTH, expand=True)

        header_font = ("Segoe UI", 11, "bold")
        row_font = ("Segoe UI", 10)

        # Frame de cabeçalho
        header_frame = tk.Frame(main_container, bg=_D["hdr"], height=45)
        header_frame.grid(row=0, column=0, columnspan=3, sticky="nsew", pady=(0, 2))
        header_frame.grid_propagate(False)

        tk.Label(header_frame, text="LABORATÓRIOS", font=header_font, 
                 bg=_D["hdr"], fg=_D["fg2"]).pack(side=tk.LEFT, expand=True, fill=tk.BOTH)
        
        tk.Label(header_frame, text="STATUS", font=header_font, 
                 bg=_D["hdr"], fg=_D["fg2"]).pack(side=tk.LEFT, expand=True, fill=tk.BOTH)

        # Coluna 3: Horário (Label)
        self.header_entry = tk.Label(header_frame, font=header_font, justify="center",
                                     bg=_D["input"], fg=_D["warning"], 
                                     height=1, width=12)
        
        # Carrega horário inicial do banco
        db_h = self.db.obter_status_horario()
        self.header_entry.config(text=db_h)
        self.header_entry.pack(side=tk.LEFT, padx=10, pady=5)
        

        # Grid de Laboratórios
        row_start = 1
        for i, lab in enumerate(self.labs):
            row = row_start + i
            
            # Fundo alternado
            row_bg = _D["card"] if i % 2 == 0 else _D["row_alt"] if "row_alt" in _D else "#2a3250"

            # Coluna 1: Nome do Lab
            f1 = tk.Frame(main_container, bg=row_bg, highlightthickness=1, highlightbackground=_D["border"])
            f1.grid(row=row, column=0, sticky="nsew", padx=1, pady=1)
            tk.Label(f1, text=lab, font=row_font, bg=row_bg, fg=_D["fg"]).pack(padx=10, pady=10)

            # Coluna 2: Estado
            # Busca status inicial do banco
            db_status = self.db.obter_todos_status_labs()
            current_status = db_status.get(lab, STATE_ORDER[0])
            try:
                self.state_index[lab] = STATE_ORDER.index(current_status)
            except ValueError:
                self.state_index[lab] = 0
                current_status = STATE_ORDER[0]

            theme = STATE_THEME[current_status]

            f2 = tk.Frame(main_container, bg=theme["bg"], cursor="hand2", 
                          highlightthickness=1, highlightbackground=_D["border"])
            f2.grid(row=row, column=1, sticky="nsew", padx=1, pady=1)
            
            lbl_state = tk.Label(f2, text=theme["text"], fg=theme["fg"],
                                 font=("Segoe UI", 10, "bold"), bg=theme["bg"],
                                 cursor="hand2")
            lbl_state.pack(padx=10, pady=10, expand=True, fill=tk.BOTH)
            
            self.state_labels[lab] = (f2, lbl_state) # Guardamos o frame e o label
            
            # Bind no label e no frame
            lbl_state.bind("<Button-1>", lambda e, ln=lab: self.toggle_state(ln))
            f2.bind("<Button-1>", lambda e, ln=lab: self.toggle_state(ln))

            # Coluna 3: Ações
            f3 = tk.Frame(main_container, bg=row_bg, highlightthickness=1, highlightbackground=_D["border"])
            f3.grid(row=row, column=2, sticky="nsew", padx=1, pady=1)
            
            actions = ["", "Organizar", "Ligar", "Desligar", "Finalizar"]
            cb = ttk.Combobox(f3, values=actions, state="readonly", width=12)
            
            # Busca ação inicial do banco
            db_acoes = self.db.obter_todas_acoes_labs()
            v_acao = db_acoes.get(lab, "")
            if v_acao in actions:
                cb.set(v_acao)
            else:
                cb.current(0)
            
            cb.pack(padx=10, pady=10)
            self.action_widgets[lab] = cb
            
            # Salva ao alterar e atualiza cor
            cb.bind("<<ComboboxSelected>>", lambda e, ln=lab: self.on_action_selected(ln))
            
            # Aplica cor inicial
            self._update_action_color(lab, v_acao)

        # Configuração de pesos
        main_container.grid_columnconfigure(0, weight=1)
        main_container.grid_columnconfigure(1, weight=1)
        main_container.grid_columnconfigure(2, weight=1)

    def toggle_state(self, lab_name: str):
        idx = self.state_index[lab_name]
        idx = (idx + 1) % len(STATE_ORDER)
        self.state_index[lab_name] = idx

        new_state = STATE_ORDER[idx]
        theme = STATE_THEME[new_state]

        frame, lbl = self.state_labels[lab_name]
        lbl.config(text=theme["text"], bg=theme["bg"], fg=theme["fg"])
        frame.config(bg=theme["bg"])
        
        # Salva no banco compartilhado
        self.db.atualizar_status_lab(lab_name, new_state)

    def on_action_selected(self, lab_name):
        """Callback quando uma ação é selecionada no Combobox."""
        action = self.action_widgets[lab_name].get()
        self.db.atualizar_acao_lab(lab_name, action)
        self._update_action_color(lab_name, action)

    def _update_action_color(self, lab_name, action):
        """Atualiza o fundo da célula (frame) baseado na ação selecionada."""
        cb = self.action_widgets.get(lab_name)
        if not cb: return
        
        frame = cb.master # O frame f3 que contém o combo
        
        # Mapeamento de cores solicitado
        colors = {
            "Ligar":     "#10b981", # Verde
            "Organizar": "#f59e0b", # Laranja (Warning)
            "Finalizar": "#000000", # Preto
            "Desligar":  "#ef4444", # Vermelho
        }
        
        # Cor padrão (alternada) se não houver ação selecionada
        # Recupera o índice do lab para saber se é par/ímpar para o fundo padrão
        try:
            i = self.labs.index(lab_name)
            default_bg = _D["card"] if i % 2 == 0 else _D.get("row_alt", "#2a3250")
        except:
            default_bg = _D["card"]

        bg = colors.get(action, default_bg)
        frame.config(bg=bg)

    def refresh_from_db(self):
        """Atualiza a UI com dados do banco (Sync rede)."""
        try:
            db_status = self.db.obter_todos_status_labs()
            for lab_name, status in db_status.items():
                if lab_name in self.state_labels:
                    # Só atualiza se for diferente do que está na tela (evita flicker)
                    if STATE_ORDER[self.state_index[lab_name]] != status:
                        try:
                            idx = STATE_ORDER.index(status)
                            self.state_index[lab_name] = idx
                            
                            theme = STATE_THEME[status]
                            frame, lbl = self.state_labels[lab_name]
                            lbl.config(text=theme["text"], bg=theme["bg"], fg=theme["fg"])
                            frame.config(bg=theme["bg"])
                        except ValueError:
                            pass
        except Exception as e:
            print(f"Erro no log_desl sync status: {e}")

        # Sincroniza as Ações (Comboboxes)
        try:
            db_acoes = self.db.obter_todas_acoes_labs()
            for lab_name, cb in self.action_widgets.items():
                v_db = db_acoes.get(lab_name, "")
                # Só atualiza se for diferente e se o usuário não estiver com o foco nele
                if cb.get() != v_db and self.focus_get() != cb:
                    cb.set(v_db)
                    self._update_action_color(lab_name, v_db)
        except Exception as e:
            print(f"Erro no log_desl sync acoes: {e}")

        # Sincroniza o campo de Horário do cabeçalho
        try:
            db_h = self.db.obter_status_horario()
            # Só atualiza se for diferente da tela para evitar flicker
            if self.header_entry.cget("text") != db_h:
                self.header_entry.config(text=db_h)
        except Exception as e:
            print(f"Erro no log_desl sync horario: {e}")
            
        # Agenda próxima atualização (2 segundos)
        self.after(2000, self.refresh_from_db)


if __name__ == "__main__":
    app = LabControlApp()
    app.mainloop()
