"""
Interface gráfica do sistema
"""

import tkinter as tk
import tksheet
from tkinter import ttk, messagebox, colorchooser, filedialog
import csv
import ctypes # Para correção do ícone no Windows
from typing import Optional, List, Any
from database import Database
from models import Aula, Laboratorio
from utils import validar_intervalo_horario, obter_cor_curso, hex_to_rgb, texto_contraste, remover_acentos
from config import (
    LABORATORIOS, DIAS_SEMANA
)
import math
import os
from datetime import datetime, timedelta
import micros
import ar_condicionado

try:
    from fpdf import FPDF
except ImportError:
    FPDF = None

try:
    from tkcalendar import DateEntry
except ImportError:
    DateEntry = None

# ---------------------------------------------------------------------------
# GradientButton — botão pill com gradiente horizontal e ícone circular
# ---------------------------------------------------------------------------
class GradientButton(tk.Canvas):
    """Botão pill com gradiente left→right e ícone em círculo branco."""

    def __init__(self, parent, text, color1, color2,
                 command=None, btn_width=180, btn_height=40,
                 icon="❯", icon_font_size=12, text_font_size=9, **kwargs):
        bg = kwargs.pop("bg", "#1e1e2e")
        super().__init__(parent, width=btn_width, height=btn_height,
                         highlightthickness=0, bd=0, cursor="hand2", bg=bg, **kwargs)
        self._cmd     = command
        self._text    = text
        self._c1      = color1
        self._c2      = color2
        self._icon    = icon
        self._btn_w   = btn_width
        self._btn_h   = btn_height
        self._icon_fs = icon_font_size
        self._text_fs = text_font_size

        self.bind("<Button-1>", self._on_click)
        self.bind("<Enter>",    lambda e: self._redraw(hover=True))
        self.bind("<Leave>",    lambda e: self._redraw(hover=False))
        self._redraw()

    @staticmethod
    def _h2rgb(h):
        h = h.lstrip("#")
        return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)

    @staticmethod
    def _lerp(c1, c2, t):
        return "#{:02x}{:02x}{:02x}".format(
            int(c1[0] + (c2[0]-c1[0]) * t),
            int(c1[1] + (c2[1]-c1[1]) * t),
            int(c1[2] + (c2[2]-c1[2]) * t),
        )

    def _redraw(self, hover=False):
        self.delete("all")
        w, h = self._btn_w, self._btn_h
        r = h // 2

        c1 = self._h2rgb(self._c1)
        c2 = self._h2rgb(self._c2)

        if hover:
            factor = 0.82
            c1 = tuple(int(v * factor) for v in c1)
            c2 = tuple(int(v * factor) for v in c2)

        # Gradiente faixa a faixa respeitando o pill
        for x in range(w):
            t = x / max(1, w - 1)
            color = self._lerp(c1, c2, t)
            if x < r:
                dx = r - x
                dy = math.sqrt(max(0, r*r - dx*dx))
                self.create_line(x, r - dy, x, r + dy, fill=color, width=1)
            elif x > w - r - 1:
                dx = x - (w - r - 1)
                if dx <= r:
                    dy = math.sqrt(max(0, r*r - dx*dx))
                    self.create_line(x, r - dy, x, r + dy, fill=color, width=1)
            else:
                self.create_line(x, 0, x, h, fill=color, width=1)

        # Círculo branco do ícone (lado direito)
        cr = (h - 10) // 2
        cx = w - r - 2
        cy = h // 2
        self.create_oval(cx - cr, cy - cr, cx + cr, cy + cr,
                         fill="white", outline="")
        icon_color = self._lerp(c1, c2, cx / max(1, w - 1))
        self.create_text(cx, cy, text=self._icon, fill=icon_color,
                         font=("Segoe UI", self._icon_fs, "bold"), anchor="center")

        # Texto centralizado na área antes do círculo
        text_end = cx - cr - 6
        self.create_text(text_end // 2, cy, text=self._text,
                         fill="white",
                         font=("Segoe UI", self._text_fs, "bold"),
                         anchor="center")

    def _on_click(self, event):
        if self._cmd:
            self._cmd()

# ---------------------------------------------------------------------------
# ImportantInfoBalloon — Ícone de balão vermelho gradiente para observações importantes
# ---------------------------------------------------------------------------
class ImportantInfoBalloon(tk.Canvas):
    """Ícone de balão colorido gradiente que expande ao ser clicado."""
    def __init__(self, parent, text, priority="Importante", **kwargs):
        self.full_text = text
        self.priority = priority
        self.expanded = False
        # Tamanho pequeno inicial (mais largo para a breve descrição)
        self.size_small = (60, 24)
        
        super().__init__(parent, width=self.size_small[0], height=self.size_small[1],
                         highlightthickness=0, bd=0, cursor="hand2", bg=parent["bg"], **kwargs)
        
        self.bind("<Button-1>", self.toggle)
        self.root = parent.winfo_toplevel()
        self._click_outside_id = None
        self._draw()

    def _draw(self):
        self.delete("all")
        w = self.size_small[0]
        h = self.size_small[1]
        
        # Cores baseadas na prioridade
        if self.priority == "Importante":
            c1 = (255, 100, 100)   # #ff6464 (mais claro pra parecer transparente-ish)
            c2 = (180, 0, 0)       # #b40000
            fg_txt = "white"
        elif self.priority == "Média":
            c1 = (173, 216, 230)   # Azul claro (LightBlue)
            c2 = (30, 144, 255)    # DodgerBlue
            fg_txt = "black"
        elif self.priority == "Baixo":
            c1 = (255, 200, 100)   # Laranja claro
            c2 = (255, 140, 0)     # DarkOrange
            fg_txt = "black"
        else: # Fallback
            c1 = (200, 200, 200)
            c2 = (100, 100, 100)
            fg_txt = "black"
        
        for i in range(h):
            t = i / max(1, h-1)
            r = int(c1[0] + (c2[0]-c1[0]) * t)
            g = int(c1[1] + (c2[1]-c1[1]) * t)
            b = int(c1[2] + (c2[2]-c1[2]) * t)
            color = f"#{r:02x}{g:02x}{b:02x}"
            # O oval cria as pontas arredondadas
            if i == 0 or i == h-1:
                self.create_line(h//2, i, w-h//2, i, fill=color)
            else:
                self.create_line(0, i, w, i, fill=color)
        
        # Breve descrição (10 chars)
        brief = self.full_text[:10] + "..." if len(self.full_text) > 10 else self.full_text
        self.create_text(w//2, h//2, text=brief, fill=fg_txt, font=("Segoe UI", 7, "bold"), width=w-4)

    def toggle(self, event=None):
        if not self.expanded:
            self.expand()
        else:
            self.collapse()
        return "break"

    def expand(self):
        if self.expanded: return
        self.expanded = True
        
        self.expand_win = tk.Toplevel(self.root)
        self.expand_win.overrideredirect(True)
        self.expand_win.attributes("-topmost", True)
        self.expand_win.attributes("-alpha", 0.95)
        
        # Cores de fundo e texto na expansão baseadas na prioridade
        if self.priority == "Importante":
            bg_color = "#991b1b" # Vermelho escuro profundo
            fg_color = "white"
        elif self.priority == "Média":
            bg_color = "#add8e6" # Azul claro
            fg_color = "black"
        elif self.priority == "Baixo":
            bg_color = "#ffcc99" # Laranja claro
            fg_color = "black"
        else:
            bg_color = "#333333"
            fg_color = "white"

        self.expand_win.configure(bg=bg_color, highlightthickness=1, highlightbackground="white")
        
        x = self.winfo_rootx()
        y = self.winfo_rooty()
        
        lbl_msg = tk.Label(self.expand_win, text=self.full_text, font=("Segoe UI", 10, "bold"), 
                          wraplength=300, bg=bg_color, fg=fg_color, padx=15, pady=15, justify="left")
        lbl_msg.pack()
        
        self.expand_win.update_idletasks()
        w = lbl_msg.winfo_width()
        h = lbl_msg.winfo_height()
        
        # Centraliza próximo ao ícone mas evita sair da tela
        self.expand_win.geometry(f"{w}x{h}+{x}+{y}")
        
        self.expand_win.bind("<Button-1>", lambda e: self.collapse())
        # Captura clique fora
        self._click_outside_id = self.root.bind("<Button-1>", self._check_click_outside, add="+")
        
    def collapse(self):
        if not self.expanded: return
        self.expanded = False
        if hasattr(self, 'expand_win'):
            try: self.expand_win.destroy()
            except: pass
        
        if self._click_outside_id:
            # Não fazemos unbind real porque pode quebrar outros, 
            # delegamos para o _check_click_outside se auto-ignorar
            self._click_outside_id = None

    def _check_click_outside(self, event):
        if not self.expanded: return
        
        # Verifica se o clique foi na janela de expansão
        try:
            x, y = event.x_root, event.y_root
            wx = self.expand_win.winfo_rootx()
            wy = self.expand_win.winfo_rooty()
            ww = self.expand_win.winfo_width()
            wh = self.expand_win.winfo_height()
            
            if not (wx <= x <= wx + ww and wy <= y <= wy + wh):
                self.collapse()
        except:
            self.collapse()

# ---------------------------------------------------------------------------
class ScheduleGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistema de Agendamento de Laboratórios")
        self.root.geometry("1280x750")
        self.root.state('zoomed') # Inicia em tela cheia (maximizado)

        # Ajuste para ícone no Windows (Taskbar)
        try:
            myappid = 'schedulelabs.alfa.1.0'
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        except Exception:
            pass

        self.db = Database()
        self.aula_selecionada: Optional[Aula] = None
        self._after_ids = {} # Controle de debouncing para buscas
        
        # UI Attributes & Settings
        self.tree: ttk.Treeview | None = None
        self.combo_filtro_dia: ttk.Combobox | None = None
        self.combo_filtro_lab: ttk.Combobox | None = None
        self.combo_filtro_turno: ttk.Combobox | None = None
        self.canvas_grade: tk.Canvas | None = None
        self.inner_grade: tk.Frame | None = None
        self.grade_window: int | None = None
        self.ent_busca_global: tk.Entry | None = None
        self.dados_btn: Any = None
        self._menu_cadastro: tk.Menu | None = None
        self.grade_frame: ttk.LabelFrame | None = None
        self.search_section: ttk.LabelFrame | None = None
        self.results_frame: ttk.LabelFrame | None = None
        self.tree_agenda: ttk.Treeview | None = None
        self.sb_agenda: ttk.Scrollbar | None = None
        self._import_cancelled: bool = False
        self.placeholder_busca: str = "Pesquisar por Professor, Curso, Disciplina, Turma..."
        self.filtro_tipo_aula: str = "TODAS"  # TODAS | SEMESTRAL | EVENTUAL
        self._filtro_btns: dict = {}  # Referências aos botões de filtro
        self.modo_compacto: bool = False
        self.btn_toggle_modo: ttk.Button | None = None
        self.btn_sel_todos: ttk.Button | None = None
        self.btn_batch_del: ttk.Button | None = None
        self.actions_frame: Any = None
        self.tooltip_window: tk.Toplevel | None = None
        
        # Sticky Grade Attributes
        self.grade_corner: tk.Frame | None = None
        self.canvas_hdr_h: tk.Canvas | None = None
        self.canvas_hdr_v: tk.Canvas | None = None
        self.inner_corner: tk.Frame | None = None
        self.inner_hdr_h: tk.Frame | None = None
        self.inner_hdr_v: tk.Frame | None = None
        self.win_hdr_h: int | None = None
        self.win_hdr_v: int | None = None
        self.grid_container: tk.Frame | None = None
        # self.cal_filtro removido conforme pedido
        # Estados Internos
        self.grade_aberta = False
        self._current_agenda_hover = -1
        self.agenda_sort_col = 0 # Índice da coluna (padrão: LAB)
        self.agenda_sort_desc = False
        self.horarios_marcados = set() # Horários (slots) marcados com linha vermelha
        self.proxima_aula_info = {"dia": "", "horario": ""} # Armazena info detalhada da próxima aula
        self.active_grade_views = [] # List de frames 'inner_data' ativos para update rápido
        self.auto_sync_grade_time = True # Por padrão, acompanha o horário atual
        self.setup_styles()
        self.criar_toolbar()
        self.criar_interface()
        self.carregar_lista()
        self.atualizar_grade()

        # Configura ícone do aplicativo (Suporte a PNG via PIL ou PhotoImage)
        try:
            icon_path = os.path.join(os.path.dirname(__file__), "icone", "ico-lab.png")
            if os.path.exists(icon_path):
                try:
                    from PIL import Image, ImageTk
                    pil_img = Image.open(icon_path)
                    icon_img = ImageTk.PhotoImage(pil_img)
                    self.root.iconphoto(True, icon_img)
                    # Manter referência para evitar garbage collection
                    self._app_icon = icon_img
                except ImportError:
                    img = tk.PhotoImage(file=icon_path)
                    self.root.iconphoto(True, img)
                    self._app_icon = img
        except Exception:
            pass

        # Auto-refresh para sincronização em rede (quando o arquivo for alterado por outro usuário)
        self._last_db_mtime = 0
        try:
            self._last_db_mtime = os.path.getmtime(self.db._path)
        except: pass
        self.iniciar_auto_refresh()

    def iniciar_auto_refresh(self):
        """Monitora mudanças no arquivo de banco de dados para sincronizar rede (Otimizado)."""
        try:
            # Se a janela está minimizada, ignoramos o refresh pesado para economizar CPU
            try:
                if self.root.state() == "iconic":
                    self.root.after(4000, self.iniciar_auto_refresh)
                    return
            except: pass

            mtime = os.path.getmtime(self.db._path)
            # Flag para saber se os dados em memória mudaram de fato
            dados_mudaram = False

            if mtime > self._last_db_mtime:
                self._last_db_mtime = mtime
                dados_mudaram = True
                
                # Preservar seleção e scroll da lista principal (Aulas Cadastradas)
                sel_ids = []
                scroll_pos = (0.0, 1.0)
                if self.tree:
                    for item in self.tree.get_children():
                        vals = self.tree.item(item, "values")
                        if vals and vals[0] == "☑":
                            tags = self.tree.item(item, "tags")
                            if tags: sel_ids.append(tags[0])
                    scroll_pos = self.tree.yview()

                # Preservar scroll da agenda (caso modo compacto)
                scroll_agenda = (0.0, 1.0)
                if self.tree_agenda and self.modo_compacto:
                    scroll_agenda = self.tree_agenda.yview()

                # Preservar scroll da grade sincronizada
                scroll_grade = (0.0, 1.0)
                if self.canvas_grade:
                     scroll_grade = self.canvas_grade.yview()

                # Recarregar e Redesenhar
                self.db.recarregar()
                self.carregar_lista()
                self.atualizar_grade()
                dados_mudaram = True
                
                # Restaurar seleção e scroll da lista principal
                if self.tree:
                    for item in self.tree.get_children():
                        tags = self.tree.item(item, "tags")
                        if tags and tags[0] in sel_ids:
                             vals = list(self.tree.item(item, "values"))
                             vals[0] = "☑"
                             self.tree.item(item, values=vals)
                    self.tree.yview_moveto(scroll_pos[0])
                    self._atualizar_botao_batch_del()
                
                # --- Lógica de Scroll da GRADE (Acompanhar Horário Atual) ---
                if self.canvas_grade and self.canvas_hdr_v:
                    dia_sel = self.combo_filtro_dia.get() if self.combo_filtro_dia else self._TODOS_DIAS
                    if getattr(self, "auto_sync_grade_time", True):
                        self.scroll_to_current_time(self.canvas_grade, self.canvas_hdr_v, dia_sel)
                    else:
                        # Senão, restaura a posição para não dar pulo visual no refresh
                        self.canvas_grade.yview_moveto(scroll_grade[0])
                        self.canvas_hdr_v.yview_moveto(scroll_grade[0])
                
                # Restaurar scroll da agenda
                if self.tree_agenda and self.modo_compacto:
                    self.tree_agenda.yview_moveto(scroll_agenda[0])
            
            # Sincroniza indicadores apenas se os dados mudaram ou se o minuto virou
            minuto_atual = datetime.now().minute
            if dados_mudaram or not hasattr(self, "_last_min_check") or self._last_min_check != minuto_atual:
                self._last_min_check = minuto_atual
                db_h_info = self.db.obter_proximo_horario_detalhado()
                if getattr(self, "proxima_aula_info", {}) != db_h_info:
                    self.proxima_aula_info = db_h_info
                    if not dados_mudaram:
                        self.atualizar_grade()
                
                if hasattr(self, 'lbl_status_horario'):
                    self.lbl_status_horario.config(text=f"Próxima Aula: {db_h_info['horario']}")

            # Se estiver na janela principal e for HOJE, o auto-scroll pode atualizar a posição a cada ciclo
            if self.canvas_grade and self.canvas_hdr_v and getattr(self, "auto_sync_grade_time", True):
                 dia_sel = self.combo_filtro_dia.get() if self.combo_filtro_dia else self._TODOS_DIAS
                 self.scroll_to_current_time(self.canvas_grade, self.canvas_hdr_v, dia_sel)

        except Exception as e:
            print(f"Erro auto-refresh: {e}")
            
        self.root.after(3000, self.iniciar_auto_refresh)

    def scroll_to_current_time(self, cg, cv, dia_selecionado):
        """Calcula a fração vertical do horário atual e rola os canvases."""
        if not cg or not cv: return
        
        # Só faz sentido autoscrollar se estivermos vendo HOJE ou TODOS OS DIAS (que mostra hoje)
        now = datetime.now()
        idx_hoje = now.weekday()
        if idx_hoje >= 6: return # Domingo não tem grade
        
        hoje_str = DIAS_SEMANA[idx_hoje]
        if dia_selecionado != hoje_str and dia_selecionado != self._TODOS_DIAS:
            return

        # Horários base fixos da grade (conforme _draw_grade)
        h_ini_m = 7 * 60 + 30  # 07:30
        h_fim_m = 23 * 60 + 15 # 23:15
        total_m = h_fim_m - h_ini_m
        
        now_m = now.hour * 60 + now.minute
        
        # Só rola se estiver dentro ou perto do intervalo da grade
        if now_m > h_ini_m - 30 and now_m < h_fim_m + 30:
            # Fração para o topo da visão (subtraímos 30 min para dar contexto da aula anterior/início)
            target_m = max(0, now_m - h_ini_m - 30)
            fraction = target_m / total_m
            cg.yview_moveto(fraction)
            cv.yview_moveto(fraction)

    # Paleta dark global — usada em todo o sistema
    _D = {
        "bg":      "#1e2540",   # fundo principal das janelas
        "card":    "#252d4a",   # fundo de cards / formulários
        "input":   "#2e3760",   # fundo dos campos de entrada
        "border":  "#3d4f82",   # borda dos campos
        "primary": "#3d5af1",   # azul primário (botão principal)
        "primary2":"#2b45d4",   # azul ao hover
        "success": "#10b981",
        "danger":  "#ef4444",
        "warning": "#f59e0b",
        "fg":      "#e8ecf8",   # texto principal
        "fg2":     "#8f9dc7",   # texto secundário
        "sel":     "#3d5af1",   # seleção na treeview
        "hdr":     "#161d36",   # cabeçalho dark
        "footer":  "#1a2238",   # rodapé dark
        "row_alt": "#232c4a",   # linhas alternadas
    }

    def setup_styles(self):
        """Configura estilos visuais dark modernos (Dark Navy)"""
        D = self._D
        style = ttk.Style()
        if 'clam' in style.theme_names():
            style.theme_use('clam')

        self.root.configure(bg=D["bg"])
        style.configure(".", font=("Segoe UI", 10), background=D["bg"], foreground=D["fg"])

        # Labels
        style.configure("TLabel", background=D["bg"], foreground=D["fg"])
        style.configure("Header.TLabel", font=("Segoe UI", 12, "bold"), foreground=D["primary"])

        # LabelFrame (cards)
        style.configure("TLabelframe", background=D["card"], bordercolor=D["border"], borderwidth=1, relief="flat")
        style.configure("TLabelframe.Label", font=("Segoe UI", 10, "bold"), foreground="#7b9cff", background=D["card"])
        style.configure("Card.TLabelframe", background=D["card"], bordercolor=D["border"])
        style.configure("Card.TLabelframe.Label", font=("Segoe UI", 10, "bold"), foreground="#7b9cff", background=D["card"])

        # Frame
        style.configure("TFrame", background=D["bg"])

        # Treeview
        style.configure("Treeview",
                        background=D["card"],
                        fieldbackground=D["card"],
                        foreground=D["fg"],
                        rowheight=35,
                        borderwidth=0)
        style.configure("Treeview.Heading",
                        font=("Segoe UI", 9, "bold"),
                        background=D["hdr"],
                        foreground=D["fg2"],
                        relief="flat",
                        padding=10)
        style.map("Treeview",
                  background=[("selected", D["sel"])],
                  foreground=[("selected", "#ffffff")])

        # Entry & Combobox
        style.configure("TEntry", padding=8,
                        fieldbackground=D["input"], foreground=D["fg"],
                        insertcolor=D["fg"], bordercolor=D["border"])
        style.map("TEntry",
                  fieldbackground=[("readonly", D["input"]), ("!disabled", D["input"])],
                  foreground=[("readonly", D["fg"])])
        style.configure("TCombobox", padding=6,
                        fieldbackground=D["input"], foreground=D["fg"],
                        selectbackground=D["sel"], selectforeground="white",
                        arrowcolor=D["fg2"])
        style.map("TCombobox",
                  fieldbackground=[("readonly", D["input"])],
                  foreground=[("readonly", D["fg"])])
        style.configure("Invalid.TCombobox", fieldbackground="#4a1c1c", bordercolor=D["danger"])
        style.map("Invalid.TCombobox", fieldbackground=[("readonly", "#4a1c1c")])

        # Botões
        btn_cfg = {"font": ("Segoe UI", 9, "bold"), "padding": (14, 8)}
        style.configure("TButton", **btn_cfg, background=D["primary"], foreground="white",
                        borderwidth=0, focuscolor="none")
        style.map("TButton", background=[("active", D["primary2"])])
        style.configure("Success.TButton", **btn_cfg, background=D["success"], foreground="white")
        style.map("Success.TButton", background=[("active", "#059669")])
        style.configure("Warning.TButton", **btn_cfg, background=D["warning"], foreground="white")
        style.map("Warning.TButton", background=[("active", "#d97706")])
        style.configure("Danger.TButton", **btn_cfg, background=D["danger"], foreground="white")
        style.map("Danger.TButton", background=[("active", "#dc2626")])

        # Scrollbar
        style.configure("TScrollbar", background=D["card"], troughcolor=D["bg"],
                        arrowcolor=D["fg2"], borderwidth=0)

        # Notebook (abas)
        style.configure("TNotebook", background=D["bg"], borderwidth=0)
        style.configure("TNotebook.Tab", font=("Segoe UI", 9, "bold"),
                        padding=[14, 7], background=D["card"], foreground=D["fg2"])
        style.map("TNotebook.Tab",
                  background=[("selected", D["primary"])],
                  foreground=[("selected", "white")])

        try:
            self.root.configure(bg=D["bg"])
        except Exception:
            pass

    # ------------------------------------------------------------------
    def criar_toolbar(self):
        """Barra de ferramentas superior com acões globais."""
        TB_BG   = "#1e1e2e"   # fundo escuro
        TB_FG   = "#cdd6f4"   # texto claro
        TB_ACC  = "#89b4fa"   # azul accent

        toolbar = tk.Frame(self.root, bg=TB_BG, height=48)
        toolbar.pack(side=tk.TOP, fill=tk.X)
        toolbar.pack_propagate(False)

        # Título à esquerda
        tk.Label(
            toolbar,
            text="🧪  Sistema de Agendamento de Laboratórios",
            bg=TB_BG, fg=TB_ACC,
            font=("Segoe UI", 12, "bold")
        ).pack(side=tk.LEFT, padx=16)

        # helper para empacotar GradientButton na toolbar
        def _gb(text, c1, c2, cmd, icon, w=170):
            b = GradientButton(toolbar, text=text, color1=c1, color2=c2,
                               command=cmd, btn_width=w, btn_height=38,
                               icon=icon, icon_font_size=11, text_font_size=9,
                               bg=TB_BG)
            b.pack(side=tk.RIGHT, padx=5, pady=5)
            return b

        # SAIR  — azul-escuro → cinza-ardósia
        _gb("SAIR",         "#475569", "#1e293b", self.sair,               "✕", w=100)
        
        # GRADE — laranja → vermelho
        _gb("GRADE",        "#f97316", "#ef4444", self.abrir_janela_grade, "⊞", w=120)

        # ATUALIZAR — violeta → índigo
        _gb("ATUALIZAR",    "#8b5cf6", "#6366f1", self.atualizar_dados,    "↺", w=140)

        # ADICIONAR AULA — esmeralda → ciano
        _gb("ADICIONAR AULA", "#10b981", "#06b6d4", self._janela_cadastro_aula, "+", w=165)

        # LABORATÓRIOS — rosa → roxo
        _gb("LABORATÓRIOS", "#ec4899", "#8b5cf6", self._janela_laboratorios, "🔬", w=150)

        # DADOS DE CADASTRO — âmbar → laranja
        self.dados_btn = GradientButton(
            toolbar, text="CADASTROS", color1="#f59e0b", color2="#f97316",
            command=None, btn_width=150, btn_height=38,
            icon="≡", icon_font_size=13, text_font_size=9, bg=TB_BG
        )
        if self.dados_btn:
            self.dados_btn.pack(side=tk.RIGHT, padx=5, pady=5)

        # Menu dropdown
        self._menu_cadastro = tk.Menu(self.root, tearoff=0,
                                      font=("Segoe UI", 10),
                                      bg="#2a2a3e", fg="white",
                                      activebackground="#7c3aed",
                                      activeforeground="white",
                                      bd=0)
        
        if self._menu_cadastro is not None:
            self._menu_cadastro.add_command(label="🏛️   Nova Faculdade",  command=self._janela_nova_faculdade)
            self._menu_cadastro.add_command(label="🎨   Novo Curso",      command=self._janela_novo_curso)
            self._menu_cadastro.add_command(label="📚   Nova Disciplina", command=self._janela_nova_disciplina)
            self._menu_cadastro.add_command(label="  🎓   Nova Turma", command=self._janela_nova_turma)
            self._menu_cadastro.add_command(label="  📊   Dados Cadastrados", command=self._janela_dados_cadastrados)
            self._menu_cadastro.add_command(label="  📊   Histórico/Relatório",   command=self._janela_historico)
            self._menu_cadastro.add_separator()
            self._menu_cadastro.add_command(label="📥   Importar Dados (CSV)", command=self._janela_importar_dados)
            self._menu_cadastro.add_separator()
            self._menu_cadastro.add_command(label="📊   Exportar para Excel (.xlsx)", command=self.exportar_xlsx)
            self._menu_cadastro.add_command(label="📄   Exportar para CSV (.csv)", command=self.exportar_csv)

        def _popup_cadastro(event=None):
            if self.dados_btn is not None and self._menu_cadastro is not None:
                x = self.dados_btn.winfo_rootx()
                y = self.dados_btn.winfo_rooty() + self.dados_btn.winfo_height()
                self._menu_cadastro.tk_popup(x, y)

        if self.dados_btn is not None:
             self.dados_btn._cmd = _popup_cadastro


    # ------------------------------------------------------------------
    def criar_interface(self):
        """Cria a interface gráfica (Layout expandido sem form lateral)"""
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(0, weight=0) # Barra de Pesquisa (Aulas Cadastradas)
        main_frame.rowconfigure(1, weight=1) # Área de Grade/Resultados (Expandir)

        # ========== ÁREA SUPERIOR: PESQUISA (Aulas Cadastradas) ==========
        self.search_section = ttk.LabelFrame(main_frame, text="📋  Aulas Cadastradas", padding="15")
        self.search_section.grid(row=0, column=0, sticky="ew", pady=(0, 5))
        self.search_section.columnconfigure(0, weight=1)

        # Bar de Pesquisa Global
        search_frame = ttk.Frame(self.search_section)
        search_frame.grid(row=0, column=0, sticky="ew", pady=(0, 5))

        # --- Busca Global Modernizada (Estilo Dark Card) ---
        D = self._D
        search_card = tk.Frame(search_frame, bg=D["input"], highlightthickness=1,
                                highlightbackground=D["border"], highlightcolor=D["primary"])
        search_card.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5, pady=5)

        tk.Label(search_card, text=" 🔍 ", bg=D["input"], fg=D["fg2"],
                    font=("Segoe UI", 12)).pack(side=tk.LEFT, padx=(8, 0))

        self.ent_busca_global = tk.Entry(search_card, font=("Segoe UI", 10),
                                            bg=D["input"], relief="flat", bd=0,
                                            fg=D["fg"], insertbackground=D["fg"],
                                            justify="left")
        self.ent_busca_global.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=8, pady=8)

        # --- Botões de Filtro por Tipo (Segmentados) ---
        filtro_container = tk.Frame(search_card, bg=D["input"])
        filtro_container.pack(side=tk.RIGHT, padx=(0, 8), pady=5)

        # Separador vertical sutil
        tk.Frame(filtro_container, bg=D["border"], width=1).pack(side=tk.LEFT, fill=tk.Y, padx=(0, 6), pady=2)

        self._filtro_btns = {}
        filtros_cfg = [
            ("TODAS",      "📋 Todas"),
            ("SEMESTRAL",  "📚 Semestrais"),
            ("EVENTUAL",   "🕒 Eventuais"),
        ]

        def _set_filtro(tipo):
            self.filtro_tipo_aula = tipo
            self._atualizar_visual_filtro_btns()
            self.carregar_lista()

        for tipo, texto in filtros_cfg:
            btn = tk.Label(
                filtro_container, text=texto, font=("Segoe UI", 9, "bold"),
                padx=10, pady=4, cursor="hand2",
                bg=D["input"], fg=D["fg2"],
                highlightthickness=0, bd=0
            )
            btn.pack(side=tk.LEFT, padx=1)
            btn.bind("<Button-1>", lambda e, t=tipo: _set_filtro(t))
            self._filtro_btns[tipo] = btn

        # Estiliza o botão ativo na inicialização
        self._atualizar_visual_filtro_btns()

        # Lógica de Placeholder
        def _on_focus_in(e):
            if self.ent_busca_global.get() == self.placeholder_busca:
                self.ent_busca_global.delete(0, tk.END)
                self.ent_busca_global.config(fg=D["fg"])

        def _on_focus_out(e):
            if not self.ent_busca_global.get():
                self.ent_busca_global.insert(0, self.placeholder_busca)
                self.ent_busca_global.config(fg=D["fg2"])

        self.ent_busca_global.insert(0, self.placeholder_busca)
        self.ent_busca_global.config(fg=D["fg2"])
        
        self.ent_busca_global.bind("<FocusIn>", _on_focus_in)
        self.ent_busca_global.bind("<FocusOut>", _on_focus_out)
        self.ent_busca_global.bind("<KeyRelease>", lambda e: self.debounce("lista", self.carregar_lista))

        # Frame de Ações (Posicionado à direita da busca)
        self.actions_frame = ttk.Frame(search_frame)
        self.actions_frame.pack(side=tk.RIGHT, padx=5)

        self.btn_sel_todos = ttk.Button(self.actions_frame, text="☑ Selecionar Tudo", command=self.selecionar_tudo)
        self.btn_sel_todos.pack(side=tk.LEFT, padx=2)

        self.btn_batch_del = ttk.Button(self.actions_frame, text="🗑 Excluir Selecionados", 
                                        command=self.excluir_selecionados)
        # Oculto inicialmente
        self.btn_batch_del.pack_forget()

        # ========== ÁREA INFERIOR: CONTEÚDO (Visualização da Grade e Resultados Overlay) ==========
        content_container = ttk.Frame(main_frame)
        content_container.grid(row=1, column=0, sticky="nsew")
        content_container.columnconfigure(0, weight=1)
        content_container.rowconfigure(0, weight=1)

        self.grade_frame = ttk.LabelFrame(content_container, text="📅  Visualização da Grade", padding="15")
        self.grade_frame.grid(row=0, column=0, sticky="nsew")

        # Container para Resultados da Pesquisa (Overlay)
        self.results_frame = ttk.LabelFrame(content_container, text="📋  Resultados da Pesquisa", padding="15")
        self.results_frame.grid(row=0, column=0, sticky="nsew")
        self.results_frame.columnconfigure(0, weight=1)
        self.results_frame.rowconfigure(0, weight=1)
        self.results_frame.grid_remove() # Oculto por padrão

        columns = ("Sel", "Lab", "Dia", "Horário", "Disciplina", "Turma", "Alunos", "Professor", "Curso", "Info")
        self.tree = ttk.Treeview(self.results_frame, columns=columns, show="headings", height=10)
        
        headers = ["SEL", "LAB", "DIA", "HORÁRIO", "DISCIPLINA", "TURMA", "ALUNOS", "PROFESSOR", "CURSO", "INFO"]
        widths  = [40, 60, 120, 100, 350, 80, 80, 150, 300, 50]
        
        for col, head, w in zip(columns, headers, widths):
            if col == "Disciplina": w = 500
            elif col == "Curso": w = 400
            self.tree.heading(col, text=head, anchor="center")
            is_flex = col in ["Disciplina", "Curso"]
            self.tree.column(col, width=w, minwidth=w//2, anchor="center", stretch=is_flex)

        if self.tree is not None:
            self.tree.grid(row=0, column=0, sticky="nsew")
            sb = ttk.Scrollbar(self.results_frame, orient=tk.VERTICAL, command=self.tree.yview)
            self.tree.configure(yscrollcommand=sb.set)
            sb.grid(row=0, column=1, sticky="ns")

            self.tree.bind("<Double-1>", self.selecionar_aula)
            self.tree.bind("<Button-1>", self._on_tree_click)
            self.tree.bind("<Button-3>", self._on_tree_right_click)
            self._setup_tree_hover(self.tree)

        self.grade_frame.columnconfigure(0, weight=1)
        self.grade_frame.rowconfigure(1, weight=1)

        # --- Novo Componente: Treeview da Agenda (Modo Compacto) ---
        # Definimos aqui para ser idêntico ao Aulas Cadastradas
        cols_agenda = ("Lab", "Dia", "Horário", "Disciplina", "Turma", "Alunos", "Professor", "Curso", "Info")
        self.tree_agenda = ttk.Treeview(self.grade_frame, columns=cols_agenda, show="headings", height=10)
        
        # Cabeçalhos e Larguras (Sincronizados com Aulas Cadastradas)
        heads_agenda = ["LAB", "DIA", "HORÁRIO", "DISCIPLINA", "TURMA", "ALUNOS", "PROFESSOR", "CURSO", "INFO"]
        # Usamos pesos proporcionais similares
        w_agenda = [50, 100, 100, 350, 80, 70, 180, 250, 45]
        
        t_ag = self.tree_agenda
        if t_ag is not None:
            for col, head, w in zip(cols_agenda, heads_agenda, w_agenda):
                t_ag.heading(col, text=head, anchor="center", 
                                       command=lambda c=col: self._sort_tree_agenda(c))
                is_flex = col in ["Disciplina", "Curso", "Professor"]
                t_ag.column(col, width=w, minwidth=w//2, anchor="center", stretch=is_flex)

            self.sb_agenda = ttk.Scrollbar(self.grade_frame, orient=tk.VERTICAL, command=t_ag.yview)
            t_ag.configure(yscrollcommand=self.sb_agenda.set)
            
            # Binds para Modo Agenda
            t_ag.bind("<Double-1>", self._on_tree_agenda_double_click)
            t_ag.bind("<Button-1>", self._on_tree_agenda_click)
            t_ag.bind("<Button-3>", self._on_tree_agenda_right_click)
            self._setup_tree_hover(t_ag)

            # Inicialmente ocultos (o modo padrão é Grade)
            t_ag.grid_remove()
            if self.sb_agenda: self.sb_agenda.grid_remove()

        # Filtros
        filtro_frame = ttk.Frame(self.grade_frame)
        filtro_frame.grid(row=0, column=0, sticky="ew", pady=(0, 5))
        
        _TODOS_DIAS = "TODOS OS DIAS"
        _TODOS_LABS = "TODOS OS LABORATÓRIOS"

        # Barra de Pesquisa em Tempo Real (Grade)
        self.grade_search_placeholder = "Pesquisar Professor, Disciplina, Curso..."
        
        search_card_grade = tk.Frame(filtro_frame, bg=self._D["input"], highlightthickness=1,
                                     highlightbackground=self._D["border"], highlightcolor=self._D["primary"])
        search_card_grade.pack(side=tk.LEFT, padx=(5, 15), pady=2)

        tk.Label(search_card_grade, text=" 🔍 ", bg=self._D["input"], fg=self._D["fg2"],
                    font=("Segoe UI", 10)).pack(side=tk.LEFT, padx=(5, 0))

        self.ent_busca_grade = tk.Entry(search_card_grade, font=("Segoe UI", 10),
                                        bg=self._D["input"], relief="flat", bd=0,
                                        fg=self._D["fg"], insertbackground=self._D["fg"], width=30)
        self.ent_busca_grade.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=8, pady=5)
        self.ent_busca_grade.insert(0, self.grade_search_placeholder)
        self.ent_busca_grade.config(fg=self._D["fg2"])
        
        def _on_grade_search_focus_in(e):
             if self.ent_busca_grade.get() == self.grade_search_placeholder:
                 self.ent_busca_grade.delete(0, tk.END)
                 self.ent_busca_grade.config(fg=self._D["fg"])
        def _on_grade_search_focus_out(e):
             if not self.ent_busca_grade.get():
                 self.ent_busca_grade.insert(0, self.grade_search_placeholder)
                 self.ent_busca_grade.config(fg=self._D["fg2"])

        self.ent_busca_grade.bind("<FocusIn>", _on_grade_search_focus_in)
        self.ent_busca_grade.bind("<FocusOut>", _on_grade_search_focus_out)
        self.ent_busca_grade.bind("<KeyRelease>", lambda e: self.debounce("grade", self.atualizar_grade))

        ttk.Label(filtro_frame, text="Dia da Semana:").pack(side=tk.LEFT, padx=5)
        self.combo_filtro_dia = ttk.Combobox(
            filtro_frame, values=[_TODOS_DIAS] + DIAS_SEMANA, state="readonly", width=22
        )
        # Seleciona o dia atual do sistema por padrão
        import datetime as dt_lib
        hoje_idx = dt_lib.datetime.now().weekday() # 0-6
        if hoje_idx < 6:
            dia_padrao = DIAS_SEMANA[hoje_idx]
        else:
            dia_padrao = _TODOS_DIAS
            
        self.combo_filtro_dia.set(dia_padrao)
        self.combo_filtro_dia.pack(side=tk.LEFT, padx=5)
        self.combo_filtro_dia.bind("<<ComboboxSelected>>", self._on_filtro_dia_selected)

        ttk.Label(filtro_frame, text="Laboratório:").pack(side=tk.LEFT, padx=5)
        self.combo_filtro_lab = ttk.Combobox(
            filtro_frame, values=[_TODOS_LABS] + LABORATORIOS, state="readonly", width=22
        )
        self.combo_filtro_lab.set(_TODOS_LABS)
        self.combo_filtro_lab.pack(side=tk.LEFT, padx=5)
        self.combo_filtro_lab.bind("<<ComboboxSelected>>", self._on_filtro_lab_selected)

        ttk.Label(filtro_frame, text="Turno:").pack(side=tk.LEFT, padx=5)
        self.combo_filtro_turno = ttk.Combobox(
            filtro_frame, values=["TODOS", "MATUTINO", "VESPERTINO", "NOTURNO"], state="readonly", width=15
        )
        self.combo_filtro_turno.set("TODOS")
        self.combo_filtro_turno.pack(side=tk.LEFT, padx=5)
        self.combo_filtro_turno.bind("<<ComboboxSelected>>", lambda e: self.atualizar_grade())

        # Seletor de Data removido
        # ttk.Label(filtro_frame, text="Semana de:").pack(side=tk.LEFT, padx=(10, 5))
        # cal_filtro logic removed conforme pedido


        # Botão de Alternância de Modo
        self.btn_toggle_modo = ttk.Button(
            filtro_frame, text="📅 MODO AGENDA", command=self.alternar_modo_visualizacao
        )
        self.btn_toggle_modo.pack(side=tk.RIGHT, padx=10)

        # Botão Gerar Mapão
        ttk.Button(
            filtro_frame, text="🗺️ Gerar Programação de Aulas", command=self.gerar_mapao
        ).pack(side=tk.RIGHT, padx=5)

        # Botão LAB STATUS (Grande e Vermelho)
        GradientButton(
            filtro_frame, text="LAB STATUS", color1="#ef4444", color2="#991b1b",
            command=self.chamar_lig_desl, btn_width=160, btn_height=45,
            icon="⚡", icon_font_size=12, text_font_size=10, bg=self._D["bg"]
        ).pack(side=tk.RIGHT, padx=5)

        # Indicador de Horário do Status (Sincronizado)
        self.lbl_status_horario = tk.Label(filtro_frame, text="", font=("Segoe UI", 10, "bold"), 
                                           bg=self._D["bg"], fg="#f59e0b")
        self.lbl_status_horario.pack(side=tk.RIGHT, padx=10)


        ttk.Label(filtro_frame, text="   (Colunas = Laboratórios / Linhas = Horários)", font=("Segoe UI", 8, "italic")).pack(side=tk.LEFT)

        # --- Estrutura de Grade Sincronizada (Sticky Headers) ---
        self.grid_container = tk.Frame(self.grade_frame, bg=self._D["border"])
        self.grid_container.grid(row=1, column=0, sticky="nsew")
        self.grid_container.columnconfigure(1, weight=1)
        self.grid_container.rowconfigure(1, weight=1)

        grid_container = self.grid_container

        # 1. Corner (Fixo)
        gc = tk.Frame(grid_container, bg=self._D["hdr"], width=70, height=40)
        self.grade_corner = gc
        gc.grid(row=0, column=0, sticky="nsew")
        gc.grid_propagate(False)

        # 2. Header Horizontal (Scroll H)
        ch = tk.Canvas(grid_container, bg=self._D["hdr"], height=40, highlightthickness=0)
        self.canvas_hdr_h = ch
        ch.grid(row=0, column=1, sticky="ew")

        # 3. Coluna de Horas (Scroll V)
        cv = tk.Canvas(grid_container, bg=self._D["card"], width=70, highlightthickness=0)
        self.canvas_hdr_v = cv
        cv.grid(row=1, column=0, sticky="ns")

        # 4. Área de Dados Principal (Scroll H/V)
        cg = tk.Canvas(grid_container, bg=self._D["card"], highlightthickness=0)
        self.canvas_grade = cg
        cg.grid(row=1, column=1, sticky="nsew")

        # Scrollbars
        v_scroll = ttk.Scrollbar(grid_container, orient=tk.VERTICAL, command=self._on_vscroll)
        h_scroll = ttk.Scrollbar(grid_container, orient=tk.HORIZONTAL, command=self._on_hscroll)
        
        cg.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        cv.configure(yscrollcommand=v_scroll.set)
        ch.configure(xscrollcommand=h_scroll.set)

        v_scroll.grid(row=1, column=2, sticky="ns")
        h_scroll.grid(row=2, column=1, sticky="ew")

        # Frames Internos
        ic = tk.Frame(gc, bg=self._D["hdr"])
        self.inner_corner = ic
        ic.pack(fill=tk.BOTH, expand=True)

        ih = tk.Frame(ch, bg=self._D["hdr"])
        self.inner_hdr_h = ih
        self.win_hdr_h = ch.create_window((0, 0), window=ih, anchor="nw")

        iv = tk.Frame(cv, bg=self._D["card"])
        self.inner_hdr_v = iv
        self.win_hdr_v = cv.create_window((0, 0), window=iv, anchor="nw")

        ig = tk.Frame(cg, bg=self._D["card"])
        self.inner_grade = ig
        self.grade_window = cg.create_window((0, 0), window=ig, anchor="nw")
        
        # Registra visualização para updates rápidos
        self.active_grade_views.append(ig)

        # Binds de configuração
        ig.bind("<Configure>", self._on_frame_configure)
        cg.bind("<Configure>", self._on_canvas_configure)
        ih.bind("<Configure>", lambda e: ch.configure(scrollregion=ch.bbox("all")))
        iv.bind("<Configure>", lambda e: cv.configure(scrollregion=cv.bbox("all")))

        # Bind de MouseWheel para scroll sincronizado global
        def _on_mousewheel(event):
            w = event.widget
            if not w.winfo_exists(): return "break"
            top = w.winfo_toplevel()
            delta = -1 * (event.delta // 120)
            is_shift = (event.state & 0x1) != 0 # Shift key pressed
            
            # Se for a janela secundária "Visualização da Grade de Horários"
            if top.title() == "Visualização da Grade de Horários":
                if hasattr(top, 'win_canvas_data'):
                    if is_shift:
                        if hasattr(top, 'win_canvas_hdr_h'): 
                            top.win_canvas_data.xview_scroll(delta, "units")
                            top.win_canvas_hdr_h.xview_scroll(delta, "units")
                    else:
                        if hasattr(top, 'win_canvas_hdr_v'):
                            setattr(top, "auto_sync_time", False) # Desativa ao scrollar manualmente
                            top.win_canvas_data.yview_scroll(delta, "units")
                            top.win_canvas_hdr_v.yview_scroll(delta, "units")
                    return "break"
            
            # Se for a janela principal, só scrollar a grade sincronizada se o mouse estiver sobre ela    
            elif top == self.root:
                # Verifica se o widget é descendente de grade_frame
                curr = w
                is_in_grid_area = False
                while curr and curr != self.root:
                    if curr == self.grade_frame:
                        is_in_grid_area = True
                        break
                    curr = curr.master
                
                if is_in_grid_area:
                    cg = self.canvas_grade
                    cv = self.canvas_hdr_v
                    ch = self.canvas_hdr_h
                    if not self.modo_compacto and cg:
                        if is_shift:
                            if ch:
                                cg.xview_scroll(delta, "units")
                                ch.xview_scroll(delta, "units")
                        else:
                            if cv:
                                self.auto_sync_grade_time = False # O usuário scrollou manualmente
                                cg.yview_scroll(delta, "units")
                                cv.yview_scroll(delta, "units")
                        return "break"
            
            # Deixa propagar ou trata outros widgets scrolláveis (Treeview, Canvas, Text, etc)
            parent = w
            while parent and parent != top:
                if hasattr(parent, "yview_scroll"):
                    # Verificamos visibilidade do scroll se for Canvas
                    can_scroll_v = True
                    can_scroll_h = True
                    if isinstance(parent, tk.Canvas):
                        bbox = parent.bbox("all")
                        if bbox:
                            can_scroll_v = (bbox[3] - bbox[1]) > parent.winfo_height()
                            can_scroll_h = (bbox[2] - bbox[0]) > parent.winfo_width()
                        else:
                            can_scroll_v = can_scroll_h = False
                    
                    if is_shift and hasattr(parent, "xview_scroll") and can_scroll_h:
                        parent.xview_scroll(delta, "units")
                        return "break"
                    elif can_scroll_v:
                        parent.yview_scroll(delta, "units")
                        return "break"
                parent = parent.master
            return None

        
        self.root.bind_all("<MouseWheel>", _on_mousewheel)


    _TODOS_DIAS = "TODOS OS DIAS"
    _TODOS_LABS = "TODOS OS LABORATÓRIOS"

    def _on_filtro_dia_selected(self, event):
        """Atualiza a grade ao selecionar o dia e restabelece auto-scroll."""
        self.auto_sync_grade_time = True # Volta a acompanhar se mudou de filtro
        self.atualizar_grade()

    def _on_filtro_lab_selected(self, event):
        """Atualiza a grade ao selecionar o laboratório."""
        self.atualizar_grade()

        # ==== métodos auxiliares do canvas da grade ====

    def _on_frame_configure(self, event):
        """Atualiza a região de scroll quando o conteúdo muda."""
        if self.canvas_grade is not None:
            # Primeiro garante que o item da janela no canvas tenha a largura correta
            cw = self.canvas_grade.winfo_width()
            rw = self.inner_grade.winfo_reqwidth() if self.inner_grade else 0
            new_w = max(cw, rw)
            if self.grade_window is not None:
                self.canvas_grade.itemconfig(self.grade_window, width=new_w)
            if self.canvas_hdr_h is not None and self.win_hdr_h is not None:
                self.canvas_hdr_h.itemconfig(self.win_hdr_h, width=new_w)
            
            # Força processamento da largura antes de ler o bounding box para o scrollregion
            self.canvas_grade.update_idletasks()
            self.canvas_grade.configure(scrollregion=self.canvas_grade.bbox("all"))
            if self.canvas_hdr_h:
                 self.canvas_hdr_h.configure(scrollregion=self.canvas_hdr_h.bbox("all"))

    def _on_canvas_configure(self, event):
        """Sincroniza a largura dos frames internos (Cabeçalho e Dados) com o canvas."""
        if self.canvas_grade is not None and self.grade_window is not None:
            canvas_width = event.width
            rw = self.inner_grade.winfo_reqwidth() if self.inner_grade else 0
            new_w = max(canvas_width, rw)
            self.canvas_grade.itemconfig(self.grade_window, width=new_w)
            if self.canvas_hdr_h is not None and self.win_hdr_h is not None:
                self.canvas_hdr_h.itemconfig(self.win_hdr_h, width=new_w)


    def converter_maiuscula(self, entry_widget):
        """Converte texto para maiúscula quando sair do campo civilizadamente."""
        texto = entry_widget.get().upper()
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, texto)

    def _mask_hora(self, event):
        """Formata HH:MM enquanto digita."""
        if event.keysym == "BackSpace":
            return
            
        w = event.widget
        val = w.get().replace(":", "")
        
        # Filtra apenas dígitos
        val = "".join([c for c in val if c.isdigit()])
        if len(val) > 4:
            val = val[:4]
            
        # Formata com :
        if len(val) >= 3:
            fmt = f"{val[:2]}:{val[2:]}"
        else:
            fmt = val
            
        # Atualiza o campo mantendo o cursor no final ou onde deve estar
        w.delete(0, tk.END)
        w.insert(0, fmt)

    def _mask_int(self, event):
        """Permite apenas números."""
        w = event.widget
        val = w.get()
        if not val.isdigit() and val != "":
            val = "".join([c for c in val if c.isdigit()])
            w.delete(0, tk.END)
            w.insert(0, val)

    def _on_vscroll(self, *args):
        """Sincroniza scroll vertical entre a coluna de horas e a grade."""
        self.auto_sync_grade_time = False # O usuário moveu a barra manual
        cg = self.canvas_grade
        cv = self.canvas_hdr_v
        if cg and cv:
            cg.yview(*args)
            cv.yview(*args)

    def _on_hscroll(self, *args):
        """Sincroniza scroll horizontal entre o cabeçalho e a grade."""
        cg = self.canvas_grade
        ch = self.canvas_hdr_h
        if cg and ch:
            cg.xview(*args)
            ch.xview(*args)

    def atualizar_dados(self):
        """Atualiza lista e grade."""
        self.db.recarregar()
        self.carregar_lista()
        self.atualizar_grade()

    def _atualizar_visual_filtro_btns(self):
        """Atualiza o visual dos botões de filtro de tipo de aula."""
        D = self._D
        for tipo, btn in self._filtro_btns.items():
            if tipo == self.filtro_tipo_aula:
                # Ativo
                if tipo == "EVENTUAL":
                    btn.config(bg="#000000", fg="#ffffff")
                elif tipo == "SEMESTRAL":
                    btn.config(bg=D["primary"], fg="#ffffff")
                else:
                    btn.config(bg=D["primary"], fg="#ffffff")
            else:
                # Inativo
                btn.config(bg=D["input"], fg=D["fg2"])

    def carregar_lista(self):
        """Monitora banco e atualiza a Treeview de aulas."""
        if self.tree is None: return

        # Limpar tree
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Lógica de filtro global
        busca = ""
        if self.ent_busca_global:
            val = self.ent_busca_global.get().strip()
            if val and val != self.placeholder_busca:
                busca = remover_acentos(val)

        filtro_tipo = self.filtro_tipo_aula  # TODAS | SEMESTRAL | EVENTUAL

        # Se não há busca E o filtro é TODAS, oculta e encerra
        if not busca and filtro_tipo == "TODAS":
            if self.results_frame:
                self.results_frame.grid_remove()
            return

        # Se há busca ou filtro ativo, mostra os resultados (Sobrepondo a grade)
        if self.results_frame:
            self.results_frame.grid()
            self.results_frame.lift()

        # Carregar aulas
        aulas = self.db.listar_todas_aulas()
        
        for aula in aulas:
            # Filtro por tipo (Eventual / Semestral)
            if filtro_tipo == "EVENTUAL" and not aula.is_eventual:
                continue
            if filtro_tipo == "SEMESTRAL" and aula.is_eventual:
                continue

            # Filtro por texto de busca
            if busca:
                campos = [
                    aula.professor,
                    aula.curso,
                    aula.disciplina,
                    aula.turma,
                    aula.faculdade
                ]
                match = any(busca in remover_acentos(str(c)) for c in campos if c)
                if not match:
                    continue
            
            horario = f"{aula.hora_inicio}-{aula.hora_fim}"

            # Inserir no tree
            tag_name = f"cor_{aula.id}"
            item_id = self.tree.insert(
                "",
                tk.END,
                values=(
                    "☐", # Checkbox vazio
                    aula.laboratorio,
                    aula.dia_semana,
                    horario,
                    aula.disciplina,
                    aula.turma,
                    aula.qtde_alunos,
                    aula.professor,
                    aula.curso,
                    "ℹ" # Ícone de info
                ),
                tags=(str(aula.id), tag_name)
            )

            # Aplicar cor de fundo e texto (contraste)
            try:
                # aula.cor_fundo agora já vem processado pelo DB (herdado ou específico)
                # Eventuais sempre usam fundo preto
                if aula.is_eventual:
                    cf = "#000000"
                else:
                    cf = aula.cor_fundo if (aula.cor_fundo and aula.cor_fundo != "#ffffff") else self._D["card"]
                ct = texto_contraste(cf)
                if self.tree is not None:
                    self.tree.tag_configure(tag_name, background=cf, foreground=ct)
            except Exception:
                pass

    def selecionar_aula(self, event):
        """Abre janela de edição para a aula selecionada na lista."""
        if self.tree is None: return
        
        item = self.tree.identify_row(event.y)
        if not item: return
        
        # Ignora se clicou na coluna de seleção (coluna #1)
        col = self.tree.identify_column(event.x)
        if col == "#1": return

        tags = self.tree.item(item, "tags")
        if tags:
            aula_id = int(tags[0])
            aulas = self.db.listar_todas_aulas()
            aula = next((a for a in aulas if a.id == aula_id), None)
            if aula:
                self._janela_cadastro_aula(aula)

    def _on_tree_click(self, event):
        """Gerencia clique na coluna de seleção."""
        if self.tree is None: return
        
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell": return
        
        column = self.tree.identify_column(event.x)
        
        # Gerencia Info (Coluna #10)
        if column == "#10":
            item = self.tree.identify_row(event.y)
            if item:
                tags = self.tree.item(item, "tags")
                if tags:
                    try:
                        # O ID está na primeira tag
                        aula_id = int(tags[0])
                        aulas = self.db.listar_todas_aulas()
                        aula = next((a for a in aulas if a.id == aula_id), None)
                        if aula:
                            self._janela_info_aula(aula)
                    except (ValueError, IndexError):
                        pass
            return

        # Gerencia Seleção (Coluna #1)
        if column != "#1": return
        
        item = self.tree.identify_row(event.y)
        if not item: return
        
        values = list(self.tree.item(item, "values"))
        current = values[0]
        
        # Inverte
        new_val = "☑" if current == "☐" else "☐"
        values[0] = new_val
        self.tree.item(item, values=values)
        
        self._atualizar_botao_batch_del()

    def _setup_tree_hover(self, tree: ttk.Treeview):
        """Implementa efeito de destaque (hover) suave em Treeviews."""
        def on_motion(event):
            item = tree.identify_row(event.y)
            # Limpa destaque anterior
            for old_item in tree.tag_has("hover"):
                tree.item(old_item, tags=[t for t in tree.item(old_item, "tags") if t != "hover"])
            
            if item:
                tags = list(tree.item(item, "tags"))
                if "hover" not in tags:
                    tags.append("hover")
                    tree.item(item, tags=tags)
                # Adiciona borda de foco (dashed line) via sistema
                tree.focus(item)

        tree.bind("<Motion>", on_motion)
        # Configura a tag de hover (Texto branco brilhante, fundo transparente)
        tree.tag_configure("hover", foreground="#ffffff")

    def _show_tooltip(self, event, aula: Aula):
        """Cria e exibe um pequeno popup informativo ao passar o mouse sobre uma aula."""
        # Se já existe um tooltip para a MESMA aula, não faz nada (evita flicker entre frame/label)
        if self.tooltip_window:
            if getattr(self.tooltip_window, 'aula_id', None) == aula.id:
                return
            self._hide_tooltip()
        
        # Posição base (padrão: direita do cursor)
        x = event.x_root + 20
        y = event.y_root + 10

        self.tooltip_window = tw = tk.Toplevel(self.root)
        tw.aula_id = aula.id # Marca qual aula está sendo exibida
        tw.wm_overrideredirect(True) # Remove as bordas da janela
        tw.attributes("-topmost", True)
        
        D = self._D
        # Estilo premium com borda sutil
        frame = tk.Frame(tw, bg=D["card"], highlightthickness=1, highlightbackground=D["primary"])
        frame.pack()
        
        # Cabeçalho do Tooltip
        header = tk.Frame(frame, bg=D["primary"])
        header.pack(fill=tk.X)
        tk.Label(header, text="  DETALHES DA AULA", font=("Segoe UI", 8, "bold"),
                 bg=D["primary"], fg="white", pady=4).pack(side=tk.LEFT)

        # Conteúdo principal
        content_frame = tk.Frame(frame, bg=D["card"], pady=10, padx=12)
        content_frame.pack()

        info_text = (
            f"👤 Professor: {aula.professor}\n"
            f"🕒 Horário: {aula.hora_inicio} - {aula.hora_fim}\n"
            f"🎓 Disciplina: {aula.disciplina}\n"
            f"📖 Curso: {aula.curso}\n"
            f"👥 Turma: {aula.turma}\n"
            f"🏢 Facul./Setor: {aula.faculdade}\n"
            f"📅 Dia da Semana: {aula.dia_semana}\n"
            f"📍 Local: LAB {aula.laboratorio}\n"
            f"👥 Alunos: {aula.qtde_alunos}"
        )

        tk.Label(content_frame, text=info_text, justify="left",
                 font=("Segoe UI", 9), bg=D["card"], fg=D["fg"]).pack()

        # Observações (respeitando a opção 'Não mostrar')
        if getattr(aula, 'observacoes', '') and aula.peso_observacao != "Não mostrar":
            # Cores para o destaque
            if aula.peso_observacao == "Importante":
                bg_obs, fg_obs = "#fee2e2", "#991b1b" # Vermelho suave
            elif aula.peso_observacao == "Média":
                bg_obs, fg_obs = "#e0f2fe", "#0369a1" # Azul suave
            elif aula.peso_observacao == "Baixo":
                bg_obs, fg_obs = "#ffedd5", "#9a3412" # Laranja suave
            else:
                bg_obs, fg_obs = "#f1f5f9", "#475569"

            obs_frame = tk.Frame(content_frame, bg=bg_obs, padx=8, pady=8, 
                                highlightthickness=1, highlightbackground=fg_obs)
            obs_frame.pack(fill=tk.X, pady=(10, 0))
            
            tk.Label(obs_frame, text=f"📝 OBSERVAÇÃO ({aula.peso_observacao}):", 
                     font=("Segoe UI", 8, "bold"), bg=bg_obs, fg=fg_obs).pack(anchor="w")
            
            tk.Label(obs_frame, text=aula.observacoes, justify="left", wraplength=250,
                     font=("Segoe UI", 9, "italic"), bg=bg_obs, fg=fg_obs).pack(anchor="w", pady=(2, 0))

        if aula.is_eventual:
            evt_lbl = tk.Label(content_frame, text=f"📅 Data Eventual: {aula.data_eventual}",
                               font=("Segoe UI", 8, "bold"), bg=D["card"], fg="#64748b")
            evt_lbl.pack(anchor="w", pady=(10, 0))

        # --- Lógica de Posicionamento Inteligente (Ajuste para borda direita/múltiplos monitores) ---
        # Força o processamento do layout para obter largura real
        tw.update_idletasks()
        w_tooltip = tw.winfo_reqwidth()
        
        labs_direita = ["12", "13", "14"]
        lab_str = str(aula.laboratorio).replace("Lab", "").replace("LAB", "").strip()
        
        if lab_str in labs_direita:
            # Posiciona à esquerda do cursor
            x = event.x_root - w_tooltip - 20
        
        tw.wm_geometry(f"+{x}+{y}")

    def _hide_tooltip(self, event=None):
        """Fecha a janela de tooltip."""
        if self.tooltip_window:
            try:
                self.tooltip_window.destroy()
            except:
                pass
            self.tooltip_window = None

    def selecionar_tudo(self):
        """Marca ou desmarca todas as aulas agendadas."""
        if self.tree is None: return
        
        items = self.tree.get_children()
        if not items: return
        
        # Se todas estiverem marcadas, desmarca; caso contrário, marca todas
        todas_marcadas = all(self.tree.item(i, "values")[0] == "☑" for i in items)
        novo_estado = "☐" if todas_marcadas else "☑"
        
        for item in items:
            vals = list(self.tree.item(item, "values"))
            vals[0] = novo_estado
            self.tree.item(item, values=vals)
            
        self._atualizar_botao_batch_del()

    def _atualizar_botao_batch_del(self):
        """Mostra/Oculta o botão de exclusão baseado na seleção."""
        if self.tree is None or self.btn_batch_del is None: return
        
        tem_selecionado = False
        for item in self.tree.get_children():
            if self.tree.item(item, "values")[0] == "☑":
                tem_selecionado = True
                break
        
        if tem_selecionado:
            self.btn_batch_del.pack(side=tk.LEFT, padx=2)
        else:
            self.btn_batch_del.pack_forget()

    def excluir_selecionados(self):
        """Exclui todas as aulas marcadas."""
        if self.tree is None: return
        
        selecionados = []
        for item in self.tree.get_children():
            if self.tree.item(item, "values")[0] == "☑":
                tags = self.tree.item(item, "tags")
                if tags:
                    selecionados.append(int(tags[0]))
        
        if not selecionados: return
        
        msg = f"Deseja realmente excluir as {len(selecionados)} aulas selecionadas?"
        if messagebox.askyesno("Confirmar Exclusão", msg):
            self.db.apagar_aulas_lote(selecionados)
            self.atualizar_dados()
            messagebox.showinfo("Sucesso", f"{len(selecionados)} aulas excluídas com sucesso.")

    def _get_aulas_com_sobrescricao(self, dia_semana_str, data_alvo):
        """
        Retorna a lista de aulas para um dia da semana e data específicos.
        Aulas eventuais são incluídas JUNTO com as fixas se coincidirem,
        conforme pedido para mostrar 'na mesma linha'.
        """
        todas = self.db.listar_todas_aulas()
        
        # Filtra fixas para aquele dia da semana
        fixas = [a for a in todas if not a.is_eventual and a.dia_semana == dia_semana_str]
        
        # Filtra eventuais para aquela data exata
        # Suporta tanto DD/MM/YYYY (banco) quanto YYYY-MM-DD (ISO)
        formatos = ["%d/%m/%Y", "%Y-%m-%d"]
        data_str_search = []
        if hasattr(data_alvo, "strftime"):
             for f in formatos:
                  data_str_search.append(data_alvo.strftime(f))
        else:
             data_str_search = [str(data_alvo)]

        eventuais = [a for a in todas if a.is_eventual and a.data_eventual in data_str_search]
        
        # Retorna ambas (fixas antes, eventuais depois para que na grade fiquem por cima)
        return fixas + eventuais

    def _janela_info_aula(self, aula: Aula):
        """Abre uma janela com informações da aula seguindo o design premium solicitado."""
        win = tk.Toplevel(self.root)
        win.title(f"Informações: {aula.disciplina}")
        
        # Largura dinâmica baseada no conteúdo
        win_w, win_h = 480, 560
        sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
        win.geometry(f"{win_w}x{win_h}+{(sw-win_w)//2}+{(sh-win_h)//2}")
        
        win.configure(bg="#1e2540") # Dark Navy
        win.transient(self.root)
        win.grab_set()

        # Header azul prêmio conforme imagem
        header = tk.Frame(win, bg="#3d5af1", height=45)
        header.pack(fill=tk.X)
        tk.Label(header, text="DETALHES DA AULA", font=("Segoe UI", 11, "bold"), 
                 bg="#3d5af1", fg="white", padx=15).pack(side=tk.LEFT, pady=10)

        body = tk.Frame(win, bg="#1e2540", padx=25, pady=20)
        body.pack(fill=tk.BOTH, expand=True)

        # Mapeamento de Ícones conforme imagem
        info_list = [
            ("👤", "Professor", (aula.professor or "Não informado").upper()),
            ("🕒", "Horário", f"{aula.hora_inicio} - {aula.hora_fim}"),
            ("🎓", "Disciplina", (aula.disciplina or "").upper()),
            ("📘", "Curso", (aula.curso or "").upper()),
            ("👥", "Turma", (aula.turma or "").upper()),
            ("🏢", "Facul./Setor", (aula.faculdade or "").upper()),
            ("📅", "Dia da Semana", aula.dia_semana),
            ("📍", "Local", f"LAB {aula.laboratorio}"),
            ("👨‍👩‍👧‍👦", "Alunos", str(aula.qtde_alunos or 0))
        ]
        
        if aula.is_eventual and aula.data_eventual:
             info_list.insert(7, ("🗓️", "Data Eventual", aula.data_eventual))

        # Adicionar campos com ícones e labels selecionáveis
        for icon, label, value in info_list:
            row = tk.Frame(body, bg="#1e2540", pady=6)
            row.pack(fill=tk.X)
            
            # Ícone + Label
            lbl_text = f" {icon}  {label}:"
            tk.Label(row, text=lbl_text, font=("Segoe UI", 10, "bold"), 
                     bg="#1e2540", fg="#8f9dc7", anchor="w").pack(side=tk.LEFT)
            
            # Valor Selecionável (Entry estilizado)
            val_ent = tk.Entry(row, font=("Segoe UI", 10), bg="#1e2540", fg="white",
                              relief="flat", highlightthickness=0, insertbackground="white",
                              readonlybackground="#1e2540")
            val_ent.insert(0, value)
            val_ent.config(state="readonly")
            val_ent.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 0))

        # Observação em estilo diferenciado se existir
        if aula.observacoes:
            obs_frame = tk.Frame(body, bg="#252d4a", padx=10, pady=10, highlightthickness=1, highlightbackground="#3d4f82")
            obs_frame.pack(fill=tk.X, pady=(15, 0))
            
            tk.Label(obs_frame, text=f"📝 Observação ({aula.peso_observacao}):", font=("Segoe UI", 9, "bold"), 
                     bg="#252d4a", fg="#f59e0b").pack(anchor="w")
            
            obs_txt = tk.Text(obs_frame, font=("Segoe UI", 10), bg="#252d4a", fg="white",
                             height=3, relief="flat", wrap="word", highlightthickness=0)
            obs_txt.insert("1.0", aula.observacoes)
            obs_txt.config(state="disabled")
            obs_txt.pack(fill=tk.X, expand=True)

        footer = tk.Frame(win, bg="#1e2540", pady=20)
        footer.pack(fill=tk.X)
        
        btn_fechar = tk.Button(footer, text="FECHAR", font=("Segoe UI", 10, "bold"),
                              bg="#3d5af1", fg="white", relief="flat", cursor="hand2",
                              padx=30, pady=8, command=win.destroy)
        btn_fechar.pack()
        btn_fechar.bind("<Enter>", lambda e: btn_fechar.configure(bg="#2b45d4"))
        btn_fechar.bind("<Leave>", lambda e: btn_fechar.configure(bg="#3d5af1"))

    def debounce(self, key, callback, ms=300):
        """Cancela agendamentos anteriores para a mesma chave e cria um novo."""
        if key in self._after_ids and self._after_ids[key]:
            self.root.after_cancel(self._after_ids[key])
        self._after_ids[key] = self.root.after(ms, callback)

    def _mostrar_menu_contexto_aula(self, event, aula: Aula):
        """Mostra menu de contexto para a aula selecionada."""
        # Garantir que o menu seja criado no Toplevel correto (Grade Principal ou Janela de Grade)
        parent_win = event.widget.winfo_toplevel()
        
        self._ctx_menu = tk.Menu(parent_win, tearoff=0, font=("Segoe UI", 10), bg="#2a2a3e", fg="white", 
                                activebackground="#3d5af1", activeforeground="white")
        
        self._ctx_menu.add_command(label="✏️  Editar Aula", command=lambda: self._janela_cadastro_aula(aula))
        self._ctx_menu.add_command(label="🔄  Alterar LAB", command=lambda: self._janela_alterar_lab(aula))
        self._ctx_menu.add_command(label="🗑️  Excluir Aula", command=lambda: self._confirmar_exclusao_aula(aula), foreground="#ef4444")
        self._ctx_menu.add_separator()
        self._ctx_menu.add_command(label="ℹ️  Informações", command=lambda: self._janela_info_aula(aula))
        
        # O uso de tk_popup é geralmente mais seguro em Windows para menus de contexto
        try:
            self._ctx_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self._ctx_menu.grab_release()

    def _confirmar_exclusao_aula(self, aula: Aula):
        """Pede confirmação e exclui a aula."""
        msg = f"Deseja realmente excluir a aula de '{aula.disciplina}'?\nLab: {aula.laboratorio} | Horário: {aula.hora_inicio}-{aula.hora_fim}"
        if messagebox.askyesno("Confirmar Exclusão", msg):
            self.db.apagar_aula(aula.id)
            self.atualizar_dados()


    def _janela_alterar_lab(self, aula: Aula):
        """Abre uma janela para mover a aula para outro laboratório."""
        win = tk.Toplevel(self.root)
        win.title("Alterar Laboratório")
        win.geometry("350x250")
        win.configure(bg=self._D["bg"])
        win.resizable(False, False)
        win.transient(self.root)
        win.grab_set()

        D = self._D
        tk.Label(win, text="🔄 Mover Aula", font=("Segoe UI", 12, "bold"), 
                 bg=D["bg"], fg=D["primary"]).pack(pady=(20, 10))
        
        tk.Label(win, text=f"Selecione o novo laboratório para:\n{aula.disciplina}", 
                 justify="center", bg=D["bg"], fg=D["fg"], font=("Segoe UI", 9)).pack(pady=5)

        combo_lab = ttk.Combobox(win, values=LABORATORIOS, state="readonly", width=25)
        combo_lab.set(aula.laboratorio)
        combo_lab.pack(pady=15)

        def _confirmar():
            novo_lab = combo_lab.get()
            if novo_lab == aula.laboratorio:
                win.destroy()
                return

            # Criar clone para testar conflito
            import copy
            aula_clone = copy.copy(aula)
            aula_clone.laboratorio = novo_lab
            
            conflito = self.db.verificar_conflito(aula_clone, ignorar_id=aula.id)
            if conflito:
                msg = (f"❌ CONFLITO DE HORÁRIO!\n\n"
                       f"O Laboratório {novo_lab} já possui uma aula neste horário:\n"
                       f"📚 {conflito.disciplina}\n"
                       f"👤 Prof: {conflito.professor}\n"
                       f"🕒 {conflito.hora_inicio} - {conflito.hora_fim}")
                messagebox.showerror("Erro de Alocação", msg, parent=win)
            else:
                aula.laboratorio = novo_lab
                self.db.alterar_aula(aula)
                self.atualizar_dados()
                messagebox.showinfo("Sucesso", f"Aula movida para o Laboratório {novo_lab} com sucesso!", parent=win)
                win.destroy()

        btn_frame = tk.Frame(win, bg=D["bg"])
        btn_frame.pack(pady=10)
        
        ttk.Button(btn_frame, text="CONFIRMAR", command=_confirmar, style="Success.TButton").pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="CANCELAR", command=win.destroy).pack(side=tk.LEFT, padx=5)

    def atualizar_grade(self):
        """Atualiza visualização da grade ou agenda na janela principal."""
        if self.combo_filtro_dia is None or self.combo_filtro_lab is None:
            return
        
        dia = self.combo_filtro_dia.get()
        lab = self.combo_filtro_lab.get()
        turno = self.combo_filtro_turno.get() if self.combo_filtro_turno else "TODOS"
        
        from datetime import datetime, timedelta
        hoje = datetime.now()
        data_ref = hoje
        if dia in DIAS_SEMANA:
            target_idx = DIAS_SEMANA.index(dia)
            hoje_idx = hoje.weekday()
            diff = target_idx - hoje_idx
            data_ref = hoje + timedelta(days=diff)

        q = self.ent_busca_grade.get().strip() if hasattr(self, 'ent_busca_grade') and self.ent_busca_grade else ""
        if hasattr(self, 'grade_search_placeholder') and q == self.grade_search_placeholder: q = ""

        if self.modo_compacto:
            self._draw_agenda(self.inner_grade, dia, lab, data_ref=data_ref, turno=turno, busca=q)
        else:
            self._draw_grade(self.inner_corner, self.inner_hdr_h, self.inner_hdr_v, self.inner_grade, dia, lab, data_ref=data_ref, turno=turno, busca=q)

    def alternar_modo_visualizacao(self):
        """Alterna entre o modo Grade e o modo Agenda Compacta."""
        self.modo_compacto = not self.modo_compacto
        if self.modo_compacto:
            self.btn_toggle_modo.config(text="📅 MODO GRADE")
            # No modo agenda, o canvas de dados ocupa quase tudo, mas mantemos o header horizontal fixo
            gc = self.grade_corner
            ch = self.canvas_hdr_h
            cv = self.canvas_hdr_v
            cg = self.canvas_grade
            if gc and ch and cv and cg:
                # Oculta elementos da Grade
                gc.grid_remove()
                ch.grid_remove()
                cv.grid_remove()
                cg.grid_remove()
                # Garante que o grid_container pai das scrollbars antigas suma
                if self.grid_container:
                     self.grid_container.grid_remove()

                # Mostra elementos da Agenda
                if self.tree_agenda and self.sb_agenda:
                    self.tree_agenda.grid(row=1, column=0, sticky="nsew")
                    self.sb_agenda.grid(row=1, column=1, sticky="ns")
        else:
            self.btn_toggle_modo.config(text="📅 MODO AGENDA")
            gc = self.grade_corner
            ch = self.canvas_hdr_h
            cv = self.canvas_hdr_v
            cg = self.canvas_grade
            if gc and ch and cv and cg:
                # Oculta Agenda
                if self.tree_agenda: self.tree_agenda.grid_remove()
                if self.sb_agenda: self.sb_agenda.grid_remove()
                
                # Restaura Grade
                if self.grid_container: self.grid_container.grid()
                gc.grid()
                ch.grid(row=0, column=1, sticky="ew")
                cv.grid(row=1, column=0, sticky="ns")
                cg.grid(row=1, column=1, sticky="nsew")
        self.atualizar_grade()

    def _on_tree_agenda_double_click(self, event):
        if self.tree_agenda is None: return
        item = self.tree_agenda.identify_row(event.y)
        if item:
            tags = self.tree_agenda.item(item, "tags")
            if tags:
                try:
                    aula_id = int(tags[0])
                    aulas = self.db.listar_todas_aulas()
                    aula = next((a for a in aulas if a.id == aula_id), None)
                    if aula: self._janela_cadastro_aula(aula)
                except (ValueError, IndexError): pass

    def _on_tree_agenda_click(self, event):
        if self.tree_agenda is None: return
        region = self.tree_agenda.identify_region(event.x, event.y)
        if region != "cell": return
        column = self.tree_agenda.identify_column(event.x)
        if column == "#9": # Coluna 'Info'
            item = self.tree_agenda.identify_row(event.y)
            if item:
                tags = self.tree_agenda.item(item, "tags")
                if tags:
                    try:
                        aula_id = int(tags[0])
                        aulas = self.db.listar_todas_aulas()
                        aula = next((a for a in aulas if a.id == aula_id), None)
                        if aula: self._janela_info_aula(aula)
                    except (ValueError, IndexError): pass

    def _on_tree_right_click(self, event):
        """Handler para clique direito na lista de resultados."""
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            tags = self.tree.item(item, "tags")
            if tags:
                try:
                    aula_id = int(tags[0])
                    aulas = self.db.listar_todas_aulas()
                    aula = next((a for a in aulas if a.id == aula_id), None)
                    if aula: self._mostrar_menu_contexto_aula(event, aula)
                except (ValueError, IndexError): pass

    def _on_tree_agenda_right_click(self, event):
        """Handler para clique direito na agenda compacta."""
        if self.tree_agenda is None: return
        item = self.tree_agenda.identify_row(event.y)
        if item:
            self.tree_agenda.selection_set(item)
            tags = self.tree_agenda.item(item, "tags")
            if tags:
                try:
                    aula_id = int(tags[0])
                    aulas = self.db.listar_todas_aulas()
                    aula = next((a for a in aulas if a.id == aula_id), None)
                    if aula: self._mostrar_menu_contexto_aula(event, aula)
                except (ValueError, IndexError): pass

    def _on_empty_cell_context_menu(self, event, dia, lab):
        """Menu de contexto para adicionar aula em espaço vazio."""
        parent_win = event.widget.winfo_toplevel()
        menu = tk.Menu(parent_win, tearoff=0, font=("Segoe UI", 10), bg="#1e1e2e", fg="white", 
                       activebackground="#3d5af1", activeforeground="white")
        
        menu.add_command(label="➕  Adicionar Aula Aqui", 
                         command=lambda: self._janela_cadastro_aula(pre_dia=dia, pre_lab=lab))
        
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def _update_all_grade_marks(self):
        """Atualiza visualmente apenas as linhas de marcação e próxima aula em todas as grades ativas (Performance)."""
        if not hasattr(self, "active_grade_views"): return
        
        # Cores e Constantes
        D = self._D
        SEP_COLOR = D["border"]
        DANGER_COLOR = D["danger"]
        ORANGE_COLOR = "#f59e0b"
        
        from datetime import datetime
        now_dt = datetime.now()
        proxima_dia = self.proxima_aula_info.get("dia", "")
        proxima_hor = self.proxima_aula_info.get("horario", "")
        
        # Limpa views que não existem mais (widget destruído)
        self.active_grade_views = [v for v in self.active_grade_views if v.winfo_exists()]
        
        for view in self.active_grade_views:
            if not hasattr(view, "_row_refs") or not hasattr(view, "_sorted_times"):
                continue
            
            row_refs     = view._row_refs
            sorted_times = view._sorted_times
            mode         = view._mode
            view_dia     = view._dia
            
            # Determina o dia da view
            dia_v = view_dia if view_dia and view_dia != self._TODOS_DIAS else DIAS_SEMANA[now_dt.weekday()]
            
            # 1. Linhas de Intervalo (Loop 0 até N-2)
            for i in range(len(sorted_times) - 1):
                t_str = sorted_times[i].strftime("%H:%M")
                t_prev_str = sorted_times[i-1].strftime("%H:%M") if i > 0 else None
                
                # Lógica de Cor (Igual ao _draw_grade)
                is_next = (t_str == proxima_hor and (mode == "LAB" or dia_v == proxima_dia))
                is_marked = (t_str in self.horarios_marcados or (t_prev_str and t_prev_str in self.horarios_marcados))
                
                c = ORANGE_COLOR if is_next else (DANGER_COLOR if is_marked else SEP_COLOR)
                h = 1 if c == SEP_COLOR else 2
                
                if i in row_refs:
                    for s in row_refs[i][2]: # row_seps
                        s.configure(bg=c, height=h)
                        if c == SEP_COLOR: s.lower()
                        else: s.lift()
            
            # 2. Linha Final (Bottom)
            t_final_str = sorted_times[-1].strftime("%H:%M")
            t_last_slot_str = sorted_times[-2].strftime("%H:%M") if len(sorted_times) > 1 else None
            
            is_next_final = (t_final_str == proxima_hor and (mode == "LAB" or dia_v == proxima_dia))
            is_marked_final = (t_final_str in self.horarios_marcados or (t_last_slot_str and t_last_slot_str in self.horarios_marcados))
            
            c_f = ORANGE_COLOR if is_next_final else (DANGER_COLOR if is_marked_final else SEP_COLOR)
            h_f = 1 if c_f == SEP_COLOR else 2
            
            if hasattr(view, "_sep_final") and view._sep_final.winfo_exists():
                view._sep_final.configure(bg=c_f, height=h_f)
                if c_f == SEP_COLOR: view._sep_final.lower()
                else: view._sep_final.lift()

    def _draw_grade(self, inner_corner, inner_hdr_h, inner_hdr_v, inner_data, dia, lab_selecionado, data_ref=None, turno="TODOS", busca=""):
        """Desenha a grade de horários distribuída nos componentes de sticky header."""
        from datetime import datetime, timedelta
        if data_ref is None:
            data_ref = datetime.now()

        # Normaliza valores "Todos" para None
        dia_real = dia if dia and dia != self._TODOS_DIAS else None
        lab_real = lab_selecionado if lab_selecionado and lab_selecionado != self._TODOS_LABS else None

        if lab_real:
            MODE = "LAB"
            COLUNAS_TITULOS = DIAS_SEMANA
        elif dia_real:
            MODE = "DIA"
            COLUNAS_TITULOS = LABORATORIOS
        else:
            # Nenhum filtro específico → mostra tudo
            MODE = "TODOS"
            COLUNAS_TITULOS = LABORATORIOS

        hora_inicio_base = datetime.strptime("07:30", "%H:%M")
        hora_fim_base    = datetime.strptime("23:15", "%H:%M")
        intervalo        = timedelta(minutes=15)

        # Cores e Constantes locais
        D           = self._D
        EMPTY_BG    = D["bg"]
        CARD_BG     = D["card"]
        TIME_BG     = D["hdr"]
        HEADER_BG   = D["hdr"]
        HOVER_BG    = "#2e3b63" # Azul escuro suave para hover no tema dark
        TEXT_FG     = D["fg"]
        TEXT_MAIN   = D["fg"]
        SEP_COLOR   = D["border"]

        # Limpar frames
        for w in inner_corner.winfo_children(): w.destroy()
        for w in inner_hdr_h.winfo_children(): w.destroy()
        for w in inner_hdr_v.winfo_children(): w.destroy()
        for w in inner_data.winfo_children(): w.destroy()

        inner_data.configure(bg=EMPTY_BG)
        inner_hdr_h.configure(bg=HEADER_BG)
        inner_hdr_v.configure(bg=TIME_BG)

        # 1. Corner
        tk.Label(inner_corner, text="HORA", bg=HEADER_BG, fg="white",
                 font=("Segoe UI", 9, "bold"), relief="flat", bd=0, width=8,
                 anchor="center"
                 ).pack(fill=tk.BOTH, expand=True)

        # 2. Header Horizontal com Status em Tempo Real
        col_hdr_refs = {}
        # Busca status dos labs para o projeto todo
        lab_status_db = self.db.obter_todos_status_labs()
        
        for col, titulo in enumerate(COLUNAS_TITULOS, start=1):
            t_format = titulo.split("-")[0].capitalize() if MODE == "LAB" else titulo
            
            # Lógica de cor e ícone baseada no Status
            bg_hdr = HEADER_BG
            fg_hdr = "white"
            status = "Desligado"
            
            if MODE in ["DIA", "TODOS"]:
                lab_key = titulo if titulo.startswith("Lab") else f"Lab {titulo}"
                status = lab_status_db.get(lab_key, "Desligado")
            elif MODE == "LAB" and lab_real:
                lab_key = lab_real if str(lab_real).startswith("Lab") else f"Lab {lab_real}"
                status = lab_status_db.get(lab_key, "Desligado")

            # Mapeamento de Status para Cores e Ícones
            status_icon = "⚪" # Padrão
            if status == "Ligado":
                bg_hdr = "#10b981"
                status_icon = "🟢"
            elif status == "Finalizado":
                bg_hdr = "#4b5563"
                status_icon = "⚪"
            elif status == "Desligado":
                bg_hdr = "#ef4444"
                status_icon = "🔴"

            # Formata o texto do cabeçalho com o ícone de status
            display_text = f"{status_icon} {t_format}"
            
            lbl_hdr = tk.Label(inner_hdr_h, text=display_text,
                     bg=bg_hdr, fg=fg_hdr, font=("Segoe UI", 9, "bold"),
                     relief="flat", bd=0, width=12, anchor="center")
            lbl_hdr.grid(row=0, column=col-1, sticky="nsew", padx=1, pady=1)
            
            # Bind para clique com botão direito (Menu de Status) - SOLICITAÇÃO DO USUÁRIO
            lbl_hdr.bind("<Button-3>", lambda e, lk=lab_key: self._mostrar_menu_status_lab(e, lk))

            # Armazena a cor original para o efeito de hover (restaurar cor do status)
            setattr(lbl_hdr, "orig_bg", bg_hdr)
            
            inner_hdr_h.columnconfigure(col-1, weight=1, minsize=110, uniform="grade_col")
            inner_data.columnconfigure(col-1, weight=1, minsize=110, uniform="grade_col")
            col_hdr_refs[col-1] = lbl_hdr

        # Buscar aulas
        aulas_para_grade = []
        # data_ref já vem por argumento ou inicializado no topo do método
        if data_ref is None:
            data_ref = datetime.now()

        if MODE == "DIA":
            aulas_para_grade = self._get_aulas_com_sobrescricao(dia_real, data_ref)
        elif MODE == "LAB":
            # Para o modo LAB, mostramos todos os dias, mas filtramos por dia_real se selecionado.
            segunda_atual = data_ref - timedelta(days=data_ref.weekday())
            aulas_para_grade = []
            for i, d_nome in enumerate(DIAS_SEMANA):
                # Se houver filtro de dia e não for este dia, pula
                if dia_real and d_nome != dia_real:
                    continue
                data_dia = segunda_atual + timedelta(days=i)
                aulas_para_grade.extend(self._get_aulas_com_sobrescricao(d_nome, data_dia))
            
            # Filtra apenas o lab selecionado (Normalizado para evitar erro de '1' vs '01')
            def norm_lab(l): return str(l).zfill(2) if str(l).isdigit() else str(l)
            aulas_para_grade = [a for a in aulas_para_grade if norm_lab(a.laboratorio) == norm_lab(lab_real)]
        else:
            # TODOS OS DIAS + TODOS OS LABS - Mostra a semana inteira achatada
            # (Útil para pesquisar Professor/Disciplina em todos os dias da grade)
            segunda_atual = data_ref - timedelta(days=data_ref.weekday())
            aulas_para_grade = []
            for i, d_nome in enumerate(DIAS_SEMANA):
                d_alvo = segunda_atual + timedelta(days=i)
                aulas_para_grade.extend(self._get_aulas_com_sobrescricao(d_nome, d_alvo))

        # Filtro de Turno
        if turno and turno != "TODOS":
            aulas_filtradas = []
            for a in aulas_para_grade:
                try:
                    h_ini = int(a.hora_inicio.split(":")[0])
                    if turno == "MATUTINO" and 7 <= h_ini < 13:
                        aulas_filtradas.append(a)
                    elif turno == "VESPERTINO" and 13 <= h_ini < 18:
                        aulas_filtradas.append(a)
                    elif turno == "NOTURNO" and h_ini >= 18:
                        aulas_filtradas.append(a)
                except:
                    aulas_filtradas.append(a)
            aulas_para_grade = aulas_filtradas

        # Lógica de filtro (busca)
        if busca:
            busca = remover_acentos(busca)

        if busca:
            aulas_filtradas = []
            for aula in aulas_para_grade:
                campos = [
                    aula.professor,
                    aula.curso,
                    aula.disciplina,
                    aula.turma,
                    aula.faculdade
                ]
                if any(busca in remover_acentos(str(c)) for c in campos if c):
                    aulas_filtradas.append(aula)
            aulas_para_grade = aulas_filtradas

        # Agrupar aulas por coluna e detectar clusters de sobreposição (incluindo parciais)
        clusters_por_coluna = {}
        aulas_col = {}
        for aula in aulas_para_grade:
            col_id = aula.dia_semana if MODE == "LAB" else aula.laboratorio
            h_ini = datetime.strptime(aula.hora_inicio, "%H:%M")
            h_fim = datetime.strptime(aula.hora_fim,    "%H:%M")
            aulas_col.setdefault(col_id, []).append((aula, h_ini, h_fim))
            
        for col_id, lista in aulas_col.items():
            lista.sort(key=lambda x: x[1]) # Ordena por início
            clusters = []
            if not lista: continue
            
            curr_cluster = [lista[0]]
            last_end = lista[0][2]
            for i in range(1, len(lista)):
                a, hi, hf = lista[i]
                if hi < last_end: # Sobreposição (parcial ou total)
                    curr_cluster.append(lista[i])
                    last_end = max(last_end, hf)
                else:
                    clusters.append(curr_cluster)
                    curr_cluster = [lista[i]]
                    last_end = hf
            clusters.append(curr_cluster)
            clusters_por_coluna[col_id] = clusters

        # 3. Coluna de Horas e Body Grid
        # 3. Determinar Horários Únicos para a Grade Dinâmica
        tempos = set()
        tempos.add(hora_inicio_base)
        tempos.add(hora_fim_base)
        for aula in aulas_para_grade:
            try:
                tempos.add(datetime.strptime(aula.hora_inicio, "%H:%M"))
                tempos.add(datetime.strptime(aula.hora_fim,    "%H:%M"))
            except: pass
        
        sorted_times = sorted(list(tempos))
        time_to_idx = {t: i for i, t in enumerate(sorted_times)}
        
        num_colunas = len(COLUNAS_TITULOS)
        row_refs = {}
        col_cells = {c: [] for c in range(num_colunas)} # Armazena refs por coluna para hover vertical
        PIXELS_PER_MIN = 1.6

        for i in range(len(sorted_times) - 1):
            t_atual = sorted_times[i]
            t_prox  = sorted_times[i+1]
            diff_min = (t_prox - t_atual).total_seconds() / 60
            h_px = int(diff_min * PIXELS_PER_MIN)
            
            # Label de Hora (Alinhado exatamente ao topo 'nw' para alinhar com o divisor)
            lbl_h = tk.Label(inner_hdr_v, text=t_atual.strftime("%H:%M"), bg=TIME_BG, fg=TEXT_MAIN,
                             font=("Segoe UI", 8, "bold"), width=8, anchor="nw", height=1)
            # Remove padding para garantir alinhamento perfeito com o topo da célula correspondente
            lbl_h.grid(row=i, column=0, sticky="nsew", padx=(5, 1), pady=0)
            inner_hdr_v.rowconfigure(i, minsize=h_px)

            # Cor do divisor horizontal
            t_str = t_atual.strftime("%H:%M")
            t_prev_str = sorted_times[i-1].strftime("%H:%M") if i > 0 else None
            
            # --- LÓGICA DE CORES DA LINHA ---
            # 1. Próxima Aula (Laranja - Prioridade)
            # t_format = dia se MODE=="LAB", senão dia_selecionado
            dia_atual_grade = dia if MODE != "LAB" else t_atual.strftime("%H:%M") # No modo LAB t_atual é dia? Não, t_atual é hora.
            # No modo LAB, MODE="LAB", o dia varia por coluna. No modo DIA/TODOS, o dia é fixo para a visualização.
            
            # Precisamos descobrir se esta linha 'i' no dia 'dia' (ou na coluna 'col') deve ser laranja.
            is_next_class = False
            if hasattr(self, "proxima_aula_info"):
                p_dia = self.proxima_aula_info.get("dia")
                p_hor = self.proxima_aula_info.get("horario")
                if p_hor == t_str:
                    # Verifica dia
                    if MODE != "LAB":
                        # Estamos vendo um dia específico (ou HOJE no modo TODOS)
                        dia_vendo = dia if dia and dia != self._TODOS_DIAS else DIAS_SEMANA[datetime.now().weekday()]
                        if dia_vendo == p_dia: is_next_class = True
                    else:
                        # Modo LAB: cada coluna é um dia. A linha de hora é a mesma pra todas as colunas.
                        # Então destacamos se a hora bater (a coluna específica será tratada se quisermos, 
                        # mas por enquanto a linha horizontal toda fica laranja no horário da próxima aula)
                        is_next_class = True

            if is_next_class:
                c_sep = "#f59e0b" # Laranja (Warning)
            elif (t_str in self.horarios_marcados or t_prev_str in self.horarios_marcados):
                c_sep = D["danger"] # Vermelho (Marcado)
            else:
                c_sep = SEP_COLOR # Padrão
            # --------------------------------
            
            # Células de fundo por coluna para permitir highlight vertical
            row_cells = []
            row_seps = []
            for c_idx in range(num_colunas):
                cell_bg = tk.Frame(inner_data, bg=EMPTY_BG)
                cell_bg.grid(row=i, column=c_idx, sticky="nsew")
                cell_bg.lower() 
                row_cells.append(cell_bg)
                col_cells[c_idx].append(cell_bg) # Armazena para hover de coluna
                
                # Divisor horizontal 
                sep = tk.Frame(inner_data, bg=c_sep, height=1 if c_sep == SEP_COLOR else 2)
                sep.grid(row=i, column=c_idx, sticky="ewn") 
                if c_sep != SEP_COLOR: sep.lift()
                else: sep.lower()
                row_seps.append(sep)

                # DETERMINAR DIA E LAB PARA O MENU DE CONTEXTO EM ESPAÇO VAZIO
                titulo_col = COLUNAS_TITULOS[c_idx]
                if MODE == "LAB":
                    d_ctx, l_ctx = titulo_col, lab_real
                elif MODE == "DIA":
                    d_ctx, l_ctx = dia_real, titulo_col
                else: # TODOS
                    import datetime
                    hoje_idx = datetime.datetime.now().weekday()
                    d_ctx = DIAS_SEMANA[hoje_idx] if hoje_idx < 6 else "SEGUNDA-FEIRA"
                    l_ctx = titulo_col
                
                # Normaliza lab para "Lab 01"
                l_ctx_norm = l_ctx if str(l_ctx).startswith("Lab") else f"Lab {l_ctx}"
                
                # Bind clique direito no espaço vazio
                cell_bg.bind("<Button-3>", lambda e, d=d_ctx, l=l_ctx_norm: self._on_empty_cell_context_menu(e, d, l))

            inner_data.rowconfigure(i, minsize=h_px)
            
            def _hover_cross(e, r=i, c=None, mode="enter"):
                """Efeito de hover em cruz: destaca linha, coluna e cabeçalhos."""
                color = HOVER_BG if mode == "enter" else EMPTY_BG
                # 1. Linha (Row)
                for rb in row_cells: rb.configure(bg=color)
                # 2. Hora (Hdr Vertical)
                lbl_h.configure(bg=HOVER_BG if mode == "enter" else TIME_BG)
                
                # 3. Coluna (Column) - Se c_idx for conhecido
                if c is not None:
                    for cb in col_cells[c]: cb.configure(bg=color)
                    # 4. Lab (Hdr Horizontal)
                    hdr_l = col_hdr_refs.get(c)
                    if hdr_l:
                         # No modo DIA usa azul, em outros pode usar a cor original do status ou azul
                         h_color = "#3d5af1" if mode == "enter" else getattr(hdr_l, "orig_bg", HEADER_BG)
                         hdr_l.configure(bg=h_color)
            
            for c_idx, rb in enumerate(row_cells):
                # Usamos default values no lambda para capturar os índices corretos da iteração
                rb.bind("<Enter>", lambda e, r=i, c=c_idx: _hover_cross(e, r, c, mode="enter"))
                rb.bind("<Leave>", lambda e, r=i, c=c_idx: _hover_cross(e, r, c, mode="leave"))
            
            # Toggle de marcação ao clicar no horário
            def _on_time_click(e, t=t_str):
                if t in self.horarios_marcados:
                    self.horarios_marcados.remove(t)
                else:
                    self.horarios_marcados.add(t)
                self._update_all_grade_marks() # Update rápido em todas as janelas

            lbl_h.configure(cursor="hand2")
            lbl_h.bind("<Button-1>", _on_time_click)
            
            row_refs[i] = (row_cells, lbl_h, row_seps)

        # Adiciona o label final (fim do dia)
        t_final_str = sorted_times[-1].strftime("%H:%M")
        t_last_slot_str = sorted_times[-2].strftime("%H:%M") if len(sorted_times) > 1 else None
        
        # Cor final é vermelha se o último horário clicado for o final do slot anterior
        c_sep_final = D["danger"] if t_final_str in self.horarios_marcados or t_last_slot_str in self.horarios_marcados else SEP_COLOR

        # Label final (fim do dia, também alinhado ao topo para bater com sep_final)
        lbl_h_final = tk.Label(inner_hdr_v, text=t_final_str, bg=TIME_BG, fg=TEXT_MAIN,
                               font=("Segoe UI", 8, "bold"), width=8, anchor="nw", height=1, cursor="hand2")
        lbl_h_final.grid(row=len(sorted_times)-1, column=0, sticky="nsew", padx=(5, 1), pady=0)
        
        def _on_final_time_click(e, t=t_final_str):
            if t in self.horarios_marcados:
                self.horarios_marcados.remove(t)
            else:
                self.horarios_marcados.add(t)
            self._update_all_grade_marks() # Update rápido em todas as janelas

        lbl_h_final.bind("<Button-1>", _on_final_time_click)

        # Linha final no corpo
        sep_final = tk.Frame(inner_data, bg=c_sep_final, height=1 if c_sep_final == SEP_COLOR else 2)
        sep_final.grid(row=len(sorted_times)-1, column=0, columnspan=num_colunas, sticky="ewn")
        if c_sep_final != SEP_COLOR: sep_final.lift()
        else: sep_final.lower()

        # Vincula referências para update rápido
        inner_data._row_refs     = row_refs
        inner_data._sorted_times = sorted_times
        inner_data._sep_final    = sep_final
        inner_data._mode         = MODE
        inner_data._dia          = dia


        # Desenhar clusters
        for col_id, clusters in clusters_por_coluna.items():
            # Normalizar ID do lab (Remover "Lab ", "Laboratório ", etc e padronizar "01") para bater com COLUNAS_TITULOS
            search_id = str(col_id)
            if MODE != "LAB":
                # Tenta extrair apenas o número se houver prefixo
                s_clean = search_id.replace("Lab ", "").replace("Laboratório ", "").replace("Lab", "").strip()
                if s_clean.isdigit():
                    search_id = f"{int(s_clean):02d}"
                else:
                    search_id = s_clean
            
            try: col_idx = COLUNAS_TITULOS.index(search_id)
            except: continue
            
            for cluster in clusters:
                # ... (lanes logic unchanged) ...
                lanes_end = []
                aula_vaga = {}
                for item in cluster:
                    aula, hi, hf = item
                    v_idx = -1
                    for i, end_t in enumerate(lanes_end):
                        if hi >= end_t:
                            v_idx = i
                            lanes_end[i] = hf
                            break
                    if v_idx == -1:
                        v_idx = len(lanes_end)
                        lanes_end.append(hf)
                    aula_vaga[aula] = v_idx

                num_vagas = len(lanes_end)
                
                # CÁLCULO DINÂMICO DE ROW E ROWSPAN
                r_start = time_to_idx.get(min(it[1] for it in cluster), 0)
                r_fim_idx = time_to_idx.get(max(it[2] for it in cluster), len(sorted_times)-1)
                r_span = max(1, r_fim_idx - r_start)

                container = tk.Frame(inner_data, bg=EMPTY_BG)
                # Remove padding 1 para a aula não ficar deslocada em relação à linha de hora
                container.grid(row=r_start, column=col_idx, rowspan=r_span, sticky="nsew", padx=0, pady=0)
                
                for (aula, hi, hf) in cluster:
                    v_idx = aula_vaga[aula]
                    
                    # Cálculo relativo ao container (em pixels)
                    # hi e hf estão garantidos em sorted_times
                    y_pos = (hi - sorted_times[r_start]).total_seconds() / 60 * PIXELS_PER_MIN
                    h_px  = (hf - hi).total_seconds() / 60 * PIXELS_PER_MIN
                    
                    prefixo_dia = f"({aula.dia_semana[:3]}) " if MODE == "TODOS" else ""
                    prefixo_eventual = "🕒 [EVENTUAL] " if aula.is_eventual else ""
                    bloco = (f"{aula.hora_inicio}-{aula.hora_fim}\n"
                             f"Prof: {aula.professor}\n"
                             f"{prefixo_dia}{prefixo_eventual}{aula.curso}\n{aula.disciplina}\n"
                             f"Turma: {aula.turma}\n"                             
                             f"Qtd: {aula.qtde_alunos}\n"
                             f"Dia: {aula.dia_semana}\n")
                    
                    from utils import formatar_cor_hex
                    # Eventuais sempre usam fundo preto
                    if aula.is_eventual:
                        bg = "#000000"
                    else:
                        bg = formatar_cor_hex(aula.cor_fundo) if (aula.cor_fundo and aula.cor_fundo != "#ffffff") else "#ddeeff"
                    card = tk.Frame(container, bg=bg, relief="flat", bd=0)
                    
                    card.place(relx=v_idx/num_vagas, y=y_pos, relwidth=1.0/num_vagas, height=max(20, h_px))
                    
                    lbl = tk.Label(card, text=bloco, bg=bg, fg=texto_contraste(bg),
                                   justify="center", font=("Segoe UI", 7 if num_vagas>1 else 8),
                                   wraplength=60 if num_vagas>1 else 100)
                    lbl.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)

                    # Balão de Observação
                    peso_obs = getattr(aula, 'peso_observacao', 'Baixo')
                    if peso_obs != "Não mostrar" and aula.observacoes:
                        balloon = ImportantInfoBalloon(card, text=aula.observacoes, priority=peso_obs)
                        # Posiciona no canto superior direito
                        balloon.place(relx=1.0, rely=0.0, x=-2, y=2, anchor="ne")
                        # Bindings para o balão também
                        balloon.bind("<Button-3>", lambda e, a=aula: self._mostrar_menu_contexto_aula(e, a))

                    # BINDINGS PARA TOOLTIP (HOVER)
                    def _enter(e, a=aula):
                        self._show_tooltip(e, a)
                    def _leave(e):
                        self._hide_tooltip(e)

                    card.bind("<Enter>", _enter)
                    card.bind("<Leave>", _leave)
                    card.bind("<Button-1>", _enter) 
                    lbl.bind("<Enter>", _enter)
                    lbl.bind("<Leave>", _leave)
                    lbl.bind("<Button-1>", _enter) 
                    
                    # Bindings para Menu de Contexto (Direito) e Edição (Double Click)
                    lbl.bind("<Double-Button-1>", lambda e, a=aula: self._janela_cadastro_aula(a))
                    card.bind("<Button-3>", lambda e, a=aula: self._mostrar_menu_contexto_aula(e, a))
                    lbl.bind("<Button-3>", lambda e, a=aula: self._mostrar_menu_contexto_aula(e, a))
                    
                    # Hover highlight: marca todos os índices de tempo cobertos
                    idx_hi = time_to_idx.get(hi)
                    idx_hf = time_to_idx.get(hf)

                    def _h_ent(e, rs=idx_hi, rf=idx_hf, ci=col_idx):
                        # 1. Highlight da Linha (apenas no intervalo da aula)
                        for i in range(rs, rf):
                            if i in row_refs:
                                for rb in row_refs[i][0]: rb.configure(bg=HOVER_BG)
                                row_refs[i][1].configure(bg=HOVER_BG)
                        
                        # 2. Highlight da Coluna (toda a extensão vertical)
                        col_hi = "#2d365a" # Hover quase transparente
                        for r_idx in row_refs:
                            row_refs[r_idx][0][ci].configure(bg=col_hi)
                        
                        # 3. Highlight do Header (e preservação do status)
                        if ci in col_hdr_refs:
                            col_hdr_refs[ci].configure(bg="#3d5af1", fg="white")

                    def _h_lev(e, rs=idx_hi, rf=idx_hf, ci=col_idx):
                        # 1. Restaura Linhas
                        for i in range(rs, rf):
                            if i in row_refs:
                                for rb in row_refs[i][0]: rb.configure(bg=EMPTY_BG)
                                row_refs[i][1].configure(bg=TIME_BG)
                        
                        # 2. Restaura Coluna
                        for r_idx in row_refs:
                            row_refs[r_idx][0][ci].configure(bg=EMPTY_BG)
                            
                        # 3. Restaura Header (cor original de status)
                        if ci in col_hdr_refs:
                            orig = getattr(col_hdr_refs[ci], "orig_bg", HEADER_BG)
                            col_hdr_refs[ci].configure(bg=orig, fg="white")

                    lbl.bind("<Enter>", _h_ent); lbl.bind("<Leave>", _h_lev)

        # ------------------------------------------------------------------
        # ROLAGEM AUTOMÁTICA PARA O HORÁRIO ATUAL
        # ------------------------------------------------------------------
        try:
            # Processa layout interno para medição
            inner_data.update_idletasks()
            canvas = inner_data.master
            if isinstance(canvas, tk.Canvas):
                agora = datetime.now()
                # Verifica se o horário atual está dentro dos limites da grade exibida
                if sorted_times and (sorted_times[0] <= agora <= sorted_times[-1]):
                    total_dur = (sorted_times[-1] - sorted_times[0]).total_seconds()
                    if total_dur > 0:
                        rel_dur = (agora - sorted_times[0]).total_seconds()
                        scroll_fraction = rel_dur / total_dur
                        # Ajustamos um pouco (ex: voltamos 10% do topo) para não ficar colado no topo
                        canvas.yview_moveto(max(0.0, float(scroll_fraction - 0.05)))
        except Exception as e:
            print(f"Erro no auto-scroll: {e}")

    def _mudar_ordem_agenda(self, col_idx):
        """Troca a coluna de ordenação ou inverte a direção."""
        if self.agenda_sort_col == col_idx:
            self.agenda_sort_desc = not self.agenda_sort_desc
        else:
            self.agenda_sort_col = col_idx
            self.agenda_sort_desc = False
        self.atualizar_grade()

    def _draw_agenda(self, inner_grade, dia, lab_selecionado, data_ref=None, turno="TODOS", busca=""):
        """Desenha a visualização de agenda compacta usando Treeview."""
        from datetime import datetime, timedelta
        if not hasattr(self, 'tree_agenda'): return
        
        if data_ref is None:
            data_ref = datetime.now()
        
        # Limpa Treeview anterior
        for item in self.tree_agenda.get_children():
            self.tree_agenda.delete(item)

        # Normaliza filtros
        dia_real = dia if dia and dia != self._TODOS_DIAS else None
        lab_real = lab_selecionado if lab_selecionado and lab_selecionado != self._TODOS_LABS else None

        # Buscar aulas
        def norm_lab(l): return str(l).zfill(2) if str(l).isdigit() else str(l)

        aulas = []
        if lab_real and dia_real:
            aulas = self._get_aulas_com_sobrescricao(dia_real, data_ref)
            aulas = [a for a in aulas if norm_lab(a.laboratorio) == norm_lab(lab_real)]
        elif lab_real or (not lab_real and not dia_real):
            segunda_atual = data_ref - timedelta(days=data_ref.weekday())
            for i, d_nome in enumerate(DIAS_SEMANA):
                data_dia = segunda_atual + timedelta(days=i)
                dia_aulas = self._get_aulas_com_sobrescricao(d_nome, data_dia)
                if lab_real: dia_aulas = [a for a in dia_aulas if norm_lab(a.laboratorio) == norm_lab(lab_real)]
                aulas.extend(dia_aulas)
        elif dia_real:
            aulas = self._get_aulas_com_sobrescricao(dia_real, data_ref)

        # Filtro de Busca
        if busca:
            busca = remover_acentos(busca)
            aulas = [a for a in aulas if any(busca in remover_acentos(str(c)) for c in [a.professor, a.curso, a.disciplina, a.turma, a.faculdade] if c)]

        # Filtro de Turno
        if turno and turno != "TODOS":
            aulas_filtradas = []
            for a in aulas:
                try:
                    h_ini = int(a.hora_inicio.split(":")[0])
                    if turno == "MATUTINO" and 7 <= h_ini < 13: aulas_filtradas.append(a)
                    elif turno == "VESPERTINO" and 13 <= h_ini < 18: aulas_filtradas.append(a)
                    elif turno == "NOTURNO" and h_ini >= 18: aulas_filtradas.append(a)
                except: aulas_filtradas.append(a)
            aulas = aulas_filtradas

        ordem_dias = {d: i for i, d in enumerate(DIAS_SEMANA, 1)}
        def get_sort_key(a):
            ord_dia = ordem_dias.get(a.dia_semana, 9)
            if self.agenda_sort_col == 0: return (a.laboratorio, ord_dia, a.hora_inicio)
            if self.agenda_sort_col == 1: return (ord_dia, a.hora_inicio, a.laboratorio)
            if self.agenda_sort_col == 2: return (a.hora_inicio, ord_dia, a.laboratorio)
            if self.agenda_sort_col == 3: return ((a.disciplina or "").upper(), ord_dia, a.hora_inicio)
            if self.agenda_sort_col == 4: return ((a.turma or "").upper(), ord_dia, a.hora_inicio)
            if self.agenda_sort_col == 5: return (a.qtde_alunos, ord_dia, a.hora_inicio)
            if self.agenda_sort_col == 6: return ((a.professor or "").upper(), ord_dia, a.hora_inicio)
            if self.agenda_sort_col == 7: return ((a.curso or "").upper(), ord_dia, a.hora_inicio)
            if self.agenda_sort_col == 8: return (a.id, ord_dia, a.hora_inicio)
            return (a.laboratorio, ord_dia, a.hora_inicio)
        
        aulas.sort(key=get_sort_key, reverse=self.agenda_sort_desc)

        if not aulas:
            # Mostra mensagem se vazio
            self.tree_agenda.insert("", "end", values=("Nenhuma aula encontrada", "", "", "", "", "", "", ""))
        else:
            for aula in aulas:
                dia_exibicao = aula.dia_semana
                if aula.is_eventual and aula.data_eventual:
                    try:
                        for fmt in ["%d/%m/%Y", "%Y-%m-%d"]:
                            try:
                                dt_ev = datetime.strptime(aula.data_eventual, fmt)
                                dia_ev_nome = DIAS_SEMANA[dt_ev.weekday()]
                                if dia_ev_nome != aula.dia_semana:
                                    dia_exibicao = f"{dia_ev_nome} ({aula.data_eventual})"
                                break
                            except ValueError: continue
                    except Exception: pass

                # Valores da linha
                values = (
                    aula.laboratorio,
                    dia_exibicao,
                    f"{aula.hora_inicio} - {aula.hora_fim}",
                    aula.disciplina,
                    aula.turma,
                    str(aula.qtde_alunos),
                    (aula.professor or "").upper(),
                    (aula.curso or "").upper(),
                    "ⓘ"
                )
                
                # Tag com ID para callbacks e Tag de cor para fundo
                tag_id = str(aula.id)
                # Eventuais sempre usam fundo preto
                if aula.is_eventual:
                    bg_color = "#000000"
                else:
                    bg_color = aula.cor_fundo if (aula.cor_fundo and aula.cor_fundo != "#ffffff") else None
                tag_color = f"color_{aula.id}" if bg_color else "default"
                
                if bg_color:
                    fg_color = texto_contraste(bg_color)
                    self.tree_agenda.tag_configure(tag_color, background=bg_color, foreground=fg_color)
                
                self.tree_agenda.insert("", "end", values=values, tags=(tag_id, tag_color))

        if not self.modo_compacto:
            inner_grade.update_idletasks()
            if self.canvas_grade:
                self.canvas_grade.configure(scrollregion=self.canvas_grade.bbox("all"))

    def _on_agenda_row_enter(self, row_idx, widgets_dict):
        """Highlight na linha inteira da agenda."""
        if row_idx in widgets_dict:
            for w in widgets_dict[row_idx]:
                w.configure(bg="#2e3b63") # Azul escuro suave

    def _check_agenda_leave(self, row_idx, widgets_dict, original_bg):
        """Verifica se realmente saiu da linha antes de resetar a cor."""
        if getattr(self, "_current_agenda_hover", -1) == row_idx:
            x, y = self.root.winfo_pointerxy()
            widget_under = self.root.winfo_containing(x, y)
            if widget_under in widgets_dict.get(row_idx, []):
                return
        
        self._current_agenda_hover = -1
        if row_idx in widgets_dict:
            for w in widgets_dict[row_idx]:
                w.configure(bg=original_bg)

    # ------------------------------------------------------------------
    def abrir_janela_grade(self):
        """Abre uma janela separada e maximizada com a Visualização da Grade."""
        win = tk.Toplevel(self.root)
        win.title("Visualização da Grade de Horários")
        win.state("zoomed")           # maximiza no Windows
        win.configure(bg=self._D["bg"])

        # --- Barra de filtros ---
        filtro_frame = ttk.Frame(win, padding="8")
        filtro_frame.pack(side=tk.TOP, fill=tk.X)

        # Barra de Pesquisa em Tempo Real (NOVO)
        search_card = tk.Frame(filtro_frame, bg=self._D["input"], highlightthickness=1,
                                highlightbackground=self._D["border"], highlightcolor=self._D["primary"])
        search_card.pack(side=tk.LEFT, padx=(5, 15), pady=2)

        tk.Label(search_card, text=" 🔍 ", bg=self._D["input"], fg=self._D["fg2"],
                    font=("Segoe UI", 10)).pack(side=tk.LEFT, padx=(5, 0))

        search_placeholder = "Pesquisar Professor, Disciplina, Curso..."
        win_ent_busca = tk.Entry(search_card, font=("Segoe UI", 10),
                                  bg=self._D["input"], relief="flat", bd=0,
                                  fg=self._D["fg"], insertbackground=self._D["fg"], width=30)
        win_ent_busca.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=8, pady=5)
        win_ent_busca.insert(0, search_placeholder)
        win_ent_busca.config(fg=self._D["fg2"])

        ttk.Label(filtro_frame, text="Dia da Semana:").pack(side=tk.LEFT, padx=5)
        win_combo_dia = ttk.Combobox(
            filtro_frame, values=[self._TODOS_DIAS] + DIAS_SEMANA,
            state="readonly", width=22)
        
        # Padrão: Dia atual do sistema
        import datetime as dt_lib
        hoje_idx = dt_lib.datetime.now().weekday()
        if hoje_idx < 6:
            dia_padrao = DIAS_SEMANA[hoje_idx]
        else:
            dia_padrao = self._TODOS_DIAS
            
        win_combo_dia.set(dia_padrao)
        win_combo_dia.pack(side=tk.LEFT, padx=5)

        ttk.Label(filtro_frame, text="Laboratório:").pack(side=tk.LEFT, padx=10)
        win_combo_lab = ttk.Combobox(
            filtro_frame, values=[self._TODOS_LABS] + LABORATORIOS,
            state="readonly", width=22)
        win_combo_lab.set(self._TODOS_LABS)
        win_combo_lab.pack(side=tk.LEFT, padx=5)

        ttk.Label(filtro_frame, text="Turno:").pack(side=tk.LEFT, padx=10)
        win_combo_turno = ttk.Combobox(
            filtro_frame, values=["TODOS", "MATUTINO", "VESPERTINO", "NOTURNO"],
            state="readonly", width=15)
        win_combo_turno.set("TODOS")
        win_combo_turno.pack(side=tk.LEFT, padx=5)

        ttk.Label(filtro_frame,
                  text="   (Colunas = Laboratórios / Linhas = Horários)"
                  ).pack(side=tk.LEFT, padx=10)

        # Botão Atualizar dentro da janela da grade
        # (Definições de funções movidas para baixo após criação de win_inner)

        def _exportar_pdf_win():
            dia = win_combo_dia.get()
            lab = win_combo_lab.get()
            
            # Agora permite exportar do jeito que está na tela
            import datetime as dt_lib
            hoje_dt = dt_lib.datetime.now()
            
            # Se for todos os dias, pedimos pra selecionar um dia
            if dia == self._TODOS_DIAS:
                messagebox.showwarning("Atenção", "Por favor, selecione um dia específico para exportar o PDF.")
                return
            
            # Chama exportar_grade_pdf que já lida com o dia
            # Se lab for 'TODOS', ele gera para todos os labs daquele dia
            self.exportar_grade_pdf(dia, lab_selecionado=lab)

        ttk.Button(filtro_frame, text="📄 PDF", command=_exportar_pdf_win,
                   style="Action.TButton").pack(side=tk.RIGHT, padx=5)

        # --- Canvas principal da janela ---
        canvas_frame = ttk.Frame(win)
        canvas_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))
        canvas_frame.rowconfigure(0, weight=1)
        canvas_frame.columnconfigure(0, weight=1)

        # --- Estrutura de Grade Sincronizada (Sticky Headers) na Janela Separada ---
        grid_container = tk.Frame(canvas_frame, bg=self._D["border"])
        grid_container.grid(row=0, column=0, sticky="nsew")
        grid_container.columnconfigure(1, weight=1)
        grid_container.rowconfigure(1, weight=1)

        # 1. Corner
        win_corner = tk.Frame(grid_container, bg=self._D["hdr"], width=70, height=30)
        win_corner.grid(row=0, column=0, sticky="nsew")
        win_corner.grid_propagate(False)

        # 2. Header Horizontal
        win_canvas_hdr_h = tk.Canvas(grid_container, bg=self._D["hdr"], height=30, highlightthickness=0)
        win_canvas_hdr_h.grid(row=0, column=1, sticky="ew")

        # 3. Coluna de Horas
        win_canvas_hdr_v = tk.Canvas(grid_container, bg=self._D["card"], width=70, highlightthickness=0)
        win_canvas_hdr_v.grid(row=1, column=0, sticky="ns")

        # 4. Área de Dados
        win_canvas_data = tk.Canvas(grid_container, bg="white", highlightthickness=0)
        win_canvas_data.grid(row=1, column=1, sticky="nsew")

        # Sincronização Local para esta janela
        def _win_vscroll(*args):
            setattr(win, "auto_sync_time", False) # Manual
            win_canvas_data.yview(*args)
            win_canvas_hdr_v.yview(*args)
        
        def _win_hscroll(*args):
            win_canvas_data.xview(*args)
            win_canvas_hdr_h.xview(*args)

        v_sb = ttk.Scrollbar(grid_container, orient=tk.VERTICAL, command=_win_vscroll)
        h_sb = ttk.Scrollbar(grid_container, orient=tk.HORIZONTAL, command=_win_hscroll)
        
        win_canvas_data.configure(yscrollcommand=v_sb.set, xscrollcommand=h_sb.set)
        win_canvas_hdr_v.configure(yscrollcommand=v_sb.set)
        # win_canvas_hdr_h.configure(xscrollcommand=h_sb.set) # Removido para evitar sincronismo reverso ambíguo

        v_sb.grid(row=1, column=2, sticky="ns")
        h_sb.grid(row=2, column=1, sticky="ew")

        # Frames Internos
        win_inner_corner = tk.Frame(win_corner, bg=self._D["hdr"])
        win_inner_corner.pack(fill=tk.BOTH, expand=True)

        win_inner_hdr_h = tk.Frame(win_canvas_hdr_h, bg=self._D["hdr"])
        win.win_hdr_h_id = win_canvas_hdr_h.create_window((0, 0), window=win_inner_hdr_h, anchor="nw")

        win_inner_hdr_v = tk.Frame(win_canvas_hdr_v, bg=self._D["card"])
        win.win_hdr_v_id = win_canvas_hdr_v.create_window((0, 0), window=win_inner_hdr_v, anchor="nw")

        win_inner_data = tk.Frame(win_canvas_data, bg=self._D["bg"])
        win.win_data_id = win_canvas_data.create_window((0, 0), window=win_inner_data, anchor="nw")

        # Registra visualização para updates rápidos
        self.active_grade_views.append(win_inner_data)
        
        def _cleanup_view(e):
             if e.widget == win:
                 if win_inner_data in self.active_grade_views:
                     self.active_grade_views.remove(win_inner_data)
        win.bind("<Destroy>", _cleanup_view)

        # Registra referencias globais no window para o Evento global de MouseWheel
        win.win_canvas_data = win_canvas_data
        win.win_canvas_hdr_v = win_canvas_hdr_v
        win.win_canvas_hdr_h = win_canvas_hdr_h

        def _on_win_inner_configure(e):
            cw = win_canvas_data.winfo_width()
            rw = win_inner_data.winfo_reqwidth()
            new_w = max(cw, rw)
            if hasattr(win, "win_data_id"):
                win_canvas_data.itemconfig(win.win_data_id, width=new_w)
            if hasattr(win, "win_hdr_h_id"):
                win_canvas_hdr_h.itemconfig(win.win_hdr_h_id, width=new_w)
            
            # Sincroniza scrollregion após ajuste de largura
            win.update_idletasks()
            win_canvas_data.configure(scrollregion=win_canvas_data.bbox("all"))
            win_canvas_hdr_h.configure(scrollregion=win_canvas_hdr_h.bbox("all"))
            win_canvas_hdr_v.configure(scrollregion=win_canvas_hdr_v.bbox("all"))

        def _on_win_canvas_configure(e):
            """Sincroniza largura do cabeçalho e dados na janela separada."""
            if not win.winfo_exists(): return
            w = e.width
            rw = win_inner_data.winfo_reqwidth()
            new_w = max(w, rw)
            if hasattr(win, "win_data_id"):
                win_canvas_data.itemconfig(win.win_data_id, width=new_w)
            if hasattr(win, "win_hdr_h_id"):
                win_canvas_hdr_h.itemconfig(win.win_hdr_h_id, width=new_w)

        win_inner_data.bind("<Configure>", _on_win_inner_configure)
        win_canvas_data.bind("<Configure>", _on_win_canvas_configure)

        # Flag de controle da janela
        win.is_active = True
        setattr(win, "auto_sync_time", True)

        def _atualizar_win(force_reload=False):
             if not getattr(win, "is_active", False) or not win.winfo_exists():
                 return
             
             if force_reload:
                 self.db.recarregar()
             from datetime import datetime, timedelta
             
             dia_sel = win_combo_dia.get()
             hoje = datetime.now()
             data_win_ref = hoje
             if dia_sel and dia_sel != self._TODOS_DIAS:
                 try:
                     target_idx = DIAS_SEMANA.index(dia_sel)
                     hoje_idx = hoje.weekday()
                     diff = target_idx - hoje_idx
                     data_win_ref = hoje + timedelta(days=diff)
                 except ValueError: pass

             q = win_ent_busca.get().strip()
             if q == search_placeholder: q = ""
             self._draw_grade(win_inner_corner, win_inner_hdr_h, win_inner_hdr_v, win_inner_data, dia_sel, win_combo_lab.get(), data_ref=data_win_ref, turno=win_combo_turno.get(), busca=q)
             if getattr(win, "auto_sync_time", True):
                 self.scroll_to_current_time(win_canvas_data, win_canvas_hdr_v, dia_sel)

        def _auto_refresh_win():
             if getattr(win, "is_active", False) and win.winfo_exists():
                 _atualizar_win()
                 win.after(5000, _auto_refresh_win)

        def _on_close_win():
             win.is_active = False
             win.destroy()

        # Auto-refresh desta janela removido
        # win.after(5000, _auto_refresh_win)

        ttk.Button(filtro_frame, text="↻ ATUALIZAR", command=lambda: _atualizar_win(True),
                   style="Refresh.TButton").pack(side=tk.RIGHT, padx=10)

        # --- Atualiza automaticamente ao mudar filtro ---
        def _on_dia_selected(e):
             setattr(win, "auto_sync_time", True) # Reset ao trocar de dia
             _atualizar_win()
        def _on_lab_selected(e):
            _atualizar_win()

        def _on_turno_selected(e):
             _atualizar_win()

        win_combo_dia.bind("<<ComboboxSelected>>", _on_dia_selected)
        win_combo_lab.bind("<<ComboboxSelected>>", _on_lab_selected)
        win_combo_turno.bind("<<ComboboxSelected>>", _on_turno_selected)

        # Eventos Pesquisa Tempo Real
        def _on_search_focus_in(e):
             if win_ent_busca.get() == search_placeholder:
                 win_ent_busca.delete(0, tk.END)
                 win_ent_busca.config(fg=self._D["fg"])
        def _on_search_focus_out(e):
             if not win_ent_busca.get():
                 win_ent_busca.insert(0, search_placeholder)
                 win_ent_busca.config(fg=self._D["fg2"])

        def _debounce_win_update(e=None):
             if hasattr(win, "_after_id") and win._after_id:
                 win.after_cancel(win._after_id)
             win._after_id = win.after(300, _atualizar_win)

        win_ent_busca.bind("<KeyRelease>", _debounce_win_update)
        win_ent_busca.bind("<FocusIn>", _on_search_focus_in)
        win_ent_busca.bind("<FocusOut>", _on_search_focus_out)

        # Pré-preenche com o mesmo filtro da janela principal (se houver)
        turno_main = self.combo_filtro_turno.get() if self.combo_filtro_turno else "TODOS"
        win_combo_turno.set(turno_main)
        
        dia_main = self.combo_filtro_dia.get()
        lab_main = self.combo_filtro_lab.get()
        if dia_main and dia_main != self._TODOS_DIAS:
            win_combo_dia.set(dia_main)
        if lab_main and lab_main != self._TODOS_LABS:
            win_combo_lab.set(lab_main)
        _atualizar_win()

    def sair(self):
        """Fecha o sistema"""
        resposta = messagebox.askyesno("Sair", "Deseja realmente sair do sistema?")
        if resposta:
            self.db.fechar()
            self.root.quit()

    # ==================================================================
    # Gerar Mapão — Imagem de Programação de Aulas (modelo PUC-SP)
    # ==================================================================

    def chamar_lig_desl(self):
        """Abre a janela de Controle de Laboratórios (lig_desl.py)."""
        import subprocess
        import sys
        try:
            # Pega o diretório do script atual
            script_dir = os.path.dirname(os.path.abspath(__file__))
            script_path = os.path.join(script_dir, "lig_desl.py")
            subprocess.Popen([sys.executable, script_path])
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível abrir o status do laboratório: {e}")

    def _mostrar_menu_status_lab(self, event, lab_nome):
        """Mostra menu de contexto para mudar status do lab (clique direito no cabeçalho)."""
        menu = tk.Menu(self.root, tearoff=0, font=("Segoe UI", 10), bg="#2a2a3e", fg="white", 
                       activebackground="#7c3aed", activeforeground="white")
        
        opcoes = [
            ("🟢 Ligado", "Ligado"),
            ("🔴 Desligado", "Desligado"),
            ("⚪ Finalizado", "Finalizado")
        ]
        
        for rotulo, status in opcoes:
            # Captura lab_nome e status via parâmetros default da lambda
            menu.add_command(label=rotulo, command=lambda ln=lab_nome, s=status: self._mudar_status_lab(ln, s))
            
        menu.tk_popup(event.x_root, event.y_root)

    def _mudar_status_lab(self, lab_nome, novo_status):
        """Atualiza o status do laboratório no BD e atualiza a interface."""
        try:
            self.db.atualizar_status_lab(lab_nome, novo_status)
            self.atualizar_grade()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao atualizar status: {e}")

    def gerar_mapao(self):
        """Abre diálogo para selecionar dia e gera imagens (Matutino/Vespertino/Noturno)."""
        try:
            from PIL import Image
            from tkcalendar import DateEntry
        except ImportError:
            messagebox.showerror("Erro", "Bibliotecas necessárias não encontradas.\nInstale com: pip install pillow tkcalendar")
            return

        dlg = tk.Toplevel(self.root)
        dlg.title("Gerar Mapão")
        dlg.resizable(False, False)
        dlg.configure(bg=self._D["bg"])
        dlg.grab_set()

        # Centraliza
        dlg.update_idletasks()
        w, h = 380, 410
        sw = dlg.winfo_screenwidth()
        sh = dlg.winfo_screenheight()
        dlg.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

        D = self._D
        tk.Label(dlg, text="🗺️  Gerar Mapão",
                 bg=D["bg"], fg=D["fg"],
                 font=("Segoe UI", 14, "bold")).pack(pady=(18, 4))
        tk.Label(dlg, text="Selecione a data para o Mapão:",
                 bg=D["bg"], fg=D["fg2"],
                 font=("Segoe UI", 10)).pack(pady=4)

        # Widget de Calendário (DateEntry)
        cal_data = DateEntry(dlg, width=28, background='darkblue',
                            foreground='white', borderwidth=2,
                            date_pattern='dd/mm/yyyy',
                            locale='pt_BR')
        cal_data.pack(pady=8)

        tk.Label(dlg, text="Buscar aulas de (Dia da Semana):",
                 bg=D["bg"], fg=D["fg2"],
                 font=("Segoe UI", 10)).pack(pady=4)

        combo_dia = ttk.Combobox(dlg, values=DIAS_SEMANA, state="readonly", width=28,
                                  font=("Segoe UI", 10))
        combo_dia.set(DIAS_SEMANA[0])
        combo_dia.pack(pady=4)

        # Sincroniza o dia da semana quando mudar a data no calendário
        def _sync_day(e):
             dt = cal_data.get_date()
             idx = dt.weekday() # 0=Monday
             if idx < 6: # Se for Seg-Sab
                  combo_dia.set(DIAS_SEMANA[idx])
        cal_data.bind("<<DateEntrySelected>>", _sync_day)

        tk.Label(dlg, text="Proporção da Imagem:",
                 bg=D["bg"], fg=D["fg2"],
                 font=("Segoe UI", 10)).pack(pady=2)

        combo_ratio = ttk.Combobox(dlg, values=["16:9 (Widescreen)", "4:3 (Padrão)"], state="readonly", width=28,
                                    font=("Segoe UI", 10))
        combo_ratio.set("16:9 (Widescreen)")
        combo_ratio.pack(pady=4)

        btn_frame_dlg = tk.Frame(dlg, bg=D["bg"])
        btn_frame_dlg.pack(pady=12)

        tk.Button(btn_frame_dlg, text="  📸 Gerar Imagens  ",
                  bg=D["primary"], fg="white",
                  font=("Segoe UI", 10, "bold"),
                  relief="flat", cursor="hand2",
                  command=lambda: _gerar("IMG")).pack(side=tk.LEFT, padx=5)

        tk.Button(btn_frame_dlg, text="  📄 Gerar PDF Completo ",
                  bg=D["success"], fg="white",
                  font=("Segoe UI", 10, "bold"),
                  relief="flat", cursor="hand2",
                  command=lambda: _gerar("PDF")).pack(side=tk.LEFT, padx=5)

        def _gerar(formato):
            dia_selecionado = combo_dia.get()
            try:
                data_dt = cal_data.get_date()
            except Exception:
                messagebox.showwarning("Atenção", "Data inválida.", parent=dlg)
                return
            
            ratio_str = combo_ratio.get()
            ratio = "16:9" if "16:9" in ratio_str else "4:3"
            
            dlg.destroy()
            self._executar_gerar_mapao(dia_selecionado, data_dt, formato, ratio)

        # --- Botão Extra: Disponibilidade MSG ---
        def _gerar_msg_direto():
            dia_selecionado = combo_dia.get()
            try:
                data_dt = cal_data.get_date()
            except Exception: return
            ratio_str = combo_ratio.get()
            ratio = "16:9" if "16:9" in ratio_str else "4:3"
            dia_label = f"{dia_selecionado.upper()} - {data_dt.strftime('%d/%m')}"
            
            ret = self._pedir_mensagem_disponibilidade()
            if not ret or not ret["msg"]: return
            
            dlg.destroy()
            self._gerar_mapao_mensagem_especial(dia_label, ret["msg"], "IMG", ratio, font_size_override=ret["tamanho"])

        tk.Button(dlg, text="  💬  Disponibilidade MSG  ",
                  bg="#8b5cf6", fg="white",
                  font=("Segoe UI", 10, "bold"),
                  relief="flat", cursor="hand2",
                  command=_gerar_msg_direto).pack(pady=(0, 12))

    def _executar_gerar_mapao(self, dia_semana: str, data_alvo, formato: str = "IMG", ratio: str = "16:9"):
        """Gera mapão baseado no dia da semana e data alvo informada."""
        import os
        from datetime import datetime, timedelta
        from PIL import Image, ImageDraw, ImageFont

        dia_label = f"{dia_semana.upper()} - {data_alvo.strftime('%d/%m')}"

        # Caminho da template
        base_dir = os.path.dirname(os.path.abspath(__file__))
        template_path = os.path.join(base_dir, "ppt", "ppt_horario_vazio.png")

        if not os.path.isfile(template_path):
            messagebox.showerror("Erro", f"Template não encontrado:\n{template_path}")
            return

        # Buscar todas as aulas do dia da semana correspondente
        self.db.recarregar()
        # Aplicar lógica de sobreposição de aulas eventuais
        # O data_alvo já vem do diálogo de geração do mapão
        aulas_dia = self._get_aulas_com_sobrescricao(dia_semana, data_alvo)
        
        if not aulas_dia:
            ret = self._pedir_mensagem_disponibilidade()
            if not ret or not ret["msg"]: return
            self._gerar_mapao_mensagem_especial(dia_label, ret["msg"], formato, ratio, font_size_override=ret["tamanho"])
            return

        # Definição dos períodos
        periodos = [
            ("Matutino",   "07:30", "12:00"),
            ("Vespertino", "12:00", "18:00"),
            ("Noturno",    "18:00", "23:15"),
        ]

        # Tentar carregar fonte do sistema
        def _fonte(tamanho, negrito=False):
            candidatas = []
            if negrito:
                candidatas = [
                    r"C:\Windows\Fonts\arialbd.ttf",
                    r"C:\Windows\Fonts\calibrib.ttf",
                    r"C:\Windows\Fonts\segoeui.ttf",
                ]
            else:
                candidatas = [
                    r"C:\Windows\Fonts\arial.ttf",
                    r"C:\Windows\Fonts\calibri.ttf",
                    r"C:\Windows\Fonts\segoeui.ttf",
                ]
            for p in candidatas:
                try:
                    return ImageFont.truetype(p, tamanho)
                except Exception:
                    pass
            return ImageFont.load_default()

        # Imagens geradas
        imagens_geradas = []

        # Cores (Paleta PUC-SP)
        COR_AMARELO     = (255, 220, 0, 255)
        COR_BRANCO      = (255, 255, 255, 255)
        COR_LINHA_PAR   = (25, 50, 100, 255)
        COR_LINHA_IMPAR = (18, 35, 75, 255)
        COR_BORDA       = (60, 100, 160, 255)

        for periodo_nome, h_ini_str, h_fim_str in periodos:
            h_ini_lim = datetime.strptime(h_ini_str, "%H:%M")
            h_fim_lim = datetime.strptime(h_fim_str, "%H:%M")

            aulas_periodo = []
            for a in aulas_dia:
                try:
                    h = datetime.strptime(a.hora_inicio, "%H:%M")
                    if h_ini_lim <= h < h_fim_lim:
                        aulas_periodo.append(a)
                except Exception: pass

            aulas_periodo.sort(key=lambda a: (int(a.laboratorio) if a.laboratorio.isdigit() else 99, a.hora_inicio))

            template_img = Image.open(template_path).convert("RGBA")
            
            # Resolução baseada no Ratio
            if ratio == "4:3":
                W, H = 1440, 1080
            else:
                W, H = 1920, 1080
            
            img = template_img.resize((W, H), Image.Resampling.LANCZOS)
            draw = ImageDraw.Draw(img)

            # ---- Configurações de Layout Dinâmico ----
            TAB_X     = 30
            TAB_Y     = 180
            TAB_W     = W - 60
            AREA_UTIL_H = H - TAB_Y - 40
            HEADER_H  = 54
            
            # Valores padrão
            row_h_default = 45
            font_size_default = 20
            font_cabec_size = 22
            
            num_aulas = len(aulas_periodo)
            if num_aulas > 0:
                h_necessario = HEADER_H + (num_aulas * row_h_default)
                if h_necessario > AREA_UTIL_H:
                    # Ajuste dinâmico para caber tudo
                    row_h = (AREA_UTIL_H - HEADER_H) // num_aulas
                    # Garante um mínimo para não ficar ilegível
                    row_h = max(22, row_h)
                    escala = row_h / row_h_default
                    font_size = max(10, int(font_size_default * escala))
                    # Ajusta cabeçalho levemente também
                    font_cabec_size = max(14, int(22 * escala))
                else:
                    row_h = row_h_default
                    font_size = font_size_default
            else:
                row_h = row_h_default
                font_size = font_size_default

            # Ajusta fontes e posições se for 4:3 para não encobrir "Programação de Aula"
            f_titulo_base = 42 if ratio == "4:3" else 60
            f_per_base = 38 if ratio == "4:3" else 55
            y_base = 20 if ratio == "4:3" else 25

            fonte_titulo_dia   = _fonte(f_titulo_base, negrito=True)
            fonte_titulo_per   = _fonte(f_per_base, negrito=True)
            fonte_cabec        = _fonte(font_cabec_size, negrito=True)
            fonte_dado         = _fonte(font_size)
            fonte_dado_neg     = _fonte(font_size, negrito=True)

            margin_right = 40
            # y_base já está definido acima
            try:
                bb_dia = draw.textbbox((0, 0), dia_label, font=fonte_titulo_dia)
                tw_dia = bb_dia[2] - bb_dia[0]
                bb_per = draw.textbbox((0, 0), periodo_nome, font=fonte_titulo_per)
                tw_per = bb_per[2] - bb_per[0]
            except Exception: tw_dia, tw_per = 600, 400

            draw.text((W - tw_dia - margin_right, y_base), dia_label, fill=COR_AMARELO, font=fonte_titulo_dia)
            draw.text((W - tw_per - margin_right, y_base + (60 if ratio == "4:3" else 70)), periodo_nome, fill=COR_AMARELO, font=fonte_titulo_per)

            col_props  = [0.06, 0.08, 0.08, 0.40, 0.15, 0.23]
            col_names = ["LAB", "INÍCIO", "FIM", "DISCIPLINA", "TURMA", "PROFESSOR"]
            col_xs, col_widths = [], []
            acc = TAB_X
            for p in col_props:
                col_xs.append(acc)
                cw = int(TAB_W * p)
                col_widths.append(cw)
                acc += cw

            draw.rectangle([TAB_X, TAB_Y, TAB_X + TAB_W, TAB_Y + HEADER_H], fill=COR_LINHA_IMPAR)
            for i, (col_x, col_w, nome) in enumerate(zip(col_xs, col_widths, col_names)):
                nome_upper = nome.upper()
                try:
                    bb_h = draw.textbbox((0, 0), nome_upper, font=fonte_cabec)
                    tw_h, th_h = bb_h[2]-bb_h[0], bb_h[3]-bb_h[1]
                except: tw_h, th_h = len(nome_upper)*10, 15
                draw.text((col_x + (col_w - tw_h)//2, TAB_Y + (HEADER_H - th_h)//2), nome_upper, fill=COR_AMARELO, font=fonte_cabec)

            y_cur = TAB_Y + HEADER_H
            for idx, aula in enumerate(aulas_periodo):
                # Se mesmo com escala ultra-pequena ainda sair do limite, paramos de desenhar (medida de segurança)
                if y_cur + row_h > H - 10: break

                cor_bg = COR_LINHA_PAR if idx % 2 == 0 else COR_LINHA_IMPAR
                draw.rectangle([TAB_X, y_cur, TAB_X + TAB_W, y_cur + row_h], fill=cor_bg)
                
                valores = [aula.laboratorio, aula.hora_inicio, aula.hora_fim, aula.disciplina, aula.turma, aula.professor]
                for i, (col_x, col_w, val) in enumerate(zip(col_xs, col_widths, valores)):
                    texto = str(val) if val else "-"
                    f_u = fonte_dado_neg if i == 0 else fonte_dado
                    while True:
                        try:
                            bb_t = draw.textbbox((0, 0), texto, font=f_u)
                            tw_t, th_t = bb_t[2]-bb_t[0], bb_t[3]-bb_t[1]
                        except: tw_t, th_t = len(texto)*8, 15; break
                        if tw_t <= col_w - 10 or len(texto) <= 1: break
                        texto = texto[:-1]
                    if val and texto != str(val): texto = texto[:-3] + "..."
                    
                    # Centraliza TODOS agora
                    tx = col_x + (col_w - tw_t) // 2
                    ty = y_cur + (row_h - th_t) // 2 - 2
                    draw.text((tx, ty), texto, fill=COR_BRANCO, font=f_u)
                y_cur += row_h

            draw.rectangle([TAB_X, TAB_Y, TAB_X + TAB_W, y_cur], outline=COR_BORDA, width=2)
            imagens_geradas.append((periodo_nome, img.convert("RGB")))

        if not imagens_geradas:
            messagebox.showinfo("Mapão", "Nenhuma aula encontrada para este dia.")
            return

        if formato == "PDF":
             self._salvar_mapao_pdf(dia_label, imagens_geradas)
        else:
             self._mostrar_janelas_mapao(dia_label, imagens_geradas)

    def _pedir_mensagem_disponibilidade(self):
        """Abre uma pequena janela para o usuário digitar a mensagem de disponibilidade e escolher o tamanho."""
        res = {"msg": None, "tamanho": 150}
        dlg = tk.Toplevel(self.root)
        dlg.title("Mensagem: Disponibilidade de Laboratórios")
        dlg.resizable(False, False)
        dlg.configure(bg=self._D["bg"])
        dlg.grab_set()

        # Centraliza
        dlg.update_idletasks()
        w, h = 540, 240
        sw = dlg.winfo_screenwidth()
        sh = dlg.winfo_screenheight()
        dlg.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

        D = self._D
        tk.Label(dlg, text="📝 Mensagem de Disponibilidade",
                 bg=D["bg"], fg=D["fg"],
                 font=("Segoe UI", 12, "bold")).pack(pady=(20, 10))

        tk.Label(dlg, text="Digite a mensagem e selecione o tamanho da fonte:",
                 bg=D["bg"], fg=D["fg2"],
                 font=("Segoe UI", 9)).pack()

        # Frame para entrada e tamanho lado a lado
        entry_row = tk.Frame(dlg, bg=D["bg"])
        entry_row.pack(pady=15, padx=20)

        entry = tk.Entry(entry_row, font=("Segoe UI", 12), width=35,
                         bg=D["input"], fg=D["fg"], insertbackground=D["fg"],
                         relief="flat", highlightthickness=1, highlightbackground=D["border"])
        entry.insert(0, "Laboratório 05 Disponível até 13h")
        entry.pack(side=tk.LEFT, ipady=5, padx=(0, 10))
        entry.focus_set()
        entry.selection_range(0, tk.END)

        tk.Label(entry_row, text="Tamanho:", bg=D["bg"], fg=D["fg2"], font=("Segoe UI", 8)).pack(side=tk.LEFT)
        # Combobox editável para seleção ou digitação
        combo_tam = ttk.Combobox(entry_row, values=["50", "80", "100", "120", "150", "180", "200", "250", "300"],
                                 width=5, font=("Segoe UI", 11))
        combo_tam.set("150")
        combo_tam.pack(side=tk.LEFT, padx=5)

        def _ok():
            res["msg"] = entry.get().strip()
            try:
                res["tamanho"] = int(combo_tam.get())
            except ValueError:
                res["tamanho"] = 150
            dlg.destroy()

        def _cancel():
            dlg.destroy()

        btn_frame = tk.Frame(dlg, bg=D["bg"])
        btn_frame.pack(pady=10)

        tk.Button(btn_frame, text="  CANCELAR  ", 
                  bg="#6c757d", fg="white", font=("Segoe UI", 10, "bold"),
                  relief="flat", cursor="hand2",
                  command=_cancel).pack(side=tk.LEFT, padx=10)

        tk.Button(btn_frame, text="      OK      ", 
                  bg=D["primary"], fg="white", font=("Segoe UI", 10, "bold"),
                  relief="flat", cursor="hand2",
                  command=_ok).pack(side=tk.LEFT, padx=10)

        # Binds
        dlg.bind("<Return>", lambda e: _ok())
        dlg.bind("<Escape>", lambda e: _cancel())

        self.root.wait_window(dlg)
        return res

    def _gerar_mapao_mensagem_especial(self, dia_label, mensagem, formato, ratio, font_size_override=None):
        """Gera imagem do mapão com mensagem centralizada grande quando não há aulas."""
        import os
        from PIL import Image, ImageDraw, ImageFont
        from datetime import datetime

        # Caminho da template
        base_dir = os.path.dirname(os.path.abspath(__file__))
        template_path = os.path.join(base_dir, "ppt", "ppt_horario_vazio.png")

        if not os.path.isfile(template_path):
            messagebox.showerror("Erro", f"Template não encontrado:\n{template_path}")
            return

        # Cores (Paleta PUC-SP)
        COR_AMARELO     = (255, 220, 0, 255)
        COR_BRANCO      = (255, 255, 255, 255)

        template_img = Image.open(template_path).convert("RGBA")
        if ratio == "4:3":
            W, H = 1440, 1080
        else:
            W, H = 1920, 1080

        img = template_img.resize((W, H), Image.Resampling.LANCZOS)
        draw = ImageDraw.Draw(img)

        # Helper para fontes (adaptado da _executar_gerar_mapao)
        def _get_f(tamanho, negrito=False):
            if negrito:
                candidatas = [
                    r"C:\Windows\Fonts\arialbd.ttf",
                    r"C:\Windows\Fonts\calibrib.ttf",
                    r"C:\Windows\Fonts\segoeui.ttf",
                ]
            else:
                candidatas = [
                    r"C:\Windows\Fonts\arial.ttf",
                    r"C:\Windows\Fonts\calibri.ttf",
                    r"C:\Windows\Fonts\segoeui.ttf",
                ]
            for p in candidatas:
                try: return ImageFont.truetype(p, tamanho)
                except: pass
            return ImageFont.load_default()

        # Desenha Título (Dia)
        f_titulo_base = 42 if ratio == "4:3" else 60
        fonte_titulo_dia = _get_f(f_titulo_base, negrito=True)
        margin_right = 40
        y_base = 20 if ratio == "4:3" else 25

        try:
            bb_dia = draw.textbbox((0, 0), dia_label, font=fonte_titulo_dia)
            tw_dia = bb_dia[2] - bb_dia[0]
        except: tw_dia = 600
        draw.text((W - tw_dia - margin_right, y_base), dia_label, fill=COR_AMARELO, font=fonte_titulo_dia)

        # Mensagem Principal
        f_init_size = font_size_override if font_size_override else (180 if ratio == "4:3" else 220)
        fonte_msg = _get_f(f_init_size, negrito=True)
        
        # Área útil maior (diminuindo as laterais)
        TAB_Y = 180
        AREA_UTIL_H = H - TAB_Y - 40
        max_w = W - 100 # Margem menor (50 de cada lado)
        max_h = AREA_UTIL_H - 10

        # Helper para quebrar o texto por largura em pixels
        def _wrap(text, font, limit_w):
            words = text.split()
            lines = []
            cur_l = []
            for w in words:
                test = " ".join(cur_l + [w])
                try:
                    bb = draw.textbbox((0, 0), test, font=font)
                    tw_test = bb[2] - bb[0]
                except: tw_test = len(test) * (f_init_size * 0.6)
                if tw_test <= limit_w:
                    cur_l.append(w)
                else:
                    if cur_l: lines.append(" ".join(cur_l))
                    cur_l = [w]
            if cur_l: lines.append(" ".join(cur_l))
            return "\n".join(lines)

        # Ajuste dinâmico de tamanho E quebra de linha
        wrapped_text = _wrap(mensagem, fonte_msg, max_w)
        while True:
            try:
                bb_m = draw.multiline_textbbox((0, 0), wrapped_text, font=fonte_msg, align="center")
                tw_m, th_m = bb_m[2]-bb_m[0], bb_m[3]-bb_m[1]
            except:
                # Fallback se multiline_textbbox falhar (versões PIL antigas)
                tw_m, th_m = max_w, 200; break
            
            if (tw_m <= max_w and th_m <= max_h) or f_init_size < 30:
                break
            f_init_size -= 5
            fonte_msg = _get_f(f_init_size, negrito=True)
            wrapped_text = _wrap(mensagem, fonte_msg, max_w)

        mx = (W - tw_m) // 2
        my = TAB_Y + (AREA_UTIL_H - th_m) // 2
        
        # Desenha texto Multiline Centralizado em AMARELO
        draw.multiline_text((mx, my), wrapped_text, fill=COR_AMARELO, font=fonte_msg, align="center")

        imagens_geradas = [("Disponibilidade", img.convert("RGB"))]

        if formato == "PDF":
             self._salvar_mapao_pdf(dia_label, imagens_geradas)
        else:
             self._mostrar_janelas_mapao(dia_label, imagens_geradas)

    def _salvar_mapao_pdf(self, dia: str, imagens: list):
        """Salva as imagens geradas em um arquivo PDF."""
        try:
            from fpdf import FPDF
        except ImportError:
            messagebox.showerror("Erro", "A biblioteca fpdf2 não está instalada.")
            return

        filename = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("Documento PDF", "*.pdf")],
            initialfile=f"mapao_{dia.lower().replace(' ', '_').replace('-', '_').replace('/', '_')}.pdf",
            title="Salvar Mapão como PDF"
        )
        if not filename:
            return

        pdf = FPDF(orientation='L', unit='mm', format='A4')
        for periodo_nome, img_pil in imagens:
            pdf.add_page()
            import io
            buf = io.BytesIO()
            img_pil.save(buf, format='PNG')
            buf.seek(0)
            # A4 Landscape (297x210). 16:9 -> 297 x 167mm. Centraliza verticalmente.
            pdf.image(buf, x=0, y=(210-167)//2, w=297)

        try:
            pdf.output(filename)
            messagebox.showinfo("Sucesso", f"PDF salvo com sucesso em:\n{filename}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar PDF: {e}")

    def _mostrar_janelas_mapao(self, dia_label: str, imagens: list):
        """Abre uma janela moderna com zoom funcional (Ctrl+Scroll) para o mapão."""
        import os
        from PIL import ImageTk, Image

        D = self._D
        win = tk.Toplevel(self.root)
        win.title(f"Mapão — {dia_label}")
        win.state("zoomed")
        win.configure(bg=D["bg"])
        win.grab_set()

        header = tk.Frame(win, bg=D["hdr"], height=60)
        header.pack(fill=tk.X)
        header.pack_propagate(False)
        tk.Label(header, text=f"🗺️  VISUALIZAÇÃO DO MAPÃO — {dia_label}", font=("Segoe UI", 12, "bold"),
                 bg=D["hdr"], fg=D["fg"]).pack(side=tk.LEFT, padx=20, pady=12)

        nb = ttk.Notebook(win)
        nb.pack(fill=tk.BOTH, expand=True, padx=10, pady=(10, 0))
        
        # Refs para não garbage collect
        win._img_refs = []
        win._zoom_states = []

        for p_nome, img_pil in imagens:
            tab = tk.Frame(nb, bg=D["bg"])
            nb.add(tab, text=f"  {p_nome}  ")

            # Toolbar do zoom
            zoom_bar = tk.Frame(tab, bg=D["hdr"], height=35)
            zoom_bar.pack(fill=tk.X)

            lbl_zoom = tk.Label(zoom_bar, text="Zoom: 100%", bg=D["hdr"], fg=D["fg2"], font=("Segoe UI", 9))
            lbl_zoom.pack(side=tk.RIGHT, padx=20)

            c_frame = tk.Frame(tab, bg=D["bg"])
            c_frame.pack(fill=tk.BOTH, expand=True)

            canvas = tk.Canvas(c_frame, bg=D["bg"], highlightthickness=0)
            sb_v = ttk.Scrollbar(c_frame, orient=tk.VERTICAL,   command=canvas.yview)
            sb_h = ttk.Scrollbar(c_frame, orient=tk.HORIZONTAL, command=canvas.xview)
            canvas.configure(yscrollcommand=sb_v.set, xscrollcommand=sb_h.set)
            
            canvas.grid(row=0, column=0, sticky="nsew")
            sb_v.grid(row=0, column=1, sticky="ns")
            sb_h.grid(row=1, column=0, sticky="ew")
            c_frame.rowconfigure(0, weight=1)
            c_frame.columnconfigure(0, weight=1)

            state = {"zoom": 1.0, "img_orig": img_pil, "canvas": canvas, "lbl": lbl_zoom, "id": None}
            win._zoom_states.append(state)

            def _render_zoom(st):
                w_z = int(st["img_orig"].width * st["zoom"])
                h_z = int(st["img_orig"].height * st["zoom"])
                img_z = st["img_orig"].resize((w_z, h_z), Image.Resampling.LANCZOS)
                tk_z = ImageTk.PhotoImage(img_z)
                win._img_refs.append(tk_z) # Mantém ref
                
                st["canvas"].delete("all")
                st["id"] = st["canvas"].create_image(0, 0, anchor="nw", image=tk_z)
                st["canvas"].configure(scrollregion=(0, 0, w_z, h_z))
                st["lbl"].config(text=f"Zoom: {int(st['zoom']*100)}%")

            _render_zoom(state)

            def _handle_zoom(event, st=state):
                if event.state & 0x0004: # CTRL pressed
                    if event.delta > 0: st["zoom"] = min(3.0, st["zoom"] + 0.1)
                    else: st["zoom"] = max(0.2, st["zoom"] - 0.1)
                    _render_zoom(st)
                    return "break"
                
            canvas.bind("<MouseWheel>", _handle_zoom)

        footer = tk.Frame(win, bg=D["hdr"], height=60)
        footer.pack(fill=tk.X, side=tk.BOTTOM)
        
        def salvar_atual():
            idx = nb.index(nb.select())
            p_nome, i_pil = imagens[idx]
            cam = filedialog.asksaveasfilename(
                defaultextension=".png",
                initialfile=f"mapao_{dia_label.lower().replace(' ','_').replace('-', '_').replace('/', '_')}_{p_nome.lower()}.png",
                title="Salvar Imagem Atual"
            )
            if cam: i_pil.save(cam); messagebox.showinfo("Sucesso", "Imagem salva!", parent=win)

        self._btn_footer(footer, "💾  SALVAR IMAGEM", D["success"], salvar_atual)
        self._btn_footer(footer, "📸  TODOS TURNOS",  D["primary"], lambda: self._salvar_mapao_img_todas(dia_label, imagens))
        self._btn_footer(footer, "📄  PDF COMPLETO",  D["primary"], lambda: self._salvar_mapao_pdf(dia_label, imagens))
        self._btn_footer(footer, "✕  FECHAR",       D["danger"],  win.destroy)

    def _salvar_mapao_img_todas(self, dia_label, imagens):
        pasta = filedialog.askdirectory(title="Selecionar Pasta para Salvar Todas")
        if not pasta: return
        import os
        for p_nome, i_pil in imagens:
            i_pil.save(os.path.join(pasta, f"mapao_{dia_label.lower().replace(' ', '_').replace('-', '_').replace('/', '_')}_{p_nome.lower()}.png"))
        messagebox.showinfo("Sucesso", f"Todas as {len(imagens)} imagens foram salvas!")

    # ==================================================================
    # Janelas de Cadastro de Dados
    # ==================================================================

    def dados_de_cadastro(self):
        """Exibe menu de opções de cadastro."""
        pass  # disparado via toolbar menu

    # ------------------------------------------------------------------
    def _criar_janela_cadastro(self, titulo: str, icone: str = "📂",
                                largura: int = 440, altura: int = 340,
                                scrollable: bool = False, resizable: bool = False):
        """Cria e retorna um Toplevel dark com identidade visual padrão."""
        D = self._D
        win = tk.Toplevel(self.root)
        win.title(titulo)
        win.geometry(f"{largura}x{altura}")
        win.resizable(resizable, resizable)
        win.grab_set()  # modal
        win.configure(bg=D["bg"])

        # Cabeçalho
        hdr = tk.Frame(win, bg=D["hdr"], height=54)
        hdr.pack(fill=tk.X)
        hdr.pack_propagate(False)
        tk.Label(hdr, text=f"{icone}  {titulo}",
                 bg=D["hdr"], fg=D["fg"],
                 font=("Segoe UI", 13, "bold")
                 ).pack(side=tk.LEFT, padx=18, pady=12)

        # Rodapé (criado antes para ficar no fundo)
        footer = tk.Frame(win, bg=D["footer"], height=56)
        footer.pack(fill=tk.X, side=tk.BOTTOM)
        footer.pack_propagate(False)

        # Separador antes do rodapé
        sep = tk.Frame(win, bg=D["border"], height=1)
        sep.pack(fill=tk.X, side=tk.BOTTOM)

        # Área de conteúdo
        if scrollable:
            container = tk.Frame(win, bg=D["bg"])
            container.pack(fill=tk.BOTH, expand=True)
            
            canvas = tk.Canvas(container, bg=D["bg"], highlightthickness=0)
            vsb = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
            canvas.configure(yscrollcommand=vsb.set)
            
            vsb.pack(side=tk.RIGHT, fill=tk.Y)
            canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            
            body = tk.Frame(canvas, bg=D["bg"], padx=24, pady=16)
            window_id = canvas.create_window((0, 0), window=body, anchor="nw")
            
            def _on_body_cfg(e):
                # Processa mudanças pendentes antes de ler bbox
                canvas.update_idletasks()
                canvas.configure(scrollregion=canvas.bbox("all"))
                # Ajusta largura do frame para preencher o canvas (resolvendo glitch visual)
                canvas.itemconfig(window_id, width=canvas.winfo_width())

            body.bind("<Configure>", _on_body_cfg)
            
            # O suporte a mousewheel agora é tratado globalmente pelo `_on_mousewheel` principal
        else:
            body = tk.Frame(win, bg=D["bg"], padx=24, pady=16)
            body.pack(fill=tk.BOTH, expand=True)

        return win, body, footer

    def _btn_footer(self, footer, texto, bg, cmd):
        """Cria botão estilizado no rodapé da janela."""
        D = self._D
        # Hover effect
        def on_enter(e): btn.config(bg=D["primary2"] if bg == D["primary"] else bg)
        def on_leave(e): btn.config(bg=bg)
        btn = tk.Button(
            footer, text=texto, command=cmd,
            bg=bg, fg="white", activebackground=bg,
            activeforeground="white",
            font=("Segoe UI", 10, "bold"),
            relief="flat", bd=0, padx=22, pady=10,
            cursor="hand2"
        )
        btn.pack(side=tk.RIGHT, padx=10, pady=10)
        btn.bind("<Enter>", on_enter)
        btn.bind("<Leave>", on_leave)

    def _lbl(self, parent, texto):
        """Label padrão dark para campos de formulário."""
        D = self._D
        tk.Label(parent, text=texto, bg=D["bg"],
                 fg=D["fg2"], font=("Segoe UI", 9, "bold"),
                 anchor="w").pack(fill=tk.X, pady=(10, 3))

    def _entry(self, parent):
        """Entry dark estilizado."""
        D = self._D
        e = tk.Entry(parent, font=("Segoe UI", 11),
                     relief="flat", bd=0,
                     bg=D["input"], fg=D["fg"],
                     insertbackground=D["fg"],
                     highlightthickness=1,
                     highlightbackground=D["border"],
                     highlightcolor=D["primary"])
        e.pack(fill=tk.X, ipady=8)
        return e

    def _combo(self, parent, valores):
        """Combobox dark estilizado."""
        c = ttk.Combobox(parent, values=valores, state="readonly",
                         font=("Segoe UI", 11))
        c.pack(fill=tk.X, ipady=5)
        return c

    # ------------------------------------------------------------------
    def _janela_nova_faculdade(self):
        """Janela para cadastrar nova faculdade."""
        win, body, footer = self._criar_janela_cadastro("Nova Faculdade", "🏛️", 420, 260)

        self._lbl(body, "Nome da Faculdade:")
        entry_nome = self._entry(body)
        entry_nome.focus()
        entry_nome.bind("<FocusOut>", lambda e: self.converter_maiuscula(entry_nome))

        def salvar():
            nome = entry_nome.get().strip().upper()
            if not nome:
                messagebox.showwarning("Atenção", "Informe o nome da faculdade.", parent=win)
                return
            self.db.adicionar_faculdade(nome)
            self.atualizar_dados()
            messagebox.showinfo("Sucesso", f"Faculdade '{nome}' cadastrada!", parent=win)
            win.destroy()

        self._btn_footer(footer, "💾  SALVAR",  "#28a745", salvar)
        self._btn_footer(footer, "✕  FECHAR", "#6c757d", win.destroy)

    # ------------------------------------------------------------------
    def _janela_novo_curso(self):
        """Janela para cadastrar novo curso (com seleção de cor)."""
        win, body, footer = self._criar_janela_cadastro("Novo Curso", "🎨", 480, 380)

        self._lbl(body, "Faculdade:")
        combo_fac = self._combo(body, self.db.listar_faculdades())
        if combo_fac["values"]:
            combo_fac.current(0)

        self._lbl(body, "Nome do Curso:")
        entry_nome = self._entry(body)
        entry_nome.bind("<FocusOut>", lambda e: self.converter_maiuscula(entry_nome))

        # Seletor de cor
        D = self._D
        cor_var = tk.StringVar(value="#4e9af1")
        cor_frame = tk.Frame(body, bg=D["bg"])
        cor_frame.pack(fill=tk.X, pady=(10, 0))
        tk.Label(cor_frame, text="Cor do Curso:", bg=D["bg"],
                 fg=D["fg2"], font=("Segoe UI", 10, "bold")
                 ).pack(side=tk.LEFT)
        preview = tk.Label(cor_frame, bg=cor_var.get(), width=8,
                           relief="solid", bd=1)
        preview.pack(side=tk.LEFT, padx=10)
        cor_label = tk.Label(cor_frame, textvariable=cor_var,
                             bg=D["bg"], fg=D["fg2"], font=("Segoe UI", 9))
        cor_label.pack(side=tk.LEFT)

        def escolher_cor():
            resultado = colorchooser.askcolor(
                color=cor_var.get(), title="Escolha a cor do curso", parent=win
            )
            if resultado and resultado[1]:
                cor_var.set(resultado[1])
                preview.configure(bg=resultado[1])

        tk.Button(cor_frame, text="Escolher cor",
                  command=escolher_cor,
                  bg="#7c3aed", fg="white",
                  font=("Segoe UI", 9), relief="flat", padx=10, pady=4,
                  cursor="hand2"
                  ).pack(side=tk.LEFT, padx=10)

        def salvar():
            faculdade = combo_fac.get()
            nome = entry_nome.get().strip().upper()
            if not faculdade:
                messagebox.showwarning("Atenção", "Selecione a faculdade.", parent=win)
                return
            if not nome:
                messagebox.showwarning("Atenção", "Informe o nome do curso.", parent=win)
                return
            self.db.adicionar_curso(faculdade, nome, cor_var.get())
            self.atualizar_dados()
            messagebox.showinfo("Sucesso", f"Curso '{nome}' cadastrado!", parent=win)
            win.destroy()

        self._btn_footer(footer, "💾  SALVAR",  "#28a745", salvar)
        self._btn_footer(footer, "✕  FECHAR", "#6c757d", win.destroy)

    # ------------------------------------------------------------------
    def _janela_nova_disciplina(self):
        """Janela para cadastrar nova disciplina vinculada a um curso."""
        win, body, footer = self._criar_janela_cadastro("Nova Disciplina", "📚", 420, 360)

        self._lbl(body, "Curso Relacionado:")
        # Listagem de todos os cursos para vincular a disciplina
        todos_cursos = []
        for fac in self.db.listar_faculdades():
            cursos = self.db.listar_cursos(fac)
            todos_cursos.extend(cursos)
        
        combo_curso = self._combo(body, sorted(list(set(todos_cursos))))

        self._lbl(body, "Nome da Disciplina:")
        entry_nome = self._entry(body)
        entry_nome.focus()
        entry_nome.bind("<FocusOut>", lambda e: self.converter_maiuscula(entry_nome))

        def salvar():
            curso = combo_curso.get()
            nome = entry_nome.get().strip().upper()
            if not curso:
                messagebox.showwarning("Atenção", "Selecione o curso.", parent=win)
                return
            if not nome:
                messagebox.showwarning("Atenção", "Informe o nome da disciplina.", parent=win)
                return
            self.db.adicionar_disciplina(curso, nome)
            self.atualizar_dados()
            messagebox.showinfo("Sucesso", f"Disciplina '{nome}' cadastrada!", parent=win)
            win.destroy()

        self._btn_footer(footer, "💾  SALVAR",  "#28a745", salvar)
        self._btn_footer(footer, "✕  FECHAR", "#6c757d", win.destroy)

    # ------------------------------------------------------------------
    def _janela_nova_turma(self):
        """Janela para cadastrar nova turma vinculada a uma disciplina."""
        win, body, footer = self._criar_janela_cadastro("Nova Turma", "🎓", 420, 360)

        self._lbl(body, "Disciplina Relacionada:")
        combo_disc = self._combo(body, self.db.listar_todas_disciplinas())

        self._lbl(body, "Nome / Código da Turma:")
        entry_nome = self._entry(body)
        entry_nome.focus()
        entry_nome.bind("<FocusOut>", lambda e: self.converter_maiuscula(entry_nome))

        self._lbl(body, "qtd Alunos:")
        entry_qtde = self._entry(body)
        entry_qtde.insert(0, "0")
        entry_qtde.bind("<KeyRelease>", self._mask_int)

        def salvar():
            disciplina = combo_disc.get()
            nome = entry_nome.get().strip().upper()
            
            # Reset colors
            combo_disc.configure(style="TCombobox")
            entry_nome.configure(highlightbackground="#ccc")
            
            erros = []
            if not disciplina:
                combo_disc.configure(style="Invalid.TCombobox")
                erros.append("- Disciplina (selecione uma)")
            if not nome:
                entry_nome.configure(highlightbackground="#ef4444")
                erros.append("- Nome da Turma (Ex: TURMA A)")

            if erros:
                msg = "Os seguintes campos precisam de atenção:\n\n" + "\n".join(erros)
                messagebox.showwarning("Campos Obrigatórios", msg, parent=win)
                return

            try:
                qtde = int(entry_qtde.get().strip())
            except ValueError:
                qtde = 0

            self.db.adicionar_turma(disciplina, nome, qtde_alunos=qtde)
            self.atualizar_dados()
            messagebox.showinfo("Sucesso", f"Turma '{nome}' cadastrada!", parent=win)
            win.destroy()

        self._btn_footer(footer, "💾  SALVAR",  "#28a745", salvar)
        self._btn_footer(footer, "✕  FECHAR", "#6c757d", win.destroy)

    def _janela_cadastro_aula(self, aula: Optional[Aula] = None, pre_dia=None, pre_lab=None):
        """Janela modal para cadastro/edição de aulas."""
        titulo = "Editar Aula" if aula else "Nova Aula"
        win, body, footer = self._criar_janela_cadastro(titulo, "📅", 500, 780)
        
        # Grid layout for fields
        D = self._D
        f = tk.Frame(body, bg=D["bg"])
        f.pack(fill=tk.BOTH, expand=True)

        is_eventual_var = tk.BooleanVar(value=aula.is_eventual if aula else False)
        
        # Definição dos campos de data (serão populados depois, mas nomes precisam existir para o toggle)
        lbl_data = tk.Label(f, text="Data exata (Eventual):", bg=D["bg"], fg=D["fg2"], font=("Segoe UI", 9, "bold"))
        row_data = tk.Frame(f, bg=D["bg"])

        def draw_switch(active):
            switch_canvas.delete("all")
            # Pill shape background
            if active:
                switch_canvas.create_oval(2, 2, 22, 22, fill=D["primary"], outline="")
                switch_canvas.create_oval(28, 2, 48, 22, fill=D["primary"], outline="")
                switch_canvas.create_rectangle(12, 2, 38, 22, fill=D["primary"], outline="")
                # White Circle
                switch_canvas.create_oval(26, 4, 46, 20, fill="white", outline="")
                # Text
                switch_canvas.create_text(16, 12, text="ON", fill="white", font=("Segoe UI", 7, "bold"))
            else:
                switch_canvas.create_oval(2, 2, 22, 22, fill="#94a3b8", outline="")
                switch_canvas.create_oval(28, 2, 48, 22, fill="#94a3b8", outline="")
                switch_canvas.create_rectangle(12, 2, 38, 22, fill="#94a3b8", outline="")
                # White Circle
                switch_canvas.create_oval(4, 4, 24, 20, fill="white", outline="")
                # Text
                switch_canvas.create_text(34, 12, text="OFF", fill="white", font=("Segoe UI", 7, "bold"))

        
        def update_toggle_ui(e=None):
            if e: # Chamado via clique
                is_eventual_var.set(not is_eventual_var.get())
            
            active = is_eventual_var.get()
            draw_switch(active)
            if active:
                layout_row = 12 # Ajustado para vir depois de Observações
                lbl_data.grid(row=layout_row, column=0, sticky="w", pady=4)
                row_data.grid(row=layout_row, column=1, sticky="ew", pady=4, padx=(10, 0))
                try: 
                    combo_cur.configure(state="normal")
                    # combo_cur.delete(0, tk.END) # Removido para não apagar ao alternar
                except: pass
            else:
                layout_row = 12 # Ajustado para vir depois de Observações
                lbl_data.grid_remove()
                row_data.grid_remove()
                try: combo_cur.configure(state="readonly")
                except: pass

        # Container do Toggle Customizado (seguindo estilo da imagem)
        toggle_frame = tk.Frame(f, bg="#eff6ff", highlightthickness=1, 
                                highlightbackground="#dbeafe", bd=0, cursor="hand2")
        toggle_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 15), ipady=6)
        
        lbl_icon = tk.Label(toggle_frame, text="📌", bg="#eff6ff", fg=D["primary"], font=("Segoe UI", 12))
        lbl_icon.pack(side=tk.LEFT, padx=(15, 5))
        
        lbl_text = tk.Label(toggle_frame, text="Inserir Aula Eventual", bg="#eff6ff", 
                           fg=D["primary"], font=("Segoe UI", 11, "bold"))
        lbl_text.pack(side=tk.LEFT)
        
        switch_canvas = tk.Canvas(toggle_frame, width=50, height=24, bg="#eff6ff", 
                                 highlightthickness=0, cursor="hand2")
        switch_canvas.pack(side=tk.RIGHT, padx=15)
        
        # Desenho inicial
        draw_switch(is_eventual_var.get())
        
        # Binds para clique em qualquer lugar do componente
        for w in [toggle_frame, lbl_icon, lbl_text, switch_canvas]:
            w.bind("<Button-1>", update_toggle_ui)
            
        # Chamada inicial movida para baixo para garantir que combo_cur exista
        pass

        def lbl(r, t): tk.Label(f, text=t, bg=D["bg"], fg=D["fg2"], font=("Segoe UI", 9, "bold")).grid(row=r+1, column=0, sticky="w", pady=4)
        def field(r, w): w.grid(row=r+1, column=1, sticky="ew", pady=4, padx=(10, 0)); return w

        f.columnconfigure(1, weight=1)

        lbl(0, "Laboratório:")
        combo_lab = ttk.Combobox(f, values=LABORATORIOS, state="readonly")
        field(0, combo_lab)

        lbl(1, "Dia da Semana:")
        combo_dia = ttk.Combobox(f, values=DIAS_SEMANA, state="readonly")
        field(1, combo_dia)

        lbl(2, "Início (HH:MM):")
        ent_ini = tk.Entry(f, highlightthickness=1, highlightbackground=D["border"],
                           highlightcolor=D["primary"], bd=0,
                           bg=D["input"], fg=D["fg"], insertbackground=D["fg"])
        ent_ini.bind("<KeyRelease>", self._mask_hora)
        field(2, ent_ini)

        lbl(3, "Fim (HH:MM):")
        ent_fim = tk.Entry(f, highlightthickness=1, highlightbackground=D["border"],
                           highlightcolor=D["primary"], bd=0,
                           bg=D["input"], fg=D["fg"], insertbackground=D["fg"])
        ent_fim.bind("<KeyRelease>", self._mask_hora)
        field(3, ent_fim)

        lbl(4, "Faculdade:")
        combo_fac = ttk.Combobox(f, values=self.db.listar_faculdades(), state="readonly")
        field(4, combo_fac)

        lbl(5, "Curso:")
        combo_cur = ttk.Combobox(f, state="readonly")
        field(5, combo_cur)

        lbl(6, "Disciplina:")
        # Container híbrido: Entry para busca + Botão para dropdown completo
        container_dis = tk.Frame(f, bg=D["bg"])
        field(6, container_dis)
        container_dis.columnconfigure(0, weight=1)
        
        ent_dis = tk.Entry(container_dis, highlightthickness=1, highlightbackground=D["border"],
                           highlightcolor=D["primary"], bd=0,
                           bg=D["input"], fg=D["fg"], insertbackground=D["fg"])
        ent_dis.grid(row=0, column=0, sticky="ew")
        
        btn_drop_dis = tk.Button(container_dis, text="▼", bg=D["card"], fg=D["primary"],
                                 activebackground=D["primary"], activeforeground="white",
                                 bd=0, relief="flat", width=2, cursor="hand2", 
                                 font=("Segoe UI", 8))
        btn_drop_dis.grid(row=0, column=1, padx=(2, 0), sticky="ns")
        
        # Dropdown (Toplevel com Listbox)
        list_win = tk.Toplevel(win)
        list_win.withdraw()
        list_win.overrideredirect(True)
        list_win.attributes("-topmost", True)
        
        list_frame = tk.Frame(list_win, bg=D["border"], bd=1)
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        lb_dis = tk.Listbox(list_frame, bg=D["card"], fg=D["fg"], 
                            selectbackground=D["primary"], selectforeground="white",
                            font=("Segoe UI", 10), bd=0, highlightthickness=0,
                            activestyle="none")
        lb_dis.pack(fill=tk.BOTH, expand=True)

        lbl(7, "Turma:")
        container_tur = tk.Frame(f, bg=D["bg"])
        field(7, container_tur)
        container_tur.columnconfigure(0, weight=1)
        
        ent_tur = tk.Entry(container_tur, highlightthickness=1, highlightbackground=D["border"],
                           highlightcolor=D["primary"], bd=0,
                           bg=D["input"], fg=D["fg"], insertbackground=D["fg"])
        ent_tur.grid(row=0, column=0, sticky="ew")
        
        btn_drop_tur = tk.Button(container_tur, text="▼", bg=D["card"], fg=D["primary"],
                                 activebackground=D["primary"], activeforeground="white",
                                 bd=0, relief="flat", width=2, cursor="hand2", 
                                 font=("Segoe UI", 8))
        btn_drop_tur.grid(row=0, column=1, padx=(2, 0), sticky="ns")
        
        # Dropdown para Turma
        list_win_tur = tk.Toplevel(win)
        list_win_tur.withdraw()
        list_win_tur.overrideredirect(True)
        list_win_tur.attributes("-topmost", True)
        
        list_frame_tur = tk.Frame(list_win_tur, bg=D["border"], bd=1)
        list_frame_tur.pack(fill=tk.BOTH, expand=True)
        
        lb_tur = tk.Listbox(list_frame_tur, bg=D["card"], fg=D["fg"], 
                            selectbackground=D["primary"], selectforeground="white",
                            font=("Segoe UI", 10), bd=0, highlightthickness=0,
                            activestyle="none")
        lb_tur.pack(fill=tk.BOTH, expand=True)

        lbl(8, "Professor:")
        container_prof = tk.Frame(f, bg=D["bg"])
        field(8, container_prof)
        container_prof.columnconfigure(0, weight=1)
        
        ent_prof = tk.Entry(container_prof, highlightthickness=1, highlightbackground=D["border"],
                            highlightcolor=D["primary"], bd=0,
                            bg=D["input"], fg=D["fg"], insertbackground=D["fg"])
        ent_prof.grid(row=0, column=0, sticky="ew")
        
        btn_drop_prof = tk.Button(container_prof, text="▼", bg=D["card"], fg=D["primary"],
                                 activebackground=D["primary"], activeforeground="white",
                                 bd=0, relief="flat", width=2, cursor="hand2", 
                                 font=("Segoe UI", 8))
        btn_drop_prof.grid(row=0, column=1, padx=(2, 0), sticky="ns")
        
        # Dropdown para Professor
        list_win_prof = tk.Toplevel(win)
        list_win_prof.withdraw()
        list_win_prof.overrideredirect(True)
        list_win_prof.attributes("-topmost", True)
        
        list_frame_prof = tk.Frame(list_win_prof, bg=D["border"], bd=1)
        list_frame_prof.pack(fill=tk.BOTH, expand=True)
        
        lb_prof = tk.Listbox(list_frame_prof, bg=D["card"], fg=D["fg"], 
                            selectbackground=D["primary"], selectforeground="white",
                            font=("Segoe UI", 10), bd=0, highlightthickness=0,
                            activestyle="none")
        lb_prof.pack(fill=tk.BOTH, expand=True)
        lbl(9, "qtd Alunos:")
        ent_qtde = tk.Entry(f, highlightthickness=1, highlightbackground=D["border"],
                            highlightcolor=D["primary"], bd=0,
                            bg=D["input"], fg=D["fg"], insertbackground=D["fg"])
        ent_qtde.insert(0, "0")
        ent_qtde.bind("<KeyRelease>", self._mask_int)
        field(9, ent_qtde)
        
        lbl(10, "Observações:")
        obs_container = tk.Frame(f, bg=D["bg"])
        field(10, obs_container)
        obs_container.columnconfigure(0, weight=1)

        from config import PESOS_OBSERVACAO
        txt_obs = tk.Text(obs_container, height=3, font=("Segoe UI", 9),
                         highlightthickness=1, highlightbackground=D["border"],
                         highlightcolor=D["primary"], bd=0,
                         bg=D["input"], fg=D["fg"], insertbackground=D["fg"])
        txt_obs.grid(row=0, column=0, sticky="ew")

        combo_peso = ttk.Combobox(obs_container, values=PESOS_OBSERVACAO, state="readonly", width=12)
        combo_peso.set("Baixo")
        combo_peso.grid(row=0, column=1, padx=(10, 0), sticky="n")

        # lbl_data e row_data já definidos no topo do escopo.
        # Ajustamos o posicionamento deles para não sobrepor
        lbl_data.grid_forget()
        row_data.grid_forget()
        
        
        from datetime import datetime
        data_ini = datetime.now()
        if aula and aula.data_eventual:
            try: data_ini = datetime.strptime(aula.data_eventual, "%Y-%m-%d")
            except: pass

        if DateEntry:
            cal_eventual = DateEntry(row_data, width=28, background=D["primary"],
                                   foreground='white', borderwidth=2,
                                   date_pattern='dd/mm/yyyy', locale='pt_BR')
            cal_eventual.set_date(data_ini)
            cal_eventual.pack(fill=tk.X)
        else:
            cal_eventual = tk.Entry(row_data, bg=D["input"], fg=D["fg"], 
                                   insertbackground=D["fg"], bd=1, relief="solid") # Fallback dark
            cal_eventual.insert(0, data_ini.strftime("%d/%m/%Y"))
            cal_eventual.pack(fill=tk.X)

        # Chamada final para sincronizar o estado visual do toggle com os dados carregados
        update_toggle_ui()
        
        def on_fac_select(e):
            cursos = self.db.listar_cursos(combo_fac.get())
            combo_cur["values"] = cursos
            combo_cur.set("")
            # Reset disciplinas list when faculty changes
            nonlocal disciplinas_full
            disciplinas_full = []
            ent_dis.delete(0, tk.END)
        combo_fac.bind("<<ComboboxSelected>>", on_fac_select)

        disciplinas_full = []
        from utils import remover_acentos

        def on_cur_select(e):
            # Normaliza o nome do curso para busca robusta no DB
            curso_raw = combo_cur.get()
            curso_nome = remover_acentos(curso_raw)
            nonlocal disciplinas_full
            
            # Tenta buscar pelo nome normalizado e se falhar tenta o original
            disciplinas_full = self.db.listar_disciplinas(curso_nome)
            if not disciplinas_full:
                disciplinas_full = self.db.listar_disciplinas(curso_raw)
                
            lb_dis.delete(0, tk.END)
            for d in disciplinas_full: lb_dis.insert(tk.END, d)
            
            # Sugestão de cor
            sugestao_cor = self.db.obter_cor_curso(curso_nome)
            if sugestao_cor and sugestao_cor != "#ffffff":
                pass

        combo_cur.bind("<<ComboboxSelected>>", on_cur_select)

        def _selecionar_dis(e=None):
            if not lb_dis.curselection(): return
            sel = lb_dis.get(lb_dis.curselection())
            ent_dis.delete(0, tk.END)
            ent_dis.insert(0, sel)
            list_win.withdraw()
            on_dis_select(None)
            ent_dis.focus_set()

        lb_dis.bind("<ButtonRelease-1>", _selecionar_dis)
        lb_dis.bind("<Return>", _selecionar_dis)

        def _filtrar_disciplinas(event):
            if event.keysym in ("Up", "Down", "Return", "Escape", "Tab"):
                if event.keysym == "Down":
                    if list_win.winfo_viewable():
                        lb_dis.focus_set()
                        lb_dis.selection_set(0)
                elif event.keysym == "Escape":
                    list_win.withdraw()
                return
            
            _exibir_dropdown(filtrar=True)

        def _exibir_dropdown(filtrar=False):
            nonlocal disciplinas_full
            val_digitado = ent_dis.get()
            
            if not filtrar or not val_digitado:
                filtrados = disciplinas_full
            else:
                u_val = remover_acentos(val_digitado.upper())
                filtrados = [d for d in disciplinas_full if u_val in remover_acentos(d.upper())]
            
            lb_dis.delete(0, tk.END)
            for f_item in filtrados:
                lb_dis.insert(tk.END, f_item)

            if filtrados:
                # Posiciona o popup relativo ao container para alinhar com o botão também
                win.update_idletasks() # Garante coordenadas atualizadas
                x = container_dis.winfo_rootx()
                y = container_dis.winfo_rooty() + container_dis.winfo_height()
                w = container_dis.winfo_width()
                h = min(200, len(filtrados) * 25 + 5)
                list_win.geometry(f"{w}x{h}+{x}+{y}")
                list_win.deiconify()
                if filtrar: # Se for busca, mantém foco no entry
                    ent_dis.focus_set()
                else: # Se for clique no botão, pode focar na lista
                    lb_dis.focus_set()
                    if lb_dis.size() > 0: lb_dis.selection_set(0)
            else:
                list_win.withdraw()

        def _toggle_drop(e=None):
            if list_win.winfo_viewable():
                list_win.withdraw()
            else:
                _exibir_dropdown(filtrar=False)

        btn_drop_dis.config(command=_toggle_drop)

        ent_dis.bind("<KeyRelease>", _filtrar_disciplinas)
        
        def _on_focus_out(e):
            # Pequeno delay para permitir o clique no listbox antes de fechar
            win.after(200, lambda: list_win.withdraw() if win.focus_get() not in (ent_dis, lb_dis) else None)
        
        ent_dis.bind("<FocusOut>", _on_focus_out)
        lb_dis.bind("<FocusOut>", _on_focus_out)

        def on_dis_select(e):
            val = ent_dis.get()
            nonlocal turmas_full
            turmas_full = self.db.listar_turmas(val)
            lb_tur.delete(0, tk.END)
            for t in turmas_full: lb_tur.insert(tk.END, t)
            ent_tur.delete(0, tk.END)

        def on_tur_select(e=None):
            qtde = self.db.obter_alunos_turma(ent_dis.get().strip().upper(), ent_tur.get())
            if qtde > 0:
                ent_qtde.delete(0, tk.END)
                ent_qtde.insert(0, str(qtde))

        # Lógica para Turma
        turmas_full = []

        def _selecionar_tur(e=None):
            if not lb_tur.curselection(): return
            sel = lb_tur.get(lb_tur.curselection())
            ent_tur.delete(0, tk.END)
            ent_tur.insert(0, sel)
            list_win_tur.withdraw()
            on_tur_select()
            ent_tur.focus_set()

        lb_tur.bind("<ButtonRelease-1>", _selecionar_tur)
        lb_tur.bind("<Return>", _selecionar_tur)

        def _exibir_dropdown_tur(filtrar=False):
            nonlocal turmas_full
            val_digitado = ent_tur.get()
            
            if not filtrar or not val_digitado:
                filtrados = turmas_full
            else:
                u_val = remover_acentos(val_digitado.upper())
                filtrados = [t for t in turmas_full if u_val in remover_acentos(t.upper())]
            
            lb_tur.delete(0, tk.END)
            for f_item in filtrados:
                lb_tur.insert(tk.END, f_item)

            if filtrados:
                win.update_idletasks()
                x = container_tur.winfo_rootx()
                y = container_tur.winfo_rooty() + container_tur.winfo_height()
                w = container_tur.winfo_width()
                h = min(200, len(filtrados) * 25 + 5)
                list_win_tur.geometry(f"{w}x{h}+{x}+{y}")
                list_win_tur.deiconify()
                if filtrar: 
                    ent_tur.focus_set()
                else: 
                    lb_tur.focus_set()
                    if lb_tur.size() > 0: lb_tur.selection_set(0)
            else:
                list_win_tur.withdraw()

        def _toggle_drop_tur(e=None):
            if list_win_tur.winfo_viewable():
                list_win_tur.withdraw()
            else:
                _exibir_dropdown_tur(filtrar=False)

        btn_drop_tur.config(command=_toggle_drop_tur)

        def _filtrar_turmas(event):
            if event.keysym in ("Up", "Down", "Return", "Escape", "Tab"):
                if event.keysym == "Down":
                    if list_win_tur.winfo_viewable():
                        lb_tur.focus_set()
                        lb_tur.selection_set(0)
                elif event.keysym == "Escape":
                    list_win_tur.withdraw()
                return
            _exibir_dropdown_tur(filtrar=True)

        ent_tur.bind("<KeyRelease>", _filtrar_turmas)
        
        def _on_focus_out_tur(e):
            win.after(200, lambda: list_win_tur.withdraw() if win.focus_get() not in (ent_tur, lb_tur) else None)
        
        ent_tur.bind("<FocusOut>", _on_focus_out_tur)
        lb_tur.bind("<FocusOut>", _on_focus_out_tur)

        # Lógica para Professor
        professores_full = self.db.listar_professores()

        def _selecionar_prof(e=None):
            if not lb_prof.curselection(): return
            sel = lb_prof.get(lb_prof.curselection())
            ent_prof.delete(0, tk.END)
            ent_prof.insert(0, sel)
            list_win_prof.withdraw()
            ent_prof.focus_set()

        lb_prof.bind("<ButtonRelease-1>", _selecionar_prof)
        lb_prof.bind("<Return>", _selecionar_prof)

        def _exibir_dropdown_prof(filtrar=False):
            nonlocal professores_full
            val_digitado = ent_prof.get()
            
            if not filtrar or not val_digitado:
                filtrados = professores_full
            else:
                u_val = remover_acentos(val_digitado.upper())
                filtrados = [p for p in professores_full if u_val in remover_acentos(p.upper())]
            
            lb_prof.delete(0, tk.END)
            for f_item in filtrados:
                lb_prof.insert(tk.END, f_item)

            if filtrados:
                win.update_idletasks()
                x = container_prof.winfo_rootx()
                y = container_prof.winfo_rooty() + container_prof.winfo_height()
                w = container_prof.winfo_width()
                h = min(200, len(filtrados) * 25 + 5)
                list_win_prof.geometry(f"{w}x{h}+{x}+{y}")
                list_win_prof.deiconify()
                if filtrar: 
                    ent_prof.focus_set()
                else: 
                    lb_prof.focus_set()
                    if lb_prof.size() > 0: lb_prof.selection_set(0)
            else:
                list_win_prof.withdraw()

        def _toggle_drop_prof(e=None):
            if list_win_prof.winfo_viewable():
                list_win_prof.withdraw()
            else:
                _exibir_dropdown_prof(filtrar=False)

        btn_drop_prof.config(command=_toggle_drop_prof)

        def _filtrar_professores(event):
            if event.keysym in ("Up", "Down", "Return", "Escape", "Tab"):
                if event.keysym == "Down":
                    if list_win_prof.winfo_viewable():
                        lb_prof.focus_set()
                        lb_prof.selection_set(0)
                elif event.keysym == "Escape":
                    list_win_prof.withdraw()
                return
            _exibir_dropdown_prof(filtrar=True)

        ent_prof.bind("<KeyRelease>", _filtrar_professores)
        
        def _on_focus_out_prof(e):
            win.after(200, lambda: list_win_prof.withdraw() if win.focus_get() not in (ent_prof, lb_prof) else None)
        
        ent_prof.bind("<FocusOut>", _on_focus_out_prof)
        lb_prof.bind("<FocusOut>", _on_focus_out_prof)

        # Preencher se for edição
        if aula:
            combo_lab.set(aula.laboratorio if str(aula.laboratorio).startswith("Lab") else f"Lab {aula.laboratorio}")
            combo_dia.set(aula.dia_semana)
            ent_ini.insert(0, aula.hora_inicio)
            ent_fim.insert(0, aula.hora_fim)
            combo_fac.set(aula.faculdade)
            on_fac_select(None)
            combo_cur.set(aula.curso)
            on_cur_select(None)
            ent_dis.delete(0, tk.END)
            ent_dis.insert(0, aula.disciplina)
            on_dis_select(None)
            ent_tur.delete(0, tk.END)
            ent_tur.insert(0, aula.turma)
            on_tur_select()
            ent_prof.insert(0, aula.professor)
            ent_qtde.delete(0, tk.END)
            ent_qtde.insert(0, str(aula.qtde_alunos))
            
            if hasattr(aula, 'peso_observacao'):
                combo_peso.set(aula.peso_observacao)
            if hasattr(aula, 'observacoes'):
                txt_obs.insert("1.0", aula.observacoes)
        else:
            # Novos valores padrão se vierem do clique na grade
            if pre_lab: combo_lab.set(pre_lab)
            if pre_dia: combo_dia.set(pre_dia)
            # Default para Qtde Alunos se for novo
            ent_qtde.insert(0, "0")
            combo_peso.set("Baixo")

        def salvar():
            D = self._D
            # Reset visual
            for e in [ent_ini, ent_fim, ent_prof]:
                e.configure(highlightbackground=D["border"])

            erros = []
            if not combo_lab.get(): 
                combo_lab.configure(style="Invalid.TCombobox")
                erros.append("- Laboratório")
            if not combo_dia.get(): 
                combo_dia.configure(style="Invalid.TCombobox")
                erros.append("- Dia da Semana")
            
            # Validação básica de hora
            h_ini = ent_ini.get().strip()
            h_fim = ent_fim.get().strip()
            if len(h_ini) < 5 or ":" not in h_ini:
                ent_ini.configure(highlightbackground="#ef4444")
                erros.append("- Horário de Início (Ex: 08:00)")
            if len(h_fim) < 5 or ":" not in h_fim:
                ent_fim.configure(highlightbackground="#ef4444")
                erros.append("- Horário de Fim (Ex: 11:30)")
                
            if not combo_fac.get(): 
                combo_fac.configure(style="Invalid.TCombobox")
                erros.append("- Faculdade")
            if not combo_cur.get(): 
                combo_cur.configure(style="Invalid.TCombobox")
                erros.append("- Curso")
            if not ent_dis.get().strip(): 
                ent_dis.configure(highlightbackground="#ef4444")
                erros.append("- Disciplina")
            if not ent_tur.get().strip(): 
                ent_tur.configure(highlightbackground="#ef4444")
                erros.append("- Turma")

            if erros:
                msg = "Os seguintes campos precisam ser preenchidos corretamente:\n\n" + "\n".join(erros)
                messagebox.showwarning("Campos Obrigatórios", msg, parent=win)
                return

            fac = combo_fac.get()
            cur = combo_cur.get().strip().upper()

            # Lógica para aula eventual: se o curso digitado não existe, cria ele
            if is_eventual_var.get() and cur:
                cursos_fac = self.db.listar_cursos(fac)
                if cur not in [c.upper() for c in cursos_fac]:
                    # Adiciona novo curso (cor padrão branca, o usuário pode mudar depois em Dados Cadastrados)
                    self.db.adicionar_curso(fac, cur)

            try:
                qtde_val = int(ent_qtde.get()) if ent_qtde.get() else 0
                nova_aula = Aula(
                    id=aula.id if aula else None,
                    laboratorio=combo_lab.get(),
                    dia_semana=combo_dia.get(),
                    hora_inicio=h_ini,
                    hora_fim=h_fim,
                    faculdade=combo_fac.get(),
                    curso=combo_cur.get(),
                    disciplina=ent_dis.get().strip().upper(),
                    turma=ent_tur.get().strip().upper(),
                    professor=ent_prof.get().upper(),
                    qtde_alunos=qtde_val,
                    cor_fundo=self.db.obter_cor_curso(combo_cur.get().strip().upper()),
                    observacoes=txt_obs.get("1.0", tk.END).strip(),
                    peso_observacao=combo_peso.get(),
                    is_eventual=is_eventual_var.get(),
                    data_eventual=None if not is_eventual_var.get() else (cal_eventual.get_date().strftime("%Y-%m-%d") if hasattr(cal_eventual, 'get_date') else cal_eventual.get())
                )
                
                # Verificação de conflito (apenas para aulas fixas)
                ignorar_id = aula.id if aula else None
                conflito = self.db.verificar_conflito(nova_aula, ignorar_id=ignorar_id)
                if conflito:
                    msg = (
                        f"Conflito de horário detectado!\n\n"
                        f"O laboratório {nova_aula.laboratorio} já possui uma aula "
                        f"nesse horário:\n\n"
                        f"  📚 {conflito.disciplina}\n"
                        f"  👥 Turma: {conflito.turma}\n"
                        f"  ⏰ {conflito.hora_inicio} – {conflito.hora_fim}\n\n"
                        f"Para cadastrar neste horário, utilize a opção "
                        f"\"Aula Eventual\"."
                    )
                    messagebox.showerror("Conflito de Horário", msg, parent=win)
                    return

                if aula: self.db.alterar_aula(nova_aula)
                else: self.db.adicionar_aula(nova_aula)
                
                self.atualizar_dados()
                win.destroy()
                messagebox.showinfo("Sucesso", "Aula salva com sucesso!", parent=self.root)
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao salvar aula: {e}", parent=win)

        self._btn_footer(footer, "💾  SALVAR",  "#10b981", salvar)
        self._btn_footer(footer, "✕  CANCELAR", "#64748b", win.destroy)
        
        if aula is not None:
            # Use local reference for lambda to satisfy type checker
            aula_id = aula.id
            self._btn_footer(footer, "🗑️  EXCLUIR", "#ef4444", 
                             lambda: [self.db.apagar_aula(aula_id), self.atualizar_dados(), win.destroy()])

    # ------------------------------------------------------------------
    # CSV Import
    # ------------------------------------------------------------------

    def _janela_importar_dados(self):
        """Abre seletor de arquivo CSV e exibe barra de progresso para a importação."""
        caminho = filedialog.askopenfilename(
            title="Selecionar arquivo CSV",
            filetypes=[("Arquivos CSV", "*.csv"), ("Todos os arquivos", "*.*")],
            parent=self.root
        )
        if not caminho:
            return

        try:
            # Tenta UTF-8-SIG (com BOM) primeiro, depois CP1252 (Windows) e Latin-1
            rows = []
            encodings = ["utf-8-sig", "cp1252", "latin-1"]
            
            for enc in encodings:
                try:
                    with open(caminho, "r", encoding=enc) as f:
                        primeira_linha = f.readline()
                        f.seek(0)
                        delimitador = ";" if ";" in primeira_linha else ","
                        reader = csv.DictReader(f, delimiter=delimitador)
                        raw_rows = list(reader)
                        if not raw_rows: continue
                        
                        for r in raw_rows:
                            cleaned_row = {str(k or "").strip(): v for k, v in r.items()}
                            rows.append(cleaned_row)
                        break
                except UnicodeDecodeError:
                    continue

            if not rows:
                messagebox.showwarning("Atenção", "O arquivo CSV está vazio ou é inválido.", parent=self.root)
                return

            # Validação robusta de cabeçalho
            colunas_obrigatorias = ["Faculdade", "Curso", "Turma", "Disciplina"]
            first_row_keys = [k.upper() for k in rows[0].keys()]
            
            for col in colunas_obrigatorias:
                if col.upper() not in first_row_keys:
                    messagebox.showerror("Erro", f"Coluna '{col}' não encontrada no CSV.", parent=self.root)
                    return

            # --- Janela de Progresso ---
            self._import_cancelled = False
            progress_win = tk.Toplevel(self.root)
            progress_win.title("Importando Dados...")
            progress_win.geometry("400x150")
            progress_win.resizable(False, False)
            progress_win.grab_set() # Modal

            tk.Label(progress_win, text="Importando registros, aguarde...", font=("Segoe UI", 10)).pack(pady=10)
            
            pb = ttk.Progressbar(progress_win, orient="horizontal", length=300, mode="determinate")
            pb.pack(pady=10)
            
            lbl_status = tk.Label(progress_win, text="0%", font=("Segoe UI", 9))
            lbl_status.pack()

            def cancelar():
                self._import_cancelled = True
                progress_win.destroy()

            tk.Button(progress_win, text="CANCELAR", command=cancelar, bg="#ef4444", fg="white").pack(pady=10)

            # --- Callback de Progresso ---
            def progress_callback(atual, total):
                if self._import_cancelled:
                    return False # Aborta no DB
                
                perc = (atual / total) * 100
                pb["value"] = perc
                lbl_status.config(text=f"{atual} de {total} ({perc:.1f}%)")
                progress_win.update()
                return True

            # Prepara os dados (mapeamento case-insensitive)
            mapeamento = {col.upper(): col for col in colunas_obrigatorias + ["Professor"]}
            final_rows = []
            for r in rows:
                f_row = {}
                for k, v in r.items():
                    ku = k.upper()
                    if ku in mapeamento: f_row[mapeamento[ku]] = v
                final_rows.append(f_row)

            # Executa a importação
            self.db.importar_dados_csv(final_rows, callback=progress_callback)
            
            if not self._import_cancelled:
                progress_win.destroy()
                messagebox.showinfo("Sucesso", "Dados inseridos com sucesso!", parent=self.root)
            
            self.atualizar_dados()

        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao importar: {str(e)}", parent=self.root)

    def exportar_xlsx(self):
        """Versão minimalista da exportação Excel."""
        try:
            caminho = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel", "*.xlsx")]
            )
            if caminho:
                self.db.exportar_para_excel(caminho)
                messagebox.showinfo("Sucesso", "Exportação concluída!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro: {e}")
            
    def exportar_relatorio_excel(self):
        """Abre seletor de arquivo e exporta os dados do dashboard para Excel."""
        try:
            from datetime import datetime
            hoje = datetime.now().strftime("%Y-%m-%d_%H-%M")
            caminho = filedialog.asksaveasfilename(
                title="Salvar Relatório de Aulas",
                defaultextension=".xlsx",
                filetypes=[("Excel", "*.xlsx")],
                initialfile=f"relatorio_aulas_{hoje}.xlsx"
            )
            if caminho:
                self.db.exportar_estatisticas_excel(caminho)
                messagebox.showinfo("Sucesso", "Relatório exportado para Excel com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao exportar relatório: {e}")

    def exportar_csv(self):
        """Versão minimalista da exportação CSV."""
        try:
            caminho = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV", "*.csv")]
            )
            if caminho:
                self.db.exportar_para_csv(caminho)
                messagebox.showinfo("Sucesso", "Exportação concluída!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro: {e}")

    def exportar_grade_pdf(self, dia, lab_selecionado=None):
        """Exporta a grade de horários de um dia específico para PDF dividido em duas páginas."""
        if lab_selecionado is None:
            lab_selecionado = self._TODOS_LABS
        if not FPDF:
            messagebox.showerror("Erro", "Biblioteca 'fpdf2' não encontrada. Verifique a instalação.")
            return

        try:
            from fpdf.enums import XPos, YPos
            
            caminho = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF", "*.pdf")],
                initialfile=f"Grade_{dia.replace('/', '-')}.pdf"
            )
            if not caminho:
                return

            from datetime import datetime, timedelta
            import math
            
            # Setup PDF (Landscape)
            pdf = FPDF(orientation='L', unit='mm', format='A4')
            pdf.set_margins(left=7, top=10, right=7)
            pdf.set_auto_page_break(False)
            
            col_hora_w = 16
            col_lab_w = 19.5 
            row_h = 6.2  # Altura reduzida para caber mais linhas de 15 min
            intervalo = timedelta(minutes=15) 

            def desenhar_pagina(hora_inicio_str, hora_fim_str, titulo_msg):
                pdf.add_page()
                # Título
                pdf.set_font("helvetica", "B", 16)
                pdf.set_text_color(30, 41, 59)
                pdf.cell(0, 10, f"GRADE DE HORÁRIOS - {dia} ({titulo_msg})", border=0, 
                         new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C")
                pdf.ln(1)
                
                # Determina quais laboratórios mostrar
                labs_para_exibir = LABORATORIOS
                if lab_selecionado and lab_selecionado != self._TODOS_LABS:
                    labs_para_exibir = [lab_selecionado]
                
                # Cabeçalho Labs
                pdf.set_font("helvetica", "B", 9)
                pdf.set_fill_color(30, 41, 59) 
                pdf.set_text_color(255, 255, 255)
                pdf.cell(col_hora_w, row_h, "HORA", border=1, fill=True, align="C")
                
                # Ajusta largura se for apenas um lab
                w_lab = col_lab_w if len(labs_para_exibir) > 1 else 200
                
                for lab in labs_para_exibir:
                    pdf.cell(w_lab, row_h, f"LAB {lab}", border=1, fill=True, align="C")
                pdf.ln()

                # Coordenada Y onde começa a grade (abaixo do cabeçalho)
                grid_y_start = pdf.get_y()
                grid_x_start = pdf.get_x()

                h_lim_ini = datetime.strptime(hora_inicio_str, "%H:%M")
                h_lim_fim = datetime.strptime(hora_fim_str,    "%H:%M")
                
                # Coleta horários para a grade dinâmica da página
                tempos_pg = set()
                tempos_pg.add(h_lim_ini)
                tempos_pg.add(h_lim_fim)
                
                # Obtém todas as aulas do dia para capturar horários
                try:
                    dt_ref = datetime.strptime(dia, "%d/%m/%Y")
                    dia_pt = DIAS_SEMANA[dt_ref.weekday()]
                except ValueError:
                    dia_pt = dia
                    hoje = datetime.now()
                    hoje_idx = hoje.weekday()
                    try:
                        target_idx = DIAS_SEMANA.index(dia)
                        diff = target_idx - hoje_idx
                        dt_ref = hoje + timedelta(days=diff)
                    except ValueError:
                        dt_ref = hoje 

                aulas_dia_real = self._get_aulas_com_sobrescricao(dia_pt, dt_ref)
                
                for a in aulas_dia_real:
                    try:
                        hi = datetime.strptime(a.hora_inicio, "%H:%M")
                        hf = datetime.strptime(a.hora_fim,    "%H:%M")
                        if hi >= h_lim_ini and hi <= h_lim_fim: tempos_pg.add(hi)
                        if hf >= h_lim_ini and hf <= h_lim_fim: tempos_pg.add(hf)
                    except: pass
                
                page_times = sorted(list(tempos_pg))
                
                # 1. Desenha o fundo da grade (DYNAMIC ROWS)
                # row_h_ref = 6.2 mm para 15 minutos
                row_h_ref = 6.2
                
                for i in range(len(page_times) - 1):
                    t_curr = page_times[i]
                    t_next = page_times[i+1]
                    diff_m = (t_next - t_curr).total_seconds() / 60
                    h_row = (diff_m / 15) * row_h_ref
                    
                    # Coluna de hora
                    pdf.set_font("helvetica", "B", 8)
                    pdf.set_fill_color(241, 245, 249) 
                    pdf.set_text_color(30, 41, 59)
                    pdf.cell(col_hora_w, h_row, t_curr.strftime("%H:%M"), border=1, fill=True, align="C")
                    
                    # Células vazias para cada lab
                    for _ in labs_para_exibir:
                        pdf.set_fill_color(255, 255, 255)
                        pdf.cell(w_lab, h_row, "", border=1)
                    pdf.ln()

                # Label final de hora (base)
                pdf.cell(col_hora_w, 4, page_times[-1].strftime("%H:%M"), border=0, align="C")
                pdf.ln()

                # 2. Desenha os blocos de aulas por cima (OVERLAY FIEL)

                from utils import hex_to_rgb, texto_contraste
                
                for idx_lab, lab in enumerate(labs_para_exibir):
                    aulas_lab = [a for a in aulas_dia_real if a.laboratorio == lab]
                    
                    # Filtra e prepara horários
                    aulas_periodo = []
                    for a in aulas_lab:
                        try:
                            hi = datetime.strptime(a.hora_inicio, "%H:%M")
                            hf = datetime.strptime(a.hora_fim,    "%H:%M")
                            # Se a aula intersecta o intervalo da página
                            if hi < h_lim_fim and hf > h_lim_ini:
                                # Clampa os horários visuais para não sair da página se exceder
                                hi_v = max(hi, h_lim_ini)
                                hf_v = min(hf, h_lim_fim)
                                aulas_periodo.append((a, hi, hf, hi_v, hf_v))
                        except: continue
                    
                    if not aulas_periodo: continue
                    aulas_periodo.sort(key=lambda x: x[1])

                    # Agrupar em clusters para lidar com sobreposições (lanes)
                    clusters = []
                    if aulas_periodo:
                        curr = [aulas_periodo[0]]
                        l_end = aulas_periodo[0][2]
                        for i in range(1, len(aulas_periodo)):
                            itm = aulas_periodo[i]
                            if itm[1] < l_end:
                                curr.append(itm)
                                l_end = max(l_end, itm[2])
                            else:
                                clusters.append(curr)
                                curr = [itm]
                                l_end = itm[2]
                        clusters.append(curr)

                    # Desenhar cada cluster no lab atual
                    x_lab_start = grid_x_start + col_hora_w + (idx_lab * w_lab)
                    
                    for cluster in clusters:
                        lanes_end = []
                        aula_to_lane = {}
                        for a, hi, hf, hiv, hfv in cluster:
                            l_idx = -1
                            for j, end_t in enumerate(lanes_end):
                                if hi >= end_t:
                                    l_idx = j
                                    lanes_end[j] = hf
                                    break
                            if l_idx == -1:
                                l_idx = len(lanes_end)
                                lanes_end.append(hf)
                            aula_to_lane[a] = l_idx
                        
                        num_lanes = len(lanes_end)
                        w_lane = w_lab / num_lanes

                        for a, hi, hf, hiv, hfv in cluster:
                            l_idx = aula_to_lane[a]
                            
                            # CÁLCULO FIEL DA POSIÇÃO Y E ALTURA
                            # (valor_minutos / 15) * row_h
                            m_desde_inicio = (hiv - h_lim_ini).total_seconds() / 60
                            y_block = grid_y_start + (m_desde_inicio / 15) * row_h
                            
                            m_duracao = (hfv - hiv).total_seconds() / 60
                            h_block = (m_duracao / 15) * row_h
                            
                            x_block = x_lab_start + (l_idx * w_lane)
                            
                            # Desenha o retângulo preenchido
                            # Eventuais sempre usam fundo preto
                            if a.is_eventual:
                                bg_h = "#000000"
                            else:
                                bg_h = a.cor_fundo if (a.cor_fundo and a.cor_fundo != "#ffffff") else "#ddeeff"
                            pdf.set_fill_color(*hex_to_rgb(bg_h))
                            pdf.rect(x_block, y_block, w_lane, h_block, style="F")
                            
                            # Desenha contorno
                            pdf.set_draw_color(100, 100, 100)
                            pdf.rect(x_block, y_block, w_lane, h_block, style="D")
                            
                            # Escreve o texto centralizado
                            pdf.set_text_color(*hex_to_rgb(texto_contraste(bg_h)))
                            pdf.set_xy(x_block, y_block + 1)
                            
                            # Escreve o texto centralizado e com fonte dinâmica
                            prefixo = "[E] " if a.is_eventual else ""
                            txt = f"{a.hora_inicio} - {a.hora_fim}\n{prefixo}{a.disciplina}\n{a.curso}\n{a.turma}\nProf: {a.professor}\nQtd: {a.qtde_alunos}"
                            
                            # Ajuste de fonte e margens
                            base_f_size = 6.0 if num_lanes == 1 else 5.0
                            if h_block < 8: base_f_size -= 1.0 # Reduz p/ aulas curtas
                            pdf.set_font("helvetica", "B", base_f_size)
                            pdf.set_text_color(*hex_to_rgb(texto_contraste(bg_h)))

                            # Tenta calcular a altura do bloco de texto para centralizar verticalmente
                            w_inner = w_lane - 1.5 # Margem interna
                            line_h = base_f_size * 0.45 # Estimativa de altura de linha em mm
                            
                            # Divide o texto em linhas para contar (fpdf2 multi_cell não retorna altura facilmente aqui)
                            linhas = pdf.multi_cell(w_inner, line_h, txt, dry_run=True, align="C", output="LINES")
                            total_txt_h = len(linhas) * line_h
                            
                            # Centraliza verticalmente
                            y_text = y_block + (h_block - total_txt_h) / 2
                            if y_text < y_block: y_text = y_block + 0.5 # Segurança
                            
                            pdf.set_xy(x_block + 0.75, y_text)
                            pdf.multi_cell(w_inner, line_h, txt, border=0, align="C")

                pdf.set_draw_color(0, 0, 0) # Reset draw color

            # Página 1: Manhã/Início de Tarde
            desenhar_pagina("07:30", "15:00", "MANHÃ / TARDE")
            
            # Página 2: Final de Tarde/Noite
            desenhar_pagina("15:15", "22:45", "TARDE / NOITE")

            pdf.output(caminho)
            messagebox.showinfo("Sucesso", f"Grade exportada com sucesso para:\n{caminho}")

        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Erro", f"Não foi possível exportar o PDF: {e}")

    def _janela_dados_cadastrados(self):
        """Janela multi-abas para gerenciar todos os dados cadastrados."""
        D = self._D
        win = tk.Toplevel(self.root)
        win.title("Dados Cadastrados")
        win.geometry("920x660")
        win.configure(bg=D["bg"])
        win.grab_set()

        # --------------- Header ---------------
        header = tk.Frame(win, bg=D["hdr"], height=56)
        header.pack(fill=tk.X)
        header.pack_propagate(False)
        tk.Label(header, text="📊  DADOS CADASTRADOS", font=("Segoe UI", 13, "bold"),
                 bg=D["hdr"], fg=D["fg"]).pack(side=tk.LEFT, padx=20, pady=14)

        # --------------- Notebook ---------------
        nb_style = ttk.Style()
        nb_style.configure("Dados.TNotebook", background=D["bg"], tabmargins=[2, 5, 2, 0],
                            borderwidth=0)
        nb_style.configure("Dados.TNotebook.Tab",
                           font=("Segoe UI", 9, "bold"),
                           padding=[14, 7],
                           background=D["card"],
                           foreground=D["fg2"])
        nb_style.map("Dados.TNotebook.Tab",
                     background=[("selected", D["primary"])],
                     foreground=[("selected", "white")])

        nb = ttk.Notebook(win, style="Dados.TNotebook")
        nb.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # ═══════════════════════════════════════════════════════════════
        # Helper: cria frame de aba com Treeview + scrollbar + footer
        # ═══════════════════════════════════════════════════════════════
        def _criar_aba(notebook, titulo, icone, colunas, larguras, ancoras=None, search_callback=None):
            """Retorna (frame_aba, tree, btn_editar, btn_excluir, footer, ent_busca)."""
            frame = ttk.Frame(notebook)
            notebook.add(frame, text=f"{icone}  {titulo}")
            frame.rowconfigure(1, weight=1) # O Treeview agora está na linha 1
            frame.columnconfigure(0, weight=1)

            # --- Barra de Busca ---
            search_frame = tk.Frame(frame, bg=D["bg"], pady=5)
            search_frame.grid(row=0, column=0, sticky="ew", padx=10)
            tk.Label(search_frame, text="🔍", bg=D["bg"], fg=D["fg2"], font=("Segoe UI", 11)).pack(side=tk.LEFT, padx=(5, 2))
            ent_busca = tk.Entry(search_frame, font=("Segoe UI", 10), bg=D["input"], fg=D["fg"],
                                  insertbackground=D["fg"], bd=1, relief="solid")
            ent_busca.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5, pady=5)
            if search_callback:
                ent_busca.bind("<KeyRelease>", lambda e: search_callback(ent_busca.get()))

            # --- Treeview ---
            tree_frame = tk.Frame(frame, bg=D["card"])
            tree_frame.grid(row=1, column=0, sticky="nsew", padx=8, pady=(4, 4))
            tree_frame.rowconfigure(0, weight=1)
            tree_frame.columnconfigure(0, weight=1)

            # Adiciona coluna de seleção no início
            colunas_com_sel = ["☐"] + list(colunas)
            tree = ttk.Treeview(tree_frame, columns=colunas_com_sel, show="headings",
                                selectmode="browse")
            
            # Configura coluna de seleção
            tree.heading("☐", text="☐")
            tree.column("☐", width=40, minwidth=40, anchor="center", stretch=False)
            
            for i, (col, larg) in enumerate(zip(colunas, larguras)):
                anc = ancoras[i] if ancoras else "w"
                tree.heading(col, text=col.upper())
                is_flex = col.upper() in ["DISCIPLINA", "CURSO", "NOME DO CURSO", "NOME DA FACULDADE"]
                tree.column(col, width=larg, minwidth=larg//2, anchor=anc, stretch=is_flex)

            sb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=tree.yview)
            tree.configure(yscrollcommand=sb.set)
            tree.grid(row=0, column=0, sticky="nsew")
            sb.grid(row=0, column=1, sticky="ns")

            self._setup_tree_hover(tree)

            # --- Footer ---
            footer = tk.Frame(frame, bg=D["footer"], pady=10)
            footer.grid(row=2, column=0, sticky="ew")

            def _btn(txt, bg, cmd, side=tk.LEFT, padx=(12, 4)):
                b = tk.Button(footer, text=txt, bg=bg, fg="white",
                              activebackground=bg, activeforeground="white",
                              font=("Segoe UI", 9, "bold"), padx=12, pady=6,
                              relief="flat", cursor="hand2", command=cmd)
                b.pack(side=side, padx=padx)
                return b

            btn_sel_todos = _btn("🗸  SEL. TODOS", "#6366f1", lambda: None)
            btn_batch_del = _btn("🗑️  EXCLUIR SELECIONADOS", D["danger"], lambda: None, side=tk.RIGHT, padx=(4, 12))
            btn_batch_del.pack_forget() # Escondido por padrão

            btn_edit = _btn("✏️  EDITAR",  D["primary"], lambda: None)
            btn_del  = _btn("🗑️  EXCLUIR", D["danger"],  lambda: None)
            btn_edit.config(state="disabled")
            btn_del.config(state="disabled")

            def _atualizar_batch_ui():
                tem_selecionados = any(tree.item(i, "values")[0] == "☑" for i in tree.get_children())
                if tem_selecionados:
                    btn_batch_del.pack(side=tk.RIGHT, padx=(4, 12))
                else:
                    btn_batch_del.pack_forget()

            def _on_click(event):
                region = tree.identify_region(event.x, event.y)
                if region == "cell":
                    col = tree.identify_column(event.x)
                    if col == "#1":
                        item = tree.identify_row(event.y)
                        if item:
                            vals = list(tree.item(item, "values"))
                            vals[0] = "☑" if vals[0] == "☐" else "☐"
                            tree.item(item, values=vals)
                            _atualizar_batch_ui()

            def _selecionar_todos():
                all_items = tree.get_children()
                if not all_items: return
                todas_marcadas = all(tree.item(i, "values")[0] == "☑" for i in all_items)
                nova_marcar = "☐" if todas_marcadas else "☑"
                for item in all_items:
                    vals = list(tree.item(item, "values"))
                    vals[0] = nova_marcar
                    tree.item(item, values=vals)
                _atualizar_batch_ui()

            btn_sel_todos.config(command=_selecionar_todos)
            tree.bind("<Button-1>", _on_click)

            def _on_select(event):
                sel = tree.selection()
                state = "normal" if sel else "disabled"
                btn_edit.config(state=state)
                btn_del.config(state=state)

            tree.bind("<<TreeviewSelect>>", _on_select)

            return frame, tree, btn_edit, btn_del, footer, ent_busca, btn_batch_del

        # ═══════════════════════════════════
        # ABA 1 — CURSOS
        # ═══════════════════════════════════
        def _buscar_cursos(texto):
            carregar_cursos(texto)

        f_cur, t_cur, be_cur, bd_cur, ft_cur, e_cur, bdel_cur = _criar_aba(
            nb, "Cursos", "🎨",
            ["Faculdade", "Curso", "Cor"],
            [230, 340, 120],
            ["w", "w", "center"],
            search_callback=_buscar_cursos
        )

        def importar_cores_csv():
            path = filedialog.askopenfilename(
                title="Selecionar CSV de Cores",
                filetypes=[("Arquivos CSV", "*.csv"), ("Todos os arquivos", "*.*")],
                parent=win
            )
            if not path: return
            
            try:
                encodings = ['utf-8-sig', 'cp1252', 'latin-1']
                rows = []
                success = False
                
                for enc in encodings:
                    try:
                        # Detecta delimitador
                        with open(path, 'r', encoding=enc) as f:
                            content = f.read(1024)
                            f.seek(0)
                            delim = ';' if ';' in content else ','
                            
                            reader = csv.DictReader(f, delimiter=delim)
                            # Normaliza nomes de colunas para minúsculas
                            temp_rows = []
                            for row in reader:
                                norm_row = {str(k or "").lower().strip(): v for k, v in row.items()}
                                if "curso" in norm_row and "cor" in norm_row:
                                    temp_rows.append(norm_row)
                            
                            if temp_rows:
                                rows = temp_rows
                                success = True
                                break
                    except UnicodeDecodeError:
                        continue
                
                if not success and not rows:
                    messagebox.showwarning("Atenção", "Não foi possível ler o arquivo (erro de codificação) ou nenhum dado válido encontrado. O CSV deve ter colunas 'curso' e 'cor'.", parent=win)
                    return
                
                if not rows:
                    messagebox.showwarning("Atenção", "Nenhum dado válido encontrado. O CSV deve ter colunas 'curso' e 'cor'.", parent=win)
                    return
                
                count = self.db.importar_cores(rows)
                messagebox.showinfo("Sucesso", f"{count} cores importadas/atualizadas com sucesso!", parent=win)
                carregar_cursos()
                self.atualizar_dados()
                
            except Exception as ex:
                messagebox.showerror("Erro", f"Erro ao importar CSV: {str(ex)}", parent=win)

        # Botão Importar Cores (específico da aba Cursos)
        btn_import_cor = tk.Button(ft_cur, text="📥  IMPORTAR CORES", bg="#8b5cf6", fg="white",
                               activebackground="#7c3aed", activeforeground="white",
                               font=("Segoe UI", 9, "bold"), padx=12, pady=6,
                               relief="flat", cursor="hand2", command=importar_cores_csv)
        btn_import_cor.pack(side=tk.LEFT, padx=(12, 4))
        # Mover os outros botões para depois deste para manter ordem visual
        be_cur.pack_forget(); be_cur.pack(side=tk.LEFT, padx=(4, 4))
        bd_cur.pack_forget(); bd_cur.pack(side=tk.LEFT, padx=(4, 4))

        def carregar_cursos(filtro=""):
            for item in t_cur.get_children(): t_cur.delete(item)
            busca = remover_acentos(filtro.upper()) if filtro else ""
            for c in self.db.obter_todos_cursos():
                fac = str(c["faculdade"]).upper()
                nome = str(c["nome"]).upper()
                
                if busca:
                    txt_norm = remover_acentos(f"{fac} {nome}")
                    if busca not in txt_norm:
                        continue

                cor_hex = str(c["cor"]).strip() if c["cor"] else "#ffffff"
                t_cur.insert("", tk.END, values=("☐", c["faculdade"], c["nome"], f"█ {cor_hex}"))

        def editar_cor_inline(event):
            region = t_cur.identify_region(event.x, event.y)
            if region != "cell": return
            
            column = t_cur.identify_column(event.x)
            if column != "#4": return # Coluna 4 é a cor
            
            rowid = t_cur.identify_row(event.y)
            if not rowid: return
            
            # Pega valores atuais
            vals = t_cur.item(rowid, "values")
            faculdade = vals[1]
            curso_nome = vals[2]
            cor_atual_vis = vals[3]
            cor_atual = cor_atual_vis.split()[-1] if " " in cor_atual_vis else cor_atual_vis
            
            # Get geometry
            x, y, width, height = t_cur.bbox(rowid, column)
            
            # Entry por cima
            ent_edit = tk.Entry(t_cur, font=("Segoe UI", 9), bd=1, relief="solid")
            ent_edit.insert(0, cor_atual)
            ent_edit.place(x=x, y=y, width=width, height=height)
            ent_edit.focus_set()
            ent_edit.selection_range(0, tk.END)
            
            def finish_edit(event=None):
                nova_cor = ent_edit.get().strip()
                ent_edit.destroy()
                
                if nova_cor and nova_cor != cor_atual:
                    # Validação básica de cor hex
                    if not nova_cor.startswith("#") or len(nova_cor) not in (4, 7):
                        messagebox.showwarning("Atenção", "Cor inválida. Use o formato HEX (Ex: #FF0000).", parent=win)
                        return
                    
                    self.db.editar_curso(curso_nome, curso_nome, faculdade, nova_cor)
                    carregar_cursos()
                    self.atualizar_dados()

            ent_edit.bind("<Return>", finish_edit)
            ent_edit.bind("<FocusOut>", finish_edit)
            ent_edit.bind("<Escape>", lambda e: ent_edit.destroy())

        t_cur.bind("<Double-1>", editar_cor_inline)

        def adicionar_curso():
            D = self._D
            d = tk.Toplevel(win); d.title("Adicionar Curso"); d.geometry("420x350")
            d.configure(bg=D["bg"]); d.grab_set()
            tk.Label(d, text="Faculdade:", bg=D["bg"], fg=D["fg2"], font=("Segoe UI", 9, "bold")).pack(pady=(14, 2))
            cb_fac = ttk.Combobox(d, values=self.db.listar_faculdades(), state="readonly", width=40)
            cb_fac.pack(padx=20, pady=4)
            tk.Label(d, text="Nome do Curso:", bg=D["bg"], fg=D["fg2"], font=("Segoe UI", 9, "bold")).pack(pady=(10, 2))
            ent = tk.Entry(d, font=("Segoe UI", 10), width=42, bg=D["input"], fg=D["fg"], insertbackground=D["fg"], bd=1, relief="solid"); ent.pack(padx=20, pady=4)
            cor_var = tk.StringVar(value=D["primary"])
            prv = tk.Label(d, text="Cor do Curso", bg=cor_var.get(), fg="white", pady=8, font=("Segoe UI", 9, "bold"))
            prv.pack(pady=12, padx=20, fill=tk.X)
            def _mudar():
                r = colorchooser.askcolor(color=cor_var.get(), parent=d)
                if r[1]: cor_var.set(r[1]); prv.configure(bg=r[1])
            tk.Button(d, text="Escolher Cor", command=_mudar, bg=D["primary"], fg="white",
                      relief="flat", padx=14, pady=6, font=("Segoe UI", 9, "bold")).pack()
            def _salvar():
                fac = cb_fac.get().strip().upper()
                nome = ent.get().strip().upper()
                if not fac or not nome:
                    messagebox.showwarning("Atenção", "Preencha Faculdade e Curso.", parent=d); return
                self.db.adicionar_curso(fac, nome, cor_var.get())
                d.destroy(); carregar_cursos(); self.atualizar_dados()
            tk.Button(d, text="SALVAR", bg=D["success"], fg="white", font=("Segoe UI", 10, "bold"),
                      command=_salvar, pady=8, width=20, relief="flat").pack(side=tk.BOTTOM, pady=15)

        def editar_curso():
            D = self._D
            sel = t_cur.selection()
            if not sel: return
            item_id = sel[0]
            vals = t_cur.item(item_id, "values")
            checkbox, fac, nome, vis_cor = str(vals[0]), str(vals[1]), str(vals[2]), str(vals[3])
            cor = vis_cor.split()[-1] if " " in vis_cor else vis_cor
            d = tk.Toplevel(win); d.title(f"Editar: {nome}"); d.geometry("420x350")
            d.configure(bg=D["bg"]); d.grab_set()
            tk.Label(d, text=f"Editando Curso: {nome}", font=("Segoe UI", 11, "bold"),
                     bg=D["bg"], fg=D["fg"]).pack(pady=(14, 4))
            tk.Label(d, text="Novo nome:", bg=D["bg"], fg=D["fg2"], font=("Segoe UI", 9, "bold")).pack()
            ent = tk.Entry(d, font=("Segoe UI", 10), width=42, bg=D["input"], fg=D["fg"], insertbackground=D["fg"], bd=1, relief="solid")
            ent.insert(0, nome); ent.pack(padx=20, pady=4)
            cor_var = tk.StringVar(value=cor)
            prv = tk.Label(d, text="Cor do Curso", bg=cor, fg="white", pady=8, font=("Segoe UI", 9, "bold"))
            prv.pack(pady=12, padx=20, fill=tk.X)
            def _mudar():
                r = colorchooser.askcolor(color=cor_var.get(), parent=d)
                if r[1]: cor_var.set(r[1]); prv.configure(bg=r[1])
            tk.Button(d, text="Mudar Cor", command=_mudar, bg="#8b5cf6", fg="white",
                      relief="flat", padx=10, pady=4).pack()
            def _salvar():
                novo = ent.get().strip().upper()
                if not novo:
                    messagebox.showwarning("Atenção", "Nome não pode ser vazio.", parent=d); return
                self.db.editar_curso(nome, novo, fac, cor_var.get())
                messagebox.showinfo("Sucesso", "Curso atualizado!", parent=d)
                d.destroy(); carregar_cursos(); self.atualizar_dados()
            tk.Button(d, text="SALVAR", bg="#10b981", fg="white", font=("Segoe UI", 10, "bold"),
                      command=_salvar, pady=6).pack(side=tk.BOTTOM, pady=10)

        def excluir_curso():
            sel = t_cur.selection()
            if not sel: return
            item_id = sel[0]
            vals = t_cur.item(item_id, "values"); fac, nome = str(vals[1]), str(vals[2])
            if messagebox.askyesno("Confirmar", f"Excluir curso '{nome}'?\n(As aulas vinculadas serão mantidas)", parent=win):
                self.db.excluir_curso(nome, fac); carregar_cursos(); self.atualizar_dados()
                messagebox.showinfo("Sucesso", "Curso excluído.", parent=win)

        def excluir_cursos_em_lote():
            selecionados = []
            for item in t_cur.get_children():
                vals = t_cur.item(item, "values")
                if vals[0] == "☑":
                    selecionados.append((vals[1], vals[2])) # (faculdade, nome)
            
            if not selecionados: return
            
            msg = f"Excluir {len(selecionados)} cursos selecionados?\n(As aulas vinculadas serão mantidas em modo órfão)"
            if messagebox.askyesno("Confirmar Exclusão em Lote", msg, parent=win):
                for fac, nome in selecionados:
                    self.db.excluir_curso(nome, fac)
                carregar_cursos()
                self.atualizar_dados()
                messagebox.showinfo("Sucesso", f"{len(selecionados)} cursos excluídos.", parent=win)

        tk.Button(ft_cur, text="➕  ADICIONAR CURSO", bg="#10b981", fg="white",
                  font=("Segoe UI", 9, "bold"), padx=12, pady=6, relief="flat",
                  cursor="hand2", command=adicionar_curso).pack(side=tk.LEFT, padx=(12, 4))
        be_cur.config(command=editar_curso)
        bd_cur.config(command=excluir_curso)
        bdel_cur.config(command=excluir_cursos_em_lote)
        carregar_cursos()

        # ═══════════════════════════════════
        # ABA 2 — FACULDADES
        # ═══════════════════════════════════
        def _buscar_faculdades(texto):
            carregar_faculdades(texto)

        f_fac, t_fac, be_fac, bd_fac, ft_fac, e_fac, bdel_fac = _criar_aba(
            nb, "Faculdades", "🏛️",
            ["Nome da Faculdade"],
            [700],
            ["w"],
            search_callback=_buscar_faculdades
        )

        def carregar_faculdades(filtro=""):
            for item in t_fac.get_children(): t_fac.delete(item)
            busca = remover_acentos(filtro.upper()) if filtro else ""
            for f in self.db.listar_faculdades():
                if busca and busca not in remover_acentos(f.upper()):
                    continue
                t_fac.insert("", tk.END, values=("☐", f))

        def adicionar_faculdade():
            D = self._D
            d = tk.Toplevel(win); d.title("Nova Faculdade"); d.geometry("400x200")
            d.configure(bg=D["bg"]); d.grab_set()
            tk.Label(d, text="Nome da Faculdade:", bg=D["bg"], fg=D["fg2"],
                     font=("Segoe UI", 10, "bold")).pack(pady=(20, 6))
            ent = tk.Entry(d, font=("Segoe UI", 11), width=40, bg=D["input"], fg=D["fg"], insertbackground=D["fg"], bd=1, relief="solid"); ent.pack(padx=20, pady=4); ent.focus()
            def _salvar():
                nome = ent.get().strip().upper()
                if not nome:
                    messagebox.showwarning("Atenção", "Digite o nome.", parent=d); return
                self.db.adicionar_faculdade(nome)
                d.destroy(); carregar_faculdades()
            tk.Button(d, text="SALVAR", bg=D["success"], fg="white", font=("Segoe UI", 10, "bold"),
                      command=_salvar, pady=8, width=15, relief="flat").pack(pady=18)

        def editar_faculdade():
            D = self._D
            sel = t_fac.selection()
            if not sel: return
            item_id = sel[0]
            nome = str(t_fac.item(item_id, "values")[1]) # Ajustado para índice 1 (pós checkbox)
            d = tk.Toplevel(win); d.title("Editar Faculdade"); d.geometry("400x200")
            d.configure(bg=D["bg"]); d.grab_set()
            tk.Label(d, text="Novo Nome:", bg=D["bg"], fg=D["fg2"],
                     font=("Segoe UI", 10, "bold")).pack(pady=(20, 6))
            ent = tk.Entry(d, font=("Segoe UI", 11), width=40, bg=D["input"], fg=D["fg"], insertbackground=D["fg"], bd=1, relief="solid"); ent.insert(0, nome)
            ent.pack(padx=20, pady=4); ent.focus()
            def _salvar():
                novo = ent.get().strip().upper()
                if not novo:
                    messagebox.showwarning("Atenção", "Nome não pode ser vazio.", parent=d); return
                self.db.editar_faculdade(nome, novo)
                messagebox.showinfo("Sucesso", "Faculdade atualizada!", parent=d)
                d.destroy(); carregar_faculdades(); self.atualizar_dados()
            tk.Button(d, text="SALVAR", bg=D["success"], fg="white", font=("Segoe UI", 10, "bold"),
                      command=_salvar, pady=8, width=15, relief="flat").pack(pady=18)

        def excluir_faculdade():
            sel = t_fac.selection()
            if not sel: return
            item_id = sel[0]
            nome = str(t_fac.item(item_id, "values")[1])
            if messagebox.askyesno("Confirmar", f"Excluir faculdade '{nome}'?\n(Todos os cursos vinculados serão removidos)", parent=win):
                self.db.excluir_faculdade(nome); carregar_faculdades()
                carregar_cursos(); self.atualizar_dados()
                messagebox.showinfo("Sucesso", "Faculdade excluída.", parent=win)

        def excluir_faculdades_em_lote():
            selecionados = []
            for item in t_fac.get_children():
                vals = t_fac.item(item, "values")
                if vals[0] == "☑":
                    selecionados.append(vals[1])
            
            if not selecionados: return
            
            msg = f"Excluir {len(selecionados)} faculdades selecionadas?\n(Cursos e aulas vinculadas também serão removidos!)"
            if messagebox.askyesno("Confirmar Exclusão em Lote", msg, parent=win):
                for f in selecionados:
                    self.db.excluir_faculdade(f)
                carregar_faculdades()
                carregar_cursos()
                self.atualizar_dados()
                messagebox.showinfo("Sucesso", f"{len(selecionados)} faculdades excluídas.", parent=win)

        tk.Button(ft_fac, text="➕  ADICIONAR FACULDADE", bg="#10b981", fg="white",
                  font=("Segoe UI", 9, "bold"), padx=12, pady=6, relief="flat",
                  cursor="hand2", command=adicionar_faculdade).pack(side=tk.LEFT, padx=(12, 4))
        be_fac.config(command=editar_faculdade)
        bd_fac.config(command=excluir_faculdade)
        bdel_fac.config(command=excluir_faculdades_em_lote)
        carregar_faculdades()

        # ═══════════════════════════════════
        # ABA 3 — TURMAS
        # ═══════════════════════════════════
        def _buscar_turmas(texto):
            carregar_turmas(texto)

        f_turm, t_turm, be_turm, bd_turm, ft_turm, e_turm, bdel_turm = _criar_aba(
            nb, "Turmas", "🎓",
            ["Disciplina", "Turma", "Alunos"],
            [320, 200, 120],
            ["w", "w", "center"],
            search_callback=_buscar_turmas
        )

        def carregar_turmas(filtro=""):
            for item in t_turm.get_children(): t_turm.delete(item)
            busca = remover_acentos(filtro.upper()) if filtro else ""
            tpd = self.db._dados.get("turmas_por_disciplina", {})
            apt = self.db._dados.get("alunos_por_turma", {})
            for disc, turmas in sorted(tpd.items()):
                for turma in sorted(turmas):
                    if busca:
                        if busca not in remover_acentos(disc.upper()) and busca not in remover_acentos(turma.upper()):
                            continue
                    qtde = apt.get(f"{disc}|{turma}", 0)
                    t_turm.insert("", tk.END, values=("☐", disc, turma, qtde))

        def adicionar_turma():
            D = self._D
            d = tk.Toplevel(win); d.title("Nova Turma"); d.geometry("420x300")
            d.configure(bg=D["bg"]); d.grab_set()
            tk.Label(d, text="Disciplina:", bg=D["bg"], fg=D["fg2"], font=("Segoe UI", 9, "bold")).pack(pady=(16, 2))
            cb_disc = ttk.Combobox(d, values=self.db.listar_todas_disciplinas(), state="readonly", width=42)
            cb_disc.pack(padx=20, pady=2)
            tk.Label(d, text="Nome da Turma:", bg=D["bg"], fg=D["fg2"], font=("Segoe UI", 9, "bold")).pack(pady=(10, 2))
            ent_t = tk.Entry(d, font=("Segoe UI", 10), width=44, bg=D["input"], fg=D["fg"], insertbackground=D["fg"], bd=1, relief="solid"); ent_t.pack(padx=20, pady=2)
            tk.Label(d, text="Qtde Alunos:", bg=D["bg"], fg=D["fg2"], font=("Segoe UI", 9, "bold")).pack(pady=(10, 2))
            ent_a = tk.Entry(d, font=("Segoe UI", 10), width=16, bg=D["input"], fg=D["fg"], insertbackground=D["fg"], bd=1, relief="solid"); ent_a.insert(0, "0"); ent_a.pack(pady=2)
            def _salvar():
                disc = cb_disc.get().strip().upper()
                nome = ent_t.get().strip().upper()
                if not disc or not nome:
                    messagebox.showwarning("Atenção", "Preencha Disciplina e Turma.", parent=d); return
                try: qtde = int(ent_a.get())
                except: qtde = 0
                self.db.adicionar_turma(disc, nome, qtde)
                d.destroy(); carregar_turmas()
            tk.Button(d, text="SALVAR", bg=D["success"], fg="white", font=("Segoe UI", 10, "bold"),
                      command=_salvar, pady=8, width=15, relief="flat").pack(pady=15)

        def editar_turma():
            D = self._D
            sel = t_turm.selection()
            if not sel: return
            item_id = sel[0]
            vals = t_turm.item(item_id, "values")
            disc, nome, qtde_str = str(vals[1]), str(vals[2]), str(vals[3]) # Índices ajustados
            d = tk.Toplevel(win); d.title(f"Editar Turma: {nome}"); d.geometry("420x280")
            d.configure(bg=D["bg"]); d.grab_set()
            tk.Label(d, text=f"Editando turma da disciplina: {disc}",
                     bg=D["bg"], fg=D["fg"], font=("Segoe UI", 9, "bold")).pack(pady=(16, 6))
            tk.Label(d, text="Novo nome da turma:", bg=D["bg"], fg=D["fg2"]).pack()
            ent_t = tk.Entry(d, font=("Segoe UI", 10), width=44, bg=D["input"], fg=D["fg"], insertbackground=D["fg"], bd=1, relief="solid")
            ent_t.insert(0, nome); ent_t.pack(padx=20, pady=4)
            tk.Label(d, text="Qtde Alunos:", bg=D["bg"], fg=D["fg2"]).pack(pady=(8, 2))
            ent_a = tk.Entry(d, font=("Segoe UI", 10), width=16, bg=D["input"], fg=D["fg"], insertbackground=D["fg"], bd=1, relief="solid")
            ent_a.insert(0, qtde_str); ent_a.pack(pady=4)
            def _salvar():
                novo = ent_t.get().strip().upper()
                if not novo:
                    messagebox.showwarning("Atenção", "Nome não pode ser vazio.", parent=d); return
                try: qtde = int(ent_a.get())
                except: qtde = 0
                self.db.editar_turma(disc, nome, novo, qtde)
                messagebox.showinfo("Sucesso", "Turma atualizada!", parent=d)
                d.destroy(); carregar_turmas(); self.atualizar_dados()
            tk.Button(d, text="SALVAR", bg=D["success"], fg="white", font=("Segoe UI", 10, "bold"),
                      command=_salvar, pady=8, width=15, relief="flat").pack(pady=15)

        def excluir_turma():
            sel = t_turm.selection()
            if not sel: return
            item_id = sel[0]
            vals = t_turm.item(item_id, "values"); disc, nome = str(vals[1]), str(vals[2])
            if messagebox.askyesno("Confirmar", f"Excluir turma '{nome}' da disciplina '{disc}'?", parent=win):
                self.db.excluir_turma(disc, nome); carregar_turmas(); self.atualizar_dados()
                messagebox.showinfo("Sucesso", "Turma excluída.", parent=win)

        def excluir_turmas_em_lote():
            selecionados = []
            for item in t_turm.get_children():
                vals = t_turm.item(item, "values")
                if vals[0] == "☑":
                    selecionados.append((vals[1], vals[2])) # (disciplina, nome)
            
            if not selecionados: return
            
            msg = f"Excluir {len(selecionados)} turmas selecionadas?"
            if messagebox.askyesno("Confirmar Exclusão em Lote", msg, parent=win):
                for disc, nome in selecionados:
                    self.db.excluir_turma(disc, nome)
                carregar_turmas()
                self.atualizar_dados()
                messagebox.showinfo("Sucesso", f"{len(selecionados)} turmas excluídas.", parent=win)

        tk.Button(ft_turm, text="➕  ADICIONAR TURMA", bg="#10b981", fg="white",
                  font=("Segoe UI", 9, "bold"), padx=12, pady=6, relief="flat",
                  cursor="hand2", command=adicionar_turma).pack(side=tk.LEFT, padx=(12, 4))
        be_turm.config(command=editar_turma)
        bd_turm.config(command=excluir_turma)
        bdel_turm.config(command=excluir_turmas_em_lote)
        carregar_turmas()

        # ═══════════════════════════════════
        # ABA 4 — DISCIPLINAS
        # ═══════════════════════════════════
        # ═══════════════════════════════════
        # ABA 4 — DISCIPLINAS
        # ═══════════════════════════════════
        def _buscar_disciplinas(texto):
            carregar_disciplinas(texto)

        f_disc, t_disc, be_disc, bd_disc, ft_disc, e_disc, bdel_disc = _criar_aba(
            nb, "Disciplinas", "📚",
            ["Curso", "Disciplina"],
            [300, 500],
            ["w", "w"],
            search_callback=_buscar_disciplinas
        )

        def carregar_disciplinas(filtro=""):
            for item in t_disc.get_children(): t_disc.delete(item)
            busca = remover_acentos(filtro.upper()) if filtro else ""
            dpc = self.db._dados.get("disciplinas_por_curso", {})
            for curso, discs in sorted(dpc.items()):
                for disc in sorted(discs):
                    if busca:
                        if busca not in remover_acentos(curso.upper()) and busca not in remover_acentos(disc.upper()):
                            continue
                    t_disc.insert("", tk.END, values=("☐", curso, disc))

        def adicionar_disciplina():
            D = self._D
            d = tk.Toplevel(win); d.title("Nova Disciplina"); d.geometry("420x250")
            d.configure(bg=D["bg"]); d.grab_set()
            tk.Label(d, text="Curso:", bg=D["bg"], fg=D["fg2"], font=("Segoe UI", 9, "bold")).pack(pady=(16, 2))
            cursos_todos = [c["nome"] for c in self.db.obter_todos_cursos()]
            cb_cur = ttk.Combobox(d, values=cursos_todos, state="readonly", width=42)
            cb_cur.pack(padx=20, pady=2)
            tk.Label(d, text="Nome da Disciplina:", bg=D["bg"], fg=D["fg2"], font=("Segoe UI", 9, "bold")).pack(pady=(10, 2))
            ent = tk.Entry(d, font=("Segoe UI", 10), width=44, bg=D["input"], fg=D["fg"], insertbackground=D["fg"], bd=1, relief="solid"); ent.pack(padx=20, pady=2); ent.focus()
            def _salvar():
                curso = cb_cur.get().strip().upper()
                nome = ent.get().strip().upper()
                if not curso or not nome:
                    messagebox.showwarning("Atenção", "Preencha Curso e Disciplina.", parent=d); return
                self.db.adicionar_disciplina(curso, nome)
                d.destroy(); carregar_disciplinas()
            tk.Button(d, text="SALVAR", bg=D["success"], fg="white", font=("Segoe UI", 10, "bold"),
                      command=_salvar, pady=8, width=15, relief="flat").pack(pady=18)

        def editar_disciplina():
            D = self._D
            sel = t_disc.selection()
            if not sel: return
            item_id = sel[0]
            vals = t_disc.item(item_id, "values"); curso, nome = str(vals[1]), str(vals[2]) # Índices ajustados
            d = tk.Toplevel(win); d.title(f"Editar: {nome}"); d.geometry("420x220")
            d.configure(bg=D["bg"]); d.grab_set()
            tk.Label(d, text=f"Editando disciplina do curso: {curso}",
                     bg=D["bg"], fg=D["fg"], font=("Segoe UI", 9, "bold")).pack(pady=(16, 6))
            tk.Label(d, text="Novo nome:", bg=D["bg"], fg=D["fg2"]).pack()
            ent = tk.Entry(d, font=("Segoe UI", 10), width=44, bg=D["input"], fg=D["fg"], insertbackground=D["fg"], bd=1, relief="solid")
            ent.insert(0, nome); ent.pack(padx=20, pady=4); ent.focus()
            def _salvar():
                novo = ent.get().strip().upper()
                if not novo:
                    messagebox.showwarning("Atenção", "Nome não pode ser vazio.", parent=d); return
                self.db.editar_disciplina(nome, novo, curso)
                messagebox.showinfo("Sucesso", "Disciplina atualizada!", parent=d)
                d.destroy(); carregar_disciplinas(); carregar_turmas(); self.atualizar_dados()
            tk.Button(d, text="SALVAR", bg=D["success"], fg="white", font=("Segoe UI", 10, "bold"),
                      command=_salvar, pady=8, width=15, relief="flat").pack(pady=18)

        def excluir_disciplina():
            sel = t_disc.selection()
            if not sel: return
            item_id = sel[0]
            vals = t_disc.item(item_id, "values"); curso, nome = str(vals[1]), str(vals[2])
            if messagebox.askyesno("Confirmar", f"Excluir disciplina '{nome}' do curso '{curso}'?\n(Turmas vinculadas também serão removidas)", parent=win):
                self.db.excluir_disciplina(nome, curso)
                carregar_disciplinas(); carregar_turmas(); self.atualizar_dados()
                messagebox.showinfo("Sucesso", "Disciplina excluída.", parent=win)

        def excluir_disciplinas_em_lote():
            selecionados = []
            for item in t_disc.get_children():
                vals = t_disc.item(item, "values")
                if vals[0] == "☑":
                    selecionados.append((vals[1], vals[2])) # (curso, nome)
            
            if not selecionados: return
            
            msg = f"Excluir {len(selecionados)} disciplinas selecionadas?\n(Turmas vinculadas também serão removidas!)"
            if messagebox.askyesno("Confirmar Exclusão em Lote", msg, parent=win):
                for curso, nome in selecionados:
                    self.db.excluir_disciplina(nome, curso)
                carregar_disciplinas()
                carregar_turmas()
                self.atualizar_dados()
                messagebox.showinfo("Sucesso", f"{len(selecionados)} disciplinas excluídas.", parent=win)

        tk.Button(ft_disc, text="➕  ADICIONAR DISCIPLINA", bg=D["success"], fg="white",
                  font=("Segoe UI", 9, "bold"), padx=12, pady=6, relief="flat",
                  cursor="hand2", command=adicionar_disciplina).pack(side=tk.LEFT, padx=(12, 4))
        be_disc.config(command=editar_disciplina)
        bd_disc.config(command=excluir_disciplina)
        bdel_disc.config(command=excluir_disciplinas_em_lote)
        carregar_disciplinas()

        # ═══════════════════════════════════
        # ABA 5 — PROFESSORES
        # ═══════════════════════════════════
        # ═══════════════════════════════════
        # ABA 5 — PROFESSORES
        # ═══════════════════════════════════
        def _buscar_professores(texto):
            carregar_professores(texto)

        f_prof, t_prof, be_prof, bd_prof, ft_prof, e_prof, bdel_prof = _criar_aba(
            nb, "Professores", "👨‍🏫",
            ["Professor"],
            [700],
            ["w"],
            search_callback=_buscar_professores
        )

        def carregar_professores(filtro=""):
            for item in t_prof.get_children(): t_prof.delete(item)
            busca = remover_acentos(filtro.upper()) if filtro else ""
            for p in self.db.listar_professores():
                if busca and busca not in remover_acentos(p.upper()):
                    continue
                t_prof.insert("", tk.END, values=("☐", p))

        def editar_professor():
            D = self._D
            sel = t_prof.selection()
            if not sel: return
            item_id = sel[0]
            nome = str(t_prof.item(item_id, "values")[1]) # Ajustado para índice 1
            d = tk.Toplevel(win); d.title(f"Editar Professor"); d.geometry("400x200")
            d.configure(bg=D["bg"]); d.grab_set()
            tk.Label(d, text=f"Editando: {nome}", bg=D["bg"], fg=D["fg"],
                     font=("Segoe UI", 10, "bold")).pack(pady=(18, 6))
            tk.Label(d, text="Novo nome:", bg=D["bg"], fg=D["fg2"]).pack()
            ent = tk.Entry(d, font=("Segoe UI", 11), width=40, bg=D["input"], fg=D["fg"], insertbackground=D["fg"], bd=1, relief="solid")
            ent.insert(0, nome); ent.pack(padx=20, pady=4); ent.focus()
            def _salvar():
                novo = ent.get().strip().upper()
                if not novo:
                    messagebox.showwarning("Atenção", "Nome não pode ser vazio.", parent=d); return
                self.db.editar_professor(nome, novo)
                messagebox.showinfo("Sucesso", "Professor renomeado em todas as aulas!", parent=d)
                d.destroy(); carregar_professores(); self.atualizar_dados()
            tk.Button(d, text="SALVAR", bg=D["success"], fg="white", font=("Segoe UI", 10, "bold"),
                      command=_salvar, pady=8, width=15, relief="flat").pack(pady=18)

        # Professores são derivados de aulas — não há exclusão direta (ou requer lógica especial)
        # Se o usuário quer o botão de excluir selecionados, eu mostro uma info ou implemento a remoção das aulas.
        def excluir_professores_em_lote():
            selecionados = []
            for item in t_prof.get_children():
                vals = t_prof.item(item, "values")
                if vals[0] == "☑":
                    selecionados.append(vals[1])
            if not selecionados: return
            messagebox.showinfo("Informação", "A exclusão de professores deve ser feita removendo suas respectivas aulas na aba 'Aulas'.", parent=win)

        bd_prof.config(state="disabled")
        tk.Label(ft_prof,
                 text="ℹ️  Professores são gerados a partir das aulas. Editar renomeia em todas as aulas vinculadas.",
                 bg=self._D["bg"], fg=self._D["fg2"], font=("Segoe UI", 8, "italic")).pack(side=tk.LEFT, padx=10)
        be_prof.config(command=editar_professor)
        bdel_prof.config(command=excluir_professores_em_lote)
        carregar_professores()

        # ═══════════════════════════════════
        # ABA 6 — AULAS
        # ═══════════════════════════════════
        def _buscar_aulas(texto):
            carregar_aulas(texto)

        f_aul, t_aul, be_aul, bd_aul, ft_aul, e_aul, bdel_aul = _criar_aba(
            nb, "Aulas", "🗓️",
            ["Lab", "Dia", "Horário", "Disciplina", "Turma", "Professor"],
            [60, 130, 120, 450, 100, 180],
            ["center", "w", "center", "w", "center", "w"],
            search_callback=_buscar_aulas
        )

        def carregar_aulas(filtro=""):
            for item in t_aul.get_children(): t_aul.delete(item)
            busca = remover_acentos(filtro.upper()) if filtro else ""
            for a in self.db.listar_todas_aulas():
                txt = f"{a.laboratorio} {a.dia_semana} {a.hora_inicio} {a.hora_fim} {a.disciplina} {a.turma} {a.professor}".upper()
                if busca and busca not in remover_acentos(txt):
                    continue
                t_aul.insert("", tk.END,
                             values=("☐", a.laboratorio, a.dia_semana,
                                     f"{a.hora_inicio}–{a.hora_fim}",
                                     a.disciplina, a.turma, a.professor or ""),
                             tags=(str(a.id),))

        def editar_aula_sel():
            sel = t_aul.selection()
            if not sel: return
            item_id = sel[0]
            tags = t_aul.item(item_id, "tags")
            if not tags: return
            aula_id = int(tags[0])
            aulas = self.db.listar_todas_aulas()
            aula = next((a for a in aulas if a.id == aula_id), None)
            if aula:
                self._janela_cadastro_aula(aula)
                win.after(300, carregar_aulas)

        def excluir_aula_sel():
            sel = t_aul.selection()
            if not sel: return
            item_id = sel[0]
            tags = t_aul.item(item_id, "tags")
            if not tags: return
            aula_id = int(tags[0])
            if messagebox.askyesno("Confirmar", "Excluir esta aula?", parent=win):
                self.db.apagar_aula(aula_id)
                carregar_aulas(); self.atualizar_dados()

        def excluir_aulas_em_lote():
            selecionados = []
            for item in t_aul.get_children():
                vals = t_aul.item(item, "values")
                if vals[0] == "☑":
                    selecionados.append(int(t_aul.item(item, "tags")[0]))
            
            if not selecionados: return
            
            msg = f"Excluir {len(selecionados)} aulas selecionadas permanentemente?"
            if messagebox.askyesno("Confirmar Exclusão em Lote", msg, parent=win):
                for aid in selecionados:
                    self.db.apagar_aula(aid)
                carregar_aulas()
                self.atualizar_dados()
                messagebox.showinfo("Sucesso", f"{len(selecionados)} aulas excluídas.", parent=win)

        def adicionar_aula_nova():
            self._janela_cadastro_aula()
            win.after(300, carregar_aulas)

        tk.Button(ft_aul, text="➕  ADICIONAR AULA", bg=self._D["success"], fg="white",
                  font=("Segoe UI", 9, "bold"), padx=12, pady=6, relief="flat",
                  cursor="hand2", command=adicionar_aula_nova).pack(side=tk.LEFT, padx=(12, 4))
        be_aul.config(command=editar_aula_sel)
        bd_aul.config(command=excluir_aula_sel)
        bdel_aul.config(command=excluir_aulas_em_lote)
        carregar_aulas()

        # ═══════════════════════════════════
        # ABA 7 — LABORATÓRIOS
        # ═══════════════════════════════════
        def _buscar_labs(texto):
            carregar_labs(texto)

        f_lab, t_lab, be_lab, bd_lab, ft_lab, e_lab, bdel_lab = _criar_aba(
            nb, "Laboratórios", "🔬",
            ["Código do Lab", "Aulas Cadastradas"],
            [220, 480],
            ["center", "center"],
            search_callback=_buscar_labs
        )

        def carregar_labs(filtro=""):
            for item in t_lab.get_children(): t_lab.delete(item)
            busca = remover_acentos(filtro.upper()) if filtro else ""
            todas_aulas = self.db.listar_todas_aulas()
            contagem = {}
            for a in todas_aulas:
                contagem[a.laboratorio] = contagem.get(a.laboratorio, 0) + 1
            for lab in LABORATORIOS:
                if busca and busca not in remover_acentos(lab.upper()):
                    continue
                qtde = contagem.get(lab, 0)
                t_lab.insert("", tk.END, values=("☐", lab, f"{qtde} aula(s)"))

        def excluir_labs_em_lote():
            selecionados = []
            for item in t_lab.get_children():
                vals = t_lab.item(item, "values")
                if vals[0] == "☑":
                    selecionados.append(vals[1])
            if not selecionados: return
            messagebox.showinfo("Informação", "Laboratórios são definidos no arquivo de configuração e não podem ser excluídos por aqui.", parent=win)

        # Laboratórios são configuração estática
        be_lab.config(state="disabled")
        bd_lab.config(state="disabled")
        bdel_lab.config(command=excluir_labs_em_lote)
        tk.Label(ft_lab,
                 text="ℹ️  Laboratórios são definidos pela configuração do sistema (config.py). Para gerenciar aulas de um lab, use a aba Aulas.",
                 bg=self._D["bg"], fg=self._D["fg2"], font=("Segoe UI", 8, "italic"),
                 wraplength=780, justify="left").pack(side=tk.LEFT, padx=10)
        carregar_labs()

        # ═══════════════════════════════════
        # Botão Fechar global
        # ═══════════════════════════════════
        btn_fechar_frame = tk.Frame(win, bg=self._D["hdr"])
        btn_fechar_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        tk.Button(btn_fechar_frame, text="✕  FECHAR", bg=self._D["danger"], fg="white",
                  font=("Segoe UI", 9, "bold"), padx=16, pady=6, relief="flat",
                  cursor="hand2", command=win.destroy).pack(side=tk.RIGHT)

    def _janela_historico(self):
        """Abre a janela de histórico e relatórios."""
        D = self._D
        win = tk.Toplevel(self.root)
        win.title("Histórico e Relatórios")
        win.geometry("1100x750")
        win.configure(bg=D["bg"])
        win.grab_set()

        # Posicionamento central na tela
        win.update_idletasks()
        w = win.winfo_width()
        h = win.winfo_height()
        sw = win.winfo_screenwidth()
        sh = win.winfo_screenheight()
        x = (sw // 2) - (w // 2)
        y = (sh // 2) - (h // 2)
        win.geometry(f"+{x}+{y}")

        # Header
        header = tk.Frame(win, bg=D["hdr"], height=70)
        header.pack(fill=tk.X)
        tk.Label(header, text="📊 Histórico & Relatórios do Sistema", font=("Segoe UI", 18, "bold"),
                 bg=D["hdr"], fg="#89b4fa").pack(side=tk.LEFT, padx=30, pady=15)

        # Notebook para as abas
        nb = ttk.Notebook(win)
        nb.pack(fill=tk.BOTH, expand=True, padx=15, pady=10)

        # --- ABA 1: HISTÓRICO ---
        tab_hist = ttk.Frame(nb)
        nb.add(tab_hist, text="  📜  Histórico  ")

        # Toolbar do histórico
        toolbar = tk.Frame(tab_hist, bg=D["card"], height=60)
        toolbar.pack(fill=tk.X, padx=10, pady=(10, 0))

        def exportar_txt():
            try:
                data_sugestao = datetime.now().strftime("%d-%m-%Y_%H-%M")
                path = filedialog.asksaveasfilename(
                    defaultextension=".txt",
                    filetypes=[("Arquivo de Texto", "*.txt")],
                    title="Exportar Histórico para TXT",
                    initialfile=f"Historico_Aulas_{data_sugestao}.txt",
                    parent=win
                )
                if not path: return
                
                logs = self.db.listar_historico()
                if not logs:
                    messagebox.showwarning("Atenção", "O histórico está vazio.", parent=win)
                    return

                with open(path, "w", encoding="utf-8") as f:
                    f.write("SISTEMA DE AGENDAMENTO DE LABORATÓRIOS\n")
                    f.write("RELATÓRIO DE HISTÓRICO DE MODIFICAÇÕES\n")
                    f.write(f"Gerado em: {datetime.now().strftime('%d/%m/%Y às %H:%M:%S')}\n")
                    f.write("="*70 + "\n\n")
                    for h in logs:
                        f.write(f"[{h.get('timestamp')}]  USUÁRIO: {h.get('usuario', '---')}\n")
                        f.write(f"AÇÃO: {h.get('acao')}\n")
                        f.write(f"DETALHES: {h.get('detalhes')}\n")
                        f.write("-" * 75 + "\n")
                messagebox.showinfo("Sucesso", f"Histórico exportado com sucesso!", parent=win)
            except Exception as ex:
                messagebox.showerror("Erro", f"Erro ao exportar: {str(ex)}", parent=win)

        tk.Button(toolbar, text="📥  EXPORTAR HISTÓRICO (.TXT)", bg=D["primary"], fg="white",
                  font=("Segoe UI", 9, "bold"), relief="flat", padx=15, pady=8,
                  cursor="hand2", command=exportar_txt).pack(side=tk.LEFT, padx=10, pady=10)

        # Container de Filtros (Pesquisa em Tempo Real)
        f_container = tk.Frame(toolbar, bg=D["card"])
        f_container.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=10)

        # 1. Busca Global
        search_card = tk.Frame(f_container, bg=D["input"], highlightthickness=1,
                               highlightbackground=D["border"], highlightcolor=D["primary"])
        search_card.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5, pady=10)

        tk.Label(search_card, text=" 🔍 ", bg=D["input"], fg=D["fg2"],
                 font=("Segoe UI", 11)).pack(side=tk.LEFT, padx=(5, 0))

        ent_busca = tk.Entry(search_card, font=("Segoe UI", 10),
                             bg=D["input"], relief="flat", bd=0,
                             fg=D["fg"], insertbackground=D["fg"])
        ent_busca.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=8, pady=5)
        
        placeholder = "Filtrar histórico..."
        ent_busca.insert(0, placeholder)
        ent_busca.config(fg=D["fg2"])
        
        # 2. Filtro Usuário
        tk.Label(f_container, text="👤 Usuário:", bg=D["card"], fg=D["fg2"], font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(10, 5))
        combo_user = ttk.Combobox(f_container, width=12, state="readonly")
        combo_user.pack(side=tk.LEFT, padx=5)
        
        # 3. Filtro Ação
        tk.Label(f_container, text="⚙️ Ação:", bg=D["card"], fg=D["fg2"], font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(10, 5))
        combo_acao = ttk.Combobox(f_container, width=12, state="readonly")
        combo_acao.pack(side=tk.LEFT, padx=5)

        # 4. Filtro Data
        tk.Label(f_container, text="📅 Data:", bg=D["card"], fg=D["fg2"], font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(10, 5))
        ent_data = tk.Entry(f_container, width=12, font=("Segoe UI", 10), bg=D["input"], fg=D["fg"], 
                            relief="flat", highlightthickness=1, highlightbackground=D["border"],
                            insertbackground=D["fg"])
        ent_data.pack(side=tk.LEFT, padx=5, pady=10)

        # Treeview do Histórico
        tree_frame = tk.Frame(tab_hist, bg=D["bg"])
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        cols = ("Data/Hora", "Usuário", "Ação", "Detalhes")
        tree = ttk.Treeview(tree_frame, columns=cols, show="headings")
        tree.heading("Data/Hora", text="📅 DATA/HORA")
        tree.heading("Usuário",   text="👤 USUÁRIO")
        tree.heading("Ação",      text="⚙️ AÇÃO")
        tree.heading("Detalhes",  text="📝 DETALHES")
        tree.column("Data/Hora", width=150, anchor="center", stretch=False)
        tree.column("Usuário",   width=120, anchor="center", stretch=False)
        tree.column("Ação",      width=100, anchor="center", stretch=False)
        tree.column("Detalhes",  width=1500, anchor="w", stretch=False)
        
        tree.grid(row=0, column=0, sticky="nsew")
        sb_v = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=tree.yview)
        sb_h = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=tree.xview)
        tree.configure(yscrollcommand=sb_v.set, xscrollcommand=sb_h.set)
        sb_v.grid(row=0, column=1, sticky="ns")
        sb_h.grid(row=1, column=0, sticky="ew")
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)

        def carregar_historico(event=None):
            termo = ent_busca.get().lower()
            if termo == placeholder.lower(): termo = ""
            
            f_user = combo_user.get()
            f_acao = combo_acao.get()
            f_data = ent_data.get().lower()
            
            for i in tree.get_children(): tree.delete(i)
            
            logs = self.db.listar_historico()
            for h in logs:
                ts = h.get("timestamp", "")
                us = h.get("usuario", "---")
                ac = h.get("acao", "")
                dt = h.get("detalhes", "")
                
                # Filtros fixos
                if f_user and f_user != "TODOS" and us.upper() != f_user.upper():
                    continue
                if f_acao and f_acao != "TODAS" and ac.upper() != f_acao.upper():
                    continue
                if f_data and f_data not in ts.lower():
                    continue
                
                # Filtro de texto global (Pesquisa tudo)
                if termo:
                    content = f"{ts} {us} {ac} {dt}".lower()
                    if termo not in content:
                        continue
                
                tree.insert("", tk.END, values=(ts, us, ac, dt))

        # Eventos para Tempo Real
        def _on_focus_in(e):
            if ent_busca.get() == placeholder:
                ent_busca.delete(0, tk.END)
                ent_busca.config(fg=D["fg"])

        def _on_focus_out(e):
            if not ent_busca.get():
                ent_busca.insert(0, placeholder)
                ent_busca.config(fg=D["fg2"])

        ent_busca.bind("<FocusIn>", _on_focus_in)
        ent_busca.bind("<FocusOut>", _on_focus_out)
        ent_busca.bind("<KeyRelease>", lambda e: self.debounce("hist", carregar_historico))
        ent_data.bind("<KeyRelease>", lambda e: self.debounce("hist", carregar_historico))
        combo_user.bind("<<ComboboxSelected>>", carregar_historico)
        combo_acao.bind("<<ComboboxSelected>>", carregar_historico)

        # Popula os filtros iniciais baseados nos dados existentes
        historico_completo = self.db.listar_historico()
        usuarios = sorted(list(set(h.get("usuario", "---").upper() for h in historico_completo)))
        acoes = sorted(list(set(h.get("acao", "").upper() for h in historico_completo)))
        
        combo_user['values'] = ["TODOS"] + usuarios
        combo_user.set("TODOS")
        combo_acao['values'] = ["TODAS"] + acoes
        combo_acao.set("TODAS")

        carregar_historico()

        # --- ABA 2: RELATÓRIOS ---
        tab_rel = ttk.Frame(nb)
        nb.add(tab_rel, text="  📊  Relatórios  ")
        
        rel_container = tk.Frame(tab_rel, bg=D["bg"], padx=40, pady=40)
        rel_container.pack(fill=tk.BOTH, expand=True)

        # Título da aba
        tk.Label(rel_container, text="Resumo Estatístico de Aulas", font=("Segoe UI", 16, "bold"),
                 bg=D["bg"], fg=D["fg"]).pack(pady=(0, 30), anchor="w")

        # Grid de cards
        cards_frame = tk.Frame(rel_container, bg=D["bg"])
        cards_frame.pack(fill=tk.X)

        stats = self.db.obter_estatisticas_aulas()

        def _criar_card(parent, title, value, color, row, col):
            f = tk.Frame(parent, bg=D["card"], highlightthickness=1, highlightbackground=D["border"], padx=20, pady=20)
            f.grid(row=row, column=col, padx=10, pady=10, sticky="nsew")
            tk.Label(f, text=title, font=("Segoe UI", 10, "bold"), bg=D["card"], fg=D["fg2"]).pack(anchor="w")
            val_text = str(value) if value is not None else "Sem dados"
            tk.Label(f, text=val_text, font=("Segoe UI", 24, "bold"), bg=D["card"], fg=color).pack(anchor="w", pady=(5, 0))
            return f

        cards_frame.columnconfigure((0, 1, 2), weight=1)

        _criar_card(cards_frame, "Aulas Semestrais", stats["semestral"], "#3d5af1", 0, 0)
        _criar_card(cards_frame, "Aulas Eventuais",  stats["eventual"],  "#10b981", 0, 1)
        _criar_card(cards_frame, "Aulas de Reposição", stats["reposicao"], "#f59e0b", 0, 2)
        
        _criar_card(cards_frame, "Aulas Canceladas", stats["cancelada"], "#ef4444", 1, 0)
        _criar_card(cards_frame, "Aulas Indeferidas", stats["indeferida"], "#94a3b8", 1, 1)

        # Espaçador
        tk.Frame(rel_container, bg=D["bg"], height=40).pack()
        
        # Informativo
        info_box = tk.Label(rel_container, text="💡 Os dados acima refletem o estado atual dos agendamentos no sistema.",
                           font=("Segoe UI", 9, "italic"), bg=D["bg"], fg=D["fg2"], justify="left")
        info_box.pack(anchor="w")
        
        # Botão Exportar Relatório para Excel
        tk.Button(rel_container, text="📊  EXPORTAR PARA EXCEL (.XLSX)", bg="#10b981", fg="white",
                  font=("Segoe UI", 10, "bold"), relief="flat", padx=20, pady=10,
                  cursor="hand2", command=self.exportar_relatorio_excel).pack(pady=20, anchor="w")

        # Botão Fechar no Rodapé
        footer = tk.Frame(win, bg=D["hdr"], height=60)
        footer.pack(fill=tk.X, side=tk.BOTTOM)
        tk.Button(footer, text="✕  FECHAR", bg=D["danger"], fg="white",
                  font=("Segoe UI", 9, "bold"), padx=20, pady=8, relief="flat",
                  cursor="hand2", command=win.destroy).pack(side=tk.RIGHT, padx=20, pady=10)


    def _resolve_image_path(self, path):
        """Resolve caminhos relativos de imagens usando o diretório da aplicação."""
        if os.path.isabs(path):
            return path
        base_dir = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(base_dir, path)

    def _renderizar_imagem_no_texto(self, txt_widget, path, pos, w=None, h=None):
        """Auxiliar para inserir imagem no widget Text do inventário/preview."""
        try:
            from PIL import Image, ImageTk
            resolved = self._resolve_image_path(path)
            if not os.path.exists(resolved): return
            img = Image.open(resolved).convert("RGBA")
            if w is None or h is None:
                max_w = 400
                if img.width > max_w:
                    h = int(img.height * (max_w / img.width))
                    w = max_w
                else:
                    w, h = img.width, img.height
            
            img_resized = img.resize((int(w), int(h)), Image.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(img_resized)
            
            img_id = f"img_{len(txt_widget.editor_images)}"
            txt_widget.editor_images.append({
                "id": img_id, "photo": photo, "path": path, "orig_pil": img,
                "orig_w": img.width, "orig_h": img.height, "cur_w": w, "cur_h": h
            })
            txt_widget.image_create(pos, image=photo, name=img_id)
        except Exception: pass

    def _render_rich_text(self, txt_widget, inventario):
        """Renderiza o conteúdo rich text no widget de texto fornecido."""
        if not inventario: return
        txt_widget.config(state=tk.NORMAL)
        txt_widget.delete("1.0", tk.END)
        if not hasattr(txt_widget, "editor_images"):
            txt_widget.editor_images = []
        else:
            txt_widget.editor_images.clear()
            
        outros_texto = inventario.get("outros_texto", "")
        outros_rich = inventario.get("outros_rich", {})
        
        txt_widget.insert("1.0", outros_texto)
        
        if outros_rich:
            # Carregar Tags
            if "tags" in outros_rich:
                for tag_name, ranges in outros_rich["tags"].items():
                    if tag_name.startswith("color_"):
                        txt_widget.tag_configure(tag_name, foreground=f"#{tag_name[6:]}")
                    for r in ranges:
                        try: txt_widget.tag_add(tag_name, r[0], r[1])
                        except: pass
            
            # Carregar Imagens
            if "images" in outros_rich:
                for img_info in outros_rich["images"]:
                    try:
                        self._renderizar_imagem_no_texto(txt_widget, img_info["path"], img_info["pos"], 
                                                        img_info.get("w"), img_info.get("h"))
                    except: pass
        txt_widget.config(state=tk.DISABLED)

    # ------------------------------------------------------------------

    def _janela_laboratorios(self):
        """Janela principal de listagem e gestão de laboratórios."""
        win = tk.Toplevel(self.root)
        win.title("Gestão de Laboratórios")
        win.geometry("1500x950")
        win.configure(bg=self._D["bg"])
        win.transient(self.root)
        win.grab_set()

        # Layout: Esquerda (Lista) | Direita (Preview Planta)
        main_paned = tk.PanedWindow(win, orient=tk.HORIZONTAL, bg=self._D["border"], sashwidth=4)
        main_paned.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # --- Lado Esquerdo: Lista ---
        left_frame = tk.Frame(main_paned, bg=self._D["bg"])
        main_paned.add(left_frame, width=600)

        # Toolbar da lista
        list_tool = tk.Frame(left_frame, bg=self._D["bg"])
        list_tool.pack(fill=tk.X, pady=(0, 10))
        
        tk.Label(list_tool, text="Laboratórios Cadastrados", font=("Segoe UI", 12, "bold"),
                 bg=self._D["bg"], fg=self._D["primary"]).pack(side=tk.LEFT)

        btn_add = tk.Button(list_tool, text="➕ Inserir", bg=self._D["success"], fg="white",
                            font=("Segoe UI", 9, "bold"), padx=10, pady=5, relief="flat",
                            command=lambda: self._form_laboratorio(callback=refresh_list))
        btn_add.pack(side=tk.RIGHT, padx=5)

        # Treeview
        cols = ("Nome", "Descrição", "Qtd Micros")
        t_lab = ttk.Treeview(left_frame, columns=cols, show="headings", height=15)
        t_lab.heading("Nome", text="NOME")
        t_lab.heading("Descrição", text="DESCRIÇÃO")
        t_lab.heading("Qtd Micros", text="QTD MICROS")
        t_lab.column("Nome", width=120, anchor="w")
        t_lab.column("Descrição", width=380, anchor="w")
        t_lab.column("Qtd Micros", width=100, anchor="center")
        t_lab.pack(fill=tk.BOTH, expand=True)

        # --- Lado Direito: Preview (Planta + Outros) ---
        right_frame = tk.Frame(main_paned, bg=self._D["card"], bd=1, relief="solid")
        main_paned.add(right_frame)

        # PanedWindow vertical para dividir Planta e Outros
        preview_paned = tk.PanedWindow(right_frame, orient=tk.VERTICAL, bg=self._D["border"], sashwidth=4)
        preview_paned.pack(fill=tk.BOTH, expand=True)

        # 1. Planta do Lab
        planta_container = tk.Frame(preview_paned, bg=self._D["card"])
        preview_paned.add(planta_container, height=350)
        
        tk.Label(planta_container, text="Planta do Lab", font=("Segoe UI", 10, "bold"),
                 bg=self._D["card"], fg=self._D["fg2"]).pack(pady=2)
        
        lbl_planta = tk.Label(planta_container, text="Selecione um laboratório\npara ver a planta", 
                              bg=self._D["input"], fg=self._D["fg2"], cursor="hand2")
        lbl_planta.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        # 2. Outros (Preview do Inventário)
        outros_container = tk.Frame(preview_paned, bg=self._D["card"])
        preview_paned.add(outros_container)

        tk.Label(outros_container, text="Outros (Extra)", font=("Segoe UI", 10, "bold"),
                 bg=self._D["card"], fg=self._D["fg2"]).pack(pady=2)
        
        f_txt = tk.Frame(outros_container, bg="white")
        f_txt.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        
        txt_preview = tk.Text(f_txt, bg="white", fg="black", font=("Segoe UI", 10), 
                              wrap=tk.WORD, padx=10, pady=10, cursor="hand2")
        txt_preview.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb_preview = ttk.Scrollbar(f_txt, orient=tk.VERTICAL, command=txt_preview.yview)
        txt_preview.configure(yscrollcommand=sb_preview.set)
        sb_preview.pack(side=tk.RIGHT, fill=tk.Y)
        txt_preview.config(state=tk.DISABLED)

        # Rodapé com ações
        footer = tk.Frame(left_frame, bg=self._D["bg"])
        footer.pack(fill=tk.X, pady=10)

        btn_edit = tk.Button(footer, text="✏️ Editar", bg=self._D["warning"], fg="white",
                             font=("Segoe UI", 9, "bold"), padx=12, pady=8, relief="flat",
                             command=lambda: editar_selecionado())
        btn_edit.pack(side=tk.LEFT, padx=5)

        def apagar_selecionado():
            sel = t_lab.selection()
            if not sel: return
            item_id = sel[0]
            lab_id = int(t_lab.item(item_id, "tags")[0])
            lab_nome = t_lab.item(item_id, "values")[0]
            if messagebox.askyesno("Confirmar", f"Deseja realmente excluir o laboratório '{lab_nome}'?", parent=win):
                self.db.apagar_laboratorio(lab_id)
                messagebox.showinfo("Sucesso", "Laboratório excluído com sucesso!", parent=win)
                refresh_list()

        btn_del = tk.Button(footer, text="🗑️ Excluir", bg=self._D["danger"], fg="white",
                            font=("Segoe UI", 9, "bold"), padx=12, pady=8, relief="flat",
                            command=apagar_selecionado)
        # btn_del.pack initially hidden if requested, but user said "when selecting", so for now not packed.

        def refresh_list():
            for item in t_lab.get_children(): t_lab.delete(item)
            labs = self.db.listar_laboratorios()
            for l in labs:
                t_lab.insert("", tk.END, values=(l.nome, l.descricao, l.qtd_micros), tags=(str(l.id),))
            btn_del.pack_forget() # Hide delete button on refresh

        def on_select(event):
            sel = t_lab.selection()
            if not sel:
                btn_del.pack_forget()
                return
            
            btn_del.pack(side=tk.LEFT, padx=5) # Show delete button on selection
            
            item_id = sel[0]
            tags = t_lab.item(item_id, "tags")
            if not tags: return
            lab_id = int(tags[0])
            labs = self.db.listar_laboratorios()
            lab = next((l for l in labs if l.id == lab_id), None)
            
            # Funções para clique nos previews
            def abrir_planta(e):
                if lab: self._janela_inventario_laboratorio(lab, aba_inicial=3, callback=refresh_list)
            def abrir_outros(e):
                if lab: self._janela_inventario_laboratorio(lab, aba_inicial=4, callback=refresh_list)
            
            lbl_planta.bind("<Button-1>", abrir_planta)
            txt_preview.bind("<Button-1>", abrir_outros)

            if lab:
                if lab.planta_path and os.path.exists(lab.planta_path):
                    self._exibir_imagem_no_label(lbl_planta, lab.planta_path)
                else:
                    lbl_planta.config(image="", text="Planta não disponível")
                
                self._render_rich_text(txt_preview, lab.inventario if lab.inventario else {})
            else:
                lbl_planta.config(image="", text="Selecione um laboratório\npara ver a planta")
                txt_preview.config(state=tk.NORMAL)
                txt_preview.delete("1.0", tk.END)
                txt_preview.config(state=tk.DISABLED)

        def editar_selecionado():
            sel = t_lab.selection()
            if not sel:
                messagebox.showwarning("Aviso", "Selecione um laboratório.", parent=win)
                return
            lab_id = int(t_lab.item(sel[0], "tags")[0])
            lab = next((l for l in self.db.listar_laboratorios() if l.id == lab_id), None)
            if lab:
                self._form_laboratorio(lab, callback=refresh_list)

        t_lab.bind("<<TreeviewSelect>>", on_select)
        
        def on_double_click(event):
            region = t_lab.identify_region(event.x, event.y)
            if region != "cell": 
                # Se não for na célula, abre inventário (comportamento antigo)
                ver_inventario()
                return
            
            column = t_lab.identify_column(event.x) # e.g. #1, #2, #3
            item = t_lab.identify_row(event.y)
            
            # Coluna 1=Nome, 2=Descrição, 3=Qtd Micros
            if column == "#1":
                ver_inventario()
                return
            
            col_idx = int(column[1:]) - 1 # 1 or 2
            col_name = cols[col_idx]
            old_val = t_lab.item(item, "values")[col_idx]
            
            # Criar editor flutuante
            x, y, w, h = t_lab.bbox(item, column)
            entry = ttk.Entry(t_lab)
            entry.insert(0, str(old_val))
            entry.place(x=x, y=y, width=w, height=h)
            entry.focus_set()
            
            def save_edit(event=None):
                new_val = entry.get()
                lab_id = int(t_lab.item(item, "tags")[0])
                lab = next((l for l in self.db.listar_laboratorios() if l.id == lab_id), None)
                if lab:
                    if col_name == "Descrição":
                        lab.descricao = new_val
                    elif col_name == "Qtd Micros":
                        try: lab.qtd_micros = int(new_val)
                        except: pass
                    self.db.alterar_laboratorio(lab)
                    # Atualiza vista
                    vals = list(t_lab.item(item, "values"))
                    vals[col_idx] = lab.descricao if col_name == "Descrição" else lab.qtd_micros
                    t_lab.item(item, values=vals)
                entry.destroy()

            entry.bind("<Return>", save_edit)
            entry.bind("<FocusOut>", lambda e: entry.destroy())
            entry.bind("<Escape>", lambda e: entry.destroy())

        t_lab.bind("<Double-1>", on_double_click)
        
        def ver_inventario():
            sel = t_lab.selection()
            if not sel: return
            lab_id = int(t_lab.item(sel[0], "tags")[0])
            lab = next((l for l in self.db.listar_laboratorios() if l.id == lab_id), None)
            if lab:
                self._janela_inventario_laboratorio(lab, callback=refresh_list)

        refresh_list()

    def _criar_planilha(self, parent, data_key, lab):
        """
        Substitui o comportamento das abas 'micros' e 'ac' para abrirem os scripts externos.
        Para qualquer outra aba, mantém o funcionamento original (gradiente tipo Excel).
        """
        # Criar pasta de inventários se não existir
        os.makedirs("inventarios", exist_ok=True)
        
        # Gerar nome de arquivo único para este lab e aba (para os novos widgets)
        lab_slug = str(lab.id) if lab.id else lab.nome.replace(" ", "_").lower()
        save_path = os.path.join("inventarios", f"inventario_{lab_slug}_{data_key}.json")

        if data_key == "micros":
            model = micros.SheetModel()
            sheet = micros.SpreadsheetWidget(parent, model, save_path, use_menubar=False)
            sheet.pack(fill="both", expand=True)
            sheet.file_mgr.load_json()
            return lambda: sheet.file_mgr.save()
            
        elif data_key == "ac":
            model = ar_condicionado.SheetModel()
            sheet = ar_condicionado.SpreadsheetWidget(parent, model, save_path, use_menubar=False)
            sheet.pack(fill="both", expand=True)
            sheet.file_mgr.load_json()
            return lambda: sheet.file_mgr.save()
            
        # --- COMPORTAMENTO ORIGINAL PARA OUTRAS ABAS ---
        container = tk.Frame(parent, bg=self._D["bg"])
        container.pack(fill=tk.BOTH, expand=True)

        # Toolbar
        toolbar = tk.Frame(container, bg=self._D["bg"])
        toolbar.pack(fill=tk.X, pady=5, padx=5)

        # Dados iniciais
        inv_data = lab.inventario.get(data_key, {
            "headers": ["Coluna 1", "Coluna 2", "Coluna 3", "Coluna 4"],
            "rows": [["", "", "", ""]],
            "col_widths": [80, 80, 80, 80],
            "row_heights": [22]
        })
        
        headers = inv_data["headers"]
        rows_data = inv_data["rows"]
        col_widths = inv_data.get("col_widths", [80] * len(headers))
        row_heights = [25] * len(rows_data)

        if len(col_widths) < len(headers): col_widths.extend([80] * (len(headers) - len(col_widths)))
        
        sel_r = [None]; sel_c = [None]; resizing_col = [None]; resizing_row = [None] 

        canvas = tk.Canvas(container, bg=self._D["bg"], highlightthickness=0)
        v_scroll = ttk.Scrollbar(container, orient=tk.VERTICAL, command=canvas.yview)
        h_scroll = ttk.Scrollbar(container, orient=tk.HORIZONTAL, command=canvas.xview)
        table_frame = tk.Frame(canvas, bg=self._D["border"])

        canvas.create_window((0, 0), window=table_frame, anchor="nw")
        canvas.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        v_scroll.pack(side=tk.RIGHT, fill=tk.Y); h_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        table_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        entries = []; header_entries = []; row_handles = []

        def clear_selection():
            sel_r[0] = None; sel_c[0] = None
            for r_idx, row in enumerate(entries):
                for c_idx, txt in enumerate(row): txt.config(bg="white")
            for e in header_entries: e.config(bg=self._D["hdr"])
            for lbl in row_handles: lbl.config(bg=self._D["hdr"])

        def select_row(r_idx):
            clear_selection(); sel_r[0] = r_idx; row_handles[r_idx].config(bg=self._D["primary"])
            for txt in entries[r_idx]: txt.config(bg="#e2e8f0")

        def select_col(c_idx):
            clear_selection(); sel_c[0] = c_idx; header_entries[c_idx].config(bg=self._D["primary"])
            for r_idx in range(len(rows_data)): entries[r_idx][c_idx].config(bg="#e2e8f0")

        def adjust_row_height(r_idx):
            max_lines = 1
            for txt in entries[r_idx]:
                d_lines = txt.count("1.0", "end", "displaylines")
                if d_lines: max_lines = max(max_lines, d_lines[0])
            content_h = max(22, max_lines * 18 + 5)
            row_heights[r_idx] = content_h
            table_frame.rowconfigure(r_idx+1, minsize=content_h)

        def start_col_resize(event, idx):
            if event.x > event.widget.winfo_width() - 8:
                resizing_col[0] = (idx, event.x_root, col_widths[idx])
                return "break"

        def do_col_resize(event):
            if resizing_col[0]:
                idx, start_x, start_w = resizing_col[0]
                new_w = max(40, start_w + (event.x_root - start_x))
                col_widths[idx] = new_w
                table_frame.columnconfigure(idx+1, minsize=new_w); return "break"

        def stop_col_resize(event): 
            resizing_col[0] = None
            for r in range(len(rows_data)): adjust_row_height(r)

        def start_row_resize(event, idx):
            if event.y > event.widget.winfo_height() - 8:
                resizing_row[0] = (idx, event.y_root, row_heights[idx]); return "break"

        def do_row_resize(event):
            if resizing_row[0]:
                idx, start_y, start_h = resizing_row[0]
                new_h = max(15, start_h + (event.y_root - start_y))
                row_heights[idx] = new_h; adjust_row_height(idx); return "break"

        def delete_item(type, idx):
            save_temp_data()
            if type == "row":
                if len(rows_data) > 1: rows_data.pop(idx); row_heights.pop(idx)
                else: rows_data[0] = [""] * len(headers)
            else:
                if len(headers) > 1:
                    headers.pop(idx); col_widths.pop(idx)
                    for r in rows_data: r.pop(idx)
                else:
                    headers[0] = "Coluna 1"
                    for r in rows_data: r[0] = ""
            render_table()

        def save_temp_data():
            headers.clear()
            for e in header_entries: headers.append(e.get())
            rows_data.clear()
            for row_group in entries: rows_data.append([t.get("1.0", "end-1c") for t in row_group])

        def render_table():
            for child in table_frame.winfo_children(): child.destroy()
            entries.clear(); header_entries.clear(); row_handles.clear()
            tk.Label(table_frame, bg=self._D["hdr"], width=4).grid(row=0, column=0, sticky="nsew", padx=1, pady=1)
            for c, h in enumerate(headers):
                f_h = tk.Frame(table_frame, bg=self._D["hdr"])
                f_h.grid(row=0, column=c+1, sticky="nsew", padx=1, pady=1)
                e = tk.Entry(f_h, bg=self._D["hdr"], fg="white", font=("Segoe UI", 9, "bold"), justify="center", bd=0, width=1)
                e.insert(0, h); e.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
                for w in (e, f_h):
                    w.bind("<Button-1>", lambda ev, idx=c: select_col(idx), add="+")
                    w.bind("<ButtonPress-1>", lambda ev, idx=c: start_col_resize(ev, idx), add="+")
                    w.bind("<B1-Motion>", do_col_resize, add="+")
                    w.bind("<ButtonRelease-1>", stop_col_resize, add="+")
                header_entries.append(e); table_frame.columnconfigure(c+1, minsize=col_widths[c])
            for r, row in enumerate(rows_data):
                lbl_h = tk.Label(table_frame, text=str(r+1), bg=self._D["hdr"], fg="white", font=("Segoe UI", 9, "bold"), width=4)
                lbl_h.grid(row=r+1, column=0, sticky="nsew", padx=1, pady=1)
                lbl_h.bind("<Button-1>", lambda ev, idx=r: select_row(idx))
                lbl_h.bind("<ButtonPress-1>", lambda ev, idx=r: start_row_resize(ev, idx), add="+")
                lbl_h.bind("<B1-Motion>", do_row_resize)
                def stop_row(e): resizing_row[0] = None
                lbl_h.bind("<ButtonRelease-1>", stop_row)
                row_handles.append(lbl_h)
                
                row_entries = []
                for c, val in enumerate(row):
                    txt = tk.Text(table_frame, bg="white", fg="black", bd=0, highlightthickness=1, 
                                  highlightcolor=self._D["primary"], highlightbackground="#cbd5e1",
                                  font=("Segoe UI", 9), wrap=tk.WORD, undo=True, width=1, height=1)
                    txt.insert("1.0", val); txt.grid(row=r+1, column=c+1, sticky="nsew", padx=1, pady=1)
                    
                    def on_mod(ev, r_idx=r, w=txt):
                        adjust_row_height(r_idx)
                        w.edit_modified(False)
                    
                    txt.bind("<<Modified>>", on_mod)
                    txt.bind("<FocusIn>", lambda ev: clear_selection())
                    row_entries.append(txt)
                entries.append(row_entries)
                table_frame.rowconfigure(r+1, minsize=row_heights[r])
            table_frame.columnconfigure(0, weight=0)

        def add_row():
            save_temp_data(); rows_data.append([""] * len(headers)); row_heights.append(25); render_table()
        def add_col():
            save_temp_data(); headers.append(f"Coluna {len(headers)+1}"); col_widths.append(100)
            for r in rows_data: r.append(""); render_table()
        def save_table_data():
            save_temp_data()
            lab.inventario[data_key] = {"headers": headers, "rows": rows_data, "col_widths": col_widths, "row_heights": row_heights}

        tk.Button(toolbar, text="➕ Nova Linha", bg=self._D["success"], fg="white", font=("Segoe UI", 9, "bold"), relief="flat", padx=10, command=add_row).pack(side=tk.LEFT, padx=5)
        tk.Button(toolbar, text="➕ Nova Coluna", bg=self._D["primary"], fg="white", font=("Segoe UI", 9, "bold"), relief="flat", padx=10, command=add_col).pack(side=tk.LEFT, padx=5)
        render_table()
        return save_table_data



    def _janela_inventario_laboratorio(self, lab, callback=None, aba_inicial=0):
        """Interface para gerenciar inventário/hardware do lab (Tabs: Projetor, Micros, AC, etc)."""
        win = tk.Toplevel(self.root)
        win.title(f"Inventário: {lab.nome}")
        win.geometry("800x600")
        win.configure(bg=self._D["bg"])
        win.grab_set()
        
        if lab.inventario is None: lab.inventario = {}
        
        nb = ttk.Notebook(win)
        nb.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # --- ABA PROJETOR ---
        tab_proj = tk.Frame(nb, bg=self._D["bg"], padx=20, pady=20)
        nb.add(tab_proj, text="Projetor")
        
        proj_data = lab.inventario.get("projetor", {"modelo": "", "patrimonio": "", "tempo_lampada": "", "outros": ""})
        
        def _field(p, label, key, val):
            f = tk.Frame(p, bg=self._D["bg"])
            f.pack(fill=tk.X, pady=5)
            tk.Label(f, text=label, width=15, anchor="w", bg=self._D["bg"], fg=self._D["fg"]).pack(side=tk.LEFT)
            e = ttk.Entry(f)
            e.insert(0, val)
            e.pack(side=tk.LEFT, fill=tk.X, expand=True)
            return e

        e_mod = _field(tab_proj, "Modelo:", "modelo", proj_data.get("modelo", ""))
        e_pat = _field(tab_proj, "Patrimônio:", "patrimonio", proj_data.get("patrimonio", ""))
        e_tmp = _field(tab_proj, "Tempo Lâmpada:", "tempo_lampada", proj_data.get("tempo_lampada", ""))
        e_out = _field(tab_proj, "Outros:", "outros", proj_data.get("outros", ""))

        # --- ABA MICROS ---
        tab_micros = tk.Frame(nb, bg=self._D["bg"])
        nb.add(tab_micros, text="Micros")
        save_micros = self._criar_planilha(tab_micros, "micros", lab)

        # --- ABA AR-CONDICIONADO ---
        tab_ac = tk.Frame(nb, bg=self._D["bg"])
        nb.add(tab_ac, text="Ar-Condicionado")
        save_ac = self._criar_planilha(tab_ac, "ac", lab)

        # --- ABA PLANTA ---
        tab_planta = tk.Frame(nb, bg=self._D["bg"], padx=10, pady=10)
        nb.add(tab_planta, text="Planta do Lab")
        lbl_p = tk.Label(tab_planta, bg=self._D["input"])
        lbl_p.pack(fill=tk.BOTH, expand=True)

        def resize_planta(event=None):
            if not lab.planta_path: return
            # Usa as dimensões atuais do label
            w = max(100, lbl_p.winfo_width())
            h = max(100, lbl_p.winfo_height())
            self._exibir_imagem_no_label(lbl_p, lab.planta_path, size=(w, h))

        if lab.planta_path: 
            # Chama após um pequeno delay para garantir que o layout inicial esteja pronto
            win.after(100, resize_planta)
            lbl_p.bind("<Configure>", lambda e: resize_planta())
        else: 
            lbl_p.config(text="Planta não disponível")

        # --- ABA OUTROS ---
        tab_outros = tk.Frame(nb, bg=self._D["bg"], padx=10, pady=10)
        nb.add(tab_outros, text="Outros")
        
        # Toolbar do Editor (WordPad-like)
        edit_toolbar = tk.Frame(tab_outros, bg=self._D["hdr"], pady=2)
        edit_toolbar.pack(fill=tk.X)

        from tkinter import font as tkfont
        
        def toggle_style(style):
            try:
                sel = txt_outros.tag_ranges("sel")
                if not sel: return
                
                # Verifica se o estilo já está aplicado em toda a seleção
                tags = txt_outros.tag_names(sel[0])
                if style in tags:
                    txt_outros.tag_remove(style, sel[0], sel[1])
                else:
                    txt_outros.tag_add(style, sel[0], sel[1])
            except tk.TclError: pass

        def set_alignment(align):
            try:
                sel = txt_outros.tag_ranges("sel")
                start = sel[0] if sel else "1.0"
                end = sel[1] if sel else tk.END
                # Remove alinhamentos anteriores
                for tag in ["left", "center", "right"]:
                    txt_outros.tag_remove(tag, start, end)
                txt_outros.tag_add(align, start, end)
            except: pass

        def choose_color():
            color = colorchooser.askcolor(title="Escolha a cor do texto")[1]
            if color:
                tag_name = f"color_{color.replace('#','')}"
                txt_outros.tag_configure(tag_name, foreground=color)
                try:
                    sel = txt_outros.tag_ranges("sel")
                    if sel: txt_outros.tag_add(tag_name, sel[0], sel[1])
                except: pass

        def _btn(p, text, cmd, w=3):
            b = tk.Button(p, text=text, command=cmd, font=("Segoe UI", 9, "bold"), 
                          bg=self._D["hdr"], fg="white", relief="flat", width=w)
            b.pack(side=tk.LEFT, padx=2)
            b.bind("<Enter>", lambda e: b.config(bg=self._D["primary"]))
            b.bind("<Leave>", lambda e: b.config(bg=self._D["hdr"]))
            return b

        _btn(edit_toolbar, "B", lambda: toggle_style("bold"))
        _btn(edit_toolbar, "I", lambda: toggle_style("italic"))
        _btn(edit_toolbar, "U", lambda: toggle_style("underline"))
        tk.Frame(edit_toolbar, width=10, bg=self._D["hdr"]).pack(side=tk.LEFT)
        _btn(edit_toolbar, "⬅️", lambda: set_alignment("left"))
        _btn(edit_toolbar, "↔️", lambda: set_alignment("center"))
        _btn(edit_toolbar, "➡️", lambda: set_alignment("right"))
        tk.Frame(edit_toolbar, width=10, bg=self._D["hdr"]).pack(side=tk.LEFT)
        _btn(edit_toolbar, "🎨", choose_color)
        tk.Frame(edit_toolbar, width=10, bg=self._D["hdr"]).pack(side=tk.LEFT)
        
        from PIL import Image, ImageTk
        txt_outros = tk.Text(tab_outros, bg="white", fg="black", height=15, 
                              font=("Segoe UI", 11), wrap=tk.WORD, undo=True, padx=10, pady=10)
        txt_outros.pack(fill=tk.BOTH, expand=True)
        txt_outros.editor_images = [] # Lista para manter referências das imagens Tkinter
        
        def insert_image(path, pos=tk.INSERT, w=None, h=None):
            self._renderizar_imagem_no_texto(txt_outros, path, pos, w, h)
            if pos == tk.INSERT: txt_outros.insert(pos, "\n")

        def choose_image():
            path = filedialog.askopenfilename(filetypes=[("Imagens", "*.png *.jpg *.jpeg *.gif *.bmp")])
            if path:
                # Copia a imagem para a pasta uploads/ compartilhada
                import uuid, shutil
                base_dir = os.path.dirname(os.path.abspath(__file__))
                uploads_dir = os.path.join(base_dir, "uploads")
                os.makedirs(uploads_dir, exist_ok=True)
                ext = os.path.splitext(path)[1]
                unique_name = f"img_{uuid.uuid4().hex[:12]}{ext}"
                dest = os.path.join(uploads_dir, unique_name)
                try:
                    shutil.copy2(path, dest)
                    # Usa caminho relativo para compartilhamento
                    rel_path = os.path.join("uploads", unique_name)
                    insert_image(rel_path)
                except Exception as e:
                    messagebox.showerror("Erro", f"Não foi possível copiar imagem: {e}")
                    insert_image(path)  # Fallback: usa caminho original

        _btn(edit_toolbar, "📷", choose_image)

        # Lógica de Redimensionamento Interativo
        resizing_data = {"active": False, "img_id": None, "start_x": 0, "start_w": 0}

        def on_text_motion(event):
            if resizing_data["active"]:
                dx = event.x - resizing_data["start_x"]
                new_w = max(50, resizing_data["start_w"] + dx)
                
                for img_info in txt_outros.editor_images:
                    if img_info["id"] == resizing_data["img_id"]:
                        # Mantém proporção
                        ratio = img_info["orig_h"] / img_info["orig_w"]
                        new_h = int(new_w * ratio)
                        
                        img_resized = img_info["orig_pil"].resize((int(new_w), new_h), Image.Resampling.LANCZOS)
                        new_photo = ImageTk.PhotoImage(img_resized)
                        
                        img_info["photo"] = new_photo # Atualiza referência para evitar GC
                        img_info["cur_w"] = new_w
                        img_info["cur_h"] = new_h
                        txt_outros.image_configure(img_info["id"], image=new_photo)
                        break
                return

            # Muda cursor se estiver nas bordas (Direita ou Inferior) da imagem
            pos = f"@{event.x},{event.y}"
            try:
                name = txt_outros.image_cget(pos, "name")
                if name:
                    bbox = txt_outros.bbox(pos)
                    if bbox:
                        # Margem de detecção de 10 pixels nas bordas
                        margin = 10
                        is_right = event.x > bbox[0] + bbox[2] - margin
                        is_bottom = event.y > bbox[1] + bbox[3] - margin
                        
                        if is_right or is_bottom:
                            txt_outros.config(cursor="size_nw_se")
                            return
            except: pass
            txt_outros.config(cursor="")

        def on_text_press(event):
            pos = f"@{event.x},{event.y}"
            try:
                name = txt_outros.image_cget(pos, "name")
                if name:
                    bbox = txt_outros.bbox(pos)
                    if bbox:
                        margin = 10
                        is_right = event.x > bbox[0] + bbox[2] - margin
                        is_bottom = event.y > bbox[1] + bbox[3] - margin
                        
                        if is_right or is_bottom:
                            resizing_data["active"] = True
                            resizing_data["img_id"] = name
                            resizing_data["start_x"] = event.x
                            for img_info in txt_outros.editor_images:
                                if img_info["id"] == name:
                                    resizing_data["start_w"] = img_info["cur_w"]
                                    break
                            return "break" # Impede mover cursor do texto
            except: pass

        def on_text_release(event):
            resizing_data["active"] = False

        txt_outros.bind("<Motion>", on_text_motion)
        txt_outros.bind("<ButtonPress-1>", on_text_press, add="+")
        txt_outros.bind("<ButtonRelease-1>", on_text_release, add="+")

        # Configuração das Tags
        txt_outros.tag_configure("bold", font=("Segoe UI", 11, "bold"))
        txt_outros.tag_configure("italic", font=("Segoe UI", 11, "italic"))
        txt_outros.tag_configure("underline", underline=True)
        txt_outros.tag_configure("left", justify="left")
        txt_outros.tag_configure("center", justify="center")
        txt_outros.tag_configure("right", justify="right")

        # Carregar dados persistidos (Rich Text Simulation)
        self._render_rich_text(txt_outros, lab.inventario if lab.inventario else {})
        txt_outros.config(state=tk.NORMAL) # Habilita edição, já que o helper desabilita

        def salvar_tudo():
            # Projetor
            lab.inventario["projetor"] = {
                "modelo": e_mod.get(),
                "patrimonio": e_pat.get(),
                "tempo_lampada": e_tmp.get(),
                "outros": e_out.get()
            }
            # Tabelas
            save_micros()
            save_ac()
            # Outros (Serialização de Tags e Imagens)
            plain_text = txt_outros.get("1.0", tk.END).strip()
            lab.inventario["outros_texto"] = plain_text
            
            rich_tags = {}
            for tag in txt_outros.tag_names():
                if tag == "sel": continue
                ranges = txt_outros.tag_ranges(tag)
                tag_ranges = []
                for i in range(0, len(ranges), 2):
                    tag_ranges.append((ranges[i].string, ranges[i+1].string))
                if tag_ranges: rich_tags[tag] = tag_ranges
            
            # Serialização de Imagens
            rich_images = []
            content_dump = txt_outros.dump("1.0", tk.END, image=True)
            for item_type, value, index in content_dump:
                if item_type == "image":
                    img_name = txt_outros.image_cget(index, "name")
                    for img_ref in txt_outros.editor_images:
                        if img_ref["id"] == img_name:
                            # Converte para caminho relativo ao salvar
                            save_path = img_ref["path"]
                            if os.path.isabs(save_path):
                                base_dir = os.path.dirname(os.path.abspath(__file__))
                                try:
                                    save_path = os.path.relpath(save_path, base_dir)
                                except ValueError:
                                    pass  # Drives diferentes, mantém absoluto
                            rich_images.append({
                                "path": save_path, 
                                "pos": index,
                                "w": img_ref["cur_w"],
                                "h": img_ref["cur_h"]
                            })
                            break
            
            lab.inventario["outros_rich"] = {"tags": rich_tags, "images": rich_images}
            
            self.db.alterar_laboratorio(lab)
            if callback: callback()
            win.destroy()
            messagebox.showinfo("Sucesso", "Inventário atualizado!")

        # Seleciona a aba inicial
        try: nb.select(aba_inicial)
        except: pass

        btn_save = tk.Button(win, text="💾 SALVAR INVENTÁRIO", bg=self._D["success"], fg="white", 
                             font=("Segoe UI", 10, "bold"), pady=10, command=salvar_tudo)
        btn_save.pack(fill=tk.X, padx=10, pady=10)

    def _exibir_imagem_no_label(self, label, path, size=(300, 300)):
        """Carrega e exibe uma imagem em um label Tkinter usando PIL."""
        try:
            from PIL import Image, ImageTk
            img = Image.open(path)
            img.thumbnail(size, Image.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(img)
            label.config(image=photo, text="")
            label.image = photo # Keep reference
        except Exception as e:
            label.config(image="", text=f"Erro ao carregar imagem")

    def _form_laboratorio(self, lab=None, callback=None):
        """Formulário para inserir ou editar um laboratório."""
        title = "Editar Laboratório" if lab else "Inserir Laboratório"
        win = tk.Toplevel(self.root)
        win.title(title)
        win.geometry("500x750")
        win.configure(bg=self._D["bg"])
        win.transient(self.root)
        win.grab_set()

        D = self._D
        main_frame = tk.Frame(win, bg=D["bg"], padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Campos
        tk.Label(main_frame, text="Nome do Laboratório:", bg=D["bg"], fg=D["fg"]).pack(anchor="w")
        ent_nome = ttk.Entry(main_frame)
        ent_nome.pack(fill=tk.X, pady=(0, 10))
        if lab: ent_nome.insert(0, lab.nome)

        tk.Label(main_frame, text="Quantidade de Micros:", bg=D["bg"], fg=D["fg"]).pack(anchor="w")
        ent_micros = ttk.Entry(main_frame)
        ent_micros.pack(fill=tk.X, pady=(0, 10))
        if lab: ent_micros.insert(0, str(lab.qtd_micros))

        tk.Label(main_frame, text="Descrição (Texto longo):", bg=D["bg"], fg=D["fg"]).pack(anchor="w")
        txt_desc = tk.Text(main_frame, height=4, bg=D["input"], fg=D["fg"], bd=1, relief="solid")
        txt_desc.pack(fill=tk.X, pady=(0, 10))
        if lab: txt_desc.insert("1.0", lab.descricao)

        tk.Label(main_frame, text="Observação (Texto longo):", bg=D["bg"], fg=D["fg"]).pack(anchor="w")
        txt_obs = tk.Text(main_frame, height=4, bg=D["input"], fg=D["fg"], bd=1, relief="solid")
        txt_obs.pack(fill=tk.X, pady=(0, 10))
        if lab: txt_obs.insert("1.0", lab.observacao)

        # Seleção de Imagens
        path_planta = tk.StringVar(value=lab.planta_path if lab else "")

        def selecionar_arquivo(var, label_preview):
            path = filedialog.askopenfilename(filetypes=[("Imagens", "*.png *.jpg *.jpeg *.bmp")])
            if path:
                if not os.path.exists("uploads"): os.makedirs("uploads")
                dest = os.path.join("uploads", os.path.basename(path))
                import shutil
                try:
                    shutil.copy(path, dest)
                    var.set(dest)
                    self._exibir_imagem_no_label(label_preview, dest, size=(100, 100))
                except Exception as e:
                    messagebox.showerror("Erro", f"Erro ao copiar imagem: {e}", parent=win)


        # Preview Planta Lab
        f_plan = tk.Frame(main_frame, bg=D["bg"])
        f_plan.pack(fill=tk.X, pady=5)
        tk.Label(f_plan, text="Planta do Lab:", bg=D["bg"], fg=D["fg"]).pack(side=tk.LEFT)
        tk.Button(f_plan, text="Selecionar", command=lambda: selecionar_arquivo(path_planta, lbl_p_plan)).pack(side=tk.RIGHT)
        lbl_p_plan = tk.Label(main_frame, text="Sem planta", bg=D["input"], height=5)
        lbl_p_plan.pack(fill=tk.X, pady=(0, 10))
        if lab and lab.planta_path: self._exibir_imagem_no_label(lbl_p_plan, lab.planta_path, size=(100, 100))

        def salvar():
            nome = ent_nome.get().strip()
            if not nome:
                messagebox.showerror("Erro", "Nome é obrigatório.", parent=win)
                return
            
            nonlocal lab
            if not lab:
                n_lab = Laboratorio(id=None, nome=nome, 
                                  descricao=txt_desc.get("1.0", tk.END).strip(),
                                  observacao=txt_obs.get("1.0", tk.END).strip(),
                                  qtd_micros=int(ent_micros.get() or 0),
                                  planta_path=path_planta.get())
                self.db.adicionar_laboratorio(n_lab)
            else:
                lab.nome = nome
                lab.descricao = txt_desc.get("1.0", tk.END).strip()
                lab.observacao = txt_obs.get("1.0", tk.END).strip()
                try: lab.qtd_micros = int(ent_micros.get() or 0)
                except: pass
                lab.planta_path = path_planta.get()
                self.db.alterar_laboratorio(lab)
            
            if callback: callback()
            win.destroy()
            messagebox.showinfo("Sucesso", "Laboratório salvo!", parent=self.root)

        tk.Button(main_frame, text="SALVAR", bg=D["success"], fg="white", font=("Segoe UI", 10, "bold"),
                  pady=10, command=salvar).pack(fill=tk.X, pady=20)

    def _janela_detalhes_laboratorio(self, lab, callback=None):
        """Exibe todas as informações de um laboratório com opções de editar/excluir/novo."""
        win = tk.Toplevel(self.root)
        win.title(f"Detalhes: {lab.nome}")
        win.geometry("900x800")
        win.configure(bg=self._D["bg"])
        win.transient(self.root)
        win.grab_set()

        D = self._D
        
        # Botões de Ação na base (fixo)
        btn_bar = tk.Frame(win, bg=D["hdr"], pady=12)
        btn_bar.pack(side=tk.BOTTOM, fill=tk.X)

        def editar():
            win.destroy()
            self._form_laboratorio(lab, callback=callback)

        def excluir():
            if messagebox.askyesno("Confirmar", f"Excluir laboratório '{lab.nome}'?", parent=win):
                self.db.apagar_laboratorio(lab.id)
                if callback: callback()
                win.destroy()
                messagebox.showinfo("Sucesso", "Laboratório excluído.")

        def gerenciar_inventario():
            self._janela_inventario_laboratorio(lab, callback=callback)

        tk.Button(btn_bar, text="✏️ Editar", bg=D["warning"], fg="white", font=("Segoe UI", 9, "bold"),
                  padx=20, pady=8, relief="flat", command=editar).pack(side=tk.LEFT, padx=15)
        
        tk.Button(btn_bar, text="🗑️ Excluir", bg=D["danger"], fg="white", font=("Segoe UI", 9, "bold"),
                  padx=20, pady=8, relief="flat", command=excluir).pack(side=tk.LEFT, padx=5)

        tk.Button(btn_bar, text="➕ Inserir Novo", bg=D["success"], fg="white", font=("Segoe UI", 9, "bold"),
                  padx=20, pady=8, relief="flat", command=gerenciar_inventario).pack(side=tk.LEFT, padx=5)
        
        tk.Button(btn_bar, text="✕ Fechar", bg="#475569", fg="white", font=("Segoe UI", 9, "bold"),
                  padx=20, pady=8, relief="flat", command=win.destroy).pack(side=tk.RIGHT, padx=15)

        # Conteúdo Rolável
        canvas = tk.Canvas(win, bg=D["bg"], highlightthickness=0)
        scroll = ttk.Scrollbar(win, orient=tk.VERTICAL, command=canvas.yview)
        scroll_frame = tk.Frame(canvas, bg=D["bg"])

        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scroll.set)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=20, pady=10)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # Título
        tk.Label(scroll_frame, text=lab.nome, font=("Segoe UI", 20, "bold"),
                 bg=D["bg"], fg=D["primary"]).pack(anchor="w", pady=(0, 10))

        # Imagens
        img_frame = tk.Frame(scroll_frame, bg=D["bg"])
        img_frame.pack(fill=tk.X, pady=10)

        f_l = tk.Frame(img_frame, bg=D["card"])
        f_l.pack(side=tk.LEFT, padx=5)
        tk.Label(f_l, text="Planta Baixa", bg=D["card"], fg=D["fg2"]).pack()
        lbl_plan = tk.Label(f_l, bg=D["input"], width=350, height=350)
        lbl_plan.pack()
        if lab.planta_path: self._exibir_imagem_no_label(lbl_plan, lab.planta_path)

        # Textos
        def _add_section(title, content):
            tk.Label(scroll_frame, text=title, font=("Segoe UI", 12, "bold"),
                     bg=D["bg"], fg=D["fg2"]).pack(anchor="w", pady=(15, 2))
            l = tk.Label(scroll_frame, text=content, bg=D["card"], fg=D["fg"],
                         wraplength=800, justify="left", padx=15, pady=15, font=("Segoe UI", 10))
            l.pack(fill=tk.X)

        _add_section("Descrição", lab.descricao or "Nenhuma descrição.")
        _add_section("Observações", lab.observacao or "Nenhuma observação.")

        # --- Aulas Vinculadas ---
        tk.Label(scroll_frame, text="Aulas no Laboratório", font=("Segoe UI", 12, "bold"),
                 bg=D["bg"], fg=D["primary"]).pack(anchor="w", pady=(20, 5))
        
        frame_aulas = tk.Frame(scroll_frame, bg=D["card"], padx=2, pady=2)
        frame_aulas.pack(fill=tk.X, pady=(0, 30))

        cols_aulas = ("Dia", "Horário", "Matéria", "Professor")
        t_aulas_lab = ttk.Treeview(frame_aulas, columns=cols_aulas, show="headings", height=8)
        for c in cols_aulas:
            t_aulas_lab.heading(c, text=c.upper())
            t_aulas_lab.column(c, width=150, anchor="center" if c in ["Dia", "Horário"] else "w")
        
        t_aulas_lab.pack(fill=tk.X)

        aulas_lab = self.db.listar_aulas_por_lab(lab.nome)
        for a in aulas_lab:
            t_aulas_lab.insert("", tk.END, values=(a.dia_semana, f"{a.hora_inicio} - {a.hora_fim}", a.disciplina, a.professor))

    def sair(self):
        """Fecha o aplicativo com confirmação."""
        if messagebox.askyesno("Sair", "Deseja realmente sair?"):
            self.root.destroy()
