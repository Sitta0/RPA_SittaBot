from playwright.sync_api import sync_playwright
import os
import sys
import time
import shutil
import configparser
from datetime import datetime
import threading
import tkinter as tk

# CONFIGURAÇÕES

EDGE_PATH_DEFAULT = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
EDGE_PATH = EDGE_PATH_DEFAULT  # sobrescrito após carregar config.ini

# URLs carregadas do config.ini (ver _carregar_config)
URL_LOGIN    = None
URL_PARCELAS = None
URL_EMITIDOS = None
SSO_BUTTON_TEXT = None

APP_NAME   = "SittaBot Reports"
APP_SUB    = "Automação de relatórios web"
APP_AUTHOR = "github.com/seu-usuario/sittabot-reports"

# DIRETÓRIO BASE

if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

PASTA_DOWNLOAD = BASE_DIR
PERFIL         = os.path.join(BASE_DIR, "sittabot_edge_profile")
os.makedirs(PERFIL, exist_ok=True)

LOG_ERRO = os.path.join(BASE_DIR, "erro.log")

# CONFIG.INI 
#
# NENHUMA credencial é armazenada no código.
# Todas as credenciais e URLs são lidas EXCLUSIVAMENTE do config.ini.
# Se o arquivo não existir ou estiver incompleto, o bot cria um
# template vazio e encerra com instrução clara ao usuário.

CONFIG_FILE = os.path.join(BASE_DIR, "config.ini")

_CONFIG_TEMPLATE = """\
[credenciais]
; Preencha os campos abaixo com as credenciais da conta de serviço.
; Nunca compartilhe este arquivo nem o versione no Git.
email = 
senha = 

[sistema]
; URLs do sistema alvo. Ajuste conforme o ambiente.
url_login     = https://seudominio.com/#/login
url_parcelas  = https://seudominio.com/#/relatorio/parcelas
url_emitidos  = https://seudominio.com/#/relatorio/meus-relatorios/emitidos

; Texto exato do botão de login SSO exibido na tela (case-sensitive)
sso_button_text = SSO Login

; Nome do relatório a ser clicado no menu de opções
nome_relatorio = Relatório

; Caminho do executável do Microsoft Edge (deixe em branco para usar o padrão)
; edge_path = C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe
"""

def _carregar_config():
    """
    Lê o config.ini e retorna credenciais + URLs.
    Se o arquivo não existir → cria o template e encerra com mensagem.
    Se campos obrigatórios estiverem vazios → encerra com mensagem.
    """
    if not os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            f.write(_CONFIG_TEMPLATE)
        msg = (
            f"Arquivo config.ini criado em:\n{CONFIG_FILE}\n\n"
            "Preencha os campos 'email', 'senha' e as URLs antes de executar o bot novamente."
        )
        try:
            import ctypes
            ctypes.windll.user32.MessageBoxW(0, msg, "SittaBot — Configuração necessária", 0x30)
        except Exception:
            print(msg)
        sys.exit(1)

    cfg = configparser.ConfigParser()
    cfg.read(CONFIG_FILE, encoding="utf-8")

    email           = cfg.get("credenciais", "email",          fallback="").strip()
    senha           = cfg.get("credenciais", "senha",          fallback="").strip()
    url_login       = cfg.get("sistema",     "url_login",      fallback="").strip()
    url_parcelas    = cfg.get("sistema",     "url_parcelas",   fallback="").strip()
    url_emitidos    = cfg.get("sistema",     "url_emitidos",   fallback="").strip()
    sso_button_text = cfg.get("sistema",     "sso_button_text",fallback="SSO Login").strip()
    nome_relatorio  = cfg.get("sistema",     "nome_relatorio", fallback="Relatório").strip()
    edge_path       = cfg.get("sistema",     "edge_path",      fallback="").strip()

    campos_vazios = [
        campo for campo, valor in [
            ("email",        email),
            ("senha",        senha),
            ("url_login",    url_login),
            ("url_parcelas", url_parcelas),
            ("url_emitidos", url_emitidos),
        ] if not valor
    ]

    if campos_vazios:
        msg = (
            f"O config.ini está incompleto.\n\n"
            f"Campo(s) vazio(s): {', '.join(campos_vazios)}\n\n"
            f"Edite o arquivo abaixo e preencha os dados:\n{CONFIG_FILE}"
        )
        try:
            import ctypes
            ctypes.windll.user32.MessageBoxW(0, msg, "SittaBot — Configuração incompleta", 0x30)
        except Exception:
            print(msg)
        sys.exit(1)

    return {
        "email":           email,
        "senha":           senha,
        "url_login":       url_login,
        "url_parcelas":    url_parcelas,
        "url_emitidos":    url_emitidos,
        "sso_button_text": sso_button_text,
        "nome_relatorio":  nome_relatorio,
        "edge_path":       edge_path,
    }


def _salvar_senha(nova_senha):
    """Persiste a nova senha no config.ini e atualiza em memória."""
    cfg = configparser.ConfigParser()
    cfg.read(CONFIG_FILE, encoding="utf-8")
    if "credenciais" not in cfg:
        cfg["credenciais"] = {}
    cfg["credenciais"]["senha"] = nova_senha
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        cfg.write(f)
    CREDENCIAIS["senha"] = nova_senha


# Carrega config e popula variáveis globais
CREDENCIAIS     = _carregar_config()
URL_LOGIN       = CREDENCIAIS["url_login"]
URL_PARCELAS    = CREDENCIAIS["url_parcelas"]
URL_EMITIDOS    = CREDENCIAIS["url_emitidos"]
SSO_BUTTON_TEXT = CREDENCIAIS["sso_button_text"]
NOME_RELATORIO  = CREDENCIAIS["nome_relatorio"]
EDGE_PATH       = CREDENCIAIS["edge_path"] or EDGE_PATH_DEFAULT

# Sinalizadores para comunicação entre threads (automação Tk)
_senha_expirada_pendente = threading.Event()
_nova_senha_disponivel   = threading.Event()
_nova_senha_valor        = None
_cancelar_automacao      = threading.Event()

# CORES 

BG_MAIN  = "#0e0e22"
BG_TITLE = "#111128"

STEP_DONE_BG     = "#0a1f0e"
STEP_DONE_BORDER = "#16a34a"
STEP_DONE_ICON   = "#16a34a"
STEP_DONE_LBL    = "#ffffff"
STEP_DONE_TABBG  = "#0a2e14"
STEP_DONE_TABFG  = "#4ade80"
STEP_DONE_TAG    = "CONCLUÍDO"

STEP_ACT_BG      = "#0a1228"
STEP_ACT_BORDER  = "#1d4ed8"
STEP_ACT_ICON    = "#1d4ed8"
STEP_ACT_LBL     = "#ffffff"
STEP_ACT_TABBG   = "#0a1628"
STEP_ACT_TABFG   = "#60a5fa"
STEP_ACT_TAG     = "EM ANDAMENTO"

STEP_WAIT_BG     = "#111128"
STEP_WAIT_BORDER = "#1a1a3a"
STEP_WAIT_ICON   = "#1a1a3a"
STEP_WAIT_LBL    = "#ffffff"
STEP_WAIT_TABBG  = "#111128"
STEP_WAIT_TABFG  = "#555577"
STEP_WAIT_TAG    = "AGUARDANDO"

C_WHITE   = "#ffffff"
C_MUTED   = "#8888aa"
C_DIMMED  = "#444466"
C_ACCENT  = "#3b82f6"
C_PURPLE  = "#6c63ff"

C_ERR_BG  = "#dc2626"
C_ERR_TXT = "#f87171"
C_MFA_BG  = "#d97706"
C_SUC_BG  = "#16a34a"

C_WARN_BORDER = "#ea580c"
C_WARN_TXT    = "#fb923c"

# ESTADO GLOBAL

popup        = None
canvas_main  = None
_prog_id     = None
_attempt_id  = None
_step_ids    = {}
_anim_pos    = 0
_anim_active = True

_erro_pendente        = None
_popup_pronto         = threading.Event()
_popup_senha_pendente = False

STEPS = [
    "Login SSO",
    "Exportação CSV solicitada",
    "Aguardando processamento",
    "Download do arquivo",
]



def _rounded_rect(c, x1, y1, x2, y2, r, **kw):
    fill    = kw.get("fill", "")
    outline = kw.get("outline", "") or ""
    points = [
        x1+r, y1,  x2-r, y1,  x2, y1,  x2, y1+r,
        x2, y2-r,  x2, y2,  x2-r, y2,  x1+r, y2,
        x1, y2,  x1, y2-r,  x1, y1+r,  x1, y1,
    ]
    return [c.create_polygon(points, smooth=True, fill=fill, outline=outline, splinesteps=32)]


def _center(win, w, h):
    sw = win.winfo_screenwidth()
    sh = win.winfo_screenheight()
    win.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")


def _titlebar(c, w, label):
    c.create_rectangle(0, 0, w, 32, fill=BG_TITLE, outline="")
    c.create_oval(12, 10, 22, 22, fill="#ff5f57", outline="")
    c.create_oval(28, 10, 38, 22, fill="#febc2e", outline="")
    c.create_oval(44, 10, 54, 22, fill="#28c840", outline="")
    c.create_text(w//2, 16, text=label, fill=C_MUTED, font=("Segoe UI", 10, "bold"))
    c.create_line(0, 32, w, 32, fill="#080818", width=1)


def _step_cfg(estado):
    if estado == "done":
        return dict(bg=STEP_DONE_BG, border=STEP_DONE_BORDER, icon_bg=STEP_DONE_ICON,
                    icon_txt="✓", lbl=STEP_DONE_LBL,
                    tag_bg=STEP_DONE_TABBG, tag_fg=STEP_DONE_TABFG, tag=STEP_DONE_TAG)
    elif estado == "active":
        return dict(bg=STEP_ACT_BG, border=STEP_ACT_BORDER, icon_bg=STEP_ACT_ICON,
                    icon_txt="⟳", lbl=STEP_ACT_LBL,
                    tag_bg=STEP_ACT_TABBG, tag_fg=STEP_ACT_TABFG, tag=STEP_ACT_TAG)
    else:
        return dict(bg=STEP_WAIT_BG, border=STEP_WAIT_BORDER, icon_bg=STEP_WAIT_ICON,
                    icon_txt="·", lbl=STEP_WAIT_LBL,
                    tag_bg=STEP_WAIT_TABBG, tag_fg=STEP_WAIT_TABFG, tag=STEP_WAIT_TAG)


def _draw_step(c, idx, label, estado, y):
    cfg = _step_cfg(estado)
    ids = {}
    ids["bg"]   = _rounded_rect(c, 16, y, 404, y+32, 8, fill=cfg["bg"], outline=cfg["border"])
    ids["ibg"]  = _rounded_rect(c, 22, y+6, 42, y+26, 5, fill=cfg["icon_bg"], outline="")
    ids["itxt"] = c.create_text(32, y+16, text=cfg["icon_txt"], fill=C_WHITE, font=("Segoe UI", 10, "bold"))
    ids["lbl"]  = c.create_text(50, y+16, anchor="w", text=label, fill=cfg["lbl"], font=("Segoe UI", 10, "bold"))
    ids["tbg"]  = _rounded_rect(c, 288, y+8, 400, y+24, 8, fill=cfg["tag_bg"], outline="")
    ids["ttxt"] = c.create_text(344, y+16, text=cfg["tag"], fill=cfg["tag_fg"], font=("Segoe UI", 7, "bold"))
    return ids


# POPUP PRINCIPAL(VALIDADO)

W, H = 420, 310


def iniciar_popup():
    global popup, canvas_main, _prog_id, _attempt_id, _step_ids

    popup = tk.Tk()
    popup.title(APP_NAME)
    popup.configure(bg=BG_TITLE)
    popup.resizable(False, False)
    popup.attributes("-topmost", True)
    _center(popup, W, H)

    c = tk.Canvas(popup, width=W, height=H, bg=BG_TITLE, highlightthickness=0)
    c.pack()
    canvas_main = c

    c.create_rectangle(0, 32, W, H, fill=BG_MAIN, outline="")
    _titlebar(c, W, APP_NAME)

    _rounded_rect(c, 16, 44, 60, 88, 10, fill=C_PURPLE, outline="")
    c.create_text(38, 66, text="🤖", font=("Segoe UI", 18))
    c.create_text(72, 57, anchor="w", text=APP_NAME, fill=C_WHITE, font=("Segoe UI", 12, "bold"))
    c.create_text(72, 75, anchor="w", text=APP_SUB,  fill=C_MUTED, font=("Segoe UI", 9))

    step_y = [100, 138, 176, 214]
    for i, lbl in enumerate(STEPS):
        _step_ids[i] = _draw_step(c, i, lbl, "wait", step_y[i])

    _rounded_rect(c, 16, 256, 404, 262, 3, fill="#1a1a35", outline="")
    _prog_id = c.create_rectangle(16, 256, 40, 262, fill=C_ACCENT, outline="")

    c.create_text(16, 282, anchor="w", text=APP_AUTHOR, fill=C_WHITE, font=("Segoe UI", 8))
    _attempt_id = c.create_text(404, 282, anchor="e", text="", fill=C_MUTED, font=("Segoe UI", 8))

    _popup_pronto.set()
    popup.after(500, _verificar_erro_pendente)
    popup.after(500, _verificar_popup_senha_pendente)
    popup.after(100, _animar_progresso)

    popup.mainloop()


def _animar_progresso():
    global _anim_pos, _anim_active
    if not _anim_active:
        return
    try:
        _anim_pos = (_anim_pos + 2) % 101
        x2 = 16 + int(388 * (_anim_pos / 100))
        canvas_main.coords(_prog_id, 16, 256, max(x2, 36), 262)
        popup.after(30, _animar_progresso)
    except Exception:
        pass


def _set_step(idx, estado):
    if popup and popup.winfo_exists():
        popup.after(0, lambda: _do_set_step(idx, estado))


def _do_set_step(idx, estado):
    c   = canvas_main
    ids = _step_ids.get(idx)
    if not ids or not c:
        return
    cfg = _step_cfg(estado)
    try:
        for item in ids["bg"]:  c.itemconfig(item, fill=cfg["bg"], outline=cfg["border"])
        for item in ids["ibg"]: c.itemconfig(item, fill=cfg["icon_bg"], outline="")
        c.itemconfig(ids["itxt"], text=cfg["icon_txt"], fill=C_WHITE)
        c.itemconfig(ids["lbl"],  fill=cfg["lbl"])
        for item in ids["tbg"]: c.itemconfig(item, fill=cfg["tag_bg"], outline="")
        c.itemconfig(ids["ttxt"], text=cfg["tag"], fill=cfg["tag_fg"])
    except Exception:
        pass


def update_status(msg):
    if _attempt_id and popup and popup.winfo_exists():
        popup.after(0, lambda: canvas_main.itemconfig(_attempt_id, text=msg))


def set_tentativa(t, max_t):
    if _attempt_id and popup and popup.winfo_exists():
        popup.after(0, lambda: canvas_main.itemconfig(
            _attempt_id, text=f"Tentativa {t} / {max_t}"))


def fechar_popup():
    global _anim_active
    _anim_active = False
    if popup and popup.winfo_exists():
        popup.after(0, popup.destroy)


def _verificar_erro_pendente():
    global _erro_pendente
    if _erro_pendente is not None:
        mensagem = _erro_pendente
        _erro_pendente = None
        _abrir_janela_erro(mensagem)
    if popup and popup.winfo_exists():
        popup.after(500, _verificar_erro_pendente)


def _verificar_popup_senha_pendente():
    global _popup_senha_pendente
    if _popup_senha_pendente:
        _popup_senha_pendente = False
        _criar_popup_senha(expirada=True)
    if popup and popup.winfo_exists():
        popup.after(500, _verificar_popup_senha_pendente)


# POPUP MFA(TESTE)

def mostrar_popup_mfa(numero):
    if popup and popup.winfo_exists():
        popup.after(0, lambda: _criar_popup_mfa(numero))


def _criar_popup_mfa(numero):
    mw, mh = 420, 280
    mfa_win = tk.Toplevel(popup)
    mfa_win.title("Autenticação MFA")
    mfa_win.configure(bg=BG_TITLE)
    mfa_win.resizable(False, False)
    mfa_win.attributes("-topmost", True)
    _center(mfa_win, mw, mh)

    c = tk.Canvas(mfa_win, width=mw, height=mh, bg=BG_TITLE, highlightthickness=0)
    c.pack()

    c.create_rectangle(0, 32, mw, mh, fill=BG_MAIN, outline="")
    _titlebar(c, mw, "Autenticação MFA")
    _rounded_rect(c, 184, 44, 236, 88, 12, fill=C_MFA_BG, outline="")
    c.create_text(210, 66, text="🔐", font=("Segoe UI", 18))
    c.create_text(210, 102, text="Autenticação necessária", fill=C_WHITE, font=("Segoe UI", 12, "bold"))
    c.create_text(210, 120, text="Abra o Microsoft Authenticator e aprove o número abaixo:",
                  fill=C_MUTED, font=("Segoe UI", 9), width=360)
    _rounded_rect(c, 90, 134, 330, 220, 14, fill="#0a1628", outline=C_ACCENT)
    c.create_text(210, 170, text=str(numero), fill=C_WHITE, font=("Segoe UI", 48, "bold"))
    c.create_text(210, 208, text="CÓDIGO DE APROVAÇÃO", fill="#1d4ed8", font=("Segoe UI", 7, "bold"))
    c.create_text(210, 256, text="Esta janela fechará automaticamente após a aprovação.",
                  fill=C_DIMMED, font=("Segoe UI", 8))

    popup._mfa_win = mfa_win


def fechar_popup_mfa():
    if popup and hasattr(popup, '_mfa_win'):
        try:
            popup.after(0, popup._mfa_win.destroy)
        except Exception:
            pass


# POPUP ERRO(VALIDADO)
def popup_erro(mensagem):
    global _erro_pendente
    _popup_pronto.wait(timeout=10)
    _erro_pendente = mensagem


def _abrir_janela_erro(mensagem):
    if not popup or not popup.winfo_exists():
        return

    ew, eh = 420, 252
    erro = tk.Toplevel(popup)
    erro.title("Erro no Processo")
    erro.configure(bg=BG_TITLE)
    erro.resizable(False, False)
    erro.attributes("-topmost", True)
    erro.protocol("WM_DELETE_WINDOW", lambda: None)
    _center(erro, ew, eh)

    c = tk.Canvas(erro, width=ew, height=eh, bg=BG_TITLE, highlightthickness=0)
    c.pack()

    c.create_rectangle(0, 32, ew, eh, fill="#120808", outline="")
    _titlebar(c, ew, "Erro no Processo")
    _rounded_rect(c, 184, 44, 236, 88, 12, fill=C_ERR_BG, outline="")
    c.create_text(210, 66, text="❌", font=("Segoe UI", 18))
    c.create_text(210, 102, text="Ocorreu um erro", fill=C_WHITE, font=("Segoe UI", 12, "bold"))
    _rounded_rect(c, 16, 114, 404, 186, 8, fill="#1f0808", outline="#5a1010")
    c.create_text(30, 124, anchor="w", text="● Detalhes do erro", fill="#7f1d1d", font=("Segoe UI", 7, "bold"))

    msg_curta = mensagem if len(mensagem) <= 160 else mensagem[:157] + "..."
    c.create_text(210, 154, text=msg_curta, fill=C_ERR_TXT,
                  font=("Courier New", 8), width=374, justify="center")

    btn_ids = _rounded_rect(c, 160, 200, 260, 226, 8, fill=C_ERR_BG, outline="")
    tid = c.create_text(210, 213, text="OK", fill=C_WHITE, font=("Segoe UI", 11, "bold"))

    def _fechar(e=None): erro.destroy(); fechar_popup()
    def _hin(e):
        for i in btn_ids:
            try: c.itemconfig(i, fill="#b91c1c")
            except: pass
    def _hout(e):
        for i in btn_ids:
            try: c.itemconfig(i, fill=C_ERR_BG)
            except: pass

    for tag in [tid] + btn_ids:
        c.tag_bind(tag, "<Button-1>", _fechar)
        c.tag_bind(tag, "<Enter>", _hin)
        c.tag_bind(tag, "<Leave>", _hout)
    c.config(cursor="hand2")


# POPUP SUCESSO(VALIDADO)

def mostrar_popup_sucesso(nome_arquivo):
    if popup and popup.winfo_exists():
        popup.after(0, lambda: _criar_popup_sucesso(nome_arquivo))


def _criar_popup_sucesso(nome_arquivo):
    global _anim_active
    _anim_active = False

    sw, sh = 420, 232
    suc = tk.Toplevel(popup)
    suc.title("Processo Finalizado")
    suc.configure(bg=BG_TITLE)
    suc.resizable(False, False)
    suc.attributes("-topmost", True)
    _center(suc, sw, sh)

    c = tk.Canvas(suc, width=sw, height=sh, bg=BG_TITLE, highlightthickness=0)
    c.pack()

    c.create_rectangle(0, 32, sw, sh, fill="#081208", outline="")
    _titlebar(c, sw, "Processo Finalizado")
    _rounded_rect(c, 184, 44, 236, 88, 12, fill=C_SUC_BG, outline="")
    c.create_text(210, 66, text="✅", font=("Segoe UI", 18))
    c.create_text(210, 102, text="Processo finalizado!", fill=C_WHITE, font=("Segoe UI", 12, "bold"))
    c.create_text(210, 118, text="O relatório foi baixado com sucesso.", fill="#166534", font=("Segoe UI", 9))
    _rounded_rect(c, 16, 130, 404, 180, 8, fill="#0a1f0e", outline="#166534")
    c.create_text(34, 155, text="📄", font=("Segoe UI", 18))
    c.create_text(58, 148, anchor="w", text=nome_arquivo, fill=C_WHITE, font=("Segoe UI", 9, "bold"))
    c.create_text(58, 164, anchor="w", text="Salvo na pasta do executável", fill="#166534", font=("Segoe UI", 8))
    c.create_text(210, 202, text=APP_AUTHOR, fill=C_WHITE, font=("Segoe UI", 8))

    countdown_id = c.create_text(210, 218, text="Fechando em 5 segundos...",
                                  fill="#2d6e3a", font=("Segoe UI", 7))

    def _fechar_tudo(e=None):
        try: suc.destroy()
        except Exception: pass
        fechar_popup()

    def _countdown(n):
        if n <= 0: _fechar_tudo(); return
        try:
            label = "segundo" if n == 1 else "segundos"
            c.itemconfig(countdown_id, text=f"Fechando em {n} {label}...")
            suc.after(1000, lambda: _countdown(n - 1))
        except Exception:
            pass

    suc.bind("<Button-1>", _fechar_tudo)
    c.bind("<Button-1>",   _fechar_tudo)
    suc.after(1000, lambda: _countdown(4))


# POPUP ATUALIZAÇÃO DE SENHA(TESTE)

def solicitar_atualizacao_senha():
    """Chamado pela thread de automação quando senha expirada é detectada."""
    global _popup_senha_pendente
    _popup_pronto.wait(timeout=10)
    _popup_senha_pendente = True


def _criar_popup_senha(expirada=False):
    """Cria o popup de atualização de senha — roda na thread Tk."""
    if not popup or not popup.winfo_exists():
        return

    ph = 430 if expirada else 380
    win = tk.Toplevel(popup)
    win.title("Atualização de Senha")
    win.configure(bg=BG_TITLE)
    win.resizable(False, False)
    win.attributes("-topmost", True)
    win.protocol("WM_DELETE_WINDOW", lambda: None)
    _center(win, 420, ph)

    c = tk.Canvas(win, width=420, height=ph, bg=BG_TITLE, highlightthickness=0)
    c.pack()

    c.create_rectangle(0, 32, 420, ph, fill=BG_MAIN, outline="")
    _titlebar(c, 420, "Atualização de Senha")

    y = 44

    if expirada:
        _rounded_rect(c, 16, y, 404, y+52, 8, fill="#1c0a00", outline=C_WARN_BORDER)
        c.create_text(36, y+16, text="⚠️", font=("Segoe UI", 14))
        c.create_text(58, y+10, anchor="w", text="SENHA EXPIRADA DETECTADA",
                      fill=C_WARN_TXT, font=("Segoe UI", 7, "bold"))
        c.create_text(58, y+26, anchor="w",
                      text="Login falhou pois a senha da conta de serviço expirou.",
                      fill="#9a3412", font=("Segoe UI", 8), width=330)
        c.create_text(58, y+40, anchor="w",
                      text="Atualize abaixo — o bot continuará automaticamente.",
                      fill="#9a3412", font=("Segoe UI", 8), width=330)
        y += 62

    _rounded_rect(c, 184, y, 236, y+44, 12, fill="#1e3a8a", outline="")
    c.create_text(210, y+22, text="🔑", font=("Segoe UI", 18))
    y += 54

    c.create_text(210, y, text="Atualizar senha do config.ini",
                  fill=C_WHITE, font=("Segoe UI", 12, "bold"))
    y += 16
    c.create_text(210, y, text="As alterações são salvas imediatamente no arquivo",
                  fill=C_MUTED, font=("Segoe UI", 8))
    y += 20

    c.create_text(20, y, anchor="w", text="E-MAIL", fill="#333355", font=("Segoe UI", 7, "bold"))
    y += 12
    _rounded_rect(c, 16, y, 404, y+28, 7, fill="#0a0a1e", outline="#1a1a3a")
    c.create_text(28, y+14, text="👤", font=("Segoe UI", 11))
    c.create_text(46, y+14, anchor="w", text=CREDENCIAIS["email"],
                  fill="#555577", font=("Courier New", 9))
    y += 36

    c.create_text(20, y, anchor="w", text="NOVA SENHA", fill="#333355", font=("Segoe UI", 7, "bold"))
    y += 12
    _rounded_rect(c, 16, y, 404, y+28, 7, fill="#0a0a1e", outline="#1d4ed8")
    c.create_text(28, y+14, text="✏️", font=("Segoe UI", 11))

    entry_var = tk.StringVar()
    entry = tk.Entry(win, textvariable=entry_var, show="•",
                     font=("Courier New", 11), bg="#0a0a1e", fg="#ffffff",
                     insertbackground="#3b82f6", relief="flat", bd=0, highlightthickness=0)
    entry.place(x=46, y=y+7, width=310, height=16)

    _mostrar = [False]
    eye_id = c.create_text(390, y+14, text="👁", font=("Segoe UI", 11), fill="#555577")

    def _toggle_eye(e=None):
        _mostrar[0] = not _mostrar[0]
        entry.config(show="" if _mostrar[0] else "•")
        c.itemconfig(eye_id, fill=C_ACCENT if _mostrar[0] else "#555577")

    c.tag_bind(eye_id, "<Button-1>", _toggle_eye)
    y += 36

    c.create_text(20, y, anchor="w", text="FORÇA DA SENHA",
                  fill="#333355", font=("Segoe UI", 7, "bold"))
    _rounded_rect(c, 16, y+12, 404, y+16, 2, fill="#1a1a3a", outline="")
    forca_bar = c.create_rectangle(16, y+12, 16, y+16, fill=C_ACCENT, outline="")
    forca_txt = c.create_text(404, y+22, anchor="e", text="",
                               fill=C_ACCENT, font=("Segoe UI", 7, "bold"))
    y += 32

    def _calcular_forca(senha):
        pts = 0
        if len(senha) >= 8:  pts += 1
        if len(senha) >= 12: pts += 1
        if any(ch.isupper() for ch in senha): pts += 1
        if any(ch.islower() for ch in senha): pts += 1
        if any(ch.isdigit() for ch in senha): pts += 1
        if any(ch in "!@#$%^&*()_+-=[]{}|;:,.<>?" for ch in senha): pts += 1
        return pts

    def _on_change(*_):
        senha = entry_var.get()
        pts   = _calcular_forca(senha)
        pct   = min(pts / 6, 1.0)
        x2    = 16 + int(388 * pct)
        cor   = "#dc2626" if pct < .4 else "#d97706" if pct < .7 else "#16a34a"
        label = "FRACA"   if pct < .4 else "MÉDIA"   if pct < .7 else "FORTE"
        try:
            c.coords(forca_bar, 16, y-20, max(x2, 17), y-16)
            c.itemconfig(forca_bar, fill=cor)
            c.itemconfig(forca_txt, text=f"FORÇA: {label}", fill=cor)
            c.itemconfig(preview_senha_id, text=entry_var.get() or "...")
        except Exception:
            pass

    entry_var.trace_add("write", _on_change)

    c.create_text(20, y, anchor="w", text="PRÉ-VISUALIZAÇÃO DO CONFIG.INI",
                  fill="#333355", font=("Segoe UI", 7, "bold"))
    y += 12
    _rounded_rect(c, 16, y, 404, y+44, 7, fill="#080818", outline="#1a1a3a")
    c.create_text(26, y+10, anchor="w", text="[credenciais]",
                  fill="#8888aa", font=("Courier New", 8))
    c.create_text(26, y+22, anchor="w", text="email = " + CREDENCIAIS["email"],
                  fill="#4ade80", font=("Courier New", 8))
    c.create_text(26, y+34, anchor="w", text="senha = ",
                  fill="#60a5fa", font=("Courier New", 8))
    preview_senha_id = c.create_text(94, y+34, anchor="w", text="...",
                                      fill="#f59e0b", font=("Courier New", 8))
    y += 52

    def _cancelar(e=None):
        _cancelar_automacao.set()
        _nova_senha_disponivel.set()
        win.destroy()
        fechar_popup()

    def _salvar(e=None):
        global _nova_senha_valor
        nova = entry_var.get().strip()
        if not nova:
            return
        _salvar_senha(nova)
        _nova_senha_valor = nova
        _nova_senha_disponivel.set()
        _mostrar_senha_salva(win, nova)

    btn_c = _rounded_rect(c, 16, y, 196, y+34, 8, fill="#1a1a35", outline="#1a1a3a")
    tid_c = c.create_text(106, y+17, text="Cancelar", fill=C_MUTED, font=("Segoe UI", 10, "bold"))

    btn_s = _rounded_rect(c, 204, y, 404, y+34, 8, fill="#1d4ed8", outline="")
    tid_s = c.create_text(304, y+17, text="💾  Salvar e continuar",
                           fill=C_WHITE, font=("Segoe UI", 10, "bold"))

    for tag in [tid_c] + btn_c:
        c.tag_bind(tag, "<Button-1>", _cancelar)

    for tag in [tid_s] + btn_s:
        c.tag_bind(tag, "<Button-1>", _salvar)
        c.tag_bind(tag, "<Enter>",
                   lambda e, ids=btn_s: [c.itemconfig(i, fill="#2563eb") for i in ids])
        c.tag_bind(tag, "<Leave>",
                   lambda e, ids=btn_s: [c.itemconfig(i, fill="#1d4ed8") for i in ids])

    entry.focus()
    popup._senha_win = win


def _mostrar_senha_salva(win_anterior, nova_senha):
    """Tela de confirmação após salvar — fecha sozinha em 3 segundos."""
    try:
        win_anterior.destroy()
    except Exception:
        pass

    if not popup or not popup.winfo_exists():
        return

    sw, sh = 420, 240
    suc = tk.Toplevel(popup)
    suc.title("Senha Atualizada")
    suc.configure(bg=BG_TITLE)
    suc.resizable(False, False)
    suc.attributes("-topmost", True)
    _center(suc, sw, sh)

    c = tk.Canvas(suc, width=sw, height=sh, bg=BG_TITLE, highlightthickness=0)
    c.pack()

    c.create_rectangle(0, 32, sw, sh, fill="#081208", outline="")
    _titlebar(c, sw, "Senha Atualizada")
    _rounded_rect(c, 184, 44, 236, 88, 12, fill=C_SUC_BG, outline="")
    c.create_text(210, 66, text="✅", font=("Segoe UI", 18))
    c.create_text(210, 102, text="Senha atualizada!", fill=C_WHITE, font=("Segoe UI", 12, "bold"))
    c.create_text(210, 118, text="config.ini salvo com sucesso",
                  fill="#166534", font=("Segoe UI", 9))

    _rounded_rect(c, 16, 130, 404, 178, 7, fill="#080818", outline="#166534")
    c.create_text(26, 144, anchor="w", text="[credenciais]",  fill="#8888aa", font=("Courier New", 8))
    c.create_text(26, 156, anchor="w", text="email = " + CREDENCIAIS["email"],
                  fill="#4ade80", font=("Courier New", 8))
    c.create_text(26, 168, anchor="w", text="senha = " + nova_senha,
                  fill="#4ade80", font=("Courier New", 8))

    countdown_id = c.create_text(210, 206, text="O bot retomará em 3 segundos…",
                                  fill="#2d6e3a", font=("Segoe UI", 8))
    c.create_text(210, 220, text=APP_AUTHOR, fill=C_MUTED, font=("Segoe UI", 7))

    def _countdown(n):
        if n <= 0:
            try: suc.destroy()
            except: pass
            return
        try:
            c.itemconfig(countdown_id,
                         text=f"O bot retomará em {n} segundo{'s' if n != 1 else ''}…")
            suc.after(1000, lambda: _countdown(n - 1))
        except Exception:
            pass

    suc.after(1000, lambda: _countdown(2))


# RECUPERAÇÃO DE PERFIL

def iniciar_contexto(playwright):
    try:
        return playwright.chromium.launch_persistent_context(
            user_data_dir=PERFIL,
            executable_path=EDGE_PATH,
            headless=True,
            args=["--disable-blink-features=AutomationControlled",
                  "--no-sandbox", "--disable-dev-shm-usage"],
            accept_downloads=True,
            viewport=None
        )
    except Exception:
        time.sleep(3)
        shutil.rmtree(PERFIL, ignore_errors=True)
        os.makedirs(PERFIL, exist_ok=True)
        return playwright.chromium.launch_persistent_context(
            user_data_dir=PERFIL,
            executable_path=EDGE_PATH,
            headless=True,
            accept_downloads=True,
            viewport=None
        )


# VERIFICAÇÃO DE MFA 

def verificar_mfa(page):
    try:
        seletores_numero = [
            "[data-testid='numberDisplay']", ".display-sign",
            "#idRichContext_DisplaySign", ".push-number", "strong.display-sign",
        ]
        numero = None
        for seletor in seletores_numero:
            try:
                elemento = page.locator(seletor)
                if elemento.count() > 0:
                    numero = elemento.first.inner_text(timeout=3000).strip()
                    if numero:
                        break
            except Exception:
                continue

        if not numero:
            return False

        update_status(f"🔐 Aprove o número {numero} no Authenticator")
        mostrar_popup_mfa(numero)

        page.wait_for_url(
            lambda url: "login" not in url and "mfa" not in url.lower() and "auth" not in url.lower(),
            timeout=120000
        )
        fechar_popup_mfa()
        return True

    except Exception:
        fechar_popup_mfa()
        return False


# DETECÇÃO DE SENHA EXPIRADA

_TEXTOS_SENHA_EXPIRADA = [
    "your password has expired",
    "sua senha expirou",
    "password expired",
    "update your password",
    "atualize sua senha",
    "change your password",
    "altere sua senha",
    "aadsts50055",
    "aadsts50056",
]

def _detectar_senha_expirada(page):
    """Retorna True se o conteúdo da página indica que a senha expirou."""
    try:
        conteudo = page.content().lower()
        return any(t in conteudo for t in _TEXTOS_SENHA_EXPIRADA)
    except Exception:
        return False

def _registrar_log_expiracao():
    msg = (
        f"{datetime.now()} - [SENHA EXPIRADA] "
        "Login falhou porque a senha da conta de servico expirou. "
        "O bot exibiu o popup de atualizacao de senha. "
        "Atualize a senha no popup ou edite manualmente o config.ini.\n"
    )
    with open(LOG_ERRO, "a", encoding="utf-8") as f:
        f.write(msg)
    print(msg.strip())


# LOGIN COM E-MAIL E SENHA

def realizar_login_credenciais(page, tentativa_recursiva=False):
    """
    Preenche e-mail e senha após redirecionamento SSO Microsoft.
    Se detectar senha expirada, abre o popup de atualização e tenta novamente.
    """
    email = CREDENCIAIS["email"]
    senha = CREDENCIAIS["senha"]

    try:
        update_status("Inserindo e-mail...")
        campo_email = page.locator(
            "input[type='email'], input[name='loginfmt'], input[id='i0116']"
        )
        campo_email.wait_for(state="visible", timeout=15000)
        campo_email.fill(email)
        page.locator("input[type='submit'], button[type='submit']").first.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(2000)
    except Exception as e:
        raise Exception(f"Falha ao preencher e-mail: {e}")

    try:
        update_status("Inserindo senha...")
        campo_senha = page.locator(
            "input[type='password'], input[name='passwd'], input[id='i0118']"
        )
        campo_senha.wait_for(state="visible", timeout=15000)
        campo_senha.fill(senha)
        page.locator("input[type='submit'], button[type='submit']").first.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(2000)
    except Exception as e:
        raise Exception(f"Falha ao preencher senha: {e}")

    if _detectar_senha_expirada(page):
        _registrar_log_expiracao()

        if tentativa_recursiva:
            raise Exception(
                "A nova senha também foi rejeitada. "
                "Verifique se a senha foi trocada corretamente no portal Microsoft."
            )

        update_status("⚠️ Senha expirada! Aguardando nova senha...")
        solicitar_atualizacao_senha()
        _nova_senha_disponivel.wait(timeout=300)

        if _cancelar_automacao.is_set():
            raise Exception("Automação cancelada pelo usuário após senha expirada.")

        update_status("Tentando login com nova senha...")
        page.goto(URL_LOGIN)
        page.wait_for_load_state("networkidle")
        try:
            page.click(f"text={SSO_BUTTON_TEXT}", timeout=8000)
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(2000)
        except Exception:
            pass
        realizar_login_credenciais(page, tentativa_recursiva=True)
        return

    try:
        btn_nao = page.locator(
            "input[id='idBtn_Back'], button:has-text('Não'), button:has-text('No')"
        )
        if btn_nao.count() > 0:
            btn_nao.first.click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(1000)
    except Exception:
        pass


# AUTOMAÇÃO

def acessar_sso():
    try:
        update_status("Iniciando...")

        with sync_playwright() as p:
            context = iniciar_contexto(p)
            page    = context.new_page()
            page.set_default_timeout(60000)

            _set_step(0, "active")
            update_status("Realizando login...")
            page.goto(URL_LOGIN)
            page.wait_for_load_state("networkidle")

            try:
                page.click(f"text={SSO_BUTTON_TEXT}", timeout=8000)
                page.wait_for_load_state("networkidle")
                page.wait_for_timeout(2000)
            except Exception:
                pass

            current_url = page.url
            if "login.microsoftonline.com" in current_url or "login.microsoft.com" in current_url:
                realizar_login_credenciais(page)

            update_status("Verificando MFA...")
            mfa_tratado = verificar_mfa(page)

            if mfa_tratado:
                update_status("MFA aprovado!")
                page.wait_for_load_state("networkidle")
                page.wait_for_timeout(3000)
            else:
                update_status("Sessão autenticada!")
                page.wait_for_timeout(5000)

            _set_step(0, "done")

            _set_step(1, "active")
            update_status("Acessando relatório...")
            page.goto(URL_PARCELAS)
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(5000)

            page.click("i.zmdi.zmdi-more-vert")
            page.click(f"text={NOME_RELATORIO}")

            update_status("Solicitando exportação CSV...")
            page.click("i.zmdi.zmdi-print")
            page.click("text=CSV em lote")
            page.click("button:has-text('Confirmar')")
            page.click("button:has-text('Ok')")

            _set_step(1, "done")

            _set_step(2, "active")
            update_status("Aguardando relatório ficar pronto...")
            page.goto(URL_EMITIDOS)
            page.wait_for_load_state("networkidle")
            page.wait_for_selector("table", timeout=30000)

            tentativas     = 0
            max_tentativas = 30

            while tentativas < max_tentativas:
                status = page.locator("td.regular.text-green:has-text('Concluído')")
                if status.count() > 0:
                    break
                tentativas += 1
                set_tentativa(tentativas, max_tentativas)
                page.wait_for_timeout(20000)
                page.reload()
                page.wait_for_load_state("networkidle")
                try:
                    page.wait_for_selector("table", timeout=15000)
                except Exception:
                    pass

            if tentativas >= max_tentativas:
                raise Exception("Timeout aguardando geração do relatório.")

            _set_step(2, "done")

            _set_step(3, "active")
            update_status("Baixando arquivo...")

            row = page.locator("tr:has(td.regular.text-green:has-text('Concluído'))").first
            row.hover()
            page.wait_for_timeout(1000)
            row.locator("i.zmdi.zmdi-more-vert").click()

            with page.expect_download(timeout=120000) as download_info:
                page.click("text=Fazer download")

            download = download_info.value
            data     = datetime.now().strftime("%Y_%m_%d")
            nome     = f"relatorio_parcelas_{data}.csv"
            caminho  = os.path.join(PASTA_DOWNLOAD, nome)
            download.save_as(caminho)

            if not os.path.exists(caminho):
                raise Exception("Arquivo CSV não foi salvo.")

            _set_step(3, "done")
            update_status("Finalizado!")

            context.close()
            mostrar_popup_sucesso(nome)
            return True

    except Exception as e:
        with open(LOG_ERRO, "a", encoding="utf-8") as f:
            f.write(f"{datetime.now()} - {str(e)}\n")
        popup_erro(str(e))
        return False


# MAIN

def executar_automacao():
    sucesso = acessar_sso()
    print("✅ Processo finalizado com sucesso!" if sucesso else "❌ Processo finalizado com erros.")


if __name__ == "__main__":
    automacao_thread = threading.Thread(target=executar_automacao)
    automacao_thread.start()
    iniciar_popup()
    automacao_thread.join()
    sys.exit(0)
