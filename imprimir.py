
# -*- coding: utf-8 -*-
import time
from pywinauto import Desktop
from pywinauto.keyboard import send_keys

def get_active_edge_window():
    """Intenta obtener la ventana activa de Edge con diferentes métodos."""
    desk = Desktop(backend="uia")
    # 1) active_window (si está disponible en tu versión)
    try:
        win = desk.active_window()
        if win:
            return win
    except Exception:
        pass
    # 2) top_window (más genérico)
    try:
        win = desk.top_window()
        if win:
            return win
    except Exception:
        pass
    # 3) HWND del foreground con Win32
    try:
        import win32gui  # pip install pywin32 si no lo tienes
        hwnd = win32gui.GetForegroundWindow()
        if hwnd:
            return desk.window(handle=hwnd)
    except Exception:
        pass
    raise RuntimeError("No se pudo obtener la ventana activa (Edge/IE Mode).")

def focus_window_hard_enter(win, pause=0.1):
    """Trae la ventana al frente y asegura foco."""
    try:
        win.restore()
    except Exception:
        pass
    try:
        win.set_focus()
        time.sleep(pause)
    except Exception:
        pass
    # Un ENTER suave a veces ayuda a despejar modales sin foco
    try:
        send_keys("{ENTER}")
    except Exception:
        pass

def print_3_copies_panel_tabs():
    """
    Secuencia rápida para el panel de impresión de Edge:
    - Ctrl+P
    - 4 x TAB (ajústalo según tu panel)
    - escribir '3'
    - ENTER (Imprimir)
    """
    #win = get_active_edge_window()
    #focus_window_hard_enter(win)

    # Abrir el panel de impresión de Edge
    send_keys("^p")
    time.sleep(0.8)

    # Navegar por TABs hasta 'Copias' (ajusta cantidad si no cae)
    for _ in range(4):
        send_keys("{TAB}")
        time.sleep(0.2)

    # Establecer 3 copias
    send_keys("^a{BACKSPACE}3")
    time.sleep(0.2)

    # Confirmar impresión
    send_keys("{ENTER}")

if __name__ == "__main__":
    # Espera breve por si acabas de abrir la pestaña del PDF
    time.sleep(1.5)
    print_3_copies_panel_tabs()
