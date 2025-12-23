
# -*- coding: utf-8 -*-
"""
Automatización SDC (RemoteApp/RDP) por imagen con 'focus hard' basado en ENTER (sin clic)
+ retorno robusto a la ventana principal tras abrir el PDF.

Flujo:
- Conecta a "UNICON - Módulo de ALMACEN ..."
- Trae la ventana al frente con restore/maximize/set_focus y ENTER
- TAB x15 hasta "Guía 7 dígitos", limpia, escribe sufijo (7 dígitos) y Enter para buscar
- Localiza "Obtener PDF" por imagen (región de la ventana) y hace clic
- Espera la apertura del PDF y **restituye el foco al SDC** (ENTER + ALT+TAB / ALT+ESC)
- Repite para el rango

Requisitos:
  python -m pip install --upgrade pillow pyscreeze opencv-python pyautogui pywinauto

Consejos:
- Mantén la escala de Windows al 100%
- No muevas/redimensiones la ventana del SDC mientras corre
"""

import os
import time
import ctypes
from datetime import datetime
import pyautogui
from pywinauto import Application, Desktop
from pywinauto.keyboard import send_keys

# ============== CONFIGURACIÓN ==============
GUIA_PREFIJO_FIJO = "195"     # lo estableces manualmente en el campo 'Guía (prefijo)' antes de ejecutar
GUIA_INICIO = 184241          # ej.: 184241 -> se convertirá en "0184241"
GUIA_FIN    = 184243

# Tabs según tu mapeo (AJUSTADO A 15)
TABS_PREFIJO_A_7D = 15        # de 'Guía (prefijo)' -> 'Guía (7 dígitos)'

# Imagen del link "Obtener PDF" (captura nítida SOLO del texto)
IM_OBTENER_PDF = r"C:\Users\ealpiste\Documents\Mis scripts\imgs\obtener_pdf.png"

# Parámetros de búsqueda por imagen
CONFIDENCE_START = 0.94       # empezamos alto
CONFIDENCE_MIN   = 0.86       # bajamos gradualmente hasta aquí
CONFIDENCE_STEP  = 0.02
GRAYSCALE_SEARCH = True       # mejora si varía levemente el color

# Tiempos de espera (ajusta si tu red/PC tardan más)
WAIT_AFTER_SEARCH   = 0.9     # tras Enter (refresco de grilla)
WAIT_AFTER_PDF_OPEN = 2.8     # tras abrir el PDF antes de volver al SDC
RETRIES_IMG         = 4       # reintentos por nivel de confianza

# Debug: guarda captura de la región de la ventana si no encuentra la imagen
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DEBUG_SHOT = os.path.join(SCRIPT_DIR, "debug_window_region.png")

# ============== UTILIDADES DE VENTANA/Foco ==============
user32 = ctypes.windll.user32

def _match_sdc_title(title: str) -> bool:
    """Identifica la ventana del RemoteApp por substrings tolerantes."""
    return bool(title) and ("unicon" in title.lower()) and ("almacen" in title.lower())

def conectar_sdc():
    """Conecta a la ventana principal del SDC (ALMACEN) probando UIA/Win32."""
    for backend in ("uia", "win32"):
        try:
            desktop = Desktop(backend=backend)
            candidates = [w for w in desktop.windows() if _match_sdc_title(w.window_text())]
            if candidates:
                target = candidates[0]
                app = Application(backend=backend).connect(handle=target.handle)
                win = app.window(handle=target.handle)
                print(f"[INFO] Conectado a: '{target.window_text()}' (backend={backend}, handle={target.handle})")
                return app, win
        except Exception:
            continue
    raise RuntimeError("No pude conectar al SDC (ALMACEN). Verifica que esté visible y maximizado.")

def get_window_region(win):
    """Región de la ventana (left, top, width, height) para locateOnScreen."""
    rect = win.rectangle()
    return (rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top)

def get_foreground_handle():
    return user32.GetForegroundWindow()

def is_sdc_foreground(win):
    try:
        return get_foreground_handle() == win.handle
    except Exception:
        return False

def focus_window_hard_enter(win, retries=4, pause=0.2):
    """
    Trae la ventana al frente de forma 'hard' basada en ENTER (sin clic):
    - restore
    - maximize
    - set_focus
    - enviar ENTER para afianzar el foco (simula interacción)
    Retorna True si logra poner SDC en foreground; False si no.
    """
    for i in range(retries):
        try: win.restore()
        except: pass
        try: win.maximize()
        except: pass
        try:
            win.set_focus()
            time.sleep(pause)
        except:
            pass
        # Empujón de interacción: ENTER (sin clic)
        try:
            send_keys("{ENTER}")
            time.sleep(pause)
        except Exception:
            pass

        if is_sdc_foreground(win):
            return True
        time.sleep(pause)
    return False

# ============== Envío de teclas ==============
DEBUG_DELAY = 0.20  # sube/baja para observar cada paso más claro

def ensure_sdc_and_send_keys_hard(win, keys: str, desc: str = ""):
    """
    Verifica/forcea foco con focus_window_hard_enter (sin clic),
    luego envía teclas y espera DEBUG_DELAY para observar en depuración.
    """
    ok = is_sdc_foreground(win)
    if not ok:
        ok = focus_window_hard_enter(win)
        # print(f"[FOCUS] {'OK' if ok else 'FAIL'} foco antes de '{desc or keys}' (fg={get_foreground_handle()}, sdc={win.handle})")
    send_keys(keys)
    # print(f"[KEYS] Enviadas: {keys}  ({desc})")
    time.sleep(DEBUG_DELAY)

def tab_hard(win, n: int, desc: str = "Tabs"):
    for i in range(n):
        ensure_sdc_and_send_keys_hard(win, "{TAB}", f"{desc} ({i+1}/{n})")

# ============== Imagen: "Obtener PDF" ==============
def save_debug_region(region):
    left, top, width, height = region
    try:
        im = pyautogui.screenshot(region=region)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        path = DEBUG_SHOT.replace(".png", f"_{ts}.png")
        im.save(path)
        print(f"[DEBUG] Captura guardada: {path}")
    except Exception as e:
        print(f"[DEBUG] No se pudo guardar captura: {e}")

def click_obtener_pdf_por_imagen(win, img_path):
    """
    Busca la imagen 'Obtener PDF' dentro de la región de la ventana y hace clic.
    Ajusta la confianza de mayor a menor; reintenta varias veces por nivel de confianza.
    Si no encuentra, guarda captura de región para diagnóstico.
    """
    if not os.path.isfile(img_path):
        raise FileNotFoundError(f"No existe la imagen: {img_path}")

    # Asegurar foco con ENTER antes de buscar imagen
    focus_window_hard_enter(win)

    region = get_window_region(win)
    conf = CONFIDENCE_START

    while conf >= CONFIDENCE_MIN:
        for i in range(RETRIES_IMG):
            try:
                box = pyautogui.locateOnScreen(img_path, confidence=conf, region=region, grayscale=GRAYSCALE_SEARCH)
            except Exception:
                box = None

            if box:
                x, y = pyautogui.center(box)
                pyautogui.click(x, y)
                # print(f"[IMG] Click 'Obtener PDF' (conf={conf}, intento={i+1}) en {x},{y}")
                time.sleep(DEBUG_DELAY)
                return True

            time.sleep(0.25)

        conf = round(conf - CONFIDENCE_STEP, 2)

    # Último intento en toda la pantalla
    try:
        box = pyautogui.locateOnScreen(img_path, confidence=CONFIDENCE_MIN, grayscale=GRAYSCALE_SEARCH)
    except Exception:
        box = None

    if box:
        x, y = pyautogui.center(box)
        pyautogui.click(x, y)
        # print(f"[IMG] Click global 'Obtener PDF' (conf={CONFIDENCE_MIN}) en {x},{y}")
        time.sleep(DEBUG_DELAY)
        return True

    save_debug_region(region)
    return False

# ============== Retorno robusto a SDC tras abrir PDF ==============
def return_to_sdc(win, timeout=6.0):
    """
    Asegura volver a la ventana del SDC tras abrir el PDF:
    1) Espera WAIT_AFTER_PDF_OPEN
    2) Intento de foco 'hard' con ENTER
    3) Si falla, ALT+TAB en bucle corto y verifica foreground
    4) Si aún falla, ALT+ESC (ciclo rápido de ventanas)
    """
    t0 = time.time()
    time.sleep(WAIT_AFTER_PDF_OPEN)

    # 1) Intento directo
    if focus_window_hard_enter(win):
        return True

    # 2) ALT+TAB hasta recuperar (máx 5 intentos)
    for _ in range(5):
        send_keys("%{TAB}")   # ALT+TAB
        time.sleep(0.25)
        if focus_window_hard_enter(win):
            return True
        if time.time() - t0 > timeout:
            break

    # 3) ALT+ESC (ciclo rápido de ventanas)
    for _ in range(5):
        send_keys("%{ESC}")   # ALT+ESC
        time.sleep(0.25)
        if focus_window_hard_enter(win):
            return True
        if time.time() - t0 > timeout:
            break

    # 4) Último intento directo
    return focus_window_hard_enter(win)

# ============== FLUJO PRINCIPAL ==============
def main():
    if not os.path.isfile(IM_OBTENER_PDF):
        raise FileNotFoundError(f"Imagen 'Obtener PDF' no existe: {IM_OBTENER_PDF}")

    app, win = conectar_sdc()

    inicio, fin = sorted((GUIA_INICIO, GUIA_FIN))
    procesadas, errores = 0, []

    for sfx in range(inicio, fin + 1):
        sufijo_7d = f"{sfx:07d}"
        print(f"\n[INFO] Procesando guía: {GUIA_PREFIJO_FIJO}-{sufijo_7d}")

        # 1) Foco 'hard' con ENTER al inicio
        _ = focus_window_hard_enter(win)
        time.sleep(DEBUG_DELAY)

        # 2) Ir de 'Guía (prefijo)' a 'Guía 7 dígitos' (TAB x15)
        tab_hard(win, TABS_PREFIJO_A_7D, desc="A Guía 7 dígitos")

        # 3) Escribir sufijo y Enter para 'Buscar'
        ensure_sdc_and_send_keys_hard(win, "^a{BACKSPACE}", "Limpiar 7 dígitos")
        ensure_sdc_and_send_keys_hard(win, sufijo_7d, "Escribir 7 dígitos")
        ensure_sdc_and_send_keys_hard(win, "{ENTER}", "Buscar (Enter)")
        time.sleep(WAIT_AFTER_SEARCH)

        # 4) Click en 'Obtener PDF' por imagen
        if not click_obtener_pdf_por_imagen(win, IM_OBTENER_PDF):
            msg = f"No se encontró 'Obtener PDF' (guía {sufijo_7d}). Revisa debug_window_region_*.png y la plantilla."
            print("[WARN]", msg)
            errores.append(msg)
            continue

        # 5) Retornar de forma robusta a SDC (evitar que los TABs se queden en IE/Edge)
        if not return_to_sdc(win):
            print("[WARN] No pude recuperar foco del SDC tras abrir PDF. Continuaré intentando en la próxima guía.")
        else:
            procesadas += 1

    print(f"\n[RESUMEN] Guías procesadas: {procesadas}")
    if errores:
        print("[ERRORES]")
        for e in errores:
            print(" -", e)

if __name__ == "__main__":
    main()
