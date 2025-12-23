
# auto_click.py
import time
import threading
import pyautogui
from pynput import keyboard

# Seguridad: si llevas el mouse a la esquina superior izquierda, se detiene con excepción
pyautogui.FAILSAFE = True

INTERVALO = 60.0  # segundos entre clics
clicking = False  # estado de clics (iniciar/pausar)
running = True    # estado del programa (salir)

def click_loop():
    """Bucle que hace clic cada INTERVALO segundos cuando clicking=True."""
    global clicking, running
    while running:
        if clicking:
            pyautogui.click()  # clic en la posición actual del mouse
            time.sleep(INTERVALO)
        else:
            time.sleep(0.1)

def on_press(key):
    """Gestión de teclas: F8 inicia/pausa; Esc sale."""
    global clicking, running
    if key == keyboard.Key.f8:
        clicking = not clicking
        print("▶️ Iniciando auto-clic..." if clicking else "⏸️ Pausado.")
    elif key == keyboard.Key.esc:
        running = False
        print("👋 Saliendo...")
        # Detiene el listener de teclado
        return False

if __name__ == "__main__":
    print("Deja el mouse donde quieres hacer clic cada 1s.")
    print("Atajos: F8 = iniciar/pausar | Esc = salir")
    print("Emergencia: mueve el cursor a la esquina superior izquierda (pyautogui.FAILSAFE)")

    # Hilo que hace los clics
    t = threading.Thread(target=click_loop, daemon=True)
    t.start()

    # Listener de teclado
    with keyboard.Listener(on_press=on_press) as listener:
        listener.join()

    print("Programa terminado.")
