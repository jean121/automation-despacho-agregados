
from pywinauto import Application, Desktop
from pywinauto.timings import TimeoutError
import re

TITULO_EXACTO = "UNICON  - Módulo de PEDIDOS_DISTRIBUCION - AGREGADOS"

def resolver_dialogo(titulo_objetivo=TITULO_EXACTO, backend="uia", timeout=10):
    """
    Devuelve el wrapper de la ventana principal usando distintas estrategias:
    - título exacto
    - regex con espacios flexibles
    - búsqueda por substrings 'UNICON' y 'PEDIDOS_DISTRIBUCION'
    """
    # 1) Intento por título exacto
    try:
        app = Application(backend=backend).connect(title=titulo_objetivo, timeout=timeout)
        dlg = app.window(title=titulo_objetivo)
        dlg.wait("visible", timeout=timeout)
        return app, dlg
    except TimeoutError:
        pass

    # 2) Regex flexible (maneja espacios variables alrededor de los guiones)
    #    Nota: el título tiene dos espacios antes del guion: "UNICON  - ..."
    patron = r"^UNICON\s+-\s+Módulo de PEDIDOS_DISTRIBUCION - AGREGADOS$"
    try:
        app = Application(backend=backend).connect(title_re=patron, timeout=timeout)
        dlg = app.window(title_re=patron)
        dlg.wait("visible", timeout=timeout)
        return app, dlg
    except TimeoutError:
        pass

    # 3) Enumeración y selección por substrings
    d = Desktop(backend=backend)
    candidatos = []
    for w in d.windows(visible_only=True):
        t = w.window_text()
        if "UNICON" in t and "PEDIDOS_DISTRIBUCION" in t and "AGREGADOS" in t:
            candidatos.append(w)

    if candidatos:
        # Tomamos el primero. Si hay varios, puedes filtrar por clase o pid.
        w = candidatos[0]
        pid = w.element_info.process_id
        app = Application(backend=backend).connect(process=pid, timeout=timeout)
        # Usar best_match puede ayudar si hay variaciones minúsculas del título:
        dlg = app.window(best_match=w.window_text())
        dlg.wait("visible", timeout=timeout)
        return app, dlg

    raise RuntimeError(f"No se pudo resolver la ventana con backend='{backend}'. "
                       f"Verifica el título exacto o prueba el otro backend ('win32'/'uia').")


def imprimir_controles(app, dlg):
    print(f"\n=== Árbol de controles de: {dlg.window_text()} ===\n")
    dlg.print_control_identifiers()

    print("\n=== Lista programática de controles ===\n")
    for c in dlg.descendants():
        info = c.element_info
        rect = info.rectangle
        print(
            f"control_type={info.control_type:<12} "
            f"name='{info.name}' "
            f"auto_id='{info.automation_id}' "
            f"class='{info.class_name}' "
            f"handle={info.handle} "
            f"rect=({rect.left},{rect.top},{rect.right},{rect.bottom})"
        )

if __name__ == "__main__":
    # Prueba primero con 'uia'. Si no detecta, cambia a 'win32'.
    backend = "uia"
    app, dlg = resolver_dialogo(backend=backend)
    imprimir_controles(app, dlg)
