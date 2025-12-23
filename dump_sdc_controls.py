
# -*- coding: utf-8 -*-
"""
Inspección de UI del SDC (UNICON - Módulo de ALMACEN):
- Conecta robustamente a la ventana (UIA/Win32)
- Recorre toda la jerarquía de controles
- Exporta a JSON y CSV con propiedades útiles
- Localiza controles clave: 'Guía:', 'Buscar', 'Obtener PDF'

Ejecuta:
  python dump_sdc_controls.py

Los archivos se guardan en la misma carpeta del script:
  - SDC_UI_dump.json
  - SDC_UI_dump.csv
"""

import json
import csv
import os
from datetime import datetime
from pywinauto import Application, Desktop
from pywinauto.findwindows import ElementNotFoundError

# ---------- Conexión robusta ----------
def _match_sdc_title(title: str) -> bool:
    if not title:
        return False
    t = title.lower()
    return ("unicon" in t) and ("módulo de almacen" in t or "modulo de almacen" in t)

def conectar_sdc():
    for backend in ("uia", "win32"):
        try:
            desktop = Desktop(backend=backend)
            candidates = [w for w in desktop.windows() if _match_sdc_title(w.window_text())]
            best = None
            for w in candidates:
                title = w.window_text().lower()
                if "almacen" in title:
                    best = w
                    break
            if not best and candidates:
                best = candidates[0]

            if best:
                app = Application(backend=backend)
                app.connect(handle=best.handle)
                win = app.window(handle=best.handle)
                try: win.set_focus()
                except: pass
                try: win.restore()
                except: pass
                try: win.maximize()
                except: pass
                print(f"[INFO] Conectado a: '{best.window_text()}' (backend={backend})")
                return app, win, backend
        except Exception:
            continue
    raise RuntimeError("No pude conectar a la ventana del SDC. Asegura que esté abierta/visible.")

# ---------- Utilidades ----------
def _safe_get(obj, attr, default=None):
    try:
        return getattr(obj, attr)
    except Exception:
        return default

def _rect_to_dict(rect):
    try:
        return {"left": rect.left, "top": rect.top, "right": rect.right, "bottom": rect.bottom}
    except Exception:
        return None

def _runtime_id_str(ei):
    rid = _safe_get(ei, "runtime_id", None)
    if rid is None:
        return None
    try:
        return "-".join(str(x) for x in rid)
    except Exception:
        return str(rid)

def _ctrl_to_dict(ctrl, depth, path):
    """Convierte un control en dict con propiedades útiles."""
    try:
        ei = ctrl.element_info
    except Exception:
        ei = None

    title = None
    try:
        title = ctrl.window_text()
    except Exception:
        pass

    rect = None
    try:
        rect = ctrl.rectangle()
    except Exception:
        rect = None

    d = {
        "depth": depth,
        "path": " > ".join(str(i) for i in path),
        "title": title,
        "control_type": _safe_get(ei, "control_type", None),
        "name": _safe_get(ei, "name", None),
        "automation_id": _safe_get(ei, "automation_id", None),
        "class_name": _safe_get(ei, "class_name", None),
        "framework_id": _safe_get(ei, "framework_id", None),
        "handle": _safe_get(ei, "handle", None),
        "runtime_id": _runtime_id_str(ei),
        "rect": _rect_to_dict(rect),
        "visible": _safe_get(ctrl, "is_visible", lambda: None)(),
        "enabled": _safe_get(ctrl, "is_enabled", lambda: None)(),
    }
    return d

def _walk_tree(root_ctrl, depth=0, path=None, out_list=None):
    """Recorre el árbol por children() y acumula dicts."""
    if path is None:
        path = []
    if out_list is None:
        out_list = []

    try:
        out_list.append(_ctrl_to_dict(root_ctrl, depth, path))
    except Exception:
        pass

    # hijos directos
    children = []
    try:
        children = root_ctrl.children()
    except Exception:
        children = []

    for idx, ch in enumerate(children):
        _walk_tree(ch, depth + 1, path + [idx], out_list)

    return out_list

# ---------- Exportación ----------
def export_json(data, filename):
    with open(filename, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"[OK] JSON guardado: {filename} (controles: {len(data)})")

def export_csv(data, filename):
    fields = [
        "depth","path","title","name","control_type","automation_id",
        "class_name","framework_id","handle","runtime_id",
        "rect","visible","enabled"
    ]
    with open(filename, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        for d in data:
            d2 = dict(d)
            # serializar rect como 'left,top,right,bottom'
            rect = d2.get("rect")
            if rect:
                d2["rect"] = f"{rect['left']},{rect['top']},{rect['right']},{rect['bottom']}"
            else:
                d2["rect"] = ""
            w.writerow(d2)
    print(f"[OK] CSV guardado: {filename} (controles: {len(data)})")

# ---------- Barrido dirigido: Guía/Buscar/Obtener PDF ----------
def find_by_text(win, text, types=("Text","Button","Hyperlink","Edit")):
    matches = []
    for ct in types:
        try:
            m = win.child_window(title=text, control_type=ct)
            matches.append(m.wrapper_object())
        except ElementNotFoundError:
            pass
        try:
            m = win.child_window(title_re=text, control_type=ct)
            matches.append(m.wrapper_object())
        except ElementNotFoundError:
            pass
    return matches

def vecinos_en_misma_fila(win, anchor_ctrl, lado="derecha", tolerancia_y=24):
    """Encuentra controles vecinos (por ejemplo, los Edit a la derecha del label 'Guía:')."""
    if not anchor_ctrl:
        return []
    try:
        a_rect = anchor_ctrl.rectangle()
    except Exception:
        return []
    cy = (a_rect.top + a_rect.bottom) // 2

    vecinos = []
    # Buscamos Edits y Buttons en toda la ventana
    cand = []
    for c in win.descendants():
        try:
            ci = c.element_info
            ct = getattr(ci, "control_type", "")
            if ct in ("Edit","Button","Hyperlink","Text"):
                cand.append(c.wrapper_object())
        except Exception:
            continue

    for w in cand:
        try:
            r = w.rectangle()
        except Exception:
            continue
        mismo_renglon = abs(((r.top + r.bottom) // 2) - cy) <= tolerancia_y
        if not mismo_renglon:
            continue
        if lado == "derecha" and r.left >= a_rect.right - 5:
            vecinos.append((r.left, w))
        elif lado == "izquierda" and r.right <= a_rect.left + 5:
            vecinos.append((r.left, w))

    vecinos.sort(key=lambda t: t[0])
    return [w for _, w in vecinos]

# ---------- Main ----------
def main():
    app, win, backend = conectar_sdc()

    root = win.wrapper_object()
    print("[INFO] Recorriendo árbol de controles… (puede tardar unos segundos)")
    data = _walk_tree(root)

    # nombres de salida
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    json_file = f"SDC_UI_dump_{stamp}.json"
    csv_file = f"SDC_UI_dump_{stamp}.csv"

    export_json(data, json_file)
    export_csv(data, csv_file)

    # Vista rápida en consola (primeros 30)
    print("\n[PREVIEW] Primeros 30 controles:")
    for i, d in enumerate(data[:30], start=1):
        print(f"{i:02d}. depth={d['depth']:2d} type={d['control_type']:<12} title={d['title']} name={d['name']} autoId={d['automation_id']} rect={d['rect']}")

    # Barrido dirigido: Guía / Buscar / Obtener PDF
    print("\n[SCAN] Buscando 'Guía:'…")
    guia_labels = find_by_text(win, r"Gu[ií]a\s*:?", types=("Text",))
    if guia_labels:
        lbl = guia_labels[0]
        print(" - Label 'Guía:' detectado.")
        vecinos = vecinos_en_misma_fila(win, lbl, lado="derecha", tolerancia_y=24)
        edits = [v for v in vecinos if getattr(v.element_info, "control_type", "") == "Edit"]
        print(f" - Edits a la derecha del label: {len(edits)}")
        for idx, ed in enumerate(edits):
            ei = ed.element_info
            print(f"   [{idx}] Edit name={ei.name} autoId={ei.automation_id} rect={ed.rectangle()}")
    else:
        print(" - No se detectó el label 'Guía:' como Text.")

    print("\n[SCAN] Buscando 'Buscar'…")
    buscar_btns = find_by_text(win, r"Buscar", types=("Button",))
    for b in buscar_btns[:1]:
        print(f" - Botón 'Buscar': rect={b.rectangle()} autoId={b.element_info.automation_id}")

    print("\n[SCAN] Buscando 'Obtener PDF'…")
    pdf_ctls = find_by_text(win, r"Obtener\s*PDF", types=("Hyperlink","Text","Button"))
    for p in pdf_ctls[:1]:
        print(f" - 'Obtener PDF': type={p.element_info.control_type} rect={p.rectangle()} autoId={p.element_info.automation_id}")

    print("\n[INFO] Listo. Revisa los archivos JSON/CSV generados para más detalles.")

if __name__ == "__main__":
    main()
