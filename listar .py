
# -*- coding: utf-8 -*-
from pywinauto import Desktop

print("\n[UIA] Ventanas visibles:")
for w in Desktop(backend="uia").windows():
    try:
        if w.is_visible():
            print(" -", w.window_text())
    except Exception:
        pass

print("\n[WIN32] Ventanas visibles:")
for w in Desktop(backend="win32").windows():
    try:
        if w.is_visible():
            print(" -", w.window_text())
    except Exception:
        pass
