
# Automation – Despacho de Agregados

Automatización en **Python** para operar módulos remotos de UNICON/SDC (Citrix/RemoteApp) y acelerar tareas críticas del **despacho de agregados**: lectura de **Excel**, interacción por **teclas**, **imagen** y **OCR**, e **impresión** de guías. Orientado a reducir tiempos y errores en procesos operativos de almacén y distribución.

## 🧩 Qué resuelve

- **Ingreso de pedidos** desde una tabla Excel al módulo remoto de PEDIDOS/DISTRIBUCIÓN por envío de teclas, navegación por TABs y botón identificado por imagen/OCR.
- **Procesamiento de placas**: lee placas de Excel y ejecuta la secuencia de despacho en la ventana remota, con confirmación manual del conductor cuando aplica.
- **Impresión de guías (batch)**: localiza “Obtener PDF” por imagen y envía impresión de 3 copias, con recuperación de foco robusta tras abrir el PDF.

> **Impacto esperado (reemplaza con tus datos reales):**
> - ↓ **Tiempo** por pedido/guía (ej. 60–70%)
> - ↓ **Errores** por tecleo manual
> - ↑ **Consistencia** en turnos y equipos

## 🚀 Scripts principales

- `scripts/pedidos_distribucion.py` – Lee una **tabla Excel** (OpenPyXL), enfoca la ventana remota y navega la UI con `send_keys`, `TAB`s, y **imagen/OCR** para confirmar “Salidas”. Incluye `DRY_RUN` y tolerancias de tiempo para Citrix.
- `scripts/despacho_placas.py` – Extrae **placas** desde una tabla Excel y ejecuta la secuencia de **despacho** (hotkeys, TABs, pegado desde portapapeles), además de utilidades para **conectar/enfocar** la ventana SDC por `pywinauto` (UIA/Win32).
- `scripts/print_guias.py` – Control de foco “hard” (restore/maximize/set_focus + **ENTER**), búsqueda por **imagen** de “Obtener PDF”, y `Ctrl+P` con navegación del diálogo para imprimir múltiples copias; incluye **capturas de depuración** si la imagen no aparece.

## 🛠️ Tecnologías

- **Python** 3.x
- **Automatización UI**: `pywinauto` (UIA/Win32), `pyautogui` (imagen/teclas), `pyscreeze`, `opencv-python`
- **Excel**: `openpyxl` (lectura directa, copia temporal si el archivo está bloqueado)
- **OCR**: `pytesseract` (fallback para localizar botones en pantalla)
- **Utilidades**: `pyperclip` (portapapeles), `ctypes` (foreground), `re`/`time`/`tempfile`/`shutil`

## 📦 Instalación

> Requisitos: Windows + Citrix/RemoteApp. Mantén la **escala de Windows al 100%** para la coincidencia de imágenes.

```bash
# Entorno (ideal en venv)
pip install -r requirements.txt
```

> Nota: `pytesseract` requiere que Tesseract OCR esté instalado en el sistema.

## ⚙️ Configuración

Edita estas constantes según tu entorno (rutas, tabla y tiempos):

**`scripts/pedidos_distribucion.py`**
- `EXCEL_PATH`, `TABLE_NAME`
- `START_ROW_IN_TABLE`, `END_ROW_IN_TABLE`
- `TARGET_COLUMN_INDEX1` (agregado-destino), `TARGET_COLUMN_INDEX2` (cubicaje)
- `WINDOW_TITLE_HINT`, `SALIDAS_IMG_PATH`, `SALIDAS_IMG_CONFIDENCE`
- `DELAY_SHORT/MED/LONG`, `WAIT_AFTER_REFRESH`, `DRY_RUN`

**`scripts/despacho_placas.py`**
- `EXCEL_PATH`, `TABLE_NAME`, `START_ROW_IN_TABLE`, `END_ROW_IN_TABLE`, `TARGET_COLUMN_INDEX`
- Parámetros de ventana remota y navegación: `SHIFT_TABS_A_BOTON_NOMBRE`, `FILTRO_NOMBRE_TEXTO`, `KEY_CONTINUAR`, `DELAY_*`

**`scripts/print_guias.py`**
- `GUIA_PREFIJO_FIJO`, `GUIA_INICIO`, `GUIA_FIN`
- Imagen y tolerancias: `IM_OBTENER_PDF`, `CONFIDENCE_*`, `RETRIES_IMG`, `GRAYSCALE_SEARCH`
- Recuperación de foco: `WAIT_AFTER_SEARCH`, `WAIT_AFTER_PDF_OPEN`

## ▶️ Uso

```bash
# 1) Pedidos (lee Excel y envía a la UI remota)
python scripts/pedidos_distribucion.py

# 2) Placas (lee Excel y ejecuta flujo de despacho)
python scripts/despacho_placas.py

# 3) Guías (batch imprimir PDF/3 copias en SDC)
python scripts/print_guias.py
```

## 📁 Estructura sugerida

```
automation-despacho-agregados/
├─ scripts/
│  ├─ pedidos_distribucion.py
│  ├─ despacho_placas.py
│  └─ print_guias.py
├─ resources/
│  ├─ salidas.png
│  └─ obtener_pdf.png
├─ docs/
│  ├─ demo-pedidos.gif
│  ├─ demo-placas.gif
│  └─ demo-guias.gif
├─ README.md
├─ requirements.txt
└─ .gitignore
```

## 🧪 Calidad y robustez

- Foco y foreground robustos: UIA/Win32 + Alt+Tab + ENTER.
- Lectura de Excel sin abrir Excel (OpenPyXL), con **copias temporales** si el archivo está bloqueado.
- Imagen/OCR con tolerancias de confianza y reintentos; capturas de depuración si no se encuentra el objetivo.
- `DRY_RUN` para validar el flujo sin enviar teclas.

## 🔒 Avisos

- No publiques credenciales ni archivos internos; usa rutas genéricas o ejemplos sintéticos.
- Si el nombre del sistema/empresa es sensible, anonimiza en el README y demos.

## 📄 Licencia

MIT.

---

### (EN) Short Overview for Recruiters

Python RPA project that automates remote UNICON modules through UI keystrokes, image/OCR recognition and window inspection. It reads orders/plates from Excel and prints guides in batch with robust focus recovery in Citrix/RemoteApp environments.

**Tech:** Python · pywinauto · pyautogui · openpyxl · pytesseract · opencv · Windows automation.

**Contact:** jean.alpiste@pucp.pe
