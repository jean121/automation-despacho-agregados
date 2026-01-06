
import os
import sys
import tempfile
import subprocess
import requests

# ---------- CONFIGURA AQUÍ SI QUIERES USAR UN NOMBRE DE IMPRESORA ESPECÍFICO ----------
# Deja printer_name = None para usar la impresora predeterminada del sistema.
printer_name = None  # p.ej., "HP_LaserJet" o "Microsoft Print to PDF"

# ---------- URL BASE Y PARÁMETROS (forma segura con requests) ----------
BASE_URL = "https://ereceipt-pe-s02.sovos.com/Facturacion/PDFServlet"
PARAMS = {
    "id": "EpQ(MaS)REwOaHPS3n28O209gg(IgU)(IgU)",
    "o": "E",
}

# ---------- CABECERAS (algunos servidores requieren user-agent/accept) ----------
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept": "application/pdf,*/*;q=0.8",
}

def assert_windows():
    if not sys.platform.startswith("win"):
        raise EnvironmentError("Este script está diseñado únicamente para Windows.")

def descargar_pdf(base_url, params, headers=None, timeout=30):
    """
    Descarga un PDF usando requests con parámetros (así se codifican bien los paréntesis).
    Devuelve la ruta del archivo temporal si es PDF; de lo contrario lanza Exception.
    """
    try:
        resp = requests.get(base_url, params=params, headers=headers, timeout=timeout)
        resp.raise_for_status()

        # Validar tipo de contenido
        content_type = resp.headers.get("Content-Type", "")
        if "application/pdf" not in content_type.lower():
            # A veces no ponen bien el header; aún así intentamos validar por contenido
            # Si el primer bytes coincide con PDF signature "%PDF", aceptamos.
            if not resp.content.startswith(b"%PDF"):
                raise Exception(f"El recurso no parece ser un PDF (Content-Type: {content_type}).")

        # Guardar en archivo temporal
        fd, temp_path = tempfile.mkstemp(suffix=".pdf")
        os.close(fd)
        with open(temp_path, "wb") as f:
            f.write(resp.content)

        return temp_path
    except Exception as e:
        raise Exception(f"Error descargando PDF: {e}")

def _buscar_lector_pdf():
    """
    Busca ejecutables comunes de lectores PDF para impresión silenciosa.
    Prioriza Adobe Acrobat/Reader. Devuelve ruta al ejecutable o None.
    """
    candidatos = [
        # Adobe Reader / Acrobat
        r"C:\Program Files\Adobe\Acrobat Reader\Reader\AcroRd32.exe",
        r"C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
        r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
        # Foxit
        r"C:\Program Files (x86)\Foxit Software\Foxit Reader\FoxitReader.exe",
        r"C:\Program Files\Foxit Software\Foxit PDF Editor\FoxitPDFEditor.exe",
        # SumatraPDF (rápido y soporta impresión por CLI)
        r"C:\Program Files\SumatraPDF\SumatraPDF.exe",
        r"C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe",
    ]
    return next((p for p in candidatos if os.path.exists(p)), None)

def imprimir_windows(pdf_path, printer=None):
    """
    Imprime en Windows.
    - Si hay Acrobat/Reader (u otro lector reconocido) y se especifica 'printer', intenta impresión silenciosa.
    - Si no, usa os.startfile(pdf_path, 'print') para la impresora predeterminada.
    """
    lector_exe = _buscar_lector_pdf()

    try:
        if printer and lector_exe:
            exe_name = os.path.basename(lector_exe).lower()

            if "acro" in exe_name:
                # Acrobat/Reader: /t archivo impresora
                subprocess.run([lector_exe, "/t", pdf_path, printer], check=True)
            elif "sumatra" in exe_name:
                # SumatraPDF: -print-to "<printer>" "<file>"
                subprocess.run([lector_exe, "-print-to", printer, pdf_path], check=True)
            elif "foxit" in exe_name:
                # Foxit (puede variar según versión; se intenta modo silencioso)
                # Algunas versiones aceptan: /p=printer /t "file"
                subprocess.run([lector_exe, "/t", pdf_path, printer], check=True)
            else:
                # Desconocido: intenta default
                os.startfile(pdf_path, "print")
        else:
            # Fallback: imprime con la aplicación asociada al PDF (no navegador) en impresora predeterminada
            os.startfile(pdf_path, "print")
    except Exception as e:
        raise Exception(f"Error imprimiendo en Windows: {e}")

def listar_impresoras():
    """
    Intenta listar impresoras instaladas para ayudar a elegir 'printer_name'.
    Usa PowerShell Get-Printer si está disponible; si falla, intenta WMIC (puede estar deprecado).
    """
    try:
        ps_cmd = ["powershell", "-NoProfile", "-Command", "Get-Printer | Select-Object -ExpandProperty Name"]
        res = subprocess.run(ps_cmd, capture_output=True, text=True, check=True)
        nombres = [line.strip() for line in res.stdout.splitlines() if line.strip()]
        return nombres
    except Exception:
        try:
            wmic_cmd = ["wmic", "printer", "get", "name"]
            res = subprocess.run(wmic_cmd, capture_output=True, text=True, check=True)
            nombres = [line.strip() for line in res.stdout.splitlines()[1:] if line.strip()]
            return nombres
        except Exception:
            return []

def main():
    assert_windows()

    print("Descargando PDF...")
    pdf_path = descargar_pdf(BASE_URL, PARAMS, headers=HEADERS)
    print(f"PDF guardado en: {pdf_path}")

    # Opcional: muestra impresoras detectadas si no definiste 'printer_name'
    if printer_name is None:
        impresoras = listar_impresoras()
        if impresoras:
            print("Impresoras disponibles detectadas:")
            for n in impresoras:
                print(f"  - {n}")
        else:
            print("No se pudieron listar las impresoras (continuando con la predeterminada).")

    print("Enviando a imprimir...")
    imprimir_windows(pdf_path, printer=printer_name)
    print("Impresión enviada correctamente.")

    # Limpieza del archivo temporal (opcional; comenta si quieres conservar el PDF)
    try:
        os.remove(pdf_path)
    except Exception:
        pass

if __name__ == "__main__":
    main()
