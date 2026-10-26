# auto_updater.py
import os
import sys
import json
import threading
import urllib.request
import subprocess
from tkinter import messagebox

# ==========================================
# CONFIGURACIÓN DEL UPDATER
# ==========================================
VERSION_ACTUAL = "1.0.0"

# URL directa en tu repositorio de GitHub:
URL_VERSION_JSON = "https://raw.githubusercontent.com/davidrf-KC/KeyCapitalHub/main/version.json"


def es_version_mayor(remota, local):
    """Compara si la versión remota (ej. '1.0.1') es superior a la local ('1.0.0')."""
    try:
        partes_remota = [int(x) for x in remota.strip().split(".")]
        partes_local = [int(x) for x in local.strip().split(".")]
        return partes_remota > partes_local
    except Exception:
        return remota != local


def consultar_remoto(url_json=URL_VERSION_JSON, timeout=5):
    """Descarga el JSON con la información de la última versión disponible."""
    try:
        req = urllib.request.Request(
            url_json,
            headers={"User-Agent": "KeyCapitalUpdater/1.0", "Cache-Control": "no-cache"}
        )
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except Exception:
        return None


def reiniciar_app():
    """Relanza la aplicación en caso de ejecución directa desde Python."""
    if getattr(sys, 'frozen', False):
        subprocess.Popen([sys.executable] + sys.argv[1:])
    else:
        subprocess.Popen([sys.executable, os.path.abspath(sys.argv[0])] + sys.argv[1:])
    sys.exit(0)


def actualizar_exe(url_exe, ruta_exe_actual):
    """
    Descarga el nuevo binario y utiliza un script .bat temporal
    para reemplazar el .exe bloqueado en Windows y relanzarlo.
    """
    try:
        ruta_temp = ruta_exe_actual + ".new"
        urllib.request.urlretrieve(url_exe, ruta_temp)

        if os.path.getsize(ruta_temp) < 10000:
            raise ValueError("El archivo descargado no parece un ejecutable válido.")

        ruta_bat = os.path.join(os.path.dirname(ruta_exe_actual), "reemplazar_app.bat")
        script_bat = f"""@echo off
timeout /t 2 /nobreak > nul
del /f /q "{ruta_exe_actual}"
move /y "{ruta_temp}" "{ruta_exe_actual}"
start "" "{ruta_exe_actual}"
del "%~f0"
"""
        with open(ruta_bat, "w", encoding="utf-8") as f:
            f.write(script_bat)

        messagebox.showinfo(
            "Actualización Lista",
            "La nueva versión se descargó correctamente. La app se reiniciará ahora."
        )

        subprocess.Popen(["cmd.exe", "/c", ruta_bat], creationflags=subprocess.CREATE_NO_WINDOW)
        sys.exit(0)

    except Exception as e:
        messagebox.showerror("Error de Actualización", f"No se pudo instalar la actualización:\n{e}")


def actualizar_py(url_script, ruta_py_actual):
    """Actualiza directamente el script .py si no está compilado."""
    try:
        ruta_temp = ruta_py_actual + ".temp"
        urllib.request.urlretrieve(url_script, ruta_temp)

        if os.path.getsize(ruta_temp) < 1000:
            raise ValueError("El script descargado parece inválido.")

        os.replace(ruta_temp, ruta_py_actual)
        messagebox.showinfo("Actualización Exitosa", "Código actualizado. La aplicación se reiniciará.")
        reiniciar_app()

    except Exception as e:
        messagebox.showerror("Error de Actualización", f"No se pudo actualizar el script:\n{e}")


def buscar_y_aplicar(silencioso=False):
    """Comprueba si hay actualizaciones y pregunta al usuario antes de descargar."""
    datos = consultar_remoto()
    if not datos:
        if not silencioso:
            messagebox.showwarning("Sin Conexión", "No se pudo contactar al servidor de actualizaciones de GitHub.")
        return

    version_remota = datos.get("version", "1.0.0")
    descripcion = datos.get("descripcion", "Mejoras generales del sistema.")

    if es_version_mayor(version_remota, VERSION_ACTUAL):
        msg = (f"Nueva versión disponible: v{version_remota}\n"
               f"Versión instalada: v{VERSION_ACTUAL}\n\n"
               f"Novedades:\n{descripcion}\n\n"
               "¿Deseas descargarla e instalarla ahora?")

        if messagebox.askyesno("Nueva Versión Disponible", msg):
            es_exe = getattr(sys, 'frozen', False)
            ruta_propia = os.path.abspath(sys.argv[0])

            if es_exe:
                actualizar_exe(datos.get("url_descarga_exe"), ruta_propia)
            else:
                actualizar_py(datos.get("url_descarga"), ruta_propia)
    else:
        if not silencioso:
            messagebox.showinfo("Aplicación Actualizada", f"Ya tienes la última versión instalada (v{VERSION_ACTUAL}).")


def iniciar_busqueda_fondo():
    """Ejecuta la verificación en un hilo secundario sin congelar la ventana."""
    threading.Thread(target=lambda: buscar_y_aplicar(silencioso=True), daemon=True).start()