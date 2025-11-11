"""
Sistema de Gestión de Lista de Espera/Cola de Producción
Conecta con base de datos PostgreSQL en Render
Vinculado por sucursal

Versión 1.6.2 - Optimización para Múltiples Coloristas:
- 🔓 Eliminado bloqueo de iniciar procesos (permite múltiples coloristas trabajar simultáneamente)
- 🔒 Mantenido bloqueo de 4 minutos SOLO para finalizar (evita finalizar procesos incompletos)
- 🛠️ Corregido error de fuente en checkbox (compatible con ttkbootstrap)
- 👥 Optimizado para entornos con múltiples coloristas trabajando diferentes facturas
- ⚡ Mayor fluidez en el inicio de procesos de producción

Versión 1.6.1 - Mejoras de Interfaz de Usuario:
- 📝 Fuentes más grandes en toda la interfaz (mejor legibilidad)
- 📊 Encabezados de tabla con fuente 12pt bold
- 🔤 Contenido de tabla con fuente 11pt 
- 📐 Anchos de columnas aumentados para mejor visualización
- 🏷️ Altura de filas aumentada (30px) para mejor espaciado
- 🧮 Campo "Base" cambiado por "Cantidad" en vista de fórmulas
- 🎨 Mejor experiencia visual general

Versión 1.6.0 - Mejoras de Seguridad y Control de Procesos:
- ✅ Confirmación obligatoria para cancelar pedidos (evita eliminaciones accidentales)
- 🔒 Bloqueo automático de 4 minutos SOLO al finalizar procesos (permite completar correctamente)
- ⚠️ Confirmación para cerrar aplicación con procesos activos
- 📋 Confirmación para iniciar/finalizar listas completas
- 🔔 Indicadores visuales del estado de bloqueo en tiempo real
- 🛡️ Protección contra interrupciones de procesos críticos
"""

# === IMPORTACIONES SEGURAS (NÚCLEO) ===
import tkinter as tk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import messagebox, simpledialog
import json
import base64
import sys
import os
import threading
import time
import tempfile
import math
from datetime import datetime, date, timedelta
import hashlib

# === SISTEMA DE CARGA DIFERIDA PARA EVITAR ERRORES DLL ===
# Variables globales para controlar disponibilidad de módulos
POSTGRES_AVAILABLE = False
WIN32_AVAILABLE = False
SONIDO_DISPONIBLE = None
PIL_AVAILABLE = False
REPORTLAB_AVAILABLE = False
PANDAS_AVAILABLE = False

# Placeholders para módulos que se cargarán bajo demanda
psycopg2 = None
win32print = None
win32api = None
Image = None
ImageTk = None
pd = None
canvas = None
colors = None
Table = None
TableStyle = None
landscape = None
A4 = None

def cargar_psycopg2():
    """Carga psycopg2 solo cuando es necesario"""
    global psycopg2, POSTGRES_AVAILABLE
    if not POSTGRES_AVAILABLE:
        try:
            import psycopg2 as pg2
            psycopg2 = pg2
            POSTGRES_AVAILABLE = True
        except Exception:
            POSTGRES_AVAILABLE = False
    return POSTGRES_AVAILABLE

def cargar_win32():
    """Carga win32 solo cuando es necesario"""
    global win32print, win32api, WIN32_AVAILABLE
    if not WIN32_AVAILABLE:
        try:
            import win32print as wp
            import win32api as wa
            win32print = wp
            win32api = wa
            WIN32_AVAILABLE = True
        except Exception:
            WIN32_AVAILABLE = False
    return WIN32_AVAILABLE

def cargar_pil():
    """Carga PIL solo cuando es necesario"""
    global Image, ImageTk, PIL_AVAILABLE
    if not PIL_AVAILABLE:
        try:
            from PIL import Image as Img, ImageTk as ImgTk
            Image = Img
            ImageTk = ImgTk
            PIL_AVAILABLE = True
        except Exception:
            PIL_AVAILABLE = False
    return PIL_AVAILABLE

def cargar_reportlab():
    """Carga ReportLab solo cuando es necesario"""
    global canvas, colors, Table, TableStyle, landscape, A4, REPORTLAB_AVAILABLE
    if not REPORTLAB_AVAILABLE:
        try:
            from reportlab.pdfgen import canvas as cv
            from reportlab.lib import colors as cl
            from reportlab.platypus import Table as Tb, TableStyle as Ts
            from reportlab.lib.pagesizes import landscape as ls, A4 as A4_size
            canvas = cv
            colors = cl
            Table = Tb
            TableStyle = Ts
            landscape = ls
            A4 = A4_size
            REPORTLAB_AVAILABLE = True
        except Exception:
            REPORTLAB_AVAILABLE = False
    return REPORTLAB_AVAILABLE

def cargar_pandas():
    """Carga pandas solo cuando es necesario"""
    global pd, PANDAS_AVAILABLE
    if not PANDAS_AVAILABLE:
        try:
            import pandas as pandas_module
            pd = pandas_module
            PANDAS_AVAILABLE = True
        except Exception:
            PANDAS_AVAILABLE = False
    return PANDAS_AVAILABLE

def cargar_sonido():
    """Configura sistema de sonido solo cuando es necesario"""
    global SONIDO_DISPONIBLE
    if SONIDO_DISPONIBLE is None:
        try:
            import winsound
            SONIDO_DISPONIBLE = "winsound"
        except Exception:
            try:
                import pygame
                pygame.mixer.init()
                SONIDO_DISPONIBLE = "pygame"
            except Exception:
                try:
                    import playsound
                    SONIDO_DISPONIBLE = "playsound"
                except Exception:
                    SONIDO_DISPONIBLE = False
    return SONIDO_DISPONIBLE

APP_VERSION = "1.7.1"
URL_VERSION = "https://gestor-flks.onrender.com/Version.txt"
URL_EXE = "https://gestor-flks.onrender.com/Gestor.exe"

def verificar_dependencias_criticas():
    """Verifica que todas las dependencias críticas estén disponibles"""
    errores = []
    advertencias = []
    
    # Verificar PostgreSQL (crítico)
    if not cargar_psycopg2():
        errores.append("• Driver PostgreSQL (psycopg2) no disponible")
    
    # Verificar dependencias opcionales
    if not cargar_win32():
        advertencias.append("• Impresión directa deshabilitada")
    
    if not cargar_sonido():
        advertencias.append("• Notificaciones de sonido deshabilitadas")
    
    # Mostrar advertencias (no bloquean)
    if advertencias:
        print("⚠️ Funcionalidades limitadas:")
        for adv in advertencias:
            print(f"  {adv}")
    
    if errores:
        mensaje_error = "ERRORES CRÍTICOS DETECTADOS:\n\n" + "\n".join(errores) + "\n\n"
        mensaje_error += "La aplicación podría no funcionar correctamente.\n"
        mensaje_error += "¿Desea continuar de todos modos?"
        
        import tkinter as tk
        from tkinter import messagebox
        
        root = tk.Tk()
        root.withdraw()  # Ocultar ventana principal
        
        continuar = messagebox.askyesno("⚠️ Advertencia de Dependencias", mensaje_error)
        root.destroy()
        
        return continuar
    
    print("✅ Todas las dependencias críticas están disponibles")
    return True

# === FUNCIONES DE ALERTBOX CON ÍCONOS ===

def mostrar_error(titulo, mensaje, parent=None):
    """Muestra mensaje de error con ícono"""
    return messagebox.showerror(f"❌ {titulo}", mensaje, parent=parent)

def mostrar_exito(titulo, mensaje, parent=None):
    """Muestra mensaje de éxito con ícono"""
    return messagebox.showinfo(f"✅ {titulo}", mensaje, parent=parent)

def mostrar_advertencia(titulo, mensaje, parent=None):
    """Muestra mensaje de advertencia con ícono"""
    return messagebox.showwarning(f"⚠️ {titulo}", mensaje, parent=parent)

def mostrar_info(titulo, mensaje, parent=None):
    """Muestra mensaje de información con ícono"""
    return messagebox.showinfo(f"ℹ️ {titulo}", mensaje, parent=parent)

def mostrar_pregunta(titulo, mensaje, parent=None):
    """Muestra pregunta Si/No con ícono"""
    return messagebox.askyesno(f"❓ {titulo}", mensaje, parent=parent)

# Las importaciones ahora se manejan con carga diferida arriba
import subprocess
import ctypes
from typing import Any

# Control de logs para evitar ruido en consola en producción
DEBUG_LOGS = False

def debug_log(*args: Any, **kwargs: Any):
    if DEBUG_LOGS:
        try:
            print(*args, **kwargs)
        except Exception:
            pass

# Silenciar stdout por defecto para evitar mensajes en consola (mantiene stderr)
if not DEBUG_LOGS:
    try:
        # Redirigir stdout a null con codificación UTF-8 para evitar UnicodeEncodeError en Windows
        sys.stdout = open(os.devnull, 'w', encoding='utf-8', errors='ignore')
    except Exception:
        pass

def version_tuple(v):
    return tuple(int(x) for x in v.strip().split(".") if x.isdigit())

def is_newer(latest, current):
    return version_tuple(latest) > version_tuple(current)

def run_windows_updater(new_exe_path, current_exe_path):
    # Creamos un .bat temporal que espera a que termine el proceso actual,
    # reemplaza el exe y lanza la nueva versión.
    print(f"🔄 DEBUG UPDATE: Iniciando actualización")
    print(f"🔄 DEBUG UPDATE: Archivo nuevo: {new_exe_path}")
    print(f"🔄 DEBUG UPDATE: Archivo actual: {current_exe_path}")
    
    bat = tempfile.NamedTemporaryFile(delete=False, suffix=".bat", mode="w", encoding="utf-8")
    new_p = new_exe_path.replace("/", "\\")
    cur_p = current_exe_path.replace("/", "\\")
    exe_name = os.path.basename(cur_p)
    backup_p = cur_p.replace(".exe", "_backup.exe")
    
    print(f"🔄 DEBUG UPDATE: Script bat: {bat.name}")
    print(f"🔄 DEBUG UPDATE: Proceso a esperar: {exe_name}")
    
    bat_contents = f"""@echo off
echo [UPDATE] Iniciando actualizacion...
timeout /t 3 /nobreak > nul

echo [UPDATE] Esperando que termine el proceso {exe_name}...
:waitloop
tasklist /FI "IMAGENAME eq {exe_name}" 2>nul | find /I "{exe_name}" > nul
if %ERRORLEVEL%==0 (
  echo [UPDATE] Proceso aun activo, esperando...
  timeout /t 1 > nul
  goto waitloop
)

echo [UPDATE] Proceso terminado, procediendo con actualizacion...

REM Intentar crear respaldo (no crítico si falla)
if exist "{cur_p}" (
    echo [UPDATE] Intentando crear respaldo...
    copy /Y "{cur_p}" "{backup_p}" > nul 2>&1
    if %ERRORLEVEL% neq 0 (
        echo [UPDATE] Advertencia: No se pudo crear respaldo, continuando...
        REM No salir con error, continuar con la actualización
    ) else (
        echo [UPDATE] Respaldo creado exitosamente
    )
)

REM Reemplazar ejecutable
echo [UPDATE] Reemplazando ejecutable...
move /Y "{new_p}" "{cur_p}"
if %ERRORLEVEL% neq 0 (
    echo [UPDATE] Error al reemplazar archivo
    if exist "{backup_p}" (
        echo [UPDATE] Restaurando desde respaldo...
        move /Y "{backup_p}" "{cur_p}" > nul 2>&1
    )
    echo [UPDATE] Puede que necesite ejecutar como administrador
    echo [UPDATE] O verificar que el archivo no esté en uso
    pause
    goto cleanup
)

echo [UPDATE] Actualizacion exitosa!

REM Eliminar respaldo si existe
if exist "{backup_p}" (
    del /Q "{backup_p}" > nul 2>&1
)

REM Limpiar archivos temporales de actualización antiguos
echo [UPDATE] Limpiando archivos temporales...
for %%f in ("{os.path.dirname(cur_p)}\\Gestor_new_*.exe") do (
    if exist "%%f" (
        del /Q "%%f" > nul 2>&1
    )
)

echo [UPDATE] Iniciando nueva version...
start "" "{cur_p}"
goto cleanup

:cleanup
echo [UPDATE] Eliminando script temporal...
del "%~f0"
"""
    bat.write(bat_contents)
    bat.close()
    
    # Hacer el bat ejecutable y lanzarlo
    print(f"🔄 DEBUG UPDATE: Ejecutando script de actualización")
    subprocess.Popen(["cmd", "/c", bat.name], creationflags=subprocess.CREATE_NEW_CONSOLE)
    sys.exit(0)

def _is_frozen_exe():
    """Indica si la app corre como ejecutable empaquetado (PyInstaller)."""
    return getattr(sys, "frozen", False)

def _current_binary_path():
    """Devuelve la ruta del binario actual (exe si frozen, script si no)."""
    return sys.executable if _is_frozen_exe() else sys.argv[0]

def mostrar_ventana_actualizacion():
    """Muestra una ventana de progreso animada durante la actualización"""
    import threading
    
    # Crear ventana de progreso
    progress_window = tk.Tk()
    progress_window.title("PaintFlow - Actualización")
    progress_window.geometry("400x200")
    progress_window.resizable(False, False)
    progress_window.configure(bg="#f0f0f0")
    
    # Centrar ventana
    progress_window.update_idletasks()
    x = (progress_window.winfo_screenwidth() // 2) - (400 // 2)
    y = (progress_window.winfo_screenheight() // 2) - (200 // 2)
    progress_window.geometry(f"400x200+{x}+{y}")
    
    try:
        aplicar_icono_y_titulo(progress_window, "Actualización")
    except:
        pass
    
    # Contenedor principal
    main_frame = tk.Frame(progress_window, bg="#f0f0f0")
    main_frame.pack(fill="both", expand=True, padx=20, pady=20)
    
    # Título
    title_label = tk.Label(
        main_frame, 
        text="🔄 Actualizando PaintFlow", 
        font=("Segoe UI", 16, "bold"),
        bg="#f0f0f0",
        fg="#1976D2"
    )
    title_label.pack(pady=(0, 10))
    
    # Mensaje de estado
    status_label = tk.Label(
        main_frame,
        text="Preparando descarga...",
        font=("Segoe UI", 10),
        bg="#f0f0f0",
        fg="#333333"
    )
    status_label.pack(pady=(0, 15))
    
    # Barra de progreso
    try:
        from tkinter import ttk
        progress_bar = ttk.Progressbar(
            main_frame,
            mode='indeterminate',
            length=300
        )
        progress_bar.pack(pady=(0, 15))
        progress_bar.start(10)  # Animación cada 10ms
    except:
        # Fallback si ttk no está disponible
        progress_label = tk.Label(
            main_frame,
            text="⏳ Descargando...",
            font=("Segoe UI", 12),
            bg="#f0f0f0",
            fg="#1976D2"
        )
        progress_label.pack(pady=(0, 15))
    
    # Texto informativo
    info_label = tk.Label(
        main_frame,
        text="Por favor espere mientras se descarga\nla nueva versión del sistema.",
        font=("Segoe UI", 9),
        bg="#f0f0f0",
        fg="#666666",
        justify="center"
    )
    info_label.pack()
    
    # Función para actualizar el estado
    def actualizar_estado(mensaje):
        status_label.config(text=mensaje)
        progress_window.update()
    
    # Función para cerrar ventana
    def cerrar_ventana():
        try:
            progress_window.destroy()
        except:
            pass
    
    # Almacenar referencias para uso externo
    progress_window.actualizar_estado = actualizar_estado
    progress_window.cerrar_ventana = cerrar_ventana
    
    return progress_window

def check_update():
    """Verifica versiones y actualiza sólo si corre como EXE en Windows.

    - Evita bloquear el arranque con timeouts. Se recomienda llamarla en hilo.
    - Omite el proceso si no es un ejecutable (útil durante desarrollo).
    """
    try:
        # Verificar si ya se intentó actualizar hoy para evitar bucles
        import tempfile
        import os
        today = datetime.now().strftime("%Y-%m-%d")
        update_flag_file = os.path.join(tempfile.gettempdir(), f"gestor_update_check_{today}.flag")
        
        if os.path.exists(update_flag_file):
            print(f"🔄 DEBUG UPDATE: Ya se verificó actualización hoy ({today}), omitiendo...")
            return
        
        is_frozen = _is_frozen_exe()

        headers = {
            "User-Agent": f"Gestor-Lista-Espera/{APP_VERSION}",
            "Cache-Control": "no-cache",
            "Pragma": "no-cache",
        }

        # Importación perezosa para no penalizar el arranque
        import requests  # type: ignore

        r = requests.get(URL_VERSION, timeout=5, headers=headers)
        r.raise_for_status()
        latest = r.text.strip()
        if not latest:
            print(f"🔄 DEBUG UPDATE: No se pudo obtener versión del servidor")
            return

        print(f"🔄 DEBUG UPDATE: Versión actual: {APP_VERSION}")
        print(f"🔄 DEBUG UPDATE: Versión en servidor: {latest}")
        print(f"🔄 DEBUG UPDATE: ¿Es más nueva? {is_newer(latest, APP_VERSION)}")

        # Crear archivo flag para indicar que ya verificamos hoy
        try:
            with open(update_flag_file, 'w') as f:
                f.write(f"Checked on {datetime.now().isoformat()}\nCurrent: {APP_VERSION}\nLatest: {latest}")
        except Exception:
            pass

        if is_newer(latest, APP_VERSION):
            # En desarrollo (no frozen), mostrar notificación elegante en lugar de actualizar
            if os.name == "nt" and not is_frozen:
                try:
                    # Crear ventana de notificación
                    notif_window = tk.Tk()
                    notif_window.title("PaintFlow - Nueva versión disponible")
                    notif_window.geometry("450x250")
                    notif_window.resizable(False, False)
                    notif_window.configure(bg="#f8f9fa")
                    
                    # Centrar ventana
                    notif_window.update_idletasks()
                    x = (notif_window.winfo_screenwidth() // 2) - (450 // 2)
                    y = (notif_window.winfo_screenheight() // 2) - (250 // 2)
                    notif_window.geometry(f"450x250+{x}+{y}")
                    
                    try:
                        aplicar_icono_y_titulo(notif_window, "Nueva versión")
                    except:
                        pass
                    
                    main_frame = tk.Frame(notif_window, bg="#f8f9fa")
                    main_frame.pack(fill="both", expand=True, padx=30, pady=20)
                    
                    # Icono y título
                    title_frame = tk.Frame(main_frame, bg="#f8f9fa")
                    title_frame.pack(fill="x", pady=(0, 15))
                    
                    tk.Label(
                        title_frame,
                        text="🚀 Nueva versión disponible",
                        font=("Segoe UI", 14, "bold"),
                        bg="#f8f9fa",
                        fg="#28a745"
                    ).pack()
                    
                    # Información de versiones
                    info_frame = tk.Frame(main_frame, bg="#f8f9fa")
                    info_frame.pack(fill="x", pady=(0, 20))
                    
                    tk.Label(
                        info_frame,
                        text=f"Nueva versión: {latest}",
                        font=("Segoe UI", 11, "bold"),
                        bg="#f8f9fa",
                        fg="#333333"
                    ).pack(anchor="w")
                    
                    tk.Label(
                        info_frame,
                        text=f"Versión actual: {APP_VERSION}",
                        font=("Segoe UI", 10),
                        bg="#f8f9fa",
                        fg="#666666"
                    ).pack(anchor="w", pady=(5, 0))
                    
                    tk.Label(
                        info_frame,
                        text="Ejecuta el archivo EXE para actualizar automáticamente\no descarga la última versión del servidor.",
                        font=("Segoe UI", 9),
                        bg="#f8f9fa",
                        fg="#666666",
                        justify="left"
                    ).pack(anchor="w", pady=(10, 0))
                    
                    # Botón cerrar
                    btn_frame = tk.Frame(main_frame, bg="#f8f9fa")
                    btn_frame.pack(fill="x")
                    
                    tk.Button(
                        btn_frame,
                        text="Entendido",
                        font=("Segoe UI", 10),
                        bg="#007bff",
                        fg="white",
                        relief="flat",
                        padx=20,
                        pady=8,
                        command=notif_window.destroy
                    ).pack(side="right")
                    
                    notif_window.mainloop()
                except Exception:
                    # Fallback al MessageBox original
                    msg = (
                        f"Hay una nueva versión disponible: {latest}\n\n"
                        f"Versión actual: {APP_VERSION}\n\n"
                        f"Ejecuta el EXE para actualizar automáticamente o descarga la última versión."
                    )
                    ctypes.windll.user32.MessageBoxW(0, msg, "Actualización disponible", 0x40)
                return

            # Flujo normal de actualización cuando es EXE empaquetado (Windows)
            if os.name == "nt" and is_frozen:
                # Mostrar ventana de progreso
                try:
                    progress_window = mostrar_ventana_actualizacion()
                    progress_window.update()
                except Exception as e:
                    print(f"⚠️ DEBUG UPDATE: No se pudo crear ventana de progreso: {e}")
                    progress_window = None
                
                base_path = os.path.dirname(_current_binary_path()) or os.getcwd()
                
                # Limpiar archivos de actualización antiguos antes de continuar
                print(f"🔄 DEBUG UPDATE: Limpiando archivos antiguos...")
                try:
                    for file in os.listdir(base_path):
                        if file.startswith("Gestor_new_") and file.endswith(".exe"):
                            old_file = os.path.join(base_path, file)
                            try:
                                os.remove(old_file)
                                print(f"🔄 DEBUG UPDATE: Eliminado archivo antiguo: {file}")
                            except Exception as e:
                                print(f"⚠️ DEBUG UPDATE: No se pudo eliminar {file}: {e}")
                except Exception as e:
                    print(f"⚠️ DEBUG UPDATE: Error limpiando archivos antiguos: {e}")
                
                # Usar nombre simple sin timestamp para evitar acumulación
                new_exe = os.path.join(base_path, "Gestor_new.exe")
                
                # Si ya existe un archivo de actualización, verificar si es válido
                if os.path.exists(new_exe):
                    try:
                        size = os.path.getsize(new_exe)
                        if size > 100_000:  # Si parece válido
                            print(f"🔄 DEBUG UPDATE: Archivo de actualización ya existe ({size} bytes)")
                            if progress_window:
                                progress_window.actualizar_estado("Usando actualización existente...")
                                time.sleep(1)
                                progress_window.cerrar_ventana()
                            print(f"🔄 DEBUG UPDATE: Iniciando proceso de reemplazo...")
                            run_windows_updater(new_exe, _current_binary_path())
                            return
                        else:
                            # Archivo corrupto, eliminarlo
                            os.remove(new_exe)
                            print(f"🔄 DEBUG UPDATE: Archivo existente corrupto eliminado")
                    except Exception as e:
                        print(f"⚠️ DEBUG UPDATE: Error verificando archivo existente: {e}")
                
                print(f"🔄 DEBUG UPDATE: Descargando a: {new_exe}")
                print(f"🔄 DEBUG UPDATE: Directorio base: {base_path}")

                if progress_window:
                    progress_window.actualizar_estado("Verificando permisos...")

                # Verificar permisos de escritura en el directorio
                try:
                    test_file = os.path.join(base_path, "test_write_permissions.tmp")
                    with open(test_file, "w") as f:
                        f.write("test")
                    os.remove(test_file)
                    print(f"🔄 DEBUG UPDATE: Permisos de escritura verificados")
                except Exception as e:
                    print(f"❌ DEBUG UPDATE: Sin permisos de escritura en {base_path}: {e}")
                    if progress_window:
                        progress_window.actualizar_estado("Error: Sin permisos de escritura")
                        time.sleep(2)
                        progress_window.cerrar_ventana()
                    return

                if progress_window:
                    progress_window.actualizar_estado("Descargando nueva versión...")

                # Descargar ejecutable
                try:
                    with requests.get(URL_EXE, stream=True, timeout=20, headers=headers) as resp:
                        resp.raise_for_status()
                        total_size = int(resp.headers.get('content-length', 0))
                        downloaded = 0
                        
                        with open(new_exe, "wb") as f:
                            for chunk in resp.iter_content(8192):
                                if chunk:
                                    f.write(chunk)
                                    downloaded += len(chunk)
                                    
                                    # Actualizar progreso si conocemos el tamaño total
                                    if progress_window and total_size > 0:
                                        percent = (downloaded / total_size) * 100
                                        progress_window.actualizar_estado(f"Descargando: {percent:.1f}%")
                                    elif progress_window:
                                        mb_downloaded = downloaded / (1024 * 1024)
                                        progress_window.actualizar_estado(f"Descargado: {mb_downloaded:.1f} MB")
                except Exception as e:
                    print(f"❌ DEBUG UPDATE: Error en descarga: {e}")
                    if progress_window:
                        progress_window.actualizar_estado("Error en la descarga")
                        time.sleep(2)
                        progress_window.cerrar_ventana()
                    return
                
                print(f"🔄 DEBUG UPDATE: Descarga completada")

                if progress_window:
                    progress_window.actualizar_estado("Verificando descarga...")

                # Verificación mínima de tamaño
                try:
                    size = os.path.getsize(new_exe)
                    print(f"🔄 DEBUG UPDATE: Tamaño del archivo descargado: {size} bytes")
                    if size < 100_000:
                        print(f"🔄 DEBUG UPDATE: Archivo muy pequeño, cancelando actualización")
                        if progress_window:
                            progress_window.actualizar_estado("Error: Archivo incompleto")
                            time.sleep(2)
                            progress_window.cerrar_ventana()
                        return
                except Exception as e:
                    print(f"🔄 DEBUG UPDATE: Error verificando tamaño: {e}")
                    if progress_window:
                        progress_window.actualizar_estado("Error verificando descarga")
                        time.sleep(2)
                        progress_window.cerrar_ventana()
                    return

                if progress_window:
                    progress_window.actualizar_estado("Preparando instalación...")
                    time.sleep(1)
                    progress_window.actualizar_estado("Cerrando aplicación actual...")
                    time.sleep(1)
                    progress_window.cerrar_ventana()

                print(f"🔄 DEBUG UPDATE: Iniciando proceso de reemplazo...")
                run_windows_updater(new_exe, _current_binary_path())
                return

            # Otros SO: permanecer en silencio
            return
        else:
            # No hay actualizaciones disponibles
            print(f"✅ DEBUG UPDATE: No hay actualizaciones disponibles (actual: {APP_VERSION}, servidor: {latest})")
    except Exception as e:
        # No bloquear inicio por fallas de actualización
        print(f"❌ DEBUG UPDATE: Error en verificación de actualización: {e}")
        return

if __name__ == "__main__":
    # Lanzar verificación en segundo plano si no se pasa --no-update
    if "--no-update" not in sys.argv:
        try:
            threading.Thread(target=check_update, daemon=True).start()
        except Exception:
            pass



# Los módulos Win32 ahora se manejan con carga diferida arriba

def obtener_sucursal_usuario(usuario_id):
    """Detecta la sucursal basándose en el ID del usuario"""
    if not usuario_id:
        return 'principal'
    
    usuario_lower = usuario_id.lower()
    
    # Mapeo completo de usuarios a sucursales
    if 'alameda' in usuario_lower:
        return 'alameda'
    elif 'churchill' in usuario_lower:
        return 'churchill'
    elif 'bavaro' in usuario_lower:
        return 'bavaro'
    elif 'bellavista' in usuario_lower:
        return 'bellavista'
    elif 'tiradentes' in usuario_lower:
        return 'tiradentes'
    elif 'la_vega' in usuario_lower or 'vega' in usuario_lower:
        return 'la_vega'
    elif 'luperon' in usuario_lower:
        return 'luperon'
    elif 'puertoplata' in usuario_lower:
        return 'puertoplata'
    elif 'puntacana' in usuario_lower:
        return 'puntacana'
    elif 'romana' in usuario_lower:
        return 'romana'
    elif 'santiago' in usuario_lower:
        return 'santiago1'
    elif 'sanisidro' in usuario_lower:
        return 'sanisidro'
    elif 'villamella' in usuario_lower:
        return 'villamella'
    elif 'terrenas' in usuario_lower:
        return 'terrenas'
    elif 'arroyohondo' in usuario_lower:
        return 'arroyohondo'
    elif 'bani' in usuario_lower:
        return 'bani'
    elif 'rafaelvidal' in usuario_lower:
        return 'rafaelvidal'
    elif 'sanfrancisco' in usuario_lower:
        return 'sanfrancisco'
    elif 'sanmartin' in usuario_lower:
        return 'sanmartin'
    elif 'zonaoriental' in usuario_lower:
        return 'zonaoriental'
    else:
        return 'principal'

def obtener_tabla_sucursal(sucursal):
    """Retorna el nombre de la tabla de pedidos para la sucursal"""
    return f'pedidos_pendientes_{sucursal}'

def obtener_sufijo_presentacion(presentacion):
    """Devuelve el sufijo correspondiente a la presentación seleccionada"""
    if not presentacion:
        return ""
    
    presentacion_lower = presentacion.lower()
    
    # Mapeo de presentaciones a sufijos (formato LabelsApp CORRECTO)
    sufijos = {
        
        "cuarto": "QT",
        "medio galón": "1/2", 
        "galón": "1",
        "cubeta": "5",
        "1/8": "1/8"
    }
    
    return sufijos.get(presentacion_lower, "")

def deducir_presentacion_desde_cantidad(cantidad):
    """Deduce la presentación más probable basándose en la cantidad"""
    if not cantidad or cantidad <= 0:
        return None
    
    # Lógica de deducción basada en patrones comunes
    if cantidad == 1:
        return "Galón"  # Lo más común para cantidad 1
    elif cantidad == 0.5:
        return "Medio Galón"  # Para cantidad 0.5
    elif cantidad in [2, 3, 4]:
        return "Cuarto"  # Cantidades pequeñas suelen ser cuartos
    elif cantidad >= 5:
        return "Cubeta" if cantidad == 5 else "Galón"
    else:
        return "Cuarto"  # Default para cantidades pequeñas

# El sistema de sonido ahora se maneja con carga diferida arriba

# Configuración de la base de datos
DB_CONFIG = {
    'host': 'dpg-d1b18u8dl3ps73e68v1g-a.oregon-postgres.render.com',
    'port': 5432,
    'database': 'labels_app_db',
    'user': 'admin',
    'password': 'KCFjzM4KYzSQx63ArufESIXq03EFXHz3',
    'sslmode': 'require'
}

def obtener_ruta_absoluta_gestor(rel_path):
    """Obtiene la ruta correcta de un archivo tanto para scripts como para ejecutables"""
    try:
        # Si es un ejecutable (PyInstaller) con recursos embebidos
        if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
            # PyInstaller crea una carpeta temporal con los recursos
            base_path = sys._MEIPASS
            ruta_recurso = os.path.join(base_path, rel_path)
            if os.path.exists(ruta_recurso):
                return ruta_recurso
        
        # Si es un ejecutable sin _MEIPASS, buscar en directorio del ejecutable
        if getattr(sys, 'frozen', False):
            base_path = os.path.dirname(sys.executable)
            ruta_recurso = os.path.join(base_path, rel_path)
            if os.path.exists(ruta_recurso):
                return ruta_recurso
        
        # Si es un script de Python normal
        base_path = os.path.dirname(os.path.abspath(__file__))
        ruta_recurso = os.path.join(base_path, rel_path)
        if os.path.exists(ruta_recurso):
            return ruta_recurso
        
        # Buscar en directorio de usuario como respaldo
        user_path = os.path.expanduser("~/.etiquetas_app")
        os.makedirs(user_path, exist_ok=True)
        return os.path.join(user_path, rel_path)
        
    except Exception as e:
        # En caso de error, usar directorio de usuario
        user_path = os.path.expanduser("~/.etiquetas_app")
        os.makedirs(user_path, exist_ok=True)
        return os.path.join(user_path, rel_path)

# Variable global para control de timers
timer_id_gestor = None

# Utilidad: aplicar icono y título unificado "PaintFlow — ..." a cualquier ventana
def aplicar_icono_y_titulo(ventana, titulo_suffix=None):
    try:
        icono_path = obtener_ruta_absoluta_gestor("icono.ico")
        if os.path.exists(icono_path):
            try:
                ventana.iconbitmap(icono_path)
            except Exception:
                pass
    except Exception:
        pass

    base = "PaintFlow"
    try:
        if titulo_suffix:
            ventana.title(f"{base} — {titulo_suffix}")
        else:
            ventana.title(base)
    except Exception:
        pass

# Configuración para impresión de etiquetas
IMPRESORA_CONF_PATH = os.path.join(os.getcwd(), 'config_impresora.txt')

def limpiar_mensaje_despues_gestor(widget, tiempo, mensaje=""):
    """Función centralizada para limpiar mensajes después de un tiempo"""
    global timer_id_gestor
    try:
        # Cancelar timer anterior si existe
        if timer_id_gestor:
            widget.after_cancel(timer_id_gestor)
            timer_id_gestor = None
        
        # Programar nueva limpieza
        timer_id_gestor = widget.after(tiempo, lambda: widget.configure(text=mensaje))
    except Exception as e:
        # Si hay error, intentar limpiar directamente
        try:
            widget.configure(text=mensaje)
        except:
            pass

# === SISTEMA DE NOTIFICACIONES SONORAS ===
class NotificacionesSonoras:
    """Maneja las notificaciones sonoras para nuevos pedidos"""
    
    def __init__(self):
        self.sonido_habilitado = True
        self.ultimo_sonido = 0  # Control de spam de sonidos
        self.crear_sonidos_predeterminados()
    
    def crear_sonidos_predeterminados(self):
        """Crea sonidos de campanas con diferentes intensidades según prioridad"""
        self.sonidos = {
            # Campanas para diferentes prioridades
            'prioridad_alta': {
                'archivo': 'campana_alta.wav',
                'secuencia': [
                    (1200, 600)  # 1 campana alta e intensa
                ],
                'descripcion': 'Campana ALTA - Toque intenso'
            },
            'prioridad_media': {
                'archivo': 'campana_media.wav',
                'secuencia': [
                    (900, 500)  # 1 campana moderada
                ],
                'descripcion': 'Campana MEDIA - Toque moderado'
            },
            'prioridad_baja': {
                'archivo': 'campana_baja.wav',
                'secuencia': [
                    (650, 400)  # 1 campana suave
                ],
                'descripcion': 'Campana BAJA - Toque suave'
            },
            # Sonidos adicionales del sistema
            'nuevo_pedido': {
                'archivo': 'notification.wav',
                'secuencia': [(800, 200)],
                'descripcion': 'Notificación general'
            },
            'pedido_completado': {
                'archivo': 'complete.wav',
                'secuencia': [(600, 150), (800, 150), (1000, 200)],
                'descripcion': 'Pedido completado exitosamente'
            },
            'pedido_vencido': {
                'archivo': 'urgent.wav',
                'secuencia': [
                    (1500, 100), (1200, 100), (1500, 100), (1200, 100),
                    (1500, 100), (1200, 100), (1500, 300)  # Urgente y repetitivo
                ],
                'descripcion': 'Alerta de pedido vencido'
            }
        }
    
    def reproducir_sonido(self, tipo_sonido='nuevo_pedido'):
        """Reproduce un sonido de notificación"""
        if not self.sonido_habilitado:
            return
        
        # Control anti-spam (máximo un sonido cada 2 segundos)
        tiempo_actual = time.time()
        if tiempo_actual - self.ultimo_sonido < 2:
            return
        
        self.ultimo_sonido = tiempo_actual
        
        # Ejecutar sonido en hilo separado para no bloquear UI
        threading.Thread(target=self._reproducir_sonido_async, 
                        args=(tipo_sonido,), daemon=True).start()
    
    def _reproducir_sonido_async(self, tipo_sonido):
        """Reproduce sonido de campanas de forma asíncrona"""
        try:
            sonido = self.sonidos.get(tipo_sonido, self.sonidos['nuevo_pedido'])
            
            # Cargar y usar sistema de sonido disponible
            sistema_sonido = cargar_sonido()
            
            if sistema_sonido == "winsound":
                # Cargar winsound dinámicamente
                try:
                    import winsound
                    archivo_sonido = obtener_ruta_absoluta_gestor(sonido['archivo'])
                    if os.path.exists(archivo_sonido):
                        winsound.PlaySound(archivo_sonido, winsound.SND_FILENAME | winsound.SND_ASYNC)
                    else:
                        # Generar secuencia de campanas con winsound
                        secuencia = sonido.get('secuencia', [(800, 200)])
                        
                        print(f"🔔 Reproduciendo: {sonido.get('descripcion', tipo_sonido)}")
                        
                        for frecuencia, duracion in secuencia:
                            try:
                                winsound.Beep(frecuencia, duracion)
                            except Exception as e:
                                print(f"⚠️ Error reproduciendo beep {frecuencia}Hz: {e}")
                                # Fallback con frecuencia estándar
                                try:
                                    winsound.Beep(800, 300)
                                except:
                                    print(f"🔔 Sonido fallback también falló - {sonido.get('descripcion', tipo_sonido)}")
                except ImportError:
                    print("⚠️ winsound no disponible")
            
            elif sistema_sonido == "pygame":
                # Cargar pygame dinámicamente
                try:
                    import pygame
                    secuencia = sonido.get('secuencia', [(800, 200)])
                    print(f"🔔 Reproduciendo: {sonido.get('descripcion', tipo_sonido)}")
                    
                    for frecuencia, duracion in secuencia:
                        # Usar pygame mixer para generar tonos
                        try:
                            # Crear un beep simple
                            pygame.mixer.Sound.play(pygame.mixer.Sound(buffer=b'\x00\x7f' * int(44100 * duracion / 1000)))
                            pygame.time.wait(duracion + 50)
                        except:
                            print(f"Campana: {frecuencia}Hz por {duracion}ms")
                except ImportError:
                    print("⚠️ pygame no disponible")
            
            elif sistema_sonido == "playsound":
                # Cargar playsound dinámicamente
                try:
                    import playsound
                    archivo_sonido = obtener_ruta_absoluta_gestor(sonido['archivo'])
                    if os.path.exists(archivo_sonido):
                        playsound.playsound(archivo_sonido, False)
                    else:
                        print(f"🔔 {sonido.get('descripcion', tipo_sonido)} (sin archivo de audio)")
                except ImportError:
                    print("⚠️ playsound no disponible")
            
            else:
                # Sin sonido disponible, solo mostrar en consola
                print(f"🔔 {sonido.get('descripcion', tipo_sonido)} - Prioridad: {tipo_sonido}")
        
        except Exception as e:
            print(f"❌ Error reproduciendo sonido {tipo_sonido}: {e}")
    
    def reproducir_sonido_por_prioridad(self, prioridad):
        """Reproduce el sonido de campana apropiado según la prioridad"""
        if not self.sonido_habilitado:
            return
        
        # Mapear prioridad a tipo de sonido
        mapeo_prioridades = {
            'Alta': 'prioridad_alta',
            'Media': 'prioridad_media', 
            'Baja': 'prioridad_baja'
        }
        
        tipo_sonido = mapeo_prioridades.get(prioridad, 'prioridad_media')
        
        debug_log(f"🔔 Reproduciendo campanas para prioridad: {prioridad}")
        self.reproducir_sonido(tipo_sonido)
    
    def alternar_sonido(self):
        """Alterna el estado del sonido (habilitar/deshabilitar)"""
        self.sonido_habilitado = not self.sonido_habilitado
        estado = "habilitado" if self.sonido_habilitado else "deshabilitado"
        debug_log(f"🔊 Sonido {estado}")
        return self.sonido_habilitado
    
    def test_sonido(self):
        """Prueba el sistema de sonido"""
        debug_log("🔧 Probando sistema de sonido...")
        self.reproducir_sonido('nuevo_pedido')

# === FUNCIONES DE IMPRESIÓN DE ETIQUETAS ===

def guardar_impresora(nombre_impresora):
    """Guarda la impresora seleccionada en el archivo de configuración con múltiples ubicaciones de respaldo"""
    global IMPRESORA_CONF_PATH
    
    ubicaciones_posibles = [
        os.path.join(os.getcwd(), 'config_impresora.txt'),  # Ubicación actual
        os.path.join(os.path.expanduser('~'), 'config_impresora.txt'),  # Carpeta usuario
        os.path.join(os.environ.get('TEMP', os.getcwd()), 'config_impresora.txt'),  # Carpeta temporal
        os.path.join(os.path.dirname(__file__), 'config_impresora.txt') if '__file__' in globals() else None  # Junto al script
    ]
    
    # Filtrar ubicaciones válidas
    ubicaciones_posibles = [loc for loc in ubicaciones_posibles if loc is not None]
    
    for ubicacion in ubicaciones_posibles:
        try:
            # Crear directorio si no existe
            directorio = os.path.dirname(ubicacion)
            if not os.path.exists(directorio):
                os.makedirs(directorio, exist_ok=True)
            
            # Intentar escribir
            with open(ubicacion, 'w', encoding='utf-8') as f:
                f.write(nombre_impresora)
            
            # Verificar que se escribió correctamente
            with open(ubicacion, 'r', encoding='utf-8') as f:
                contenido = f.read().strip()
                if contenido == nombre_impresora:
                    debug_log(f"💾 Impresora guardada: {nombre_impresora} en {ubicacion}")
                    # Actualizar la ruta global para futuras operaciones
                    IMPRESORA_CONF_PATH = ubicacion
                    return True
        except PermissionError:
            debug_log(f"⚠️ Sin permisos para escribir en: {ubicacion}")
            continue
        except Exception as e:
            debug_log(f"⚠️ Error en ubicación {ubicacion}: {e}")
            continue
    
    # Si llegamos aquí, no se pudo guardar en ninguna ubicación
    debug_log(f"⚠️ No se pudo guardar la configuración de impresora en ninguna ubicación")
    debug_log(f"📋 Ubicaciones intentadas: {', '.join(ubicaciones_posibles)}")
    debug_log(f"💡 Sugerencia: Ejecutar como administrador o mover programa a carpeta con permisos")
    return False

def obtener_impresoras_disponibles():
    """Obtiene la lista de impresoras disponibles en el sistema"""
    if not WIN32_AVAILABLE:
        return []
    
    try:
        impresoras = [p[2] for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
        return impresoras
    except Exception as e:
        debug_log(f"❌ Error obteniendo impresoras: {e}")
        return []

def mostrar_seleccionador_impresora():
    """Muestra una ventana para seleccionar la impresora"""
    if not WIN32_AVAILABLE:
        mostrar_error("Error", "Módulos de impresión no disponibles")
        return None
    
    impresoras = obtener_impresoras_disponibles()
    
    if not impresoras:
        mostrar_advertencia("Sin Impresoras", "No se detectaron impresoras en el sistema")
        return None
    
    # Crear ventana de selección
    ventana = tk.Toplevel()
    aplicar_icono_y_titulo(ventana, "Seleccionar Impresora")
    ventana.geometry("400x300")
    ventana.resizable(False, False)
    ventana.grab_set()  # Modal
    
    # Centrar ventana
    ventana.update_idletasks()
    x = (ventana.winfo_screenwidth() // 2) - (400 // 2)
    y = (ventana.winfo_screenheight() // 2) - (300 // 2)
    ventana.geometry(f"400x300+{x}+{y}")
    
    resultado = {'impresora': None}
    
    # Header
    header_frame = ttk.Frame(ventana)
    header_frame.pack(fill="x", padx=20, pady=10)
    
    ttk.Label(header_frame, text="🖨️ Seleccionar Impresora", 
             font=("Segoe UI", 14, "bold")).pack()
    
    ttk.Label(header_frame, text="Seleccione la impresora para las etiquetas:", 
             font=("Segoe UI", 10)).pack(pady=(5, 0))
    
    # Lista de impresoras
    lista_frame = ttk.Frame(ventana)
    lista_frame.pack(fill="both", expand=True, padx=20, pady=10)
    
    # Variable para la selección
    impresora_var = tk.StringVar()
    impresora_actual = cargar_impresora_guardada()
    if impresora_actual in impresoras:
        impresora_var.set(impresora_actual)
    
    # Crear radiobuttons para cada impresora
    for impresora in impresoras:
        rb = ttk.Radiobutton(lista_frame, text=impresora, 
                            variable=impresora_var, value=impresora,
                            style="TRadiobutton")
        rb.pack(anchor="w", pady=2)
    
    # Botones
    botones_frame = ttk.Frame(ventana)
    botones_frame.pack(fill="x", padx=20, pady=10)
    
    def confirmar():
        seleccion = impresora_var.get()
        if seleccion:
            # SIEMPRE configurar la impresora para esta sesión
            resultado['impresora'] = seleccion
            
            # Intentar guardar, pero no bloquear si falla
            try:
                guardar_impresora(seleccion)
            except Exception:
                pass
        else:
            mostrar_advertencia("Selección requerida", "Por favor seleccione una impresora")
        ventana.destroy()
    
    def cancelar():
        ventana.destroy()
    
    ttk.Button(botones_frame, text="✅ Confirmar", command=confirmar,
              style="success.TButton").pack(side="right", padx=(5, 0))
    
    ttk.Button(botones_frame, text="❌ Cancelar", command=cancelar,
              style="secondary.TButton").pack(side="right")
    
    # Información adicional
    info_frame = ttk.Frame(ventana)
    info_frame.pack(fill="x", padx=20, pady=(0, 10))
    
    if impresora_actual:
        ttk.Label(info_frame, text=f"Actual: {impresora_actual}", 
                 font=("Segoe UI", 9), foreground="gray").pack()
    
    ventana.wait_window()
    return resultado['impresora']

def cargar_impresora_guardada():
    """Carga la impresora guardada desde el archivo de configuración buscando en múltiples ubicaciones"""
    global IMPRESORA_CONF_PATH
    
    if not WIN32_AVAILABLE:
        return ''
    
    ubicaciones_posibles = [
        IMPRESORA_CONF_PATH,  # Ubicación actual configurada
        os.path.join(os.getcwd(), 'config_impresora.txt'),  # Ubicación actual
        os.path.join(os.path.expanduser('~'), 'config_impresora.txt'),  # Carpeta usuario
        os.path.join(os.environ.get('TEMP', os.getcwd()), 'config_impresora.txt'),  # Carpeta temporal
        os.path.join(os.path.dirname(__file__), 'config_impresora.txt') if '__file__' in globals() else None  # Junto al script
    ]
    
    # Filtrar ubicaciones válidas y eliminar duplicados
    ubicaciones_posibles = list(set([loc for loc in ubicaciones_posibles if loc is not None]))
    
    for ubicacion in ubicaciones_posibles:
        try:
            if os.path.exists(ubicacion):
                with open(ubicacion, 'r', encoding='utf-8') as f:
                    impresora = f.read().strip()
                    if impresora:  # Solo retornar si hay contenido
                        debug_log(f"📄 Configuración de impresora cargada desde: {ubicacion}")
                        # Actualizar la ruta global para futuras operaciones
                        IMPRESORA_CONF_PATH = ubicacion
                        return impresora
        except Exception as e:
            print(f"⚠️ Error leyendo {ubicacion}: {e}")
            continue
    
    print("📄 No se encontró configuración de impresora guardada")
    return ''

def diagnosticar_permisos_escritura():
    """Diagnostica problemas de permisos de escritura para configuración"""
    ubicaciones_prueba = [
        os.path.join(os.getcwd(), 'test_permisos.txt'),
        os.path.join(os.path.expanduser('~'), 'test_permisos.txt'),
        os.path.join(os.environ.get('TEMP', os.getcwd()), 'test_permisos.txt')
    ]
    
    resultado = []
    for ubicacion in ubicaciones_prueba:
        try:
            # Intentar crear directorio si no existe
            directorio = os.path.dirname(ubicacion)
            if not os.path.exists(directorio):
                os.makedirs(directorio, exist_ok=True)
            
            # Intentar escribir un archivo de prueba
            with open(ubicacion, 'w') as f:
                f.write("test")
            
            # Intentar leerlo
            with open(ubicacion, 'r') as f:
                contenido = f.read()
            
            # Eliminar archivo de prueba
            os.remove(ubicacion)
            
            resultado.append(f"✅ {ubicacion} - Escritura OK")
        except PermissionError:
            resultado.append(f"❌ {ubicacion} - Sin permisos")
        except Exception as e:
            resultado.append(f"⚠️ {ubicacion} - Error: {e}")
    
    return resultado

def generar_zpl_gestor(codigo, descripcion, producto, terminacion, presentacion, cantidad=1, base="", ubicacion="", sucursal="", id_profesional="", operador="", codigo_base=None, nombre_cliente: str = ""):
    """Genera código ZPL para impresión de etiquetas desde el gestor usando el mismo diseño de LabelsApp.

    Si el producto es personalizado (según CSV o base=='custom'), agrega una banda de texto 'PERSONALIZADO'.
    """
    w, h = 406, 203  # 2x1 pulgadas a 203 dpi
    
    # Detectar si el código sigue el patrón SW 0000 (SW + espacio + 4 números exactos)
    import re
    patron_sw = re.compile(r'^SW \d{4}$')
    # Solo considerar el código sin espacios adicionales para mayor precisión
    codigo_limpio = codigo.strip()
    es_codigo_sw = bool(patron_sw.match(codigo_limpio)) and len(codigo_limpio) == 7
    
    # Ajuste dinámico de fuentes según longitud del contenido y patrón
    if es_codigo_sw:
        # Para códigos SW 0000: usar tamaños normales
        font_codigo = 70 if len(codigo) <= 6 else 70 if len(codigo) <= 8 else 30
        font_desc = 20 if len(descripcion) > 25 else 24  # Descripción más pequeña
    else:
        # Para productos que NO siguen el patrón SW 0000: código y descripción más pequeños
        font_codigo = 50 if len(codigo) <= 6 else 45 if len(codigo) <= 8 else 25
        font_desc = 18 if len(descripcion) > 25 else 20  # Descripción aún más pequeña
    
    font_producto = 22 if len(f"{producto}/{terminacion}") > 20 else 26
    
    # Posiciones optimizadas - igual que LabelsApp
    margin = 65  # Incrementado de 55 a 65
    y_cod_base = 32  # Bajado un poco más para separar mejor el código
    # y_desc se calculará más adelante tras ajustar y_cod si es personalizado
    
    # Preparar descripción a imprimir (ajustable para personalizados)
    desc_text = descripcion or ""
    
    # === Borde decorativo ===
    border_thickness = 2
    
    # === Sucursal lateral vertical optimizada ===
    sucursal_font_size = 16  # Reducido de 20 a 16
    x_sucursal = 9  # Último pequeño ajuste hacia la izquierda (de 10 a 9)
    y_sucursal_start = 30
    
    # === Base/Ubicación en la parte inferior ===
    # Productos que no deben mostrar la base
    productos_sin_base = ['laca', 'uretano', 'esmalte kem', 'esmalte multiuso', 'monocapa']
    mostrar_base = not any(prod.lower() in producto.lower() for prod in productos_sin_base)
    
    if mostrar_base:
        info_adicional = f"{base} | {ubicacion}" if base and ubicacion else base or ubicacion
    else:
        # Solo mostrar ubicación para productos sin base
        info_adicional = ubicacion if ubicacion else ""
    
    font_info = 16
    y_info = h - 25

    # --- Detección de producto personalizado (simplificada) ---
    # Regla solicitada:
    # - Es personalizado si: codigo_base == 'N/A' (o variantes) OR base == 'custom'.
    # - En cualquier otro caso: no es personalizado. Sin CSV ni otras heurísticas.
    base_str = (base or '').strip().lower()
    cb = (codigo_base or '').strip().lower() if codigo_base is not None else ''
    # Considerar 'n/a' con posibles sufijos de presentación (ej. 'n/a-qt')
    es_na_cb = cb.startswith('n/a') or cb in {'na', 'n.a'}
    is_personalizado = True if (es_na_cb or base_str == 'custom') else False

    # Para personalizados: ocultar la descripción (el usuario no quiere descripción en estos casos)
    if is_personalizado:
        desc_text = ""

    # Calcular posición de producto/terminación en función de la descripción final
    # (se calculará después de fijar y_cod e y_desc)

    zpl = "^XA\n"
    zpl += "^CI28\n"  # Codificación UTF-8
    zpl += f"^PW{w}\n^LL{h}\n^LH0,0\n"
    
    # === BORDE DECORATIVO ===
    zpl += f"^FO0,0^GB{w},{border_thickness},B^FS\n"  # Borde superior
    zpl += f"^FO0,{h-border_thickness}^GB{w},{border_thickness},B^FS\n"  # Borde inferior
    zpl += f"^FO{w-border_thickness-2},0^GB{border_thickness},{h},B^FS\n"  # Borde derecho movido ligeramente a la izquierda
    zpl += f"^FO6,0^GB{border_thickness},{h},B^FS\n"  # Borde izquierdo último pequeño ajuste a la izquierda
    
    # === LÍNEA DECORATIVA SUPERIOR ===
    zpl += f"^FO8,15^GB{w-16},1,B^FS\n"  # Línea superior movida aún más a la izquierda y ajustada en ancho

    # Calcular posiciones finales según si es personalizado (banner bajo la línea decorativa)
    font_banner = 16
    y_banner = 18  # Justo debajo de la línea en y=15
    y_cod = y_cod_base + ((font_banner + 4) if is_personalizado else 0)
    # Subir un poco más la descripción en etiquetas normales (no personalizadas)
    # Subimos la descripción unos píxeles para que quede más cerca del código
    # (antes: +5 si personalizado, +2 si normal). Ahora: +2 si personalizado, -1 si normal
    y_desc = y_cod + font_codigo + (2 if is_personalizado else -1)
    
    # Ahora que y_desc está definido, calcular líneas de descripción
    desc_lines = (1 if len(desc_text) <= 32 else 2) if desc_text else 0
    # Fijar PRODUCTO/TERMINACIÓN justo encima del OP (anclado al fondo)
    y_producto = y_info - (font_producto + 4)

    # === BANDA 'PERSONALIZADO' (solo si aplica) ===
    if is_personalizado:
        zpl += f"^CF0,{font_banner}\n^FO{margin},{y_banner}^FB{w-margin*2-5},1,0,C,0^FDPERSONALIZADO^FS\n"
    
    # === CÓDIGO PRINCIPAL (Destacado y centrado) ===
    zpl += f"^CF0,{font_codigo}\n"
    zpl += f"^FO{margin},{y_cod}^FB{w-margin*2-5},1,0,C,0^FD{codigo}^FS\n"
    
    # === DESCRIPCIÓN (Centrada, máximo 2 líneas) — solo si hay texto ===
    if desc_text:
        zpl += f"^CF0,{font_desc}\n"
        zpl += f"^FO{margin},{y_desc}^FB{w-margin*2-5},{desc_lines},0,C,0^FD{desc_text}^FS\n"
    
    # === PRODUCTO/TERMINACIÓN (sin presentación) ===
    producto_texto = f"{producto.upper()}/{terminacion.upper()}"
    zpl += f"^CF0,{font_producto}\n"
    zpl += f"^FO{margin-10},{y_producto}^FB{w-margin*2+15},1,0,C,0^FD{producto_texto}^FS\n"
    
    # === OPERADOR CON PRESENTACIÓN (si existe) — ahora más grande y más abajo (última línea) ===
    if operador:
        # Obtener sufijo de presentación
        sufijo_presentacion = obtener_sufijo_presentacion(presentacion) if presentacion else ""
        
        # Construir texto del operador con sufijo
        if sufijo_presentacion:
            op_text = f"OP: {operador} - {sufijo_presentacion}"
        else:
            op_text = f"OP: {operador}"
        
        # Limitar longitud
        op_text = op_text[:28]
        
        zpl += (
            f"^CF0,{max(18, font_info+2)}\n"
            f"^FO{margin},{y_info}^FB{w-margin*2-5},1,0,C,0^FD{op_text}^FS\n"
        )

    # === INFORMACIÓN ADICIONAL (Base/Ubicación) — por encima de PRODUCTO/TERMINACIÓN ===
    if info_adicional:
        y_info_adic = y_producto - (font_info + 6)
        zpl += (
            f"^CF0,{font_info}\n"
            f"^FO{margin},{y_info_adic}^FB{w-margin*2-5},1,0,C,0^FD{info_adicional}^FS\n"
        )
    
    # === LÍNEA SEPARADORA ENTRE BASE (si hay) Y PRODUCTO ===
    if info_adicional:
        y_linea_separadora = (y_producto - 4)
    else:
        # si no hay base/ubicación, colocar una línea sutil por encima del producto
        y_linea_separadora = max(y_desc + (font_desc * desc_lines) + 6, y_producto - 8)
    zpl += f"^FO{margin+20},{y_linea_separadora}^GB{w-margin*2-50},1,B^FS\n"  # Línea más pequeña
    
    # === SUCURSAL LATERAL (Rotada 90°) ===
    if sucursal:
        # Calcular el centro vertical real de la etiqueta
        centro_etiqueta = h // 2  # Centro absoluto de la etiqueta (203/2 = 101.5)
        
        # Calcular la longitud del texto para centrarlo perfectamente
        longitud_texto = len(sucursal) * (sucursal_font_size * 0.5)  # Ajustado de 0.6 a 0.5
        y_inicio_centrado = centro_etiqueta - (longitud_texto // 2)  # Cambiado + por -
        
        zpl += (
            f"^A0R,{sucursal_font_size},{sucursal_font_size}\n"
            f"^FO{x_sucursal},{y_inicio_centrado}^FD{sucursal.upper()}^FS\n"
        )
    
    # === CLIENTE/ID LATERAL DERECHO (Rotado 90°) ===
    texto_lateral_derecho = (nombre_cliente or "").strip() or (id_profesional or "").strip()
    if texto_lateral_derecho:
        # Configuración para el texto en el lado derecho
        id_font_size = 14  # Tamaño de fuente (un poco más pequeño que la sucursal)
        x_id = w - 25  # Posición en el lado derecho (cerca del borde derecho)

        # Limitar longitud visible para evitar solapamiento excesivo
        texto_lateral_derecho = texto_lateral_derecho[:22]
        
        # Calcular el centro vertical para el texto rotado
        centro_etiqueta_id = h // 2
        longitud_texto_id = len(texto_lateral_derecho) * (id_font_size * 0.5)
        y_inicio_centrado_id = centro_etiqueta_id - (longitud_texto_id // 2)
        
        zpl += (
            f"^A0R,{id_font_size},{id_font_size}\n"
            f"^FO{x_id},{y_inicio_centrado_id}^FD{texto_lateral_derecho.upper()}^FS\n"
        )
    
    zpl += "^XZ\n"
    
    return zpl * int(cantidad)

def imprimir_zebra_zpl_gestor(zpl_code):
    """Imprime código ZPL en impresora Zebra desde el gestor"""
    # Cargar módulos de impresión dinámicamente
    if not cargar_win32():
        mostrar_error("Error de Impresión", "Módulos de impresión no disponibles en el sistema")
        return False

    try:
        pr = cargar_impresora_guardada()
        
        # Verificar lista de impresoras disponibles (opcional)
        try:
            available = [p[2] for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
        except Exception as e:
            available = []
            debug_log(f"❌ Error obteniendo lista de impresoras: {e}")

        if not pr:
            nueva_impresora = mostrar_seleccionador_impresora()
            if nueva_impresora:
                pr = nueva_impresora
            else:
                return False

        if available and pr not in available:
            nueva_impresora = mostrar_seleccionador_impresora()
            if nueva_impresora:
                pr = nueva_impresora
            else:
                return False

        # Enviar ZPL en RAW
        h = win32print.OpenPrinter(pr)
        try:
            win32print.StartDocPrinter(h, 1, ("Etiqueta_Gestor", None, "RAW"))
            win32print.StartPagePrinter(h)
            win32print.WritePrinter(h, zpl_code.encode())
            win32print.EndPagePrinter(h)
            win32print.EndDocPrinter(h)
            return True
        finally:
            try:
                win32print.ClosePrinter(h)
            except Exception as e:
                debug_log(f"⚠️ Error cerrando impresora: {e}")
                
    except Exception as e:
        error_msg = f"Error de impresión: {str(e)}"
        # Guardar ZPL en archivo temporal como fallback y notificar
        try:
            tmp = tempfile.mktemp(suffix='.zpl')
            with open(tmp, 'w', encoding='utf-8') as f:
                f.write(zpl_code)
            messagebox.showerror(
                "Error de Impresión",
                f"{error_msg}\n\nEl código ZPL se guardó en:\n{tmp}\n\nPuede enviarlo manualmente a la impresora."
            )
        except Exception as e2:
            messagebox.showerror("Error Crítico", f"{error_msg}\n\nNo se pudo crear archivo de respaldo: {e2}")
        return False

# === SISTEMA DE LOGIN INTEGRADO ===
class SistemaLoginColorista:
    """Sistema de login integrado para Coloristas"""
    
    def __init__(self):
        self.db_config = DB_CONFIG
        self.usuario_actual = None
        self.sucursal = None
        
    def conectar_bd(self):
        """Conecta a la base de datos"""
        try:
            return psycopg2.connect(**self.db_config)
        except Exception as e:
            print(f"❌ Error conectando a BD: {e}")
            return None
    
    def hash_password(self, password):
        """Encripta la contraseña usando SHA-256"""
        return hashlib.sha256(password.encode()).hexdigest()
    
    def verificar_credenciales(self, username, password):
        """Verifica las credenciales del usuario"""
        conn = self.conectar_bd()
        if not conn:
            return {"error": "No se pudo conectar a la base de datos"}
        
        try:
            cur = conn.cursor()
            password_hash = self.hash_password(password)
            
            query = """
            SELECT u.id, u.username, u.password_hash, u.nombre_completo, u.rol, 
                   u.activo, u.sucursal_id, s.nombre as sucursal_nombre
            FROM usuarios u
            LEFT JOIN sucursales s ON u.sucursal_id = s.id
            WHERE u.username = %s AND u.activo = true
            """
            
            cur.execute(query, (username,))
            usuario = cur.fetchone()
            
            if not usuario:
                return {"error": "Usuario no encontrado o inactivo"}
            
            if usuario[2] != password_hash:
                return {"error": "Contraseña incorrecta"}
            
            # Verificar que el rol sea apropiado para el Gestor de Lista de Espera
            roles_permitidos = ['colorista', 'administrador']
            if usuario[4] not in roles_permitidos:
                return {"error": f"Rol '{usuario[4]}' no tiene acceso al Gestor de Lista de Espera. Se requiere rol de colorista o administrador."}
            
            return {
                "id": usuario[0],
                "username": usuario[1],
                "nombre_completo": usuario[3],
                "rol": usuario[4],
                "sucursal_id": usuario[6],
                "sucursal_nombre": usuario[7] or "SUCURSAL PRINCIPAL"
            }
            
        except Exception as e:
            return {"error": f"Error en verificación: {e}"}
        finally:
            cur.close()
            conn.close()
    
    def mostrar_login(self, master=None):
        """Muestra la ventana de login.

        Si se proporciona `master`, se crea como una Toplevel modal para usar un único root.
        """
        # Cargar PIL antes de usar imágenes
        if not cargar_pil():
            print("⚠️ DEBUG LOGIN: PIL no está disponible")
        else:
            print("✅ DEBUG LOGIN: PIL cargado correctamente")
            
        if master is not None:
            ventana_login = tk.Toplevel(master)
            try:
                ventana_login.transient(master)
                ventana_login.grab_set()
            except Exception:
                pass
        else:
            ventana_login = tk.Tk()

        aplicar_icono_y_titulo(ventana_login, "Login")
        ventana_login.geometry("600x330")
        ventana_login.resizable(False, False)
        ventana_login.configure(bg="#f5f5f5")

        # Aplicar estilo ttkbootstrap
        try:
            ttk.Style(theme="flatly")
            try:
                style = ttk.Style()
                style.configure('primary.TEntry', foreground="#1b1f23", insertcolor="#0D47A1")
                style.map('primary.TEntry', bordercolor=[('focus', '#1565C0'), ('!focus', '#1976D2')])
            except Exception:
                pass
        except Exception:
            try:
                ttk.Style()
            except Exception:
                pass

        # Cargar preferencias de login
        saved_username = ""
        saved_password = ""
        remember_access_saved = False
        try:
            cfg_path = obtener_ruta_absoluta_gestor("gestor_login.json")
            if os.path.exists(cfg_path):
                with open(cfg_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    saved_username = data.get('usuario', "") or ""
                    remember_access_saved = bool(data.get('recordar', False) or data.get('recordar_pass', False))
                    enc_pwd = data.get('password')
                    if remember_access_saved and enc_pwd:
                        try:
                            saved_password = base64.b64decode(enc_pwd.encode('utf-8')).decode('utf-8')
                        except Exception:
                            saved_password = ""
        except Exception:
            pass

        # Icono
        # icono y título ya asignados por aplicar_icono_y_titulo

        # Centrar ventana
        ventana_login.update_idletasks()
        x = (ventana_login.winfo_screenwidth() // 2) - (600 // 2)
        y = (ventana_login.winfo_screenheight() // 2) - (330 // 2)
        ventana_login.geometry(f"600x330+{x}+{y}")

        # Contenedor principal
        main_frame = tk.Frame(ventana_login, bg="white", relief="flat", bd=0)
        main_frame.pack(fill="both", expand=True, padx=16, pady=12)

        content_frame = tk.Frame(main_frame, bg="white")
        content_frame.pack(fill="both", expand=True, padx=8, pady=8)

        # Card horizontal: logo | divisor | formulario
        card = tk.Frame(content_frame, bg="white", bd=0, highlightthickness=0)
        card.pack(fill="both", expand=True, padx=2, pady=4)

        card.grid_columnconfigure(0, weight=0)
        card.grid_columnconfigure(1, weight=0)
        card.grid_columnconfigure(2, weight=1)
        card.grid_rowconfigure(0, weight=1)

        # Logo
        left_panel = tk.Frame(card, bg="white")
        left_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 10), pady=(2, 6))

        logo_path = obtener_ruta_absoluta_gestor("logo.png")
        print(f"🔍 DEBUG LOGIN: Buscando logo en: {logo_path}")
        print(f"🔍 DEBUG LOGIN: Logo existe: {os.path.exists(logo_path)}")
        if os.path.exists(logo_path):
            try:
                # Asegurar que PIL está disponible
                if not PIL_AVAILABLE:
                    cargar_pil()
                
                logo_img = Image.open(logo_path)
                logo_img = logo_img.resize((260, 260), Image.Resampling.LANCZOS)
                self.logo_photo = ImageTk.PhotoImage(logo_img)
                tk.Label(left_panel, image=self.logo_photo, bg="white").pack(anchor="nw", padx=0, pady=0)
                print(f"✅ DEBUG LOGIN: Logo cargado exitosamente")
            except Exception as e:
                print(f"❌ DEBUG LOGIN: Error cargando logo: {e}")
                tk.Label(left_panel, text="LISTA DE ESPERA", font=("Segoe UI", 18, "bold"), fg="#1976D2", bg="white").pack(anchor="nw", padx=0, pady=0)
        else:
            print(f"⚠️ DEBUG LOGIN: Logo no encontrado, usando texto")
            tk.Label(left_panel, text="LISTA DE ESPERA", font=("Segoe UI", 18, "bold"), fg="#1976D2", bg="white").pack(anchor="nw", padx=0, pady=0)

        # Divisor
        divider = ttk.Separator(card, orient="vertical")
        divider.grid(row=0, column=1, sticky="ns", pady=8)

        # Formulario
        right_panel = tk.Frame(card, bg="white")
        right_panel.grid(row=0, column=2, sticky="nsew", padx=(10, 12), pady=12)

        form_frame = tk.Frame(right_panel, bg="white")
        form_frame.pack(fill="both", expand=True)

        # Contenedor angosto para reducir ancho de entradas y botón
        fields_frame = tk.Frame(form_frame, bg="white")
        fields_frame.pack(anchor="w")

        tk.Label(fields_frame, text="Usuario", font=("Segoe UI", 10, "normal"), fg="#333333", bg="white").pack(anchor="w", pady=(6, 4))
        entry_usuario = ttk.Entry(fields_frame, bootstyle="primary", font=("Segoe UI", 10), width=28)
        entry_usuario.pack(anchor="w", ipady=1, pady=(0, 6))
        if remember_access_saved and saved_username:
            try:
                entry_usuario.insert(0, saved_username)
            except Exception:
                pass

        tk.Label(fields_frame, text="Contraseña", font=("Segoe UI", 10, "normal"), fg="#333333", bg="white").pack(anchor="w", pady=(0, 4))
        entry_password = ttk.Entry(fields_frame, bootstyle="primary", font=("Segoe UI", 10), show="*", width=28)
        entry_password.pack(anchor="w", ipady=1, pady=(0, 4))
        if remember_access_saved and saved_password:
            try:
                entry_password.insert(0, saved_password)
            except Exception:
                pass

        controls_frame = tk.Frame(fields_frame, bg="white")
        controls_frame.pack(anchor="w", pady=(0, 4))
        mostrar_var = tk.BooleanVar(value=False)
        recordar_acceso_var = tk.BooleanVar(value=remember_access_saved)

        def toggle_password():
            try:
                entry_password.configure(show="" if mostrar_var.get() else "*")
            except Exception:
                pass

        # Toggle mostrar con texto más pequeño
        mostrar_row = tk.Frame(controls_frame, bg="white")
        mostrar_row.pack(fill="x")
        chk_mostrar = ttk.Checkbutton(mostrar_row, text="", variable=mostrar_var, bootstyle="primary-round-toggle", command=toggle_password)
        chk_mostrar.pack(side="left")
        tk.Label(mostrar_row, text="Mostrar contraseña", font=("Segoe UI", 10), bg="white", fg="#333333").pack(side="left", padx=(6, 0))

        remember_row = tk.Frame(controls_frame, bg="white")
        remember_row.pack(fill="x", pady=(4, 0))
        chk_recordar_inline = ttk.Checkbutton(remember_row, text="", variable=recordar_acceso_var, bootstyle="primary-round-toggle")
        chk_recordar_inline.pack(side="left")
        tk.Label(remember_row, text="Recordar usuario", font=("Segoe UI", 10), bg="white", fg="#333333").pack(side="left", padx=(6, 0))

        mensaje_frame = tk.Frame(fields_frame, bg="white")
        mensaje_frame.pack(anchor="w", pady=(0, 4))
        label_mensaje = tk.Label(mensaje_frame, text="", font=("Segoe UI", 10), bg="white", wraplength=260, justify="left")
        label_mensaje.pack()

        pb_login = ttk.Progressbar(mensaje_frame, mode="indeterminate", bootstyle="info-striped")
        pb_login.pack(anchor="w", pady=(4, 0))
        pb_login.stop()
        pb_login.pack_forget()

        def mostrar_mensaje(mensaje, tipo="error"):
            if tipo == "error":
                label_mensaje.configure(text=mensaje, fg="#d32f2f")
            elif tipo == "exito":
                label_mensaje.configure(text=mensaje, fg="#388e3c")
            tiempo = 5000 if tipo == "error" else 2000
            limpiar_mensaje_despues_gestor(label_mensaje, tiempo, "")

        def procesar_login():
            username = entry_usuario.get().strip()
            password = entry_password.get()
            if not username or not password:
                mostrar_mensaje("Por favor ingresa usuario y contraseña")
                return
            try:
                btn_login.configure(state="disabled", text="Verificando…")
                try:
                    pb_login.pack(fill="x", pady=(6, 0))
                    pb_login.start(10)
                except Exception:
                    pass
                ventana_login.update_idletasks()
            except Exception:
                pass

            resultado = self.verificar_credenciales(username, password)
            if "error" in resultado:
                mostrar_mensaje(resultado["error"])
                try:
                    btn_login.configure(state="normal", text="INICIAR SESIÓN")
                    try:
                        pb_login.stop()
                        pb_login.pack_forget()
                    except Exception:
                        pass
                except Exception:
                    pass
                return

            self.usuario_actual = resultado
            self.sucursal = resultado.get('sucursal_nombre', 'SUCURSAL PRINCIPAL')

            try:
                cfg_path_local = obtener_ruta_absoluta_gestor("gestor_login.json")
                if recordar_acceso_var.get():
                    try:
                        enc_pwd = base64.b64encode(password.encode('utf-8')).decode('utf-8')
                    except Exception:
                        enc_pwd = ""
                    payload = {"usuario": username, "recordar": True, "recordar_pass": True, "password": enc_pwd}
                    with open(cfg_path_local, 'w', encoding='utf-8') as f:
                        json.dump(payload, f, ensure_ascii=False)
                else:
                    if os.path.exists(cfg_path_local):
                        os.remove(cfg_path_local)
            except Exception:
                pass

            mostrar_mensaje(f"¡Bienvenido {resultado['nombre_completo']}!", "exito")
            ventana_login.after(1500, ventana_login.destroy)
            try:
                btn_login.configure(state="normal", text="INICIAR SESIÓN")
                try:
                    pb_login.stop()
                    pb_login.pack_forget()
                except Exception:
                    pass
            except Exception:
                pass

        # Botón un poco más estrecho y texto ligeramente más alto
        btn_login = ttk.Button(fields_frame, text="INICIAR SESIÓN", bootstyle="info", padding=(10, 8, 10, 14), command=procesar_login, width=26)
        btn_login.pack(anchor="w", ipady=3, pady=(2, 6))

        try:
            entry_usuario.focus_set()
            entry_password.bind("<Return>", lambda e: procesar_login())
            ventana_login.bind("<Escape>", lambda e: ventana_login.destroy())
        except Exception:
            pass

        entry_password.bind("<Return>", lambda e: procesar_login())
        entry_usuario.bind("<Return>", lambda e: entry_password.focus())
        entry_usuario.focus()

        if master is not None:
            try:
                master.wait_window(ventana_login)
            except Exception:
                pass
        else:
            ventana_login.mainloop()

        return self.usuario_actual is not None
    
    def debug_verificar_bd(self):
        """Método de debug para verificar la base de datos"""
        print("🔧 === DEBUG: Verificando Base de Datos ===")
        
        conn = self.conectar_bd()
        if not conn:
            print("❌ No se pudo conectar a la base de datos")
            return
        
        try:
            cur = conn.cursor()
            
            # Verificar tabla usuarios
            cur.execute("SELECT COUNT(*) FROM usuarios WHERE activo = true")
            usuarios_activos = cur.fetchone()[0]
            print(f"✅ Usuarios activos: {usuarios_activos}")
            
            # Verificar usuarios específicos
            cur.execute("SELECT username, rol FROM usuarios WHERE activo = true ORDER BY username")
            usuarios = cur.fetchall()
            print("📋 Usuarios disponibles:")
            for username, rol in usuarios:
                print(f"   - {username} ({rol})")
            
            cur.close()
            conn.close()
            
            print("✅ Base de datos funcionando correctamente")
                
        except Exception as e:
            print(f"❌ Error verificando BD: {e}")
        finally:
            if conn:
                conn.close()

# Función para ejecutar el login del colorista
def ejecutar_login_colorista(master=None):
    """Ejecuta el sistema de login y retorna la información del usuario.

    Si se pasa `master`, el login se muestra como Toplevel modal sobre ese root.
    """
    sistema_login = SistemaLoginColorista()
    ok = sistema_login.mostrar_login(master=master)
    if ok:
        return sistema_login.usuario_actual, sistema_login.sucursal
    return None, None

class GestorListaEspera:
    def __init__(self, usuario_info, sucursal_info, master=None):
        """Inicializa el gestor con información del usuario autenticado.

        Reutiliza un root existente si se proporciona `master` para mantener un solo intérprete Tk.
        """
        if master is not None:
            self.root = master
            try:
                # Aplicar tema de ttkbootstrap al root existente
                ttk.Style(theme="cosmo")
            except Exception:
                try:
                    ttk.Style()
                except Exception:
                    pass
        else:
            # Fallback: crear una ventana de ttkbootstrap si no se proporcionó master
            self.root = ttk.Window(themename="cosmo")
        
        # Información del usuario desde login integrado
        self.usuario_info = usuario_info
        self.usuario_id = str(usuario_info['id'])
        self.usuario_username = usuario_info['username']
        self.usuario_rol = usuario_info['rol']
        self.sucursal_actual = sucursal_info
        
        # Sistema de notificaciones sonoras
        self.notificaciones = NotificacionesSonoras()
        self.ultimo_conteo_pedidos = 0  # Para detectar nuevos pedidos
        self.actualizacion_en_progreso = False  # Controlar actualizaciones
        self.timer_id = None  # Para rastrear el timer de actualización
        
        # Caché para optimización de rendimiento
        self.ultimo_conteo_verificacion = -1
        self.ultima_modificacion_verificacion = None
        self.cache_sucursal = {}  # Caché para consultas de sucursal
        self.cache_codigo_base = {}  # Caché para códigos base
        self.cargando_datos = False  # Flag para evitar cargas simultáneas
        
        # Sistema de archivado automático
        self.ultima_verificacion_archivado = None
        self.archivado_ejecutado_hoy = False
        # Opción: archivar cancelados en histórico antes de borrar (por defecto: OFF)
        self.archivar_cancelados = False

        # Preferencias de impresión
        self.imprimir_al_iniciar = True   # Imprimir automáticamente al iniciar
        self.imprimir_al_finalizar = False  # No imprimir automáticamente al finalizar

        # Interacción del usuario (para evitar refrescos durante gestos)
        self.interactuando = False
        self._interaccion_timer = None
        
        # Control de bloqueo de procesos por factura (4 minutos)
        self.bloqueos_por_factura = {}  # {id_factura: {'bloqueado': bool, 'inicio': timestamp, 'timer': timer_id}}
        self.duracion_bloqueo = 240  # 4 minutos en segundos
        
        # Título con información del usuario autenticado (unificado con PaintFlow)
        titulo_completo = f"PaintFlow — Producción | Usuario: {self.usuario_username} ({self.usuario_rol.title()}) | {self.sucursal_actual}"
        try:
            aplicar_icono_y_titulo(self.root, None)  # Solo icono; el título se establece abajo
        except Exception:
            pass
        self.root.title(titulo_completo)
        self.root.geometry("1500x800")
        
        # Asegurar índices que mejoran el rendimiento de consultas por factura/estado
        try:
            self._asegurar_indices_sucursal()
        except Exception:
            pass

        self.crear_interfaz()
        self.cargar_datos()
        
        # Configurar protocolo de cierre
        self.root.protocol("WM_DELETE_WINDOW", self.cerrar_aplicacion)
        
        # Programar actualización automática cada minuto
        self.actualizar_tiempos_automatico()
        
        # Recordatorio sonoro cada 30s mientras existan pendientes
        self._recordatorio_timer_id = None
        self._recordatorio_intervalo_ms = 30000  # 30 segundos
        self.iniciar_recordatorio_pendientes()

    def _asegurar_indices_sucursal(self):
        """Crea índices claves (id_factura, (id_factura, estado), codigo, presentacion) si no existen.
        Esto acelera la carga por listas y operaciones de grupo.
        """
        try:
            conn = self.conectar_db()
            if not conn:
                return
            sucursal_usuario = obtener_sucursal_usuario(self.usuario_username)
            tabla = obtener_tabla_sucursal(sucursal_usuario)
            cur = conn.cursor()

            # Detectar columnas existentes
            cur.execute("""
                SELECT column_name FROM information_schema.columns
                WHERE table_schema='public' AND table_name=%s
            """, (tabla,))
            cols = {r[0] for r in cur.fetchall()}

            # Índices siempre útiles
            cur.execute(f"CREATE INDEX IF NOT EXISTS idx_{tabla}_id_factura ON {tabla}(id_factura);")
            cur.execute(f"CREATE INDEX IF NOT EXISTS idx_{tabla}_factura_estado ON {tabla}(id_factura, estado);")
            cur.execute(f"CREATE INDEX IF NOT EXISTS idx_{tabla}_codigo ON {tabla}(codigo);")
            if 'presentacion' in cols:
                cur.execute(f"CREATE INDEX IF NOT EXISTS idx_{tabla}_presentacion ON {tabla}(presentacion);")

            conn.commit()
            cur.close(); conn.close()
        except Exception as e:
            try:
                conn.rollback()
            except Exception:
                pass
            try:
                cur.close(); conn.close()
            except Exception:
                pass
    
    def calcular_tiempo_restante(self, fecha_compromiso, estado, tiempo_estimado, row_id):
        """Calcula tiempo restante basado en estado y tiempo estimado"""
        # Si está finalizado, cancelado o completado, no mostrar tiempo
        if estado in ['Finalizado', 'Completado', 'Cancelado']:
            return "N/A"
        
        # Si no hay fecha de compromiso, no calcular
        if not fecha_compromiso:
            return "N/A"
        
        # Usar tiempo estimado como base, o generar uno consistente basado en ID
        if tiempo_estimado and tiempo_estimado > 0:
            base_tiempo = tiempo_estimado
        else:
            # Generar tiempo consistente basado en row_id para que no cambie
            base_tiempo = (row_id * 7) % 90 + 15  # Entre 15 y 105 minutos
        
        # Simular tiempo restante como un porcentaje del tiempo estimado
        # (En producción real, esto sería la diferencia entre fecha_compromiso y now())
        tiempo_restante = int(base_tiempo * 0.7)  # 70% del tiempo estimado
        
        if tiempo_restante > 60:
            horas = tiempo_restante // 60
            minutos = tiempo_restante % 60
            return f"{horas}h {minutos}m"
        else:
            return f"{tiempo_restante}m"
    
    def actualizar_tiempos_automatico(self):
        """Actualiza los tiempos restantes automáticamente cada 10 segundos"""
        try:
            # Verificar que la ventana aún existe y no está cerrada
            if not hasattr(self, 'root') or not self.root or not self.root.winfo_exists():
                print("🔍 Ventana cerrada, deteniendo actualizaciones automáticas")
                return
            
            # Evitar múltiples actualizaciones simultáneas
            if self.actualizacion_en_progreso:
                self.timer_id = self.root.after(3000, self.actualizar_tiempos_automatico)  # Comunicación más rápida
                return
            
            # Solo actualizar si la ventana no está cerrada
            self.actualizacion_en_progreso = True
            
            # Verificar archivado automático (6pm)
            self.verificar_archivado_automatico()
            
            self.cargar_datos()
            self.actualizacion_en_progreso = False
            
            # Programar la próxima actualización solo si la ventana sigue existiendo (más frecuente)
            if hasattr(self, 'root') and self.root and self.root.winfo_exists():
                self.timer_id = self.root.after(3000, self.actualizar_tiempos_automatico)  # 3 segundos para comunicación instantánea
                
        except tk.TclError:
            # Error común cuando la ventana se cierra
            print("🔍 Ventana cerrada durante actualización, deteniendo timer")
            self.actualizacion_en_progreso = False
        except Exception as e:
            print(f"Error en actualización automática: {e}")
            self.actualizacion_en_progreso = False
            # Intentar reprogramar solo si la ventana existe
            try:
                if hasattr(self, 'root') and self.root and self.root.winfo_exists():
                    self.timer_id = self.root.after(5000, self.actualizar_tiempos_automatico)  # Reintento más rápido en 5s
            except:
                print("🔍 No se puede reprogramar actualización, ventana no disponible")
    
    def on_filtro_change(self, event=None):
        """Maneja el cambio de filtros"""
        # Forzar recarga para que el filtro se aplique inmediatamente
        self.cargar_datos(forzar_recarga=True)
    
    def obtener_nombre_sucursal(self, sucursal_id):
        """Obtiene el nombre de la sucursal desde la base de datos con caché"""
        if not sucursal_id:
            return "SUCURSAL PRINCIPAL"
        
        # Verificar caché primero
        if sucursal_id in self.cache_sucursal:
            return self.cache_sucursal[sucursal_id]
        
        try:
            conn = self.conectar_db()
            if not conn:
                return "SUCURSAL PRINCIPAL"
            
            cur = conn.cursor()
            cur.execute("SELECT nombre FROM sucursales WHERE id = %s", (sucursal_id,))
            resultado = cur.fetchone()
            
            cur.close()
            conn.close()
            
            nombre_sucursal = resultado[0] if resultado else "SUCURSAL PRINCIPAL"
            
            # Guardar en caché
            self.cache_sucursal[sucursal_id] = nombre_sucursal
            
            return nombre_sucursal
                
        except Exception as e:
            print(f"Error obteniendo nombre de sucursal: {e}")
            return "SUCURSAL PRINCIPAL"
    
    def conectar_db(self):
        """Conecta a la base de datos PostgreSQL con reintentos y optimización"""
        # Cargar psycopg2 dinámicamente cuando sea necesario
        if not cargar_psycopg2():
            print("❌ No se puede conectar: Driver PostgreSQL no disponible")
            return None
            
        max_intentos = 3
        tiempo_espera = 1
        
        for intento in range(max_intentos):
            try:
                # Conexión optimizada con parámetros de rendimiento
                conn = psycopg2.connect(
                    **DB_CONFIG,
                    connect_timeout=10,
                    application_name='GestorListaEspera'
                )
                # Configurar para mejor rendimiento
                conn.set_session(autocommit=False)
                return conn
            except Exception as e:
                if intento < max_intentos - 1:
                    print(f"⚠️ Intento {intento + 1} fallido, reintentando en {tiempo_espera}s...")
                    time.sleep(tiempo_espera)
                    tiempo_espera *= 2  # Backoff exponencial
                else:
                    messagebox.showerror("Error de Conexión", f"No se pudo conectar a la base de datos tras {max_intentos} intentos:\n{e}")
                    return None
            except Exception as e:
                messagebox.showerror("Error de Conexión", f"Error inesperado al conectar a la base de datos:\n{e}")
                return None
    
    def crear_interfaz(self):
        """Crea la interfaz principal con diseño minimalista y moderno"""
        # Estilos ligeros y modernos con fuentes más grandes
        try:
            style = ttk.Style()
            style.configure("Modern.Treeview", rowheight=30, font=("Segoe UI", 11))
            style.configure("Treeview.Heading", font=("Segoe UI", 12, "bold"))
            # Mejorar contraste de selección
            style.map("Modern.Treeview",
                      background=[("selected", "#1976D2")],
                      foreground=[("selected", "white")])
        except Exception:
            pass
        # Header agrupado y moderno
        header_frame = ttk.Frame(self.root, style="Card.TFrame")
        header_frame.pack(fill="x", padx=10, pady=8)

        ttk.Label(header_frame, text="🏭 Sistema de Lista de Espera - Producción", font=("Segoe UI", 18, "bold"), style="Card.TLabel").pack(side="left", padx=8)

        # Agrupar info usuario, sucursal y config impresora en el header (derecha)
        right_header = ttk.Frame(header_frame, style="Card.TFrame")
        right_header.pack(side="right")
        if hasattr(self, 'usuario_id') and self.usuario_id != "sistema":
            ttk.Label(right_header, text=f"👤 {self.usuario_id} ({self.usuario_rol.title()})", font=("Segoe UI", 11), style="Card.TLabel", bootstyle="success").pack(side="right", padx=5)
        ttk.Label(right_header, text=f"🏢 {self.sucursal_actual}", font=("Segoe UI", 13), style="Card.TLabel", bootstyle="info").pack(side="right", padx=5)
        impresora_actual = cargar_impresora_guardada()
        self.label_impresora = ttk.Label(right_header, text=f"🖨️ {impresora_actual[:20] if impresora_actual else 'Sin configurar'}", font=("Segoe UI", 10), style="Card.TLabel", bootstyle="success" if impresora_actual else "warning")
        self.label_impresora.pack(side="right", padx=5)
        ttk.Button(right_header, text="🖨️ Config", command=self.configurar_impresora, style="Modern.TButton", bootstyle="secondary").pack(side="right", padx=5)

        # Filtros y actualizar agrupados en una sola barra
        filtros_bar = ttk.Frame(self.root, style="Card.TFrame")
        filtros_bar.pack(fill="x", padx=10, pady=4)
        ttk.Label(filtros_bar, text="Estado:", style="Card.TLabel", font=("Segoe UI", 11)).pack(side="left", padx=5)
        self.filtro_estado = ttk.Combobox(filtros_bar, values=["Todos", "Pendiente", "En Espera", "En Proceso", "Finalizados"], state="readonly", width=12, font=("Segoe UI", 10))
        self.filtro_estado.set("Todos")
        self.filtro_estado.pack(side="left", padx=5)
        self.filtro_estado.bind('<<ComboboxSelected>>', self.on_filtro_change)
        ttk.Label(filtros_bar, text="Prioridad:", style="Card.TLabel", font=("Segoe UI", 11)).pack(side="left", padx=5)
        self.filtro_prioridad = ttk.Combobox(filtros_bar, values=["Todas", "Alta", "Media", "Baja"], state="readonly", width=12, font=("Segoe UI", 10))
        self.filtro_prioridad.set("Todas")
        self.filtro_prioridad.pack(side="left", padx=5)
        self.filtro_prioridad.bind('<<ComboboxSelected>>', self.on_filtro_change)
        ttk.Button(filtros_bar, text="🔄 Actualizar", command=self.cargar_datos, style="Modern.TButton", bootstyle="primary").pack(side="left", padx=16)
        self.label_ultima_actualizacion = ttk.Label(filtros_bar, text="🕐 Actualizando...", font=("Segoe UI", 10), style="Card.TLabel", bootstyle="secondary")
        self.label_ultima_actualizacion.pack(side="left", padx=16)

        # Mensajes de impresión
        self.label_mensaje = ttk.Label(self.root, text="", font=("Segoe UI", 11), style="Card.TLabel", bootstyle="info")
        self.label_mensaje.pack(pady=2)
        # Barra de acciones (se eliminan botones duplicados por menú contextual)
        acciones_bar = ttk.Frame(self.root, style="Card.TFrame")
        acciones_bar.pack(fill="x", padx=10, pady=4)
        # Nota: Las acciones Iniciar, Finalizar, Cancelar, Próximos a Vencer y Sonido
        # ahora se acceden desde el menú contextual (clic derecho) o atajos.
    # (Se removieron botones utilitarios; se mantendrá solo el toggle de archivado)
        # Placeholder para compatibilidad cuando no exista botón de sonido
        self.btn_sonido = None

        # Toggle opcional para archivar cancelados en histórico
        self.var_archivar_cancelados = tk.BooleanVar(value=self.archivar_cancelados)
        chk_archivar = ttk.Checkbutton(
            acciones_bar,
            text="Guardar cancelados en histórico",
            variable=self.var_archivar_cancelados,
            command=self._toggle_archivar_cancelados,
            bootstyle="primary-round-toggle"
        )
        chk_archivar.pack(side="right", padx=10)

        # Tabla minimalista con menos columnas y estilo moderno
        tree_frame = ttk.Frame(self.root, style="Card.TFrame")
        tree_frame.pack(fill="both", expand=True, padx=10, pady=6)
        self.tree = ttk.Treeview(tree_frame, style="Modern.Treeview")
        columnas = ("ID Prof.", "Factura", "Código", "Producto", "Terminación", "Código Base", "Prioridad", "Cantidad", "Estado", "Tiempo Est.", "Operador")
        self.tree["columns"] = columnas
        self.tree["show"] = "headings"
        anchos = {
            "ID Prof.": 100,
            "Factura": 100,
            "Código": 100,
            "Producto": 180,
            "Terminación": 105,
            "Código Base": 120,
            "Prioridad": 90,
            "Cantidad": 80,
            "Estado": 110,
            "Tiempo Est.": 105,
            "Operador": 120
        }
        # Alineación por columna para mejor lectura
        anchors = {
            "ID Prof.": "center",
            "Factura": "center",
            "Código": "center",
            "Producto": "w",
            "Terminación": "center",
            "Código Base": "center",
            "Prioridad": "center",
            "Cantidad": "center",
            "Estado": "center",
            "Tiempo Est.": "center",
            "Operador": "w"
        }
        for col in columnas:
            self.tree.heading(col, text=col, anchor=anchors.get(col, "w"))
            self.tree.column(col, width=anchos.get(col, 100), anchor=anchors.get(col, "w"))
        scrollbar_v = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        scrollbar_h = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=scrollbar_v.set, xscrollcommand=scrollbar_h.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        scrollbar_v.grid(row=0, column=1, sticky="ns")
        scrollbar_h.grid(row=1, column=0, sticky="ew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        self.tree.tag_configure('alta', background='#ffebee')
        self.tree.tag_configure('media', background='#fff3e0')
        self.tree.tag_configure('baja', background='#e8f5e8')
        self.tree.tag_configure('proceso', background='#e3f2fd')
        self.tree.tag_configure('finalizado', background='#f3e5f5')

        # Marcar interacción de usuario para pausar refrescos pesados
        try:
            eventos_interaccion = ["<ButtonPress-1>", "<B1-Motion>", "<MouseWheel>", "<KeyPress>"]
            for ev in eventos_interaccion:
                self.tree.bind(ev, self._marcar_interaccion)
        except Exception:
            pass

    def _marcar_interaccion(self, event=None):
        """Marca interacción por un breve periodo para evitar refrescos durante gestos."""
        try:
            self.interactuando = True
            # Cancelar temporizador previo si existe
            if self._interaccion_timer:
                try:
                    self.root.after_cancel(self._interaccion_timer)
                except Exception:
                    pass
            # Resetear flag tras 800ms sin nuevos eventos
            self._interaccion_timer = self.root.after(800, self._reset_interaccion)
        except Exception:
            pass

    def _reset_interaccion(self):
        try:
            self.interactuando = False
            self._interaccion_timer = None
        except Exception:
            pass

        # Menú contextual (clic derecho) sobre la tabla
        self.menu_contextual = tk.Menu(self.root, tearoff=0)
        self.menu_contextual.add_command(label="▶️ Iniciar Producción", command=self.iniciar_produccion, accelerator="Ctrl+Enter")
        self.menu_contextual.add_command(label="✅ Finalizar", command=self.finalizar_pedido, accelerator="Ctrl+F")
        self.menu_contextual.add_command(label="🖨️ Imprimir Etiqueta", command=self.imprimir_etiqueta, accelerator="Ctrl+P")
        # Nueva opción: ver fórmula de la presentación seleccionada
        self.menu_contextual.add_command(label="📘 Fórmula", command=self.mostrar_formula, accelerator="Ctrl+L")
        # Simplificado: una sola opción por acción; se aplicará a pedido o lista según contexto
        # Acciones utilitarias también disponibles por clic derecho
        self.menu_contextual.add_command(label="⏰ Próximos a Vencer", command=self.mostrar_proximos_vencer)
        self.menu_contextual.add_command(label="🔔 Test Sonido", command=self.test_notificacion)
        self.menu_contextual.add_command(label="🔊 Alternar Sonido", command=self.alternar_sonido)
        self.menu_contextual.add_separator()
        self.menu_contextual.add_command(label="❌ Cancelar", command=self.cancelar_pedido, accelerator="Del")

        # Asegurar selección del renglón bajo el cursor y mostrar menú
        self.tree.bind("<Button-3>", self._mostrar_menu_contextual)
        # Atajos de teclado
        self._configurar_atajos()
    
    def verificar_cambios_pendientes(self):
        """Verifica si hay cambios en la tabla antes de recargar completamente"""
        conn = self.conectar_db()
        if not conn:
            return True  # Si no hay conexión, forzar recarga
        
        try:
            sucursal_usuario = obtener_sucursal_usuario(self.usuario_username)
            tabla_sucursal = obtener_tabla_sucursal(sucursal_usuario)
            
            cur = conn.cursor()
            # Consulta más sensible para detectar nuevos pedidos incluso si las fechas máximas no cambian
            cur.execute(f"""
                SELECT 
                    COUNT(*)                                  AS total,
                    COALESCE(MAX(id), 0)                       AS max_id,
                    COALESCE(SUM(CASE WHEN TRIM(prioridad) = 'Alta'  THEN 1 ELSE 0 END), 0) AS alta,
                    COALESCE(SUM(CASE WHEN TRIM(prioridad) = 'Media' THEN 1 ELSE 0 END), 0) AS media,
                    COALESCE(SUM(CASE WHEN TRIM(prioridad) = 'Baja'  THEN 1 ELSE 0 END), 0) AS baja,
                    COALESCE(SUM(CASE WHEN TRIM(estado) IN ('Pendiente','En Espera','En Proceso') THEN 1 ELSE 0 END), 0) AS activos,
                    COALESCE(MAX(fecha_creacion),   '1970-01-01'::timestamp) AS ultima_creacion,
                    COALESCE(MAX(fecha_asignacion), '1970-01-01'::timestamp) AS ultima_asignacion,
                    COALESCE(MAX(fecha_completado), '1970-01-01'::timestamp) AS ultima_finalizacion
                FROM {tabla_sucursal}
                WHERE TRIM(estado) <> 'Cancelado'
            """)
            
            resultado = cur.fetchone()
            total, max_id, alta, media, baja, activos, ultima_creacion, ultima_asignacion, ultima_finalizacion = resultado
            
            # Crear una huella digital más robusta de los datos actuales
            huella_actual = f"{total}|{max_id}|{alta}|{media}|{baja}|{activos}|{ultima_creacion}|{ultima_asignacion}|{ultima_finalizacion}"
            
            cur.close()
            conn.close()
            
            # Verificar si hay cambios comparando la huella digital
            if hasattr(self, 'huella_datos_anterior'):
                if huella_actual == self.huella_datos_anterior:
                    return False  # No hay cambios
            
            # Guardar huella actual para próxima verificación
            self.huella_datos_anterior = huella_actual
            
            print(f"🔍 DEBUG: Detectados cambios en datos - Recargando")
            return True  # Hay cambios
            
        except Exception as e:
            print(f"Error verificando cambios: {e}")
            return True  # En caso de error, forzar recarga

    def obtener_base_desde_codigo(self, codigo):
        """Obtiene la base de un código desde ProductSW"""
        if not codigo:
            return ""
            
        try:
            conn = self.conectar_db()
            if not conn:
                return ""
            
            cur = conn.cursor()
            cur.execute("SELECT base FROM ProductSW WHERE codigo = %s AND activo = TRUE", (codigo,))
            resultado = cur.fetchone()
            
            cur.close()
            conn.close()
            
            if resultado:
                print(f"🔍 DEBUG: Base encontrada para {codigo}: {resultado[0]}")
                return resultado[0]
            else:
                print(f"🔍 DEBUG: No se encontró base para código {codigo}")
                return ""
                
        except Exception as e:
            print(f"Error obteniendo base desde código: {e}")
            return ""

    def obtener_descripcion_codigo(self, codigo):
        """Obtiene la descripción (nombre) de un código desde ProductSW"""
        if not codigo:
            return codigo  # Si no hay código, retornar el código mismo
            
        try:
            conn = self.conectar_db()
            if not conn:
                return codigo
            
            cur = conn.cursor()
            cur.execute("SELECT nombre FROM ProductSW WHERE codigo = %s AND activo = TRUE", (codigo,))
            resultado = cur.fetchone()
            
            cur.close()
            conn.close()
            
            if resultado and resultado[0]:
                print(f"🔍 DEBUG: Descripción encontrada para {codigo}: {resultado[0]}")
                return resultado[0]
            else:
                print(f"🔍 DEBUG: No se encontró descripción para código {codigo}, usando código")
                return codigo  # Si no se encuentra descripción, usar el código
                
        except Exception as e:
            print(f"Error obteniendo descripción desde código: {e}")
            return codigo  # En caso de error, usar el código

    def obtener_codigo_base_desde_db(self, base, producto, terminacion, presentacion=None):
        """Obtiene el código base usando la misma lógica exacta de LabelsApp con caché + sufijo presentación"""
        if not base or not producto or not terminacion:
            return "N/A"
        
        # Crear clave de caché incluyendo presentación
        presentacion_str = presentacion or ""
        cache_key = f"{base.lower()}|{producto.lower()}|{terminacion.lower()}|{presentacion_str}"
        
        # Verificar caché primero
        if cache_key in self.cache_codigo_base:
            return self.cache_codigo_base[cache_key]
        
        try:
            # Solo cargar tabla CodigoBase una vez y cachearla
            if not hasattr(self, 'codigo_base_data'):
                self.cargar_tabla_codigo_base()
            
            # Buscar en datos cacheados
            base_lower = base.lower()
            for row_data in self.codigo_base_data:
                if row_data['base'].lower() == base_lower:
                    resultado = self.calcular_codigo_base_logica(row_data, producto, terminacion, base)
                    
                    # Agregar sufijo de presentación si está disponible
                    if presentacion and resultado not in ["No encontrado", "No Aplica", "Error"]:
                        sufijo_presentacion = obtener_sufijo_presentacion(presentacion)
                        if sufijo_presentacion:
                            # Si el código base ya termina con guión, no agregar otro guión
                            if resultado.endswith('-'):
                                resultado += sufijo_presentacion
                            else:
                                resultado += '-' + sufijo_presentacion
                    
                    # Guardar en caché
                    self.cache_codigo_base[cache_key] = resultado
                    return resultado
            
            # Si no se encuentra
            self.cache_codigo_base[cache_key] = "No encontrado"
            return "No encontrado"
            
        except Exception as e:
            print(f"Error obteniendo código base: {e}")
            return "Error"
    
    def cargar_tabla_codigo_base(self):
        """Carga toda la tabla CodigoBase en memoria una sola vez"""
        try:
            conn = self.conectar_db()
            if not conn:
                self.codigo_base_data = []
                return
            
            cur = conn.cursor()
            cur.execute("""
                SELECT base, tath, tath2, tath3, flat, satin, sgi, flat2, satin3, sg4, 
                       satinkq, flatkp, flatmp, flatcov, flatpas, satinem, sgem, 
                       flatsp, satinsp, glossp, flatap, satinap, satinsan 
                FROM CodigoBase 
            """)
            
            rows = cur.fetchall()
            self.codigo_base_data = []
            
            for row in rows:
                self.codigo_base_data.append({
                    'base': row[0],
                    'tath': row[1], 'tath2': row[2], 'tath3': row[3],
                    'flat': row[4], 'satin': row[5], 'sgi': row[6],
                    'flat2': row[7], 'satin3': row[8], 'sg4': row[9],
                    'satinkq': row[10], 'flatkp': row[11], 'flatmp': row[12],
                    'flatcov': row[13], 'flatpas': row[14], 'satinem': row[15],
                    'sgem': row[16], 'flatsp': row[17], 'satinsp': row[18],
                    'glossp': row[19], 'flatap': row[20], 'satinap': row[21],
                    'satinsan': row[22]
                })
            
            cur.close()
            conn.close()
            print(f"🔍 DEBUG: Tabla CodigoBase cargada en memoria: {len(self.codigo_base_data)} registros")
            
        except Exception as e:
            print(f"Error cargando tabla CodigoBase: {e}")
            self.codigo_base_data = []
    
    def calcular_codigo_base_logica(self, row_data, producto, terminacion, base):
        """Calcula el código base usando la lógica COMPLETA de LabelsApp"""
        try:
            producto = producto.lower()
            terminacion = terminacion.lower()
            base_color = base.lower()
            
            # Identificar TODOS los tipos de producto (lógica exacta de LabelsApp)
            es_esmalte = "esmalte multiuso" in producto
            es_kempro = "kem pro" in producto
            es_kemaqua = "kem aqua" in producto
            es_masterpaint = "master paint" in producto
            es_pastel = "excello pastel" in producto
            es_emerald = "emerald" in producto
            es_superpaint = "super paint" in producto
            es_superpaintAP = "airpurtec" in producto
            es_sanitizing = "sanitizing" in producto
            es_laca = "laca" in producto
            es_EsmalteIndustrial = "esmalte kem" in producto
            es_uretano = "uretano" in producto
            es_tintealthinner = "tinte al thinner" in producto
            es_monocapa = "monocapa" in producto
            es_excellocov = "excello voc" in producto
            es_excellopremium = "excello premium" in producto
            es_waterblocking = "water blocking" in producto
            es_airpuretec = "airpuretec" in producto
            es_hcsiloconeacr = "h&c silicone-acrylic" in producto
            es_hcheavyshield = "h&c heavy-shield" in producto
            es_ProMarEgShel = "promar® 200 voc" in producto
            es_ProMarEgShel400 = "promar® 400 voc" in producto
            es_proindustrialDTM = "pro industrial dtm" in producto
            es_armoseal = "armoseal 1000hs" in producto
            es_armosealtp = "armoseal t-p" in producto
            es_scufftuff = "scuff tuff-wb" in producto
            
            # LÓGICA COMPLETA COPIADA DE LABELSAPP
            
            if es_kemaqua:

                if terminacion == "satin":
                    return row_data.get('satinkq') or "No Aplica"
                else:
                    return "No Aplica"
            
            if es_airpuretec:

                if terminacion == "mate":
                    if base_color == "extra white":
                        return "A86W00061-"
                    elif base_color == "deep":
                        return "A86W00063-"
                elif terminacion == "satin":
                    if base_color == "extra white":
                        return "A87W00061-"
                    elif base_color == "deep":
                        return "A87W00063-"
                else:
                    return "No Aplica"
                    
            if es_waterblocking:

                if terminacion == "mate":
                    return "LX12WDR50-"
                else:
                    return "No Aplica"
                    
            if es_excellocov:

                if terminacion == "mate":
                    return row_data.get('flatcov') or "No Aplica"
                elif terminacion == "satin" and base_color == "extra white":
                    return "A20DR2651-"
                else:
                    return "No Aplica"
                    
            if es_laca:

                if terminacion in ["mate", "semimate", "brillo"]:
                    return "L15-"
                else:
                    return "No Aplica"
                    
            if es_EsmalteIndustrial:

                if terminacion in ["mate", "semimate", "brillo"]:
                    return "F300-"
                else:
                    return "No Aplica"
                    
            if es_hcsiloconeacr:

                if terminacion == "mate":
                    if base_color == "extra white":
                        return "20.101214-"
                    elif base_color == "deep":
                        return "20.102214-1"
                    elif base_color == "ultra deep":
                        return "20.103214-"
                    else:
                        return "No aplica"
                else:
                    return "No Aplica"
                    
            if es_proindustrialDTM:

                if terminacion == "gloss":
                    if base_color == "extra white":
                        return "B66W1051-"
                    elif base_color == "ultra deep":
                        return "B66T1054-"
                    else:
                        return "No aplica"
                else:
                    return "No Aplica"
                    
            if es_scufftuff:
                if terminacion == "mate":
                    if base_color == "extra white":
                        return "S23W00051-"
                    elif base_color == "ultra deep":
                        return "S23T00154-"
                    elif base_color == "deep":
                        return "S23W00153-"
                    else:
                        return "No aplica"
                elif terminacion == "satin":
                    return "S24W00051-"
                elif terminacion == "semigloss":
                    return "S26W00051-"
                else:
                    return "No Aplica"
                    
            if es_hcheavyshield:

                if terminacion == "gloss":
                    if base_color == "extra white":
                        return "35.100214-"
                    elif base_color == "deep":
                        return "35.100314-"
                    elif base_color == "ultra deep":
                        return "35.100414-"
                    else:
                        return "No aplica"
                else:
                    return "No Aplica"
                    
            if es_ProMarEgShel:

                if terminacion == "satin":
                    if base_color == "deep":
                        return "B20W02653-"
                    elif base_color == "extra white":
                        return "B20W12651-"
                    else:
                        return "No aplica"
                elif terminacion == "mate":
                    if base_color == "ultra deep":
                        return "B30T02654-"
                    elif base_color == "extra White":
                        return "B30W02651-1"
                    elif base_color == "deep":
                        return "B30W02653-"
                elif terminacion == "semigloss":
                    if base_color == "extra White":
                        return "B31W02651-"
                else:
                    return "No aplica"
                    
            if es_ProMarEgShel400:
                if terminacion == "satin":
                    if base_color == "extra white":
                        return "B20W04651-"
                    else:
                        return "No aplica"
                else:
                    return "No Aplica"
                    
            if es_armoseal:
                if terminacion == "gloss":
                    if base_color == "extra white":
                        return "B67W2001-"
                    elif base_color == "ultra deep":
                        return "B67T2004-"
                    else:
                        return "No aplica"
                else:
                    return "No Aplica"
                    
            if es_armosealtp:
                if terminacion == "semigloss":
                    if base_color == "extra white":
                        return "B90T104-"
                    elif base_color == "ultra deep":
                        return "B90W111-"
                    else:
                        return "No aplica"
                else:
                    return "No Aplica"
                    
            if es_uretano:
                if terminacion in ["mate", "semimate", "brillo"]:
                    if base_color == "extra white":
                        return "ASPPA-"
                    elif base_color in ["deep", "ultra deep"]:
                        return "ASPPB-"
                    else:
                        return "ASPPD-"
                else:
                    return "No Aplica"
                    
            if es_tintealthinner:
                if terminacion == "claro":
                    return row_data.get('tath') or "No Aplica"
                elif terminacion == "intermedio":
                    return row_data.get('tath2') or "No Aplica"
                elif terminacion == "especial":
                    return row_data.get('tath3') or "No Aplica"
                else:
                    return "No Aplica"
                    
            if es_monocapa:
                if terminacion in ["mate", "semimate", "brillo"]:
                    if base_color == "extra white":
                        return "ASMCA-"
                    elif base_color in ["deep", "ultra deep"]:
                        return "ASMCB-"
                    else:
                        return "ASMCD-"
                else:
                    return "No Aplica"
                    
            if es_esmalte:

                if terminacion == "mate":
                    return row_data.get('flat2') or "No Aplica"
                elif terminacion == "satin":
                    return row_data.get('satin3') or "No Aplica"
                elif terminacion == "gloss":
                    return row_data.get('sg4') or "No Aplica"
                else:
                    return "No Aplica"
                    
            if es_kempro:

                if terminacion == "mate":
                    return row_data.get('flatkp') or "No Aplica"
                else:
                    return "No Aplica"
                    
            if es_masterpaint:

                if terminacion == "mate":
                    return row_data.get('flatmp') or "No Aplica"
                else:
                    return "No Aplica"
                    
            if es_pastel:
                if terminacion == "mate":
                    return row_data.get('flatpas') or "No Aplica"
                else:
                    return "No Aplica"
                    
                
            if es_emerald:

                if terminacion == "satin":
                    return (row_data.get('satinem') or "") + " K37W02751-"
                elif terminacion == "gloss":
                    return (row_data.get('sgem') or "") + "K38W02751-"
                else:
                    return "No Aplica"
                    
            if es_superpaint:

                if terminacion == "mate":
                    return row_data.get('flatsp') or "No Aplica"
                elif terminacion == "satin":
                    return row_data.get('satinsp') or "No Aplica"
                elif terminacion == "gloss":
                    return row_data.get('glossp') or "No Aplica"
                else:
                    return "No Aplica"
                    
            if es_superpaintAP:

                if terminacion == "mate":
                    return row_data.get('flatap') or "No Aplica"
                elif terminacion == "satin":
                    return row_data.get('satinap') or "No Aplica"
                else:
                    return "No Aplica"
                    
            if es_sanitizing:

                if terminacion == "satin":
                    return row_data.get('satinsan') or "No Aplica"
                else:
                    return "No Aplica"
                    
            if es_excellopremium:
                # Reglas especiales Excello Premium
                # 1) Ultra Deep II -> código PP4-
                if any(k in base_color for k in ["ultra deep ii", "ultradeep ii", "ultra-deep ii", "ultra deep 2"]):
                    return "PP4-"

                # 2) Ultra Deep (no II) -> solo Semisatin con código A27WDR03-
                if ("ultra deep" in base_color) and not any(k in base_color for k in ["ultra deep ii", "ultradeep ii", "ultra-deep ii", "ultra deep 2"]):
                    if terminacion == "semisatin":
                        return "A27WDR03-"
                    else:
                        return "No Aplica"

                if terminacion == "mate":
                    return row_data.get('flat') or "No Aplica"
                elif terminacion == "satin":
                    return row_data.get('satin') or "No Aplica"
                elif terminacion == "semigloss":
                    return row_data.get('sgi') or "No Aplica"
                else:
                    return "No Aplica"
            
            # Casos generales como fallback
            if terminacion in ["mate", "flat"]:
                return row_data.get('flat') or row_data.get('flat2') or "No Aplica"
            elif terminacion in ["satin", "semisatinado"]:
                return row_data.get('satin') or row_data.get('satin3') or "No Aplica"
            elif terminacion in ["semigloss", "semi gloss"]:
                return row_data.get('sgi') or row_data.get('sg4') or "No Aplica"
            
            return "No Aplica"
            
        except Exception as e:
            print(f"Error en calcular_codigo_base_logica: {e}")
            return "Error"

    def deducir_presentacion_desde_codigo(self, codigo):
        """Deduce la presentación desde el código del producto si es posible"""
        if not codigo:
            return None
        
        # Buscar patrones comunes en códigos que indiquen presentación
        codigo_upper = codigo.upper()
        
        # Patrones específicos que podrían indicar tamaño
        if "QT" in codigo_upper or "QUART" in codigo_upper:
            return "Cuarto"
        elif "1/2" in codigo or "HALF" in codigo_upper or "MEDIO" in codigo_upper:
            return "Medio Galón"
        elif "GAL" in codigo_upper or codigo_upper.endswith("1"):
            return "Galón" 
        elif "5G" in codigo_upper or "BUCKET" in codigo_upper:
            return "Cubeta"
        elif "1/8" in codigo or "SAMPLE" in codigo_upper:
            return "1/8"
        
        # Si no hay patrones claros, retorna None para usar lógica de cantidad
        return None

    def seleccionar_operador(self):
        """Ventana para digitar el nombre del operador/colorista"""
        try:
            # Crear ventana de entrada con mejor diseño
            ventana = tk.Toplevel(self.root)
            aplicar_icono_y_titulo(ventana, "Asignar Operador")
            ventana.geometry("500x350")
            ventana.resizable(False, False)
            ventana.transient(self.root)
            ventana.grab_set()
            ventana.configure(bg="white")
            
            # Centrar ventana
            ventana.update_idletasks()
            x = (ventana.winfo_screenwidth() // 2) - (500 // 2)
            y = (ventana.winfo_screenheight() // 2) - (350 // 2)
            ventana.geometry(f"500x350+{x}+{y}")
            
            # Variable para el resultado
            operador_seleccionado = [None]
            
            # Frame principal con padding
            main_frame = tk.Frame(ventana, bg="white", padx=30, pady=25)
            main_frame.pack(fill="both", expand=True)
            
            # Encabezado: reemplazar título por LOGO
            titulo_frame = tk.Frame(main_frame, bg="white")
            titulo_frame.pack(fill="x", pady=(0, 10))
            
            try:
                logo_path = obtener_ruta_absoluta_gestor("logo.png")
                if os.path.exists(logo_path):
                    from PIL import Image, ImageTk  # asegurado arriba, pero por seguridad en ejecutables
                    logo_img = Image.open(logo_path)
                    # Tamaño adecuado para el diálogo
                    logo_img = logo_img.resize((180, 80), Image.Resampling.LANCZOS)
                    ventana.logo_operador = ImageTk.PhotoImage(logo_img)
                    tk.Label(titulo_frame, image=ventana.logo_operador, bg="white").pack()
                else:
                    tk.Label(titulo_frame, text="PaintFlow", font=("Segoe UI", 16, "bold"), fg="#1976D2", bg="white").pack()
            except Exception:
                tk.Label(titulo_frame, text="PaintFlow", font=("Segoe UI", 16, "bold"), fg="#1976D2", bg="white").pack()
            
            # Línea decorativa
            linea_decorativa = tk.Frame(main_frame, height=2, bg="#1976D2")
            linea_decorativa.pack(fill="x", pady=(0, 15))
            
            # Información de sucursal con mejor estilo
            sucursal_usuario = obtener_sucursal_usuario(self.usuario_username)
            info_frame = tk.Frame(main_frame, bg="#f0f7ff", relief="solid", bd=1)
            info_frame.pack(fill="x", pady=(0, 20))
            
            info_label = tk.Label(
                info_frame,
                text=f"📍 Sucursal: {sucursal_usuario}",
                font=("Segoe UI", 10, "bold"),
                fg="#1565C0",
                bg="#f0f7ff"
            )
            info_label.pack(pady=8)
            
            # Etiqueta del campo
            label_operador = tk.Label(
                main_frame,
                text="Nombre del Operador/Colorista:",
                font=("Segoe UI", 11, "normal"),
                fg="#333333",
                bg="white"
            )
            label_operador.pack(anchor="w", pady=(0, 8))
            
            # Frame para el campo de entrada con borde mejorado
            entry_frame = tk.Frame(main_frame, bg="#f8f9fa", relief="solid", bd=1)
            entry_frame.pack(fill="x", pady=(0, 5))
            
            # Campo de entrada para el nombre con mejor estilo
            entry_operador = tk.Entry(
                entry_frame,
                font=("Segoe UI", 13),
                relief="flat",
                bd=0,
                highlightthickness=0,
                bg="#f8f9fa",
                fg="#333333",
                insertbackground="#1976D2",
                selectbackground="#1976D2",
                selectforeground="white"
            )
            entry_operador.pack(fill="both", padx=12, pady=12)
            
            # Efectos focus para el campo de entrada
            def on_focus_in(event):
                entry_frame.config(bg="#e3f2fd", bd=2, relief="solid")
                entry_operador.config(bg="#e3f2fd")
            
            def on_focus_out(event):
                entry_frame.config(bg="#f8f9fa", bd=1, relief="solid")
                entry_operador.config(bg="#f8f9fa")
            
            entry_operador.bind("<FocusIn>", on_focus_in)
            entry_operador.bind("<FocusOut>", on_focus_out)
            
            # Precargar con el usuario actual
            try:
                conn = self.conectar_db()
                if conn:
                    cur = conn.cursor()
                    cur.execute("""
                        SELECT nombre_completo, rol 
                        FROM usuarios 
                        WHERE username = %s
                    """, (self.usuario_username,))
                    
                    result = cur.fetchone()
                    if result and result[0]:  # Si tiene nombre completo
                        entry_operador.insert(0, result[0])
                    else:  # Usar username si no hay nombre completo
                        entry_operador.insert(0, self.usuario_username)
                    
                    cur.close()
                    conn.close()
            except Exception:
                # Si hay error, usar el username
                entry_operador.insert(0, self.usuario_username)
            
            # Texto de ayuda con mejor estilo
            ayuda_frame = tk.Frame(main_frame, bg="#fff3e0", relief="solid", bd=1)
            ayuda_frame.pack(fill="x", pady=(5, 25))
            
            ayuda_label = tk.Label(
                ayuda_frame,
                text="💡 Digite el nombre completo del colorista que realizará el trabajo",
                font=("Segoe UI", 10),
                fg="#f57c00",
                bg="#fff3e0",
                wraplength=380
            )
            ayuda_label.pack(pady=6, padx=10)
            
            # Frame para botones con separador
            separator = tk.Frame(main_frame, height=1, bg="#e0e0e0")
            separator.pack(fill="x", pady=(0, 20))
            
            botones_frame = tk.Frame(main_frame, bg="white")
            botones_frame.pack(fill="x")
            
            def confirmar_operador():
                nombre_operador = entry_operador.get().strip()
                if nombre_operador:
                    operador_seleccionado[0] = nombre_operador
                    ventana.destroy()
                else:
                    messagebox.showwarning("Campo Requerido", "Por favor digite el nombre del operador")
                    entry_operador.focus()
            
            def cancelar_seleccion():
                ventana.destroy()
            
            # Efectos hover para botones
            def on_enter_confirmar(event):
                btn_confirmar.config(bg="#1565C0")
            
            def on_leave_confirmar(event):
                btn_confirmar.config(bg="#1976D2")
            
            def on_enter_cancelar(event):
                btn_cancelar.config(bg="#eeeeee")
            
            def on_leave_cancelar(event):
                btn_cancelar.config(bg="#f5f5f5")
            
            # Botón Confirmar mejorado
            btn_confirmar = tk.Button(
                botones_frame,
                text="✓ Confirmar",
                command=confirmar_operador,
                font=("Segoe UI", 11, "bold"),
                bg="#1976D2",
                fg="white",
                relief="flat",
                padx=30,
                pady=12,
                cursor="hand2",
                borderwidth=0,
                activebackground="#1565C0",
                activeforeground="white"
            )
            btn_confirmar.pack(side="right", padx=(15, 0))
            btn_confirmar.bind("<Enter>", on_enter_confirmar)
            btn_confirmar.bind("<Leave>", on_leave_confirmar)
            
            # Botón Cancelar mejorado
            btn_cancelar = tk.Button(
                botones_frame,
                text="✕ Cancelar",
                command=cancelar_seleccion,
                font=("Segoe UI", 11),
                bg="#f5f5f5",
                fg="#666666",
                relief="flat",
                padx=30,
                pady=12,
                cursor="hand2",
                borderwidth=0,
                activebackground="#eeeeee",
                activeforeground="#555555"
            )
            btn_cancelar.pack(side="right")
            
            # Enfocar el campo de entrada y seleccionar todo el texto
            entry_operador.focus()
            entry_operador.select_range(0, tk.END)
            
            # Permitir confirmar con Enter
            def on_enter(event):
                confirmar_operador()
            
            entry_operador.bind('<Return>', on_enter)
            
            # Esperar hasta que se cierre la ventana
            ventana.wait_window()
            
            return operador_seleccionado[0]
            
        except Exception as e:
            print(f"Error en seleccionar_operador: {e}")
            messagebox.showerror("Error", f"Error al asignar operador: {e}")
            return None

    def cargar_datos(self, forzar_recarga=False):
        """Carga los datos de la lista de espera específica de la sucursal"""
        # Evitar cargas simultáneas
        if self.cargando_datos:
            print("🔍 DEBUG: Carga ya en progreso, ignorando solicitud")
            return
        # Evitar recargas durante una interacción activa del usuario (a menos que se fuerce)
        if not forzar_recarga and getattr(self, 'interactuando', False):
            print("⏸️ DEBUG: Interacción activa, diferimos recarga")
            return
        
        # Fallback: forzar recarga periódica aunque no se detecten cambios (evita quedarse desincronizado)
        try:
            ahora_ts = time.time()
            ultima_ts = getattr(self, 'ultima_carga_forzada', 0)
            if not forzar_recarga and (ahora_ts - ultima_ts) > 15:
                print("🔁 DEBUG: Forzando recarga periódica (15s)")
                forzar_recarga = True
                self.ultima_carga_forzada = ahora_ts
        except Exception:
            pass

        # Verificar si realmente necesitamos recargar los datos (a menos que se fuerce)
        if not forzar_recarga and not self.verificar_cambios_pendientes():
            # Solo actualizar el timestamp sin recargar datos
            if hasattr(self, 'label_ultima_actualizacion'):
                hora_actual = datetime.now().strftime("%H:%M:%S")
                self.label_ultima_actualizacion.config(
                    text=f"🕐 Sin cambios: {hora_actual} (optimizado)"
                )
            return
        
        # Marcar como cargando
        self.cargando_datos = True
        # Registrar última carga forzada/real
        try:
            self.ultima_carga_forzada = time.time()
        except Exception:
            pass
        
        # Usar threading para evitar congelamiento
        threading.Thread(
            target=self._cargar_datos_async, 
            args=(forzar_recarga,), 
            daemon=True
        ).start()
    
    def _cargar_datos_async(self, forzar_recarga=False):
        """Carga los datos de forma asíncrona para evitar congelamiento"""
        try:
            start_time = time.time()  # Para medir rendimiento
            conn = self.conectar_db()
            if not conn:
                return
            
            # Detectar sucursal del usuario y tabla correspondiente
            sucursal_usuario = obtener_sucursal_usuario(self.usuario_username)
            tabla_sucursal = obtener_tabla_sucursal(sucursal_usuario)
            
            print(f"🔍 DEBUG: Cargando datos de tabla {tabla_sucursal}")
            
            # Consulta optimizada (calculamos código base en tiempo real)
            query = f"""
                SELECT id, id_orden_profesional, id_factura, codigo, producto, terminacion, 
                       prioridad, cantidad, estado, tiempo_estimado, 
                       operador, fecha_creacion, base, ubicacion, presentacion, fecha_asignacion
                FROM {tabla_sucursal} 
                WHERE 1=1 AND COALESCE(TRIM(estado), '') <> 'Cancelado'
            """
            params = []
            
            # Aplicar filtros de forma eficiente
            filtro_estado = self.filtro_estado.get() if hasattr(self, 'filtro_estado') else "Todos"
            filtro_prioridad = self.filtro_prioridad.get() if hasattr(self, 'filtro_prioridad') else "Todas"
            
            print(f"🔍 DEBUG: Filtros aplicados - Estado: {filtro_estado}, Prioridad: {filtro_prioridad}")
            
            if filtro_estado != "Todos":
                if filtro_estado == "Finalizados":
                    query += " AND COALESCE(TRIM(estado), '') IN ('Finalizado', 'Completado')"
                else:
                    query += " AND COALESCE(TRIM(estado), 'Pendiente') = %s"
                    params.append(filtro_estado.strip())
            else:
                # En "Todos", excluir pedidos finalizados
                query += " AND (estado IS NULL OR estado NOT IN ('Finalizado', 'Completado'))"
            
            if filtro_prioridad != "Todas":
                query += " AND TRIM(prioridad) = %s"
                params.append(filtro_prioridad.strip())
            
            # Orden optimizado
            query += " ORDER BY CASE TRIM(prioridad) WHEN 'Alta' THEN 1 WHEN 'Media' THEN 2 WHEN 'Baja' THEN 3 ELSE 4 END, fecha_creacion DESC"
            
            print(f"🔍 DEBUG: Ejecutando consulta con filtros: {params}")
            
            cur = conn.cursor()
            cur.execute(query, params)
            
            results = cur.fetchall()
            print(f"🔍 DEBUG: Obtenidos {len(results)} registros de la base de datos")
            try:
                # Métricas rápidas por prioridad para validar listas grandes
                pri_counts = {"Alta": 0, "Media": 0, "Baja": 0, "Otras": 0}
                for r in results:
                    p = (r[6] or "").strip()
                    if p in pri_counts:
                        pri_counts[p] += 1
                    else:
                        pri_counts["Otras"] += 1
                print(f"📊 DEBUG: Prioridades -> Alta: {pri_counts['Alta']}, Media: {pri_counts['Media']}, Baja: {pri_counts['Baja']}, Otras: {pri_counts['Otras']}")
            except Exception:
                pass
            
            # Optimizar: obtener bases faltantes para códigos en un solo query
            codigos_sin_base = list({row[3] for row in results if (not row[12]) and row[3]})
            bases_dict = {}
            try:
                if codigos_sin_base:
                    cur.execute(
                        "SELECT codigo, base FROM ProductSW WHERE activo = TRUE AND codigo = ANY(%s)",
                        (codigos_sin_base,)
                    )
                    for codigo, base in cur.fetchall():
                        bases_dict[codigo] = base or ""
            except Exception as e:
                print(f"⚠️ No se pudo obtener bases en lote: {e}")
                bases_dict = {}

            # Precalcular campos costosos (presentación y código base) en hilo de fondo
            try:
                precalc_cb = []  # alineado por índice de results
                for row in results:
                    codigo = row[3]
                    producto = row[4]
                    terminacion = row[5]
                    cantidad = row[7]
                    base_producto = row[12] or (bases_dict.get(codigo, "") if codigo else "")
                    presentacion = row[14] or (deducir_presentacion_desde_cantidad(cantidad) if cantidad else "")
                    codigo_base_calc = self.obtener_codigo_base_desde_db(base_producto, producto, terminacion, presentacion)
                    precalc_cb.append((codigo_base_calc, base_producto, presentacion))
            except Exception as e:
                print(f"⚠️ No se pudo precalcular código base en lote: {e}")
                precalc_cb = None
            
            cur.close()
            conn.close()
            
            # Procesar datos y actualizar UI en el hilo principal (pasando precálculos)
            self.root.after(0, self._actualizar_ui_con_datos, results, start_time, sucursal_usuario, bases_dict, precalc_cb)
            
        except Exception as e:
            print(f"❌ Error al cargar datos: {e}")
            self.root.after(0, self._mostrar_error_carga, str(e))
        finally:
            self.cargando_datos = False
    
    def _actualizar_ui_con_datos(self, results, start_time, sucursal_usuario, bases_precalculadas=None, precalc_cb=None):
        """Actualiza la UI con los datos cargados (ejecutado en hilo principal)"""
        try:
            # Preservar scroll y selección
            y_top = self.tree.yview()
            sel = set(self.tree.selection())

            # Índices actuales y deseados
            actuales = set(self.tree.get_children(""))
            deseados = set()

            # Procesar resultados de forma más eficiente (actualización incremental)
            for i, row in enumerate(results):
                prioridad = row[6]
                estado = row[8]
                iid = str(row[1]) if row[1] is not None else None  # id_orden_profesional como iid estable
                if not iid:
                    continue
                deseados.add(iid)

                if estado == 'En Proceso':
                    tag = 'proceso'
                elif estado in ['Finalizado', 'Completado']:
                    tag = 'finalizado'
                elif prioridad == 'Alta':
                    tag = 'alta'
                elif prioridad == 'Media':
                    tag = 'media'
                else:
                    tag = 'baja'

                tiempo_estimado_valor = row[9] or 0
                tiempo_estimado = f"{tiempo_estimado_valor}min" if tiempo_estimado_valor else "N/A"

                # Formatear fecha (no se muestra en columnas, puede usarse futuro)
                if row[11]:
                    try:
                        _ = datetime.fromisoformat(str(row[11]).replace('Z', '+00:00'))
                    except:
                        _ = None

                operador = row[10] or "No Asignado"
                base_producto = row[12] or ""
                presentacion = row[14] or ""
                if not presentacion and row[7]:
                    presentacion = deducir_presentacion_desde_cantidad(row[7])

                if not base_producto and row[3]:
                    if bases_precalculadas and row[3] in bases_precalculadas:
                        base_producto = bases_precalculadas[row[3]]
                    else:
                        base_producto = self.obtener_base_desde_codigo(row[3])

                if precalc_cb and i < len(precalc_cb) and precalc_cb[i]:
                    codigo_base = precalc_cb[i][0]
                    if not base_producto:
                        base_producto = precalc_cb[i][1]
                    if not presentacion:
                        presentacion = precalc_cb[i][2]
                else:
                    codigo_base = self.obtener_codigo_base_desde_db(base_producto, row[4], row[5], presentacion)

                prioridad_display = self._prioridad_badge(prioridad)
                codigo_base_display = (codigo_base or "").upper()

                valores = (
                    row[1],         # ID Prof.
                    row[2],         # Factura
                    row[3],         # Código
                    row[4],         # Producto
                    row[5],         # Terminación
                    codigo_base_display,    # Código Base
                    prioridad_display,      # Prioridad
                    row[7],         # Cantidad
                    estado,         # Estado
                    tiempo_estimado, # Tiempo Est.
                    operador        # Operador
                )

                if iid in actuales:
                    actuales_vals = self.tree.item(iid, 'values')
                    if tuple(actuales_vals) != tuple(valores):
                        self.tree.item(iid, values=valores, tags=(tag,))
                    else:
                        self.tree.item(iid, tags=(tag,))
                else:
                    self.tree.insert("", "end", iid=iid, values=valores, tags=(tag,))

            # Eliminar los que ya no están
            para_eliminar = list(actuales - deseados)
            if para_eliminar:
                self.tree.delete(*para_eliminar)

            # Métricas de rendimiento
            load_time = time.time() - start_time
            conteo_actual = len(self.tree.get_children())
            self.detectar_nuevos_pedidos(conteo_actual)
            hora_actual = datetime.now().strftime("%H:%M:%S")
            if hasattr(self, 'label_ultima_actualizacion'):
                self.label_ultima_actualizacion.config(
                    text=f"🕐 Última actualización: {hora_actual} ({len(results)} registros en {load_time:.2f}s)"
                )

            # Restaurar selección y scroll
            try:
                if sel:
                    for iid in sel:
                        if self.tree.exists(iid):
                            self.tree.selection_add(iid)
                if y_top:
                    self.tree.yview_moveto(y_top[0])
            except Exception:
                pass

        except Exception as e:
            print(f"❌ Error actualizando UI: {e}")

    def _prioridad_badge(self, prioridad):
        """Devuelve una representación visual ligera de la prioridad."""
        try:
            p = (prioridad or "").strip()
            if p == "Alta":
                return "🔴 Alta"
            if p == "Media":
                return "🟠 Media"
            if p == "Baja":
                return "🟢 Baja"
            return p or ""
        except Exception:
            return prioridad
    
    def _mostrar_error_carga(self, error_msg):
        """Muestra error de carga en la UI"""
        if hasattr(self, 'label_ultima_actualizacion'):
            hora_actual = datetime.now().strftime("%H:%M:%S")
            self.label_ultima_actualizacion.config(
                text=f"🕐 Error: {hora_actual}"
            )
        messagebox.showerror("Error", f"Error al cargar datos: {error_msg}")
    
    def _calcular_tiempo_restante_rapido(self, fecha_compromiso, estado, tiempo_estimado, row_id, fecha_asignacion=None):
        """Calcula tiempo restante basado en fecha de asignación y tiempo estimado real"""
        # Solo mostrar tiempo restante para pedidos en proceso
        if estado != 'En Proceso':
            return "N/A"
        
        # Si no hay fecha de asignación, no podemos calcular
        if not fecha_asignacion:
            return "N/A"
        
        try:
            # Obtener tiempo actual
            ahora = datetime.now()
            
            # Convertir fecha_asignacion a datetime
            if isinstance(fecha_asignacion, str):
                # Manejar diferentes formatos de fecha
                try:
                    fecha_inicio = datetime.fromisoformat(fecha_asignacion.replace('Z', '+00:00'))
                    if fecha_inicio.tzinfo:
                        fecha_inicio = fecha_inicio.replace(tzinfo=None)
                except:
                    fecha_inicio = datetime.strptime(fecha_asignacion[:19], '%Y-%m-%d %H:%M:%S')
            else:
                fecha_inicio = fecha_asignacion
            
            # Calcular tiempo transcurrido desde que se inició
            tiempo_transcurrido = (ahora - fecha_inicio).total_seconds() / 60  # en minutos
            
            # Usar tiempo estimado real del pedido, o un default si no existe
            if tiempo_estimado and tiempo_estimado > 0:
                tiempo_total_estimado = tiempo_estimado
            else:
                tiempo_total_estimado = 60  # Default 60 minutos
            
            # Calcular tiempo restante
            tiempo_restante_mins = max(0, tiempo_total_estimado - tiempo_transcurrido)
            
            # Si ya se pasó del tiempo estimado, mostrar tiempo extra
            if tiempo_restante_mins <= 0:
                tiempo_extra = int(tiempo_transcurrido - tiempo_total_estimado)
                if tiempo_extra > 60:
                    horas = tiempo_extra // 60
                    minutos = tiempo_extra % 60
                    return f"⚠️ +{horas}h {minutos}m"
                else:
                    return f"⚠️ +{tiempo_extra}m"
            
            # Formatear tiempo restante
            tiempo_restante_int = int(tiempo_restante_mins)
            if tiempo_restante_int > 60:
                horas = tiempo_restante_int // 60
                minutos = tiempo_restante_int % 60
                return f"⏱️ {horas}h {minutos}m"
            else:
                return f"⏱️ {tiempo_restante_int}m"
                
        except Exception as e:
            print(f"Error calculando tiempo restante: {e}")
            return "N/A"
    
    def iniciar_produccion(self):
        """Inicia la producción de un pedido seleccionado"""
        # Evitar ejecuciones simultáneas
        if self.cargando_datos:
            messagebox.showwarning("Procesando", "Espere a que termine la operación en curso")
            return
        
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Selección", "Selecciona un pedido para iniciar producción.")
            return
        
        item = self.tree.item(selection[0])
        id_profesional = item['values'][0]  # Este es el ID profesional (texto)
        estado_actual = item['values'][8]
        
        if estado_actual not in ["Pendiente", "En Espera"]:
            messagebox.showwarning("Estado", f"El pedido ya está en estado: {estado_actual}")
            return
        
        # Si pertenece a una lista (misma factura con más de un elemento pendiente/en espera), aplicar a toda la lista
        factura_sel = self._obtener_factura_seleccionada()
        try:
            if factura_sel:
                cantidad_en_lista = self._contar_items_factura(factura_sel, estados=["Pendiente", "En Espera"])
                if cantidad_en_lista and cantidad_en_lista > 1:
                    # Reutiliza el flujo de lista (pide operador una vez)
                    self.iniciar_lista_completa()
                    return
        except Exception:
            pass

        # Caso individual: solicitar operador y procesar solo el seleccionado
        operador = self.seleccionar_operador()
        if not operador or not str(operador).strip():
            return
        
        # Obtener la factura del pedido seleccionado
        factura_pedido = self._obtener_factura_seleccionada()
        if factura_pedido:
            # Iniciar bloqueo de 4 minutos desde el inicio del proceso para esta factura
            self._iniciar_bloqueo_proceso(factura_pedido)
        
        threading.Thread(target=self._iniciar_produccion_async, args=(id_profesional, operador), daemon=True).start()

    def _resolver_operador_por_defecto(self):
        """Obtiene el operador por defecto: nombre completo del usuario o su username."""
        try:
            nombre = None
            if hasattr(self, 'usuario_info') and isinstance(self.usuario_info, dict):
                nombre = self.usuario_info.get('nombre_completo')
            if nombre and isinstance(nombre, str) and nombre.strip():
                return nombre.strip()
            return getattr(self, 'usuario_username', 'operador')
        except Exception:
            return getattr(self, 'usuario_username', 'operador')

    def _mostrar_menu_contextual(self, event):
        """Selecciona la fila bajo el cursor y muestra el menú contextual"""
        try:
            item_id = self.tree.identify_row(event.y)
            if item_id:
                self.tree.selection_set(item_id)
                self.tree.focus(item_id)
            # Actualizar los estados habilitado/deshabilitado según estado del pedido
            self._actualizar_menu_contextual_estado()
            # Mostrar menú donde hizo clic el usuario
            self.menu_contextual.tk_popup(event.x_root, event.y_root)
        finally:
            try:
                self.menu_contextual.grab_release()
            except:
                pass

    def _configurar_atajos(self):
        """Configura atajos de teclado para acciones rápidas"""
        try:
            self.tree.bind('<Control-Return>', self._atajo_iniciar)
            self.tree.bind('<Control-f>', self._atajo_finalizar)
            self.tree.bind('<Control-F>', self._atajo_finalizar)
            self.tree.bind('<Control-p>', self._atajo_imprimir)
            self.tree.bind('<Control-P>', self._atajo_imprimir)
            self.tree.bind('<Delete>', self._atajo_cancelar)
            self.tree.bind('<Control-l>', lambda e: self.mostrar_formula())
            self.tree.bind('<Control-L>', lambda e: self.mostrar_formula())
        except Exception as e:
            print(f"⚠️ No se pudieron configurar atajos: {e}")

    def _atajo_iniciar(self, event=None):
        self.iniciar_produccion()

    def _atajo_finalizar(self, event=None):
        self.finalizar_pedido()

    def _atajo_imprimir(self, event=None):
        self.imprimir_etiqueta()

    def _atajo_cancelar(self, event=None):
        self.cancelar_pedido()

    def _actualizar_menu_contextual_estado(self):
        """Habilita o deshabilita opciones del menú según el estado del pedido"""
        try:
            selection = self.tree.selection()
            if not selection:
                # Sin selección: deshabilitar todas las entradas
                total = self.menu_contextual.index("end") or 0
                for idx in range(total + 1):
                    try:
                        self.menu_contextual.entryconfig(idx, state=tk.DISABLED)
                    except Exception:
                        pass
                return

            vals = self.tree.item(selection[0], 'values')
            estado = vals[8] if len(vals) > 8 else ""

            # Estados por defecto
            iniciar_state = tk.NORMAL
            finalizar_state = tk.NORMAL
            cancelar_state = tk.NORMAL
            imprimir_state = tk.DISABLED
            formula_state = tk.NORMAL

            if estado in ("Pendiente", "En Espera"):
                iniciar_state = tk.NORMAL
                finalizar_state = tk.DISABLED
                cancelar_state = tk.NORMAL
                imprimir_state = tk.DISABLED
            elif estado == "En Proceso":
                iniciar_state = tk.DISABLED
                finalizar_state = tk.NORMAL
                cancelar_state = tk.NORMAL
                imprimir_state = tk.NORMAL
            elif estado in ("Finalizado", "Completado"):
                iniciar_state = tk.DISABLED
                finalizar_state = tk.DISABLED
                cancelar_state = tk.DISABLED
                imprimir_state = tk.NORMAL  # Habilitar impresión para finalizados
            else:  # Cancelado u otros
                iniciar_state = tk.DISABLED
                finalizar_state = tk.DISABLED
                cancelar_state = tk.DISABLED
                imprimir_state = tk.DISABLED

            # Aplicar
            self.menu_contextual.entryconfig(0, state=iniciar_state)
            self.menu_contextual.entryconfig(1, state=finalizar_state)
            self.menu_contextual.entryconfig(2, state=imprimir_state)
            # Fórmula está en índice 3
            try:
                self.menu_contextual.entryconfig(3, state=formula_state)
            except Exception:
                pass
            # ÍNDICE DE CANCELAR: último
            try:
                cancel_idx = self.menu_contextual.index("end")
                if cancel_idx is not None:
                    self.menu_contextual.entryconfig(cancel_idx, state=cancelar_state)
            except Exception:
                pass
        except Exception as e:
            print(f"⚠️ No se pudo actualizar el menú contextual: {e}")

    def _mapear_tipo_presentacion(self, presentacion):
        """Mapea la presentación a tipo de fórmula. Solo para pintura 'Excello Premium'.
        Acepta: cuarto, galón/galon, cubeta. Otras (1/2, 1/8) no aplican.
        """
        try:
            pr = (presentacion or '').strip().lower()
            if pr in ('cuarto', 'qt'):
                return 'cuarto'
            if pr in ('galón', 'galon', 'galon.', 'galón.'):
                return 'galon'
            if pr == 'cubeta':
                return 'cubeta'
        except Exception:
            pass
        return None

    def _mapear_presentacion_tinte(self, presentacion):
        """Mapea presentación de UI a claves de presentacion_tintes."""
        pr = (presentacion or '').strip().lower()
        mapa = {
            '1/8': '1/8',
            'octavo': '1/8',
            'cuarto': 'QT',
            'qt': 'QT',
            'medio galón': '1/2',
            'medio galon': '1/2',
            '1/2': '1/2',
            'galón': 'GALON',
            'galon': 'GALON',
        }
        return mapa.get(pr)

    def _obtener_datos_por_pintura(self, pintura_id):
        """Obtiene filas de fórmula para una pintura (colorante, tipo, oz, 32s, 64s, 128s)."""
        try:
            conn = self.conectar_db()
            if not conn:
                return []
            with conn.cursor() as cur:
                cur.execute(
                    """
                    SELECT 
                        c.nombre AS colorante,
                        pr.tipo,
                        pr.oz,
                        pr._32s,
                        pr._64s,
                        pr._128s
                    FROM presentacion pr
                    JOIN pintura p ON pr.id_pintura = p.id
                    JOIN colorante c ON pr.id_colorante = c.id
                    WHERE p.id = %s;
                    """,
                    (pintura_id,)
                )
                rows = cur.fetchall()
            conn.close()
            return rows
        except Exception:
            try:
                conn.close()
            except Exception:
                pass
            return []

    def _obtener_datos_por_tinte(self, tinte_id):
        """Obtiene fórmula de tintes (colorante, tipo, cantidad)."""
        try:
            conn = self.conectar_db()
            if not conn:
                return []
            with conn.cursor() as cur:
                cur.execute(
                    """
                    SELECT 
                        c.nombre AS colorante,
                        p.tipo,
                        p.cantidad
                    FROM presentacion_tintes p
                    JOIN tintes t ON p.id_tinte = t.id
                    JOIN colorantes_tintes c ON p.id_colorante_tinte = c.codigo
                    WHERE t.id = %s;
                    """,
                    (tinte_id,)
                )
                rows = cur.fetchall()
            conn.close()
            return rows
        except Exception:
            try:
                conn.close()
            except Exception:
                pass
            return []

    def _crear_tabla_formula(self, parent, filas, titulo_cols):
        """Crea un Treeview tabla con las filas y títulos de columnas dados."""
        frame = ttk.Frame(parent)
        frame.pack(fill="both", expand=True)
        tree = ttk.Treeview(frame, columns=titulo_cols, show="headings", height=12)
        for col in titulo_cols:
            tree.heading(col, text=col)
            ancho = 120 if col.lower() == 'colorante' else 70
            tree.column(col, width=ancho, anchor='center')
        vs = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vs.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)
        return tree

    def mostrar_formula(self):
        """Muestra una ventana con la fórmula correspondiente a la presentación seleccionada."""
        try:
            selection = self.tree.selection()
            if not selection:
                messagebox.showwarning("Selección", "Selecciona un pedido para ver su fórmula.")
                return
            iid = selection[0]
            vals = self.tree.item(iid, 'values')
            codigo = vals[2] if len(vals) > 2 else None
            producto = (vals[3] if len(vals) > 3 else '') or ''

            presentacion = ''
            cantidad = ''
            try:
                conn = self.conectar_db()
                if conn:
                    sucursal_usuario = obtener_sucursal_usuario(self.usuario_username)
                    tabla_sucursal = obtener_tabla_sucursal(sucursal_usuario)
                    with conn.cursor() as cur:
                        cur.execute(
                            f"""
                            SELECT presentacion, cantidad
                            FROM {tabla_sucursal}
                            WHERE id_orden_profesional = %s
                            LIMIT 1
                            """,
                            (iid,)
                        )
                        row = cur.fetchone()
                        if row:
                            presentacion = row[0] or ''
                            cantidad = row[1] or ''
                    conn.close()
            except Exception:
                try:
                    conn.close()
                except Exception:
                    pass

            win = tk.Toplevel(self.root)
            aplicar_icono_y_titulo(win, "Fórmula")
            win.transient(self.root)
            win.grab_set()
            win.geometry("520x420")

            info = ttk.Label(win, text=f"Código: {codigo}    Producto: {producto}    Presentación: {presentacion}    Cantidad: {cantidad}",
                             font=("Segoe UI", 11))
            info.pack(padx=10, pady=(10, 6))

            producto_l = producto.strip().lower()
            # Pinturas: soportamos SOLO 'excello premium' como en LabelsApp
            if producto_l == 'excello premium':
                datos = self._obtener_datos_por_pintura(codigo)
                if not datos:
                    ttk.Label(win, text="No hay fórmula disponible para este código.").pack(pady=14)
                    return
                tipo = self._mapear_tipo_presentacion(presentacion)
                if not tipo:
                    ttk.Label(win, text="Presentación no disponible para este producto.").pack(pady=14)
                    return
                filas = []
                for colorante, t, oz, s32, s64, s128 in datos:
                    if (t or '').strip().lower() == tipo:
                        def fmt(v):
                            if v is None:
                                return ''
                            try:
                                num = float(v)
                                if math.isfinite(num):
                                    return str(int(num)) if num.is_integer() else f"{num:.2f}".rstrip('0').rstrip('.')
                            except Exception:
                                pass
                            return ''
                        filas.append((colorante, fmt(oz), fmt(s32), fmt(s64), fmt(s128)))
                tree = self._crear_tabla_formula(win, filas, ("Colorante", "oz", "32s", "64s", "128s"))
                for f in filas:
                    tree.insert('', 'end', values=f)
            elif producto_l == 'tinte al thinner':
                datos = self._obtener_datos_por_tinte(codigo)
                if not datos:
                    ttk.Label(win, text="No hay fórmula disponible para este tinte.").pack(pady=14)
                    return
                tipo_deseado = self._mapear_presentacion_tinte(presentacion)
                if not tipo_deseado:
                    ttk.Label(win, text="Presentación no soportada para este tinte.").pack(pady=14)
                    return
                filas = []
                for colorante, tipo, cantidad in datos:
                    if (tipo or '').strip().upper() == tipo_deseado:
                        try:
                            num = float(cantidad)
                            cantidad_str = str(int(num)) if num.is_integer() else f"{num:.2f}".rstrip('0').rstrip('.')
                        except Exception:
                            cantidad_str = str(cantidad)
                        filas.append((colorante, cantidad_str))
                tree = self._crear_tabla_formula(win, filas, ("Colorante", tipo_deseado))
                for f in filas:
                    tree.insert('', 'end', values=f)
            else:
                ttk.Label(win, text=f"Producto no soportado para fórmula: {producto}").pack(pady=14)

            ttk.Button(win, text="Cerrar", command=win.destroy).pack(pady=10)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo mostrar la fórmula: {e}")
    
    def _iniciar_produccion_async(self, id_profesional, usuario):
        """Proceso asíncrono para iniciar producción"""
        try:
            conn = self.conectar_db()
            if not conn:
                self.root.after(0, lambda: messagebox.showerror("Error", "No se pudo conectar a la base de datos"))
                return
            
            # Detectar sucursal del usuario y tabla correspondiente
            sucursal_usuario = obtener_sucursal_usuario(self.usuario_username)
            tabla_sucursal = obtener_tabla_sucursal(sucursal_usuario)
            
            print(f"🔍 DEBUG: Iniciando producción en tabla {tabla_sucursal} para pedido {id_profesional}")
            
            cur = conn.cursor()
            
            try:
                cur.execute("BEGIN")
                
                # Verificar que el pedido existe
                cur.execute(f"""
                    SELECT estado FROM {tabla_sucursal} 
                    WHERE id_orden_profesional = %s
                """, (id_profesional,))
                
                resultado = cur.fetchone()
                if not resultado:
                    cur.execute("ROLLBACK")
                    self.root.after(0, lambda: messagebox.showerror("Error", f"No se encontró el pedido {id_profesional}"))
                    return
                
                print(f"🔍 DEBUG: Estado actual del pedido: {resultado[0]}")
                
                # Obtener datos del pedido para imprimir etiqueta (sin operador aún)
                cur.execute(f"""
                    SELECT codigo, producto, terminacion, presentacion, cantidad, base, ubicacion
                    FROM {tabla_sucursal} 
                    WHERE id_orden_profesional = %s
                """, (id_profesional,))
                datos_pedido = cur.fetchone()
                
                # Actualizar en la tabla específica de la sucursal
                cur.execute(f"""
                    UPDATE {tabla_sucursal} 
                    SET estado = 'En Proceso', fecha_asignacion = %s, operador = %s
                    WHERE id_orden_profesional = %s
                """, (datetime.now(), usuario, id_profesional))
                
                filas_afectadas = cur.rowcount
                print(f"🔍 DEBUG: Filas afectadas por la actualización: {filas_afectadas}")
                
                if filas_afectadas == 0:
                    cur.execute("ROLLBACK")
                    self.root.after(0, lambda: messagebox.showerror("Error", f"No se pudo actualizar el pedido {id_profesional}"))
                    return
                
                cur.execute("COMMIT")
                print(f"✅ DEBUG: Cambios confirmados en la base de datos")
                
                # Imprimir etiqueta automáticamente solo si está habilitado
                print(f"🔍 DEBUG: imprimir_al_iniciar = {self.imprimir_al_iniciar}")
                print(f"🔍 DEBUG: datos_pedido = {datos_pedido}")
                if self.imprimir_al_iniciar and datos_pedido:
                    print(f"🖨️ DEBUG: Iniciando impresión automática para pedido {id_profesional}")
                    # Pasar el operador asignado explícitamente
                    self._imprimir_etiqueta_pedido(datos_pedido, sucursal_usuario, id_profesional, operador=usuario)
                else:
                    print(f"⚠️ DEBUG: No se imprime - imprimir_al_iniciar={self.imprimir_al_iniciar}, datos_pedido={bool(datos_pedido)}")
                
                # Actualizar UI en el hilo principal
                self.root.after(0, lambda: self._mostrar_exito_inicio(id_profesional, usuario))
                self.root.after(0, lambda: self.cargar_datos(forzar_recarga=True))
                
                # Iniciar verificación de estado de bloqueo
                self.root.after(0, self._verificar_y_mostrar_estado_bloqueo)
                
            except Exception as e:
                cur.execute("ROLLBACK")
                print(f"❌ Error en transacción: {e}")
                self.root.after(0, lambda m=str(e): messagebox.showerror("Error", f"Error en la transacción: {m}"))
            
            cur.close()
            conn.close()
            
        except Exception as e:
            print(f"❌ ERROR en iniciar_produccion: {e}")
            self.root.after(0, lambda m=str(e): messagebox.showerror("Error", f"Error al iniciar producción: {m}"))
    
    def _mostrar_exito_inicio(self, id_profesional, usuario):
        """Muestra mensaje de éxito en el hilo principal"""
        # Evitar popups: mostrar en barra de mensajes
        try:
            self._mostrar_mensaje_impresion(f"✅ Pedido {id_profesional} iniciado por {usuario}")
        except Exception:
            pass
    
    def _imprimir_etiqueta_pedido(self, datos_pedido, sucursal, id_profesional, operador=""):
        """Imprime la etiqueta del pedido (usada al finalizar o bajo demanda)."""
        print(f"🖨️ DEBUG: _imprimir_etiqueta_pedido llamada con operador='{operador}' para ID {id_profesional}")
        try:
            # datos_pedido puede traer 7 u 8 campos (con operador al final)
            id_factura = None
            if len(datos_pedido) >= 9:
                # Incluye operador e id_factura
                codigo, producto, terminacion, presentacion, cantidad, base, ubicacion, operador_db, id_factura = datos_pedido
                if not operador:
                    operador = operador_db or ""
            elif len(datos_pedido) == 8:
                # Incluye operador
                codigo, producto, terminacion, presentacion, cantidad, base, ubicacion, operador_db = datos_pedido
                if not operador:
                    operador = operador_db or ""
            else:
                codigo, producto, terminacion, presentacion, cantidad, base, ubicacion = datos_pedido
            
            print(f"🖨️ Imprimiendo etiqueta para: {codigo} - {producto} (ID: {id_profesional})")
            
            # Obtener descripción real del código desde ProductSW
            descripcion = self.obtener_descripcion_codigo(codigo)
            
            # Calcular código base para decidir si es personalizado (según regla)
            base_para_calc = base or self.obtener_base_desde_codigo(codigo)
            codigo_base_calc = self.obtener_codigo_base_desde_db(
                base_para_calc or "", producto or "", terminacion or "", presentacion or ""
            )

            # Resolver nombre del cliente (si es posible)
            nombre_cliente = ""
            try:
                suc = obtener_sucursal_usuario(self.usuario_username)
                tabla = obtener_tabla_sucursal(suc)
                nombre_cliente = self.obtener_nombre_cliente(tabla, id_profesional, id_factura)
            except Exception:
                nombre_cliente = ""

            # Generar ZPL
            zpl_code = generar_zpl_gestor(
                codigo=codigo,
                descripcion=descripcion,
                producto=producto or "",
                terminacion=terminacion or "",
                presentacion=presentacion or "",
                cantidad=cantidad or 1,
                base=base or "",
                ubicacion=ubicacion or "",
                sucursal=sucursal,
                id_profesional=id_profesional,
                operador=operador or "",
                codigo_base=codigo_base_calc,
                nombre_cliente=nombre_cliente
            )
            
            # Intentar imprimir
            exito = imprimir_zebra_zpl_gestor(zpl_code)
            
            if exito:
                print(f"✅ Etiqueta impresa correctamente para {codigo}")
                # Mostrar mensaje en la UI
                self.root.after(0, lambda: self._mostrar_mensaje_impresion(f"✅ Etiqueta impresa: {codigo}"))
            else:
                print(f"⚠️ No se pudo imprimir etiqueta para {codigo}")
                self.root.after(0, lambda: self._mostrar_mensaje_impresion(f"⚠️ Error impresión: {codigo}"))
                
        except Exception as e:
            print(f"❌ Error imprimiendo etiqueta: {e}")
            self.root.after(0, lambda m=str(e): self._mostrar_mensaje_impresion(f"❌ Error: {m}"))
    
    def _mostrar_mensaje_impresion(self, mensaje):
        """Muestra mensaje de estado de impresión en la UI"""
        if hasattr(self, 'label_mensaje'):
            self.label_mensaje.config(text=mensaje)
            # Limpiar mensaje después de 5 segundos
            limpiar_mensaje_despues_gestor(self.label_mensaje, 5000, "")

    def obtener_nombre_cliente(self, tabla_sucursal: str, id_profesional: str = None, id_factura: str = None) -> str:
        """Intenta obtener el nombre del cliente desde la tabla de pedidos.

        Busca columnas comunes (cliente, nombre_cliente, cliente_nombre, nombre) y usa
        id_orden_profesional o id_factura para localizar el registro. Si no existe la
        columna o no se encuentra valor, retorna cadena vacía.
        """
        try:
            conn = self.conectar_db()
            if not conn:
                return ""
            cur = conn.cursor()

            # Descubrir si hay columnas de cliente en la tabla
            posibles = ['cliente', 'nombre_cliente', 'cliente_nombre', 'nombre']
            col_cliente = None
            try:
                cur.execute(
                    """
                    SELECT column_name
                    FROM information_schema.columns
                    WHERE table_schema = 'public' AND table_name = %s AND column_name = ANY(%s)
                    """,
                    (tabla_sucursal, posibles)
                )
                cols = [r[0] for r in cur.fetchall()] if cur.rowcount is not None else []
                for c in posibles:
                    if c in cols:
                        col_cliente = c
                        break
            except Exception:
                col_cliente = None

            if not col_cliente:
                cur.close(); conn.close()
                return ""

            # Intentar por id_orden_profesional
            if id_profesional:
                try:
                    cur.execute(
                        f"SELECT {col_cliente} FROM {tabla_sucursal} WHERE id_orden_profesional = %s",
                        (id_profesional,)
                    )
                    row = cur.fetchone()
                    if row and row[0]:
                        nombre = str(row[0]).strip()
                        cur.close(); conn.close()
                        return nombre
                except Exception:
                    pass

            # Intentar por id_factura
            if id_factura:
                try:
                    cur.execute(
                        f"SELECT {col_cliente} FROM {tabla_sucursal} WHERE id_factura = %s",
                        (id_factura,)
                    )
                    row = cur.fetchone()
                    if row and row[0]:
                        nombre = str(row[0]).strip()
                        cur.close(); conn.close()
                        return nombre
                except Exception:
                    pass

            cur.close(); conn.close()
            return ""
        except Exception:
            return ""

    def imprimir_etiqueta(self):
        """Imprime la etiqueta del pedido seleccionado sin cambiar su estado"""
        try:
            selection = self.tree.selection()
            if not selection:
                messagebox.showwarning("Selección", "Selecciona un pedido para imprimir su etiqueta.")
                return

            item = self.tree.item(selection[0])
            id_profesional = item['values'][0]
            estado_actual = item['values'][8] if len(item['values']) > 8 else ''

            # Regla: solo imprimir si está en proceso o finalizado
            if estado_actual not in ['En Proceso', 'Finalizado', 'Completado']:
                messagebox.showinfo("Impresión restringida", "Solo se puede imprimir cuando el pedido está En Proceso, Finalizado o Completado.")
                return

            # Si pertenece a lista con más de un elemento en proceso, imprimir toda la lista
            factura_sel = self._obtener_factura_seleccionada()
            try:
                if factura_sel:
                    cnt_proc = self._contar_items_factura(factura_sel, estados=['En Proceso'])
                    if cnt_proc and cnt_proc > 1:
                        self.imprimir_pendientes_lista()
                        return
            except Exception:
                pass

            # Obtener datos del pedido desde la BD
            sucursal = obtener_sucursal_usuario(self.usuario_username)
            tabla = obtener_tabla_sucursal(sucursal)

            conn = self.conectar_db()
            if not conn:
                messagebox.showerror("Error", "No se pudo conectar a la base de datos")
                return
            cur = conn.cursor()
            cur.execute(f"""
                SELECT codigo, producto, terminacion, presentacion, cantidad, base, ubicacion, operador, id_factura
                FROM {tabla}
                WHERE id_orden_profesional = %s
            """, (id_profesional,))
            datos = cur.fetchone()
            cur.close(); conn.close()

            if not datos:
                messagebox.showerror("Error", f"No se encontraron datos del pedido {id_profesional}")
                return

            # Reutilizar el flujo de impresión existente (incluye operador si está disponible)
            self._imprimir_etiqueta_pedido(datos, sucursal, id_profesional)

        except Exception as e:
            print(f"❌ Error en imprimir_etiqueta: {e}")
            messagebox.showerror("Error", f"Error al imprimir etiqueta: {e}")

    # Eliminado: acción unificada por diálogo. Los handlers deciden si aplican a pedido o lista.

    # ====== Acciones por lista (misma factura) ======
    def _obtener_factura_seleccionada(self):
        try:
            sel = self.tree.selection()
            if not sel:
                return None
            vals = self.tree.item(sel[0], 'values')
            # Índice 1 = Factura (según definición de columnas)
            return vals[1] if len(vals) > 1 else None
        except Exception:
            return None

    def iniciar_lista_completa(self):
        """Inicia todos los pedidos 'Pendiente/En Espera' de la misma factura que la fila seleccionada."""
        factura = self._obtener_factura_seleccionada()
        if not factura:
            messagebox.showwarning("Lista", "No se pudo determinar la factura de la selección.")
            return
        
        # Contar pedidos pendientes en la lista
        cantidad_pendientes = self._contar_items_factura(factura, estados=["Pendiente", "En Espera"])
        
        # Confirmación para iniciar lista completa
        confirmacion = mostrar_pregunta("Confirmar Inicio de Lista", 
                                      f"¿Desea iniciar TODA la lista de producción?\n\n"
                                      f"Factura: {factura}\n"
                                      f"Pedidos a iniciar: {cantidad_pendientes}\n\n"
                                      f"Esto iniciará todos los pedidos pendientes de esta factura.")
        
        if not confirmacion:
            return  # Usuario canceló la operación
        
        # Pedir operador una vez
        operador = self.seleccionar_operador()
        if not operador or not str(operador).strip():
            return
        
        # Iniciar bloqueo de 4 minutos desde el inicio del proceso para esta factura
        self._iniciar_bloqueo_proceso(factura)
        threading.Thread(target=self._iniciar_lista_completa_async, args=(factura, operador), daemon=True).start()

    def _iniciar_lista_completa_async(self, id_factura, operador):
        try:
            conn = self.conectar_db()
            if not conn:
                self.root.after(0, lambda: messagebox.showerror("Error", "No se pudo conectar a la base de datos"))
                return
            suc = obtener_sucursal_usuario(self.usuario_username)
            tabla = obtener_tabla_sucursal(suc)
            cur = conn.cursor()
            try:
                cur.execute("BEGIN")
                cur.execute(f"""
                    UPDATE {tabla}
                    SET estado='En Proceso', fecha_asignacion=%s, operador=%s
                    WHERE id_factura=%s AND (estado IS NULL OR estado IN ('Pendiente','En Espera'))
                """, (datetime.now(), operador, id_factura))
                n = cur.rowcount
                cur.execute("COMMIT")
                
                # Si está habilitada la impresión al iniciar, imprimir todas las etiquetas de la lista
                if self.imprimir_al_iniciar:
                    print(f"🖨️ DEBUG: Imprimiendo etiquetas automáticamente para lista {id_factura}")
                    # Obtener todos los pedidos de la lista que acabamos de iniciar
                    cur.execute(f"""
                        SELECT id_orden_profesional, codigo, producto, terminacion, presentacion, cantidad, base, ubicacion
                        FROM {tabla}
                        WHERE id_factura=%s AND estado='En Proceso'
                    """, (id_factura,))
                    pedidos_lista = cur.fetchall()
                    
                    # Imprimir etiqueta para cada pedido
                    for pedido in pedidos_lista:
                        id_prof, codigo, producto, terminacion, presentacion, cantidad, base, ubicacion = pedido
                        datos_pedido = (codigo, producto, terminacion, presentacion, cantidad, base, ubicacion)
                        self._imprimir_etiqueta_pedido(datos_pedido, suc, id_prof, operador=operador)
                else:
                    print(f"⚠️ DEBUG: Impresión al iniciar deshabilitada para lista {id_factura}")
                
                self.root.after(0, lambda: self._mostrar_mensaje_impresion(f"▶️ Lista {id_factura}: {n} pedidos iniciados por {operador}"))
                self.root.after(0, lambda: self.cargar_datos(forzar_recarga=True))
                
                # Iniciar verificación de estado de bloqueo
                self.root.after(0, self._verificar_y_mostrar_estado_bloqueo)
            except Exception as e:
                cur.execute("ROLLBACK")
                print(f"❌ Error iniciar lista: {e}")
                self.root.after(0, lambda: messagebox.showerror("Error", f"Error al iniciar lista: {e}"))
            finally:
                cur.close(); conn.close()
        except Exception as e:
            print(f"❌ ERROR _iniciar_lista_completa_async: {e}")

    def imprimir_pendientes_lista(self):
        """Imprime todas las etiquetas de pedidos pendientes de la lista (misma factura)."""
        factura = self._obtener_factura_seleccionada()
        if not factura:
            messagebox.showwarning("Lista", "No se pudo determinar la factura de la selección.")
            return
        threading.Thread(target=self._imprimir_pendientes_lista_async, args=(factura,), daemon=True).start()

    def _imprimir_pendientes_lista_async(self, id_factura):
        try:
            suc = obtener_sucursal_usuario(self.usuario_username)
            tabla = obtener_tabla_sucursal(suc)
            conn = self.conectar_db()
            if not conn:
                self.root.after(0, lambda: messagebox.showerror("Error", "No se pudo conectar a la base de datos"))
                return
            cur = conn.cursor()
            cur.execute(f"""
                SELECT id_orden_profesional, codigo, producto, terminacion, presentacion, cantidad, base, ubicacion, operador, id_factura
                FROM {tabla}
                WHERE id_factura=%s AND estado = 'En Proceso'
                ORDER BY fecha_asignacion ASC, fecha_creacion ASC
            """, (id_factura,))
            filas = cur.fetchall()
            cur.close(); conn.close()

            if not filas:
                self.root.after(0, lambda: self._mostrar_mensaje_impresion(f"ℹ️ Lista {id_factura}: no hay pedidos en proceso para imprimir"))
                return

            for (id_prof, codigo, producto, terminacion, presentacion, cantidad, base, ubicacion, operador, factura) in filas:
                datos = (codigo, producto, terminacion, presentacion, cantidad, base, ubicacion, operador, factura)
                try:
                    self._imprimir_etiqueta_pedido(datos, suc, id_prof)
                except Exception as e:
                    print(f"⚠️ Error imprimiendo de lista {id_factura} ({id_prof}): {e}")

            self.root.after(0, lambda: self._mostrar_mensaje_impresion(f"🖨️ Lista {id_factura}: {len(filas)} etiquetas enviadas"))
        except Exception as e:
            print(f"❌ ERROR _imprimir_pendientes_lista_async: {e}")

    def finalizar_lista(self):
        """Finaliza todos los pedidos en proceso de la lista (misma factura)."""
        factura = self._obtener_factura_seleccionada()
        if not factura:
            messagebox.showwarning("Lista", "No se pudo determinar la factura de la selección.")
            return
        
        # Verificar bloqueo de proceso activo para esta factura específica
        if factura in self.bloqueos_por_factura:
            tiempo_restante = self._obtener_tiempo_bloqueo_restante(factura)
            if tiempo_restante > 0:
                mostrar_advertencia("Lista Bloqueada", 
                                  f"La lista {factura} está en proceso. Debe esperar {tiempo_restante // 60}m {tiempo_restante % 60}s antes de finalizar operaciones.")
                return
        
        # Contar pedidos en proceso en la lista
        cantidad_proceso = self._contar_items_factura(factura, estados=["En Proceso"])
        
        # Confirmación para finalizar lista completa
        confirmacion = mostrar_pregunta("Confirmar Finalización de Lista", 
                                      f"¿Desea finalizar TODA la lista de producción?\n\n"
                                      f"Factura: {factura}\n"
                                      f"Pedidos a finalizar: {cantidad_proceso}\n\n"
                                      f"Esto marcará como completados todos los pedidos en proceso de esta factura.")
        
        if not confirmacion:
            return  # Usuario canceló la operación
        
        threading.Thread(target=self._finalizar_lista_async, args=(factura,), daemon=True).start()

    def _finalizar_lista_async(self, id_factura):
        try:
            suc = obtener_sucursal_usuario(self.usuario_username)
            tabla = obtener_tabla_sucursal(suc)
            conn = self.conectar_db()
            if not conn:
                self.root.after(0, lambda: messagebox.showerror("Error", "No se pudo conectar a la base de datos"))
                return
            cur = conn.cursor()
            try:
                # Obtener datos de los que están en proceso (para impresión posterior)
                cur.execute(f"""
                    SELECT id_orden_profesional, codigo, producto, terminacion, presentacion, cantidad, base, ubicacion, operador, id_factura
                    FROM {tabla}
                    WHERE id_factura=%s AND estado='En Proceso'
                    ORDER BY fecha_asignacion ASC
                """, (id_factura,))
                en_proceso = cur.fetchall()

                cur.execute("BEGIN")
                cur.execute(f"""
                    UPDATE {tabla}
                    SET estado='Finalizado', fecha_completado=%s
                    WHERE id_factura=%s AND estado='En Proceso'
                """, (datetime.now(), id_factura))
                n = cur.rowcount
                cur.execute("COMMIT")

                # Imprimir al finalizar si está habilitado
                if n > 0 and getattr(self, 'imprimir_al_finalizar', False):
                    for (id_prof, codigo, producto, terminacion, presentacion, cantidad, base, ubicacion, operador, factura) in en_proceso:
                        datos = (codigo, producto, terminacion, presentacion, cantidad, base, ubicacion, operador, factura)
                        try:
                            self._imprimir_etiqueta_pedido(datos, suc, id_prof)
                        except Exception as e:
                            print(f"⚠️ Error imprimiendo al finalizar lista {id_factura} ({id_prof}): {e}")

                self.root.after(0, lambda: self._mostrar_mensaje_impresion(f"✅ Lista {id_factura}: {n} pedidos finalizados"))
                self.root.after(0, lambda: self.cargar_datos(forzar_recarga=True))
                
                # Iniciar verificación de estado de bloqueo
                self.root.after(0, self._verificar_y_mostrar_estado_bloqueo)
            except Exception as e:
                cur.execute("ROLLBACK")
                print(f"❌ Error finalizar lista: {e}")
                self.root.after(0, lambda: messagebox.showerror("Error", f"Error al finalizar lista: {e}"))
            finally:
                cur.close(); conn.close()
        except Exception as e:
            print(f"❌ ERROR _finalizar_lista_async: {e}")

    def finalizar_pedido(self):
        """Finaliza un pedido en proceso"""
        # Obtener la factura del pedido seleccionado para verificar su bloqueo específico
        factura_pedido = self._obtener_factura_seleccionada()
        if factura_pedido and factura_pedido in self.bloqueos_por_factura:
            tiempo_restante = self._obtener_tiempo_bloqueo_restante(factura_pedido)
            if tiempo_restante > 0:
                mostrar_advertencia("Lista Bloqueada", 
                                  f"La lista {factura_pedido} está en proceso. Debe esperar {tiempo_restante // 60}m {tiempo_restante % 60}s antes de finalizar operaciones.")
                return
        
        # Evitar ejecuciones simultáneas
        if self.cargando_datos:
            messagebox.showwarning("Procesando", "Espere a que termine la operación en curso")
            return
        
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Selección", "Selecciona un pedido para finalizar.")
            return
        
        item = self.tree.item(selection[0])
        id_profesional = item['values'][0]  # Este es el ID profesional (texto)
        estado_actual = item['values'][8]
        
        if estado_actual != "En Proceso":
            messagebox.showwarning("Estado", f"Solo se pueden finalizar pedidos en proceso. Estado actual: {estado_actual}")
            return

        # Si pertenece a lista con más de un elemento en proceso, finalizar toda la lista
        factura_sel = self._obtener_factura_seleccionada()
        try:
            if factura_sel:
                cnt_proc = self._contar_items_factura(factura_sel, estados=['En Proceso'])
                if cnt_proc and cnt_proc > 1:
                    self.finalizar_lista()
                    return
        except Exception:
            pass
        
        # Usar threading para evitar congelamiento
        threading.Thread(
            target=self._finalizar_pedido_async,
            args=(id_profesional,),
            daemon=True
        ).start()
    
    def _finalizar_pedido_async(self, id_profesional):
        """Proceso asíncrono para finalizar pedido"""
        try:
            # Obtener la tabla correspondiente a la sucursal del usuario
            sucursal = obtener_sucursal_usuario(self.usuario_username)
            tabla_sucursal = obtener_tabla_sucursal(sucursal)
            
            conn = self.conectar_db()
            if not conn:
                self.root.after(0, lambda: messagebox.showerror("Error", "No se pudo conectar a la base de datos"))
                return
            
            cur = conn.cursor()
            
            try:
                cur.execute("BEGIN")
                
                cur.execute(f"""
                    UPDATE {tabla_sucursal} 
                    SET estado = 'Finalizado', fecha_completado = %s
                    WHERE id_orden_profesional = %s AND estado = 'En Proceso'
                """, (datetime.now(), id_profesional))
                
                filas_afectadas = cur.rowcount
                
                if filas_afectadas == 0:
                    cur.execute("ROLLBACK")
                    self.root.after(0, lambda: messagebox.showerror("Error", f"No se pudo finalizar el pedido {id_profesional}"))
                    return
                
                cur.execute("COMMIT")

                # Imprimir etiqueta automáticamente al finalizar si está habilitado
                try:
                    if self.imprimir_al_finalizar:
                        # Obtener datos para impresión (incluyendo operador e id_factura)
                        conn2 = self.conectar_db()
                        if conn2:
                            cur2 = conn2.cursor()
                            cur2.execute(f"""
                                SELECT codigo, producto, terminacion, presentacion, cantidad, base, ubicacion, operador, id_factura
                                FROM {tabla_sucursal}
                                WHERE id_orden_profesional = %s
                            """, (id_profesional,))
                            datos_imp = cur2.fetchone()
                            cur2.close(); conn2.close()
                            if datos_imp:
                                self._imprimir_etiqueta_pedido(datos_imp, sucursal, id_profesional)
                except Exception as e_imp:
                    print(f"⚠️ No se pudo imprimir al finalizar: {e_imp}")
                
                # Reproducir sonido de pedido completado en el hilo principal
                self.root.after(0, self.notificar_pedido_completado)
                
                # Actualizar UI en el hilo principal
                self.root.after(0, lambda: self._mostrar_exito_finalizacion(id_profesional))
                self.root.after(0, lambda: self.cargar_datos(forzar_recarga=True))
                
                # Iniciar verificación de estado de bloqueo
                self.root.after(0, self._verificar_y_mostrar_estado_bloqueo)
                
            except Exception as e:
                cur.execute("ROLLBACK")
                print(f"❌ Error en transacción de finalización: {e}")
                self.root.after(0, lambda m=str(e): messagebox.showerror("Error", f"Error en la transacción: {m}"))
            
            cur.close()
            conn.close()
            
        except Exception as e:
            print(f"❌ ERROR en finalizar_pedido: {e}")
            self.root.after(0, lambda m=str(e): messagebox.showerror("Error", f"Error al finalizar pedido: {m}"))
    
    def _mostrar_exito_finalizacion(self, id_profesional):
        """Muestra mensaje de éxito de finalización en el hilo principal"""
        try:
            self._mostrar_mensaje_impresion(f"✅ Pedido {id_profesional} finalizado")
        except Exception:
            pass
    
    def _iniciar_bloqueo_proceso(self, id_factura):
        """Inicia el bloqueo de proceso por 4 minutos para una factura específica"""
        try:
            # Cancelar timer anterior para esta factura si existe
            if id_factura in self.bloqueos_por_factura:
                timer_anterior = self.bloqueos_por_factura[id_factura].get('timer')
                if timer_anterior:
                    try:
                        self.root.after_cancel(timer_anterior)
                    except Exception:
                        pass
            
            # Inicializar bloqueo para esta factura
            tiempo_inicio = time.time()
            timer_id = self.root.after(
                self.duracion_bloqueo * 1000,  # convertir a milliseconds
                lambda f=id_factura: self._liberar_bloqueo_proceso(f)
            )
            
            self.bloqueos_por_factura[id_factura] = {
                'bloqueado': True,
                'inicio': tiempo_inicio,
                'timer': timer_id
            }
            
            # Mostrar mensaje informativo
            self._mostrar_mensaje_impresion(f"🔒 Lista {id_factura} - Bloqueada por {self.duracion_bloqueo // 60} minutos para completar operación")
            
            print(f"🔒 Bloqueo activado para factura {id_factura} por {self.duracion_bloqueo} segundos")
            
        except Exception as e:
            print(f"⚠️ Error iniciando bloqueo para factura {id_factura}: {e}")
    
    def _liberar_bloqueo_proceso(self, id_factura):
        """Libera el bloqueo de proceso para una factura específica"""
        try:
            if id_factura in self.bloqueos_por_factura:
                del self.bloqueos_por_factura[id_factura]
                
                # Mostrar mensaje de liberación
                self._mostrar_mensaje_impresion(f"🔓 Lista {id_factura} desbloqueada - Puede finalizar operaciones")
                
                print(f"🔓 Bloqueo liberado para factura {id_factura}")
            
        except Exception as e:
            print(f"⚠️ Error liberando bloqueo para factura {id_factura}: {e}")
    
    def _obtener_tiempo_bloqueo_restante(self, id_factura):
        """Obtiene el tiempo restante del bloqueo en segundos para una factura específica"""
        try:
            if id_factura not in self.bloqueos_por_factura:
                return 0
            
            bloqueo_info = self.bloqueos_por_factura[id_factura]
            if not bloqueo_info['bloqueado'] or not bloqueo_info['inicio']:
                return 0
            
            tiempo_transcurrido = time.time() - bloqueo_info['inicio']
            tiempo_restante = max(0, self.duracion_bloqueo - tiempo_transcurrido)
            
            # Si el tiempo se agotó, liberar automáticamente
            if tiempo_restante <= 0:
                self._liberar_bloqueo_proceso(id_factura)
                return 0
            
            return int(tiempo_restante)
            
        except Exception:
            return 0
    
    def _verificar_y_mostrar_estado_bloqueo(self, id_factura=None):
        """Verifica el estado del bloqueo para una factura específica y actualiza la UI si es necesario"""
        try:
            # Si no se especifica factura, verificar todas las facturas bloqueadas
            if id_factura is None:
                facturas_bloqueadas = list(self.bloqueos_por_factura.keys())
                if facturas_bloqueadas:
                    # Mostrar estado de la primera factura bloqueada para el mensaje general
                    self._verificar_y_mostrar_estado_bloqueo(facturas_bloqueadas[0])
                return
            
            if id_factura in self.bloqueos_por_factura:
                tiempo_restante = self._obtener_tiempo_bloqueo_restante(id_factura)
                if tiempo_restante > 0:
                    minutos = tiempo_restante // 60
                    segundos = tiempo_restante % 60
                    
                    # Actualizar mensaje en la interfaz
                    if hasattr(self, 'label_mensaje'):
                        self.label_mensaje.config(text=f"🔒 Lista {id_factura} bloqueada - Tiempo restante: {minutos}m {segundos}s")
                        
                    # Reprogramar verificación cada 10 segundos
                    self.root.after(10000, lambda f=id_factura: self._verificar_y_mostrar_estado_bloqueo(f))
                else:
                    # El bloqueo expiró, limpiar estado
                    self._liberar_bloqueo_proceso(id_factura)
        except Exception as e:
            print(f"⚠️ Error verificando estado de bloqueo para factura {id_factura}: {e}")

    def _contar_items_factura(self, id_factura, estados=None):
        """Cuenta pedidos por factura en la tabla de la sucursal.
        - estados: lista de estados para filtrar (exactos, con TRIM). Si es None, cuenta todos.
        Devuelve 0 ante error.
        """
        try:
            if not id_factura:
                return 0
            sucursal = obtener_sucursal_usuario(self.usuario_username)
            tabla = obtener_tabla_sucursal(sucursal)
            conn = self.conectar_db()
            if not conn:
                return 0
            cur = conn.cursor()
            if estados:
                placeholders = ",".join(["%s"] * len(estados))
                cur.execute(
                    f"""
                    SELECT COUNT(*) FROM {tabla}
                    WHERE id_factura = %s AND TRIM(COALESCE(estado,'')) IN ({placeholders})
                    """,
                    tuple([id_factura] + estados),
                )
            else:
                cur.execute(
                    f"SELECT COUNT(*) FROM {tabla} WHERE id_factura = %s",
                    (id_factura,),
                )
            n = cur.fetchone()[0]
            cur.close(); conn.close()
            try:
                return int(n)
            except Exception:
                return 0
        except Exception:
            return 0
    
    def cancelar_pedido(self):
        """Cancela un pedido"""
        # Evitar ejecuciones simultáneas
        if self.cargando_datos:
            messagebox.showwarning("Procesando", "Espere a que termine la operación en curso")
            return
        
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Selección", "Selecciona un pedido para cancelar.")
            return
        
        item = self.tree.item(selection[0])
        id_profesional = item['values'][0]  # Este es el ID profesional (texto)
        estado_actual = item['values'][8]  # Estado actual
        codigo = item['values'][2] if len(item['values']) > 2 else 'N/A'
        
        # Validar que se puede cancelar
        if estado_actual in ['Finalizado', 'Cancelado']:
            messagebox.showwarning("Estado", f"No se puede cancelar un pedido que ya está {estado_actual.lower()}")
            return
        
        # CONFIRMACIÓN OBLIGATORIA PARA CANCELAR
        confirmacion = mostrar_pregunta("Confirmar Cancelación", 
                                      f"¿Está seguro de que desea CANCELAR el pedido?\n\n"
                                      f"ID: {id_profesional}\n"
                                      f"Código: {codigo}\n"
                                      f"Estado actual: {estado_actual}\n\n"
                                      f"Esta acción eliminará permanentemente el pedido.")
        
        if not confirmacion:
            return  # Usuario canceló la operación
        
        # Usar threading para evitar congelamiento
        threading.Thread(
            target=self._cancelar_pedido_async,
            args=(id_profesional,),
            daemon=True
        ).start()
    
    def _cancelar_pedido_async(self, id_profesional):
        """Proceso asíncrono para cancelar pedido"""
        try:
            # Detectar sucursal del usuario y tabla correspondiente
            sucursal_usuario = obtener_sucursal_usuario(self.usuario_username)
            tabla_sucursal = obtener_tabla_sucursal(sucursal_usuario)
            
            conn = self.conectar_db()
            if not conn:
                self.root.after(0, lambda: messagebox.showerror("Error", "No se pudo conectar a la base de datos"))
                return
            
            print(f"🔍 DEBUG: Cancelando pedido en tabla {tabla_sucursal} para pedido {id_profesional}")
            
            cur = conn.cursor()
            
            try:
                cur.execute("BEGIN")
                
                # Opcional: archivar en histórico antes de borrar
                try:
                    if getattr(self, 'var_archivar_cancelados', None) and self.var_archivar_cancelados.get():
                        # Obtener datos del pedido para histórico
                        cur.execute(f"""
                            SELECT id_orden_profesional, codigo, producto, terminacion, id_factura,
                                   prioridad, cantidad, tiempo_estimado, base, ubicacion, sucursal,
                                   COALESCE(operador, %s) as operador, fecha_creacion, fecha_asignacion
                            FROM {tabla_sucursal}
                            WHERE id_orden_profesional = %s
                            FOR UPDATE
                        """, (self.usuario_username, id_profesional))
                        fila = cur.fetchone()
                        if fila:
                            (
                                v_id_prof, v_codigo, v_producto, v_terminacion, v_id_factura,
                                v_prioridad, v_cantidad, v_tiempo_est, v_base, v_ubicacion, v_sucursal,
                                v_operador, v_fecha_creacion, v_fecha_asignacion
                            ) = fila

                            fecha_arch = datetime.now()
                            semana = fecha_arch.isocalendar()[1]
                            mes = fecha_arch.month
                            ano = fecha_arch.year

                            # Calcular tiempo total procesamiento si hay fecha_creacion
                            tiempo_total = None
                            try:
                                if v_fecha_creacion:
                                    tiempo_total = (fecha_arch - v_fecha_creacion).total_seconds() / 60.0
                            except Exception:
                                tiempo_total = None

                            notas = f"Cancelado por {self.usuario_username}"

                            cur.execute(
                                """
                                INSERT INTO pedidos_historicos 
                                (id_orden_profesional, codigo, producto, terminacion, id_factura,
                                 prioridad, cantidad, tiempo_estimado, base, ubicacion, sucursal,
                                 operador, estado, fecha_creacion, fecha_asignacion, fecha_completado,
                                 notas, fecha_archivado, semana_archivado, mes_archivado, ano_archivado,
                                 tiempo_total_procesamiento)
                                VALUES (%s, %s, %s, %s, %s,
                                        %s, %s, %s, %s, %s, %s,
                                        %s, %s, %s, %s, %s,
                                        %s, %s, %s, %s, %s,
                                        %s)
                                ON CONFLICT (id_orden_profesional, fecha_archivado) DO NOTHING
                                """,
                                (
                                    v_id_prof, v_codigo, v_producto, v_terminacion, v_id_factura,
                                    v_prioridad, v_cantidad, v_tiempo_est, v_base, v_ubicacion, v_sucursal,
                                    v_operador, 'Cancelado', v_fecha_creacion, v_fecha_asignacion, fecha_arch,
                                    notas, fecha_arch, semana, mes, ano,
                                    tiempo_total
                                )
                            )
                except Exception as e_hist:
                    # No bloquear la cancelación por errores de histórico
                    print(f"⚠️ No se pudo archivar el cancelado en histórico: {e_hist}")

                # Eliminar directamente el pedido cancelado para que no aparezca en ningún filtro
                cur.execute(f"""
                    DELETE FROM {tabla_sucursal}
                    WHERE id_orden_profesional = %s
                """, (id_profesional,))
                
                filas_afectadas = cur.rowcount
                print(f"🔍 DEBUG: Filas afectadas al cancelar: {filas_afectadas}")
                
                if filas_afectadas == 0:
                    cur.execute("ROLLBACK")
                    self.root.after(0, lambda: messagebox.showerror("Error", f"No se encontró el pedido {id_profesional} para cancelar."))
                    return
                
                cur.execute("COMMIT")
                print(f"✅ DEBUG: Pedido cancelado exitosamente")
                
                # Actualizar UI en el hilo principal
                self.root.after(0, lambda: self._mostrar_exito_cancelacion(id_profesional))
                # Eliminar inmediatamente de la tabla visual
                self.root.after(0, lambda: self._eliminar_item_de_tree(id_profesional))
                self.root.after(0, lambda: self.cargar_datos(forzar_recarga=True))
                
            except Exception as e:
                cur.execute("ROLLBACK")
                print(f"❌ Error en transacción de cancelación: {e}")
                self.root.after(0, lambda: messagebox.showerror("Error", f"Error en la transacción: {e}"))
            
            cur.close()
            conn.close()
            
        except Exception as e:
            print(f"❌ ERROR en cancelar_pedido: {e}")
            self.root.after(0, lambda: messagebox.showerror("Error", f"Error al cancelar pedido: {e}"))
    
    def _mostrar_exito_cancelacion(self, id_profesional):
        """Muestra mensaje de éxito de cancelación en el hilo principal"""
        try:
            if getattr(self, 'var_archivar_cancelados', None) and self.var_archivar_cancelados.get():
                self._mostrar_mensaje_impresion(f"🗂️ Pedido {id_profesional} cancelado y archivado")
            else:
                self._mostrar_mensaje_impresion(f"🗑️ Pedido {id_profesional} cancelado y eliminado")
        except Exception:
            pass

    def _toggle_archivar_cancelados(self):
        """Sincroniza el flag de archivado de cancelados con el checkbox."""
        try:
            self.archivar_cancelados = bool(self.var_archivar_cancelados.get())
        except Exception:
            self.archivar_cancelados = False

    def _eliminar_item_de_tree(self, id_profesional):
        """Elimina de la vista el pedido con el ID profesional indicado"""
        try:
            for iid in self.tree.get_children():
                vals = self.tree.item(iid, 'values')
                if vals and str(vals[0]) == str(id_profesional):
                    self.tree.delete(iid)
                    break
        except Exception as e:
            print(f"⚠️ No se pudo eliminar de la vista: {e}")
    
    def verificar_archivado_automatico(self):
        """Verifica si debe ejecutar el archivado automático a las 6pm"""
        hora_actual = datetime.now().time()
        fecha_actual = datetime.now().date()
        
        # Solo ejecutar después de las 6pm (18:00) y una vez por día
        if hora_actual.hour >= 18 and not self.archivado_ejecutado_hoy:
            print("🕕 Hora de archivado automático - Ejecutando limpieza...")
            self.ejecutar_archivado_automatico()
            self.archivado_ejecutado_hoy = True
            
        # Resetear flag al cambio de día
        if self.ultima_verificacion_archivado and self.ultima_verificacion_archivado.date() != fecha_actual:
            self.archivado_ejecutado_hoy = False
            
        self.ultima_verificacion_archivado = datetime.now()
    
    def ejecutar_archivado_automatico(self):
        """Ejecuta el archivado automático en threading"""
        threading.Thread(
            target=self._archivar_solicitudes_async,
            daemon=True
        ).start()
    
    def _archivar_solicitudes_async(self):
        """Archiva todas las solicitudes de forma asíncrona con histórico para reportes"""
        try:
            print("📁 Iniciando archivado automático con histórico...")
            
            conn = self.conectar_db()
            if not conn:
                print("❌ No se pudo conectar para archivado")
                return
            
            cur = conn.cursor()
            
            # Detectar tabla de sucursal actual
            sucursal_usuario = obtener_sucursal_usuario(self.usuario_username)
            tabla_sucursal = obtener_tabla_sucursal(sucursal_usuario)
            
            # Contar registros antes del archivado
            cur.execute(f"SELECT COUNT(*) FROM {tabla_sucursal}")
            count_antes = cur.fetchone()[0]
            
            if count_antes == 0:
                print("📭 No hay solicitudes para archivar")
                return
            
            # Preparar datos para archivado histórico
            fecha_archivado = datetime.now()
            semana = fecha_archivado.isocalendar()[1]  # Semana ISO
            mes = fecha_archivado.month
            ano = fecha_archivado.year
            
            # Mover a tabla histórica con cálculos para reportes
            cur.execute(f"""
                INSERT INTO pedidos_historicos 
                (id_orden_profesional, codigo, producto, terminacion, id_factura,
                 prioridad, cantidad, tiempo_estimado, base, ubicacion, sucursal,
                 operador, estado, fecha_creacion, fecha_asignacion, fecha_completado,
                 notas, fecha_archivado, semana_archivado, mes_archivado, ano_archivado,
                 tiempo_total_procesamiento)
                SELECT 
                    id_orden_profesional, codigo, producto, terminacion, id_factura,
                    prioridad, cantidad, tiempo_estimado, base, ubicacion, sucursal,
                    operador, estado, fecha_creacion, fecha_asignacion, fecha_completado,
                    notas, %s, %s, %s, %s,
                    CASE 
                        WHEN fecha_completado IS NOT NULL AND fecha_creacion IS NOT NULL 
                        THEN EXTRACT(EPOCH FROM (fecha_completado - fecha_creacion))/60
                        ELSE NULL 
                    END as tiempo_total_procesamiento
                FROM {tabla_sucursal}
                ON CONFLICT (id_orden_profesional, fecha_archivado) DO NOTHING
            """, (fecha_archivado, semana, mes, ano))
            
            registros_archivados = cur.rowcount
            
            # Limpiar tabla de trabajo
            cur.execute(f"DELETE FROM {tabla_sucursal}")
            
            conn.commit()
            cur.close()
            conn.close()
            
            print(f"✅ Archivado completado: {registros_archivados} solicitudes → histórico")
            
            # Actualizar UI en hilo principal
            self.root.after(0, lambda: self._mostrar_notificacion_archivado(registros_archivados))
            self.root.after(0, lambda: self.cargar_datos(forzar_recarga=True))
            
        except Exception as e:
            print(f"❌ Error en archivado automático: {e}")
    
    def _mostrar_notificacion_archivado(self, cantidad):
        """Muestra notificación del archivado ejecutado"""
        if hasattr(self, 'label_ultima_actualizacion'):
            self.label_ultima_actualizacion.config(
                text=f"📁 Archivado automático: {cantidad} solicitudes archivadas a las 6pm"
            )
    
    def mostrar_reportes(self):
        """Muestra ventana de reportes históricos"""
        ventana_reportes = ttk.Toplevel(self.root)
        aplicar_icono_y_titulo(ventana_reportes, "Reportes Históricos")
        ventana_reportes.geometry("800x600")
        
        # Frame principal
        main_frame = ttk.Frame(ventana_reportes, padding="20")
        main_frame.pack(fill="both", expand=True)
        
        # Título
        ttk.Label(main_frame, text="📊 Reportes de Productividad", 
                 font=("Segoe UI", 16, "bold")).pack(pady=(0, 20))
        
        # Frame de botones de reportes
        botones_frame = ttk.Frame(main_frame)
        botones_frame.pack(pady=(0, 20))
        
        ttk.Button(botones_frame, text="📈 Reporte Semanal", 
                  command=lambda: self.generar_reporte("semanal"), 
                  bootstyle="primary", width=20).pack(side="left", padx=10)
        
        ttk.Button(botones_frame, text="📅 Reporte Mensual", 
                  command=lambda: self.generar_reporte("mensual"), 
                  bootstyle="info", width=20).pack(side="left", padx=10)
        
        # Área de texto para mostrar reportes
        self.texto_reporte = tk.Text(main_frame, wrap="word", width=80, height=25)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=self.texto_reporte.yview)
        self.texto_reporte.configure(yscrollcommand=scrollbar.set)
        
        self.texto_reporte.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
    
    def generar_reporte(self, tipo):
        """Genera reporte semanal o mensual"""
        try:
            conn = self.conectar_db()
            if not conn:
                messagebox.showerror("Error", "No se pudo conectar a la base de datos")
                return
            
            cur = conn.cursor()
            fecha_actual = datetime.now()
            
            if tipo == "semanal":
                semana_actual = fecha_actual.isocalendar()[1]
                ano_actual = fecha_actual.year
                
                cur.execute("""
                    SELECT * FROM reporte_semanal 
                    WHERE ano_archivado = %s
                    ORDER BY semana_archivado DESC, sucursal
                    LIMIT 15
                """, (ano_actual,))
                
                self.texto_reporte.delete(1.0, tk.END)
                self.texto_reporte.insert(tk.END, f"📈 REPORTES SEMANALES {ano_actual}\n")
                self.texto_reporte.insert(tk.END, "="*60 + "\n\n")
                
            elif tipo == "mensual":
                ano_actual = fecha_actual.year
                
                cur.execute("""
                    SELECT * FROM reporte_mensual 
                    WHERE ano_archivado = %s
                    ORDER BY mes_archivado DESC, sucursal
                    LIMIT 12
                """, (ano_actual,))
                
                self.texto_reporte.delete(1.0, tk.END)
                self.texto_reporte.insert(tk.END, f"📅 REPORTES MENSUALES {ano_actual}\n")
                self.texto_reporte.insert(tk.END, "="*60 + "\n\n")
            
            resultados = cur.fetchall()
            
            if not resultados:
                self.texto_reporte.insert(tk.END, "📭 No hay datos históricos disponibles\n")
                self.texto_reporte.insert(tk.END, "💡 Los reportes aparecerán después del primer archivado automático\n")
                return
            
            for fila in resultados:
                if tipo == "semanal":
                    sucursal, semana, ano, total, completados, cancelados, pendientes, tiempo_prom, operadores, _, _ = fila
                    periodo = f"Semana {semana}/{ano}"
                else:  # mensual
                    sucursal, mes, ano, total, completados, cancelados, alta_p, media_p, baja_p, tiempo_prom, operadores, dias = fila
                    meses = ["", "Ene", "Feb", "Mar", "Abr", "May", "Jun", 
                            "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
                    periodo = f"{meses[mes]} {ano}"
                
                self.texto_reporte.insert(tk.END, f"🏢 {sucursal} - {periodo}\n")
                self.texto_reporte.insert(tk.END, f"   📋 Total: {total} pedidos\n")
                self.texto_reporte.insert(tk.END, f"   ✅ Completados: {completados}\n")
                self.texto_reporte.insert(tk.END, f"   ❌ Cancelados: {cancelados}\n")
                
                if tipo == "semanal":
                    self.texto_reporte.insert(tk.END, f"   ⏳ Pendientes: {pendientes}\n")
                else:
                    self.texto_reporte.insert(tk.END, f"   🔴 Alta prioridad: {alta_p}\n")
                    self.texto_reporte.insert(tk.END, f"   🟡 Media prioridad: {media_p}\n")
                    self.texto_reporte.insert(tk.END, f"   🟢 Baja prioridad: {baja_p}\n")
                    self.texto_reporte.insert(tk.END, f"   📅 Días activos: {dias}\n")
                
                self.texto_reporte.insert(tk.END, f"   ⏱️ Tiempo promedio: {tiempo_prom:.1f} min\n")
                self.texto_reporte.insert(tk.END, f"   👥 Operadores: {operadores}\n")
                self.texto_reporte.insert(tk.END, "\n")
            
            cur.close()
            conn.close()
            
        except Exception as e:
            messagebox.showerror("Error", f"Error generando reporte: {e}")
            print(f"❌ Error en reporte: {e}")
    
    def editar_pedido(self):
        """Edita un pedido seleccionado"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Selección", "Selecciona un pedido para editar.")
            return
        
        item = self.tree.item(selection[0])
        pedido_id = item['values'][0]
        
        # Crear ventana de edición (implementar según necesidades)
        messagebox.showinfo("En Desarrollo", "Función de edición en desarrollo...")
    
    def mostrar_estadisticas(self):
        """Muestra estadísticas de la cola de producción"""
        conn = self.conectar_db()
        if not conn:
            return
        
        # Obtener la tabla correspondiente a la sucursal del usuario
        sucursal = obtener_sucursal_usuario(self.usuario_username)
        tabla_sucursal = obtener_tabla_sucursal(sucursal)
        
        try:
            cur = conn.cursor()
            
            # Consultar estadísticas básicas
            cur.execute(f"""
                SELECT 
                    COUNT(*) as total,
                    SUM(CASE WHEN estado = 'Pendiente' THEN 1 ELSE 0 END) as pendientes,
                    SUM(CASE WHEN estado = 'En Proceso' THEN 1 ELSE 0 END) as en_proceso,
                    SUM(CASE WHEN estado = 'Completado' THEN 1 ELSE 0 END) as completados,
                    SUM(CASE WHEN prioridad = 'Crítica' THEN 1 ELSE 0 END) as criticos,
                    SUM(CASE WHEN prioridad = 'Urgente' THEN 1 ELSE 0 END) as urgentes
                FROM {tabla_sucursal}
                WHERE estado <> 'Cancelado'
            """)
            
            stats = cur.fetchone()
            
            # Consultar estadísticas de tiempo
            cur.execute(f"""
                SELECT 
                    COALESCE(SUM(tiempo_estimado), 0) as tiempo_total_pendiente,
                    COALESCE(AVG(tiempo_estimado), 0) as tiempo_promedio,
                    COUNT(*) as pedidos_vencidos
                FROM {tabla_sucursal} 
                WHERE estado IN ('Pendiente', 'En Proceso')
                AND fecha_compromiso < NOW()
            """)
            
            tiempo_stats = cur.fetchone()
            
            # Tiempo total acumulado para nuevos pedidos
            cur.execute(f"""
                SELECT COALESCE(SUM(tiempo_estimado), 0)
                FROM {tabla_sucursal} 
                WHERE estado IN ('Pendiente', 'En Proceso')
            """)
            
            tiempo_acumulado = cur.fetchone()[0]
            
            mensaje = f"""📊 Estadísticas de {self.sucursal_actual}:

📋 PEDIDOS:
• Total de pedidos: {stats[0]}
• Pendientes: {stats[1]}
• En proceso: {stats[2]}
• Completados: {stats[3]}

⚡ PRIORIDADES:
• Críticos: {stats[4]}
• Urgentes: {stats[5]}

⏱️ TIEMPOS:
• Tiempo acumulado pendiente: {tiempo_acumulado} minutos ({tiempo_acumulado//60}h {tiempo_acumulado%60}m)
• Tiempo promedio por pedido: {tiempo_stats[1]:.1f} minutos
• Pedidos vencidos: {tiempo_stats[2]}

📅 Próximo pedido se compromete en: {tiempo_acumulado} minutos
"""
            
            messagebox.showinfo("Estadísticas", mensaje)
            
            cur.close()
            conn.close()
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al obtener estadísticas: {e}")
    
    def mostrar_proximos_vencer(self):
        """Muestra pedidos que están próximos a vencer (menos de 30 minutos)"""
        conn = self.conectar_db()
        if not conn:
            return
        
        try:
            cur = conn.cursor()
            
            # Consultar pedidos próximos a vencer (menos de 30 minutos)
            cur.execute("""
                SELECT id_factura, codigo, producto, terminacion, prioridad, 
                       fecha_compromiso, tiempo_estimado,
                       EXTRACT(EPOCH FROM (fecha_compromiso - NOW()))/60 as minutos_restantes
                FROM lista_espera 
                WHERE sucursal = %s 
                AND estado IN ('Pendiente', 'En Proceso')
                AND fecha_compromiso > NOW()
                AND fecha_compromiso <= NOW() + INTERVAL '30 minutes'
                ORDER BY fecha_compromiso ASC
            """, (self.sucursal_actual,))
            
            pedidos_proximos = cur.fetchall()
            
            if not pedidos_proximos:
                messagebox.showinfo("Próximos a Vencer", "✅ No hay pedidos próximos a vencer en los próximos 30 minutos.")
                cur.close()
                conn.close()
                return
            
            # Crear ventana emergente con detalles
            ventana = ttk.Toplevel(self.root)
            aplicar_icono_y_titulo(ventana, "Pedidos Próximos a Vencer")
            ventana.geometry("800x400")
            
            ttk.Label(ventana, text="⚠️ PEDIDOS PRÓXIMOS A VENCER (Menos de 30 minutos)", 
                     font=("Arial", 14, "bold")).pack(pady=10)
            
            # Frame con scrollbar para la lista
            frame_lista = ttk.Frame(ventana)
            frame_lista.pack(fill="both", expand=True, padx=20, pady=10)
            
            # Treeview para mostrar los pedidos
            tree_proximos = ttk.Treeview(frame_lista)
            tree_proximos["columns"] = ("Factura", "Código", "Producto", "Prioridad", "Minutos Restantes", "Compromiso")
            tree_proximos["show"] = "headings"
            
            # Configurar headers
            tree_proximos.heading("Factura", text="ID Factura")
            tree_proximos.heading("Código", text="Código")
            tree_proximos.heading("Producto", text="Producto")
            tree_proximos.heading("Prioridad", text="Prioridad")
            tree_proximos.heading("Minutos Restantes", text="Min. Rest.")
            tree_proximos.heading("Compromiso", text="Fecha Compromiso")
            
            # Configurar anchos
            tree_proximos.column("Factura", width=100)
            tree_proximos.column("Código", width=100)
            tree_proximos.column("Producto", width=150)
            tree_proximos.column("Prioridad", width=80)
            tree_proximos.column("Minutos Restantes", width=80)
            tree_proximos.column("Compromiso", width=150)
            
            # Agregar datos
            for pedido in pedidos_proximos:
                minutos_rest = int(pedido[7])
                fecha_comp = pedido[5].strftime("%d/%m/%Y %H:%M")
                
                # Color según urgencia
                tag = ""
                if minutos_rest <= 5:
                    tag = "critico"
                elif minutos_rest <= 15:
                    tag = "urgente"
                else:
                    tag = "proximo"
                
                tree_proximos.insert("", "end", values=(
                    pedido[1], pedido[2], pedido[3], pedido[4], 
                    f"{minutos_rest} min", fecha_comp
                ), tags=(tag,))
            
            # Configurar colores
            tree_proximos.tag_configure('critico', background='#ffcdd2')  # Rojo
            tree_proximos.tag_configure('urgente', background='#ffe0b2')  # Naranja
            tree_proximos.tag_configure('proximo', background='#fff9c4')  # Amarillo
            
            # Scrollbar
            scroll_proximos = ttk.Scrollbar(frame_lista, orient="vertical", command=tree_proximos.yview)
            tree_proximos.configure(yscrollcommand=scroll_proximos.set)
            
            tree_proximos.pack(side="left", fill="both", expand=True)
            scroll_proximos.pack(side="right", fill="y")
            
            # Botón cerrar
            ttk.Button(ventana, text="Cerrar", command=ventana.destroy,
                      bootstyle="secondary").pack(pady=10)
            
            cur.close()
            conn.close()
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al consultar pedidos próximos a vencer: {e}")
    
    def abrir_reporte_performance(self):
        """Abre el sistema de reporte de performance"""
        try:
            import subprocess
            import sys
            import os
            
            script_dir = os.path.dirname(os.path.abspath(__file__))
            reporte_path = os.path.join(script_dir, "reporte_performance.py")
            
            subprocess.Popen([sys.executable, reporte_path])
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir el reporte de performance: {e}")
    
    def agregar_a_cola(self, datos_pedido):
        """Agrega un nuevo pedido a la cola de espera"""
        conn = self.conectar_db()
        if not conn:
            return False
        
        # Obtener la tabla correspondiente a la sucursal del pedido
        sucursal = datos_pedido.get('sucursal', obtener_sucursal_usuario(self.usuario_username))
        tabla_sucursal = obtener_tabla_sucursal(sucursal)
        
        try:
            cur = conn.cursor()
            cur.execute(f"""
                INSERT INTO {tabla_sucursal} 
                (codigo, producto, terminacion, id_factura, codigo_base, prioridad, 
                 cantidad, sucursal, usuario_creador, base, ubicacion, fecha_compromiso, observaciones)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                RETURNING id
            """, (
                datos_pedido['codigo'], datos_pedido['producto'], datos_pedido['terminacion'],
                datos_pedido['id_factura'], datos_pedido['codigo_base'], datos_pedido['prioridad'],
                datos_pedido['cantidad'], datos_pedido['sucursal'], datos_pedido['usuario_creador'],
                datos_pedido.get('base', ''), datos_pedido.get('ubicacion', ''),
                datos_pedido.get('fecha_compromiso'), datos_pedido.get('observaciones', '')
            ))
            
            pedido_id = cur.fetchone()[0]
            conn.commit()
            cur.close()
            conn.close()
            
            return pedido_id
            
        except Exception as e:
            print(f"Error al agregar pedido a cola: {e}")
            return False
    
    # === FUNCIONES DE CONTROL DE SONIDO ===
    
    def alternar_sonido(self):
        """Alterna el estado del sonido (ON/OFF)"""
        estado = self.notificaciones.alternar_sonido()
        
        # Actualizar label del botón solo si existe en la interfaz
        if hasattr(self, 'btn_sonido') and self.btn_sonido is not None:
            if estado:
                self.btn_sonido.configure(text="🔊 Sonido ON", bootstyle="outline-success")
            else:
                self.btn_sonido.configure(text="🔇 Sonido OFF", bootstyle="outline-secondary")
        # Sin popups
    
    def configurar_impresora(self):
        """Abre el configurador de impresora"""
        try:
            impresora_seleccionada = mostrar_seleccionador_impresora()
            if impresora_seleccionada:
                print(f"✅ Impresora configurada: {impresora_seleccionada}")
                # Actualizar label de impresora en la interfaz
                if hasattr(self, 'label_impresora'):
                    self.label_impresora.config(text=f"🖨️ {impresora_seleccionada[:20]}", 
                                              bootstyle="success")
        except Exception as e:
            print(f"❌ Error configurando impresora: {e}")
            messagebox.showerror("Error", f"Error configurando impresora: {e}")

    def test_impresion(self):
        """Realiza una impresión de prueba"""
        try:
            # Verificar que hay impresora configurada
            impresora = cargar_impresora_guardada()
            if not impresora:
                respuesta = messagebox.askyesno("Sin Impresora", 
                                               "No hay impresora configurada. ¿Desea configurar una ahora?")
                if respuesta:
                    self.configurar_impresora()
                    impresora = cargar_impresora_guardada()
                    if not impresora:
                        return
                else:
                    return
            
            # Generar etiqueta de prueba
            zpl_prueba = generar_zpl_gestor(
                codigo="TEST001",
                descripcion="ETIQUETA DE PRUEBA",
                producto="Test",
                terminacion="Prueba",
                presentacion="",
                cantidad=1,
                base="TEST",
                ubicacion="PRUEBA",
                sucursal=self.sucursal_actual,
                id_profesional="T001"
            )
            
            print("🖨️ Iniciando impresión de prueba...")
            
            # Intentar imprimir
            exito = imprimir_zebra_zpl_gestor(zpl_prueba)
            
            if exito:
                print("✅ Prueba de impresión exitosa")
            else:
                print("❌ Prueba de impresión falló")
                
        except Exception as e:
            print(f"❌ Error en prueba de impresión: {e}")
            messagebox.showerror("Error", f"Error en prueba de impresión: {e}")

    def test_notificacion(self):
        """Prueba el sistema de notificaciones"""
        self.notificaciones.test_sonido()
        # Sin popup: solo log en consola
    
    def detectar_nuevos_pedidos(self, conteo_actual):
        """Detecta si hay nuevos pedidos y reproduce notificación según prioridad"""        
        if self.ultimo_conteo_pedidos > 0 and conteo_actual > self.ultimo_conteo_pedidos:
            # Hay nuevos pedidos
            nuevos_pedidos = conteo_actual - self.ultimo_conteo_pedidos
            
            # Determinar prioridad más alta de los pedidos actuales
            prioridad_maxima = self.obtener_prioridad_maxima()
            
            # Reproducir campanas según prioridad
            if prioridad_maxima:
                self.notificaciones.reproducir_sonido_por_prioridad(prioridad_maxima)
                print(f"🔔 {nuevos_pedidos} nuevo(s) pedido(s) recibido(s) para {self.sucursal_actual} - Prioridad: {prioridad_maxima}")
            else:
                # Fallback al sonido normal si no se puede determinar prioridad
                self.notificaciones.reproducir_sonido('nuevo_pedido')
                print(f"🔔 {nuevos_pedidos} nuevo(s) pedido(s) recibido(s) para {self.sucursal_actual}")
        
        self.ultimo_conteo_pedidos = conteo_actual
    
    def hay_pedidos_urgentes(self):
        """Verifica si hay pedidos con prioridad Alta en la lista actual"""
        try:
            for item in self.tree.get_children():
                valores = self.tree.item(item)['values']
                if len(valores) > 6 and valores[6] == "Alta":  # Columna de prioridad
                    return True
            return False
        except:
            return False
    
    def obtener_prioridad_maxima(self):
        """Obtiene la prioridad más alta entre todos los pedidos actuales"""
        try:
            prioridades = []
            for item in self.tree.get_children():
                valores = self.tree.item(item)['values']
                if len(valores) > 6:
                    prioridad = valores[6]  # Columna de prioridad
                    if prioridad in ['Alta', 'Media', 'Baja']:
                        prioridades.append(prioridad)
            
            # Determinar prioridad máxima según jerarquía
            if 'Alta' in prioridades:
                return 'Alta'
            elif 'Media' in prioridades:
                return 'Media'
            elif 'Baja' in prioridades:
                return 'Baja'
            else:
                return 'Media'  # Prioridad por defecto
        except Exception as e:
            print(f"❌ Error obteniendo prioridad máxima: {e}")
            return 'Media'  # Prioridad por defecto
    
    def notificar_pedido_completado(self):
        """Notifica cuando un pedido se completa"""
        self.notificaciones.reproducir_sonido('pedido_completado')
    
    # ====== Recordatorio sonoro de pendientes cada 30s ======
    def iniciar_recordatorio_pendientes(self):
        """Inicia el bucle de recordatorio sonoro si hay pendientes."""
        try:
            # Cancelar anterior si existe
            if hasattr(self, "_recordatorio_timer_id") and self._recordatorio_timer_id:
                try:
                    self.root.after_cancel(self._recordatorio_timer_id)
                except Exception:
                    pass
            # Programar primer tick inmediato para evaluar estado actual
            self._recordatorio_timer_id = self.root.after(500, self._recordatorio_pendientes_tick)
        except Exception:
            pass

    def _recordatorio_pendientes_tick(self):
        """Tick del recordatorio: si hay pendientes, reproducir un pitido."""
        try:
            # Verificar que la ventana sigue viva
            try:
                if not self.root.winfo_exists():
                    return
            except Exception:
                return

            if self._hay_pedidos_pendientes():
                # Elegir intensidad según prioridad máxima visible
                try:
                    prioridad = self.obtener_prioridad_maxima()
                except Exception:
                    prioridad = None

                if prioridad:
                    self.notificaciones.reproducir_sonido_por_prioridad(prioridad)
                else:
                    self.notificaciones.reproducir_sonido('nuevo_pedido')

            # Reprogramar siguiente tick
            self._recordatorio_timer_id = self.root.after(self._recordatorio_intervalo_ms, self._recordatorio_pendientes_tick)
        except tk.TclError:
            # La ventana puede haberse destruido; no reprogramar
            return
        except Exception:
            # Ante cualquier error, intentar reprogramar para resiliencia
            try:
                self._recordatorio_timer_id = self.root.after(self._recordatorio_intervalo_ms, self._recordatorio_pendientes_tick)
            except Exception:
                pass

    def _hay_pedidos_pendientes(self) -> bool:
        """Determina si hay pedidos con estado 'Pendiente' o 'En Espera' en la vista actual.

        Nota: Se considera 'pendiente' todo lo que no está Finalizado ni Cancelado y aún no está en proceso,
        por lo que se cuentan específicamente los estados 'Pendiente' y 'En Espera'.
        """
        try:
            # Si el Tree aún no existe, no hay pendientes detectables
            if not hasattr(self, 'tree') or self.tree is None:
                return False

            for item in self.tree.get_children():
                valores = self.tree.item(item)['values']
                if not valores:
                    continue
                # Columna "Estado" es el índice 8 según la definición de columnas
                if len(valores) > 8:
                    estado = str(valores[8]).strip()
                    if estado in ("Pendiente", "En Espera"):
                        return True
            return False
        except Exception:
            return False
    
    def cerrar_aplicacion(self):
        """Cierra la aplicación de manera segura"""
        try:
            # Si hay procesos bloqueados activos, preguntar confirmación
            facturas_bloqueadas = list(self.bloqueos_por_factura.keys())
            if facturas_bloqueadas:
                # Obtener tiempo restante de la primera factura bloqueada para mostrar
                factura_ejemplo = facturas_bloqueadas[0]
                tiempo_restante = self._obtener_tiempo_bloqueo_restante(factura_ejemplo)
                if tiempo_restante > 0:
                    lista_facturas = ", ".join(facturas_bloqueadas)
                    confirmacion = mostrar_pregunta("Listas en Proceso", 
                                                  f"Hay listas en proceso: {lista_facturas}\n"
                                                  f"Tiempo restante (ejemplo): {tiempo_restante // 60}m {tiempo_restante % 60}s.\n\n"
                                                  f"¿Está seguro de que desea cerrar la aplicación ahora?\n"
                                                  f"Esto podría interrumpir operaciones importantes.")
                    if not confirmacion:
                        return  # No cerrar si el usuario cancela
            
            print("🔍 Cerrando aplicación...")
            # Detener actualizaciones automáticas
            self.actualizacion_en_progreso = False
            
            # Cancelar timers de bloqueo por factura si existen
            for factura, bloqueo_info in list(self.bloqueos_por_factura.items()):
                timer_id = bloqueo_info.get('timer')
                if timer_id:
                    try:
                        self.root.after_cancel(timer_id)
                        print(f"🔍 Timer de bloqueo cancelado para factura {factura}")
                    except:
                        pass
            self.bloqueos_por_factura.clear()
            
            # Cancelar timer específico si existe
            if hasattr(self, 'timer_id') and self.timer_id and hasattr(self, 'root') and self.root:
                try:
                    self.root.after_cancel(self.timer_id)
                    print("🔍 Timer de actualización cancelado")
                except:
                    pass
            
            # Cancelar recordatorio si existe
            if hasattr(self, '_recordatorio_timer_id') and self._recordatorio_timer_id and hasattr(self, 'root') and self.root:
                try:
                    self.root.after_cancel(self._recordatorio_timer_id)
                except:
                    pass
            
            # Los bloqueos por factura ya fueron limpiados arriba
            
            # Cerrar ventana
            if hasattr(self, 'root') and self.root:
                self.root.quit()
                self.root.destroy()
                
            print("✅ Aplicación cerrada correctamente")
        except Exception as e:
            print(f"Error al cerrar: {e}")
            # Forzar cierre si es necesario
            try:
                import sys
                sys.exit()
            except:
                pass
    
    def run(self):
        """Inicia la aplicación"""
        self.root.mainloop()

if __name__ == "__main__":
    # Verificar dependencias críticas antes de continuar
    print("🔍 Verificando dependencias del sistema...")
    if not verificar_dependencias_criticas():
        print("❌ Usuario decidió no continuar debido a dependencias faltantes")
        sys.exit(1)
    
    # Crear root base y ejecutar login como Toplevel para mantener un único mainloop
    debug_log("Iniciando sistema de login para coloristas...")
    # Usar ttkbootstrap Window para que el login herede correctamente el tema azul
    try:
        base_root = ttk.Window(themename="flatly")
    except Exception:
        base_root = tk.Tk()
    try:
        base_root.withdraw()  # ocultar mientras se muestra el login
    except Exception:
        pass

    usuario_info, sucursal_info = ejecutar_login_colorista(master=base_root)
    
    if usuario_info:
        print(f"✅ Usuario autenticado: {usuario_info['username']} ({usuario_info['rol']}) - {sucursal_info}")
        try:
            base_root.deiconify()
        except Exception:
            pass
        app = GestorListaEspera(usuario_info, sucursal_info, master=base_root)
        app.run()
    else:
        print("❌ No se pudo autenticar usuario")
        try:
            base_root.destroy()
        except Exception:
            pass
        sys.exit(0)