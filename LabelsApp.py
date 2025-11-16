import tkinter as tk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import messagebox, simpledialog
import psycopg2
import os, sys, subprocess
import json
import base64
import os
import sys
import pandas as pd
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.platypus import Table, TableStyle
import tempfile
import math
from collections import defaultdict
import hashlib
import time
from datetime import datetime, timedelta
from PIL import Image, ImageTk
from typing import Any
import ctypes
import threading

# Control de logs para evitar ruido en consola en producción
DEBUG_LOGS = True  # Habilitado temporalmente para verificar sucursal

def debug_log(*args: Any, **kwargs: Any):
    if DEBUG_LOGS:
        try:
            print(*args, **kwargs)
        except Exception:
            pass

# Silenciar stdout por defecto para evitar mensajes en consola (mantiene stderr)
if not DEBUG_LOGS:
    try:
        sys.stdout = open(os.devnull, 'w')
    except Exception:
        pass

APP_VERSION = "5.5.0"
URL_VERSION = "https://labelsapp.onrender.com/Version2.txt"
URL_EXE = "https://labelsapp.onrender.com/LabelsApp2.exe"

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

def version_tuple(v):
    return tuple(int(x) for x in v.strip().split(".") if x.isdigit())

def is_newer(latest, current):
    return version_tuple(latest) > version_tuple(current)

def run_windows_updater(new_exe_path, current_exe_path):
    # Creamos un .bat temporal que espera a que termine el proceso actual,
    # reemplaza el exe y lanza la nueva versión.
    bat = tempfile.NamedTemporaryFile(delete=False, suffix=".bat", mode="w", encoding="utf-8")
    new_p = new_exe_path.replace("/", "\\")
    cur_p = current_exe_path.replace("/", "\\")
    exe_name = os.path.basename(cur_p)
    bat_contents = f"""@echo off
timeout /t 2 /nobreak > nul
:waitloop
tasklist /FI "IMAGENAME eq {exe_name}" | find /I "{exe_name}" > nul
if %ERRORLEVEL%==0 (
  timeout /t 1 > nul
  goto waitloop
)
move /Y "{new_p}" "{cur_p}"
start "" "{cur_p}"
del "%~f0"
"""
    bat.write(bat_contents)
    bat.close()
    # lanzar el .bat y salir
    subprocess.Popen(["cmd", "/c", bat.name], creationflags=subprocess.CREATE_NEW_CONSOLE)
    sys.exit(0)

def _is_frozen_exe():
    return getattr(sys, "frozen", False)

def _current_binary_path():
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
    
    # Intentar aplicar icono si está disponible
    try:
        if ICONO_PATH and os.path.exists(ICONO_PATH):
            progress_window.iconbitmap(ICONO_PATH)
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
        is_frozen = _is_frozen_exe()
        headers = {
            "User-Agent": f"LabelsApp/{APP_VERSION}",
            "Cache-Control": "no-cache",
            "Pragma": "no-cache",
        }
        # Importación perezosa
        import requests  # type: ignore

        r = requests.get(URL_VERSION, timeout=5, headers=headers)
        r.raise_for_status()
        latest = r.text.strip()
        if not latest:
            return

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
                        if ICONO_PATH and os.path.exists(ICONO_PATH):
                            notif_window.iconbitmap(ICONO_PATH)
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
                    debug_log(f"⚠️ DEBUG UPDATE: No se pudo crear ventana de progreso: {e}")
                    progress_window = None
                
                base_path = os.path.dirname(_current_binary_path()) or os.getcwd()
                
                # Usar un nombre único para evitar conflictos
                timestamp = str(int(time.time()))
                new_exe = os.path.join(base_path, f"LabelsApp_new_{timestamp}.exe")
                
                debug_log(f"🔄 DEBUG UPDATE: Descargando a: {new_exe}")
                debug_log(f"🔄 DEBUG UPDATE: Directorio base: {base_path}")

                if progress_window:
                    progress_window.actualizar_estado("Verificando permisos...")

                # Verificar permisos de escritura en el directorio
                try:
                    test_file = os.path.join(base_path, "test_write_permissions.tmp")
                    with open(test_file, "w") as f:
                        f.write("test")
                    os.remove(test_file)
                    debug_log(f"🔄 DEBUG UPDATE: Permisos de escritura verificados")
                except Exception as e:
                    debug_log(f"❌ DEBUG UPDATE: Sin permisos de escritura en {base_path}: {e}")
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
                    debug_log(f"❌ DEBUG UPDATE: Error en descarga: {e}")
                    if progress_window:
                        progress_window.actualizar_estado("Error en la descarga")
                        time.sleep(2)
                        progress_window.cerrar_ventana()
                    return
                
                debug_log(f"🔄 DEBUG UPDATE: Descarga completada")

                if progress_window:
                    progress_window.actualizar_estado("Verificando descarga...")

                # Verificación mínima de tamaño
                try:
                    size = os.path.getsize(new_exe)
                    debug_log(f"🔄 DEBUG UPDATE: Tamaño del archivo descargado: {size} bytes")
                    if size < 100_000:
                        debug_log(f"🔄 DEBUG UPDATE: Archivo muy pequeño, cancelando actualización")
                        if progress_window:
                            progress_window.actualizar_estado("Error: Archivo incompleto")
                            time.sleep(2)
                            progress_window.cerrar_ventana()
                        return
                except Exception as e:
                    debug_log(f"🔄 DEBUG UPDATE: Error verificando tamaño: {e}")
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

                debug_log(f"🔄 DEBUG UPDATE: Iniciando proceso de reemplazo...")
                run_windows_updater(new_exe, _current_binary_path())
                return

            # Otros SO: permanecer en silencio
            return
    except Exception:
        # No bloquear inicio por fallas de actualización
        return

if __name__ == "__main__":
    # Verificación en segundo plano si no se pasa --no-update
    if "--no-update" not in sys.argv:
        try:
            threading.Thread(target=check_update, daemon=True).start()
        except Exception:
            pass

# === SISTEMA DE LOGIN INTEGRADO ===
class SistemaLoginIntegrado:
    """Sistema de login integrado para LabelsApp"""
    
    def __init__(self):
        self.db_config = {
            "host": "dpg-d1b18u8dl3ps73e68v1g-a.oregon-postgres.render.com",
            "port": 5432,
            "database": "labels_app_db",
            "user": "admin",
            "password": "KCFjzM4KYzSQx63ArufESIXq03EFXHz3",
            "sslmode": "require"
        }
        self.usuario_actual = None
        self.sucursal = None
        
    def conectar_bd(self):
        """Conecta a la base de datos"""
        try:
            return psycopg2.connect(**self.db_config)
        except Exception as e:
            debug_log(f"❌ Error conectando a BD: {e}")
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
            
            # Verificar que el rol sea apropiado para LabelsApp
            roles_permitidos = ['facturador', 'administrador']
            if usuario[4] not in roles_permitidos:
                return {"error": f"Rol '{usuario[4]}' no tiene acceso a PaintFlow. Se requiere rol de facturador o administrador."}
            
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

        Si se pasa `master`, se crea como Toplevel modal para usar un único root.
        """
        if master is not None:
            ventana_login = tk.Toplevel(master)
            try:
                ventana_login.transient(master)
                ventana_login.grab_set()
            except Exception:
                pass
        else:
            ventana_login = tk.Tk()
        ventana_login.title("PaintFlow — Login")
        ventana_login.geometry("600x330")
        ventana_login.resizable(False, False)
        ventana_login.configure(bg="#f5f5f5")
        # Aplicar estilo ttkbootstrap "flatly" (primary azul) en el login para bordes azules
        try:
            ttk.Style(theme="flatly")
            try:
                style = ttk.Style()
                style.configure('primary.TEntry', foreground="#1b1f23", insertcolor="#0D47A1")
                style.map('primary.TEntry', 
                          bordercolor=[('focus', '#1565C0'), ('!focus', '#1976D2')],
                          lightcolor=[('focus', '#1565C0')])
            except Exception:
                pass
        except Exception:
            try:
                ttk.Style()
            except Exception:
                pass
        
        # Cargar preferencias de login (recordar acceso: usuario + contraseña)
        saved_username = ""
        saved_password = ""
        remember_access_saved = False
        try:
            cfg_path = obtener_ruta_absoluta("paintflow_login.json")
            if os.path.exists(cfg_path):
                with open(cfg_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    saved_username = data.get('usuario', "") or ""
                    # Back-compat: si existe cualquiera de las banderas anteriores, recordar acceso
                    remember_access_saved = bool(data.get('recordar', False) or data.get('recordar_pass', False))
                    enc_pwd = data.get('password')
                    if remember_access_saved and enc_pwd:
                        try:
                            saved_password = base64.b64decode(enc_pwd.encode('utf-8')).decode('utf-8')
                        except Exception:
                            saved_password = ""
        except Exception:
            pass
        
        # Configurar icono si existe usando ruta absoluta
        icono_path = obtener_ruta_absoluta("icono.ico")
        if os.path.exists(icono_path):
            try:
                ventana_login.iconbitmap(icono_path)
            except:
                pass
        
        # Centrar ventana
        ventana_login.update_idletasks()
        x = (ventana_login.winfo_screenwidth() // 2) - (600 // 2)
        y = (ventana_login.winfo_screenheight() // 2) - (330 // 2)
        ventana_login.geometry(f"600x330+{x}+{y}")

        # Contenedor principal (card horizontal: logo | divisor | formulario)
        main_frame = tk.Frame(ventana_login, bg="white", relief="flat", bd=0)
        main_frame.pack(fill="both", expand=True, padx=16, pady=12)

        content_frame = tk.Frame(main_frame, bg="white")
        content_frame.pack(fill="both", expand=True, padx=8, pady=8)

        card = tk.Frame(content_frame, bg="white", bd=0, highlightthickness=0)
        card.pack(fill="both", expand=True, padx=2, pady=4)
        card.grid_columnconfigure(0, weight=0)
        card.grid_columnconfigure(1, weight=0)
        card.grid_columnconfigure(2, weight=1)
        card.grid_rowconfigure(0, weight=1)

        # Panel izquierdo con logo
        left_panel = tk.Frame(card, bg="white")
        left_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 10), pady=(2, 6))

        logo_path = obtener_ruta_absoluta("logo.png")
        if os.path.exists(logo_path):
            try:
                logo_img = Image.open(logo_path)
                logo_img = logo_img.resize((260, 260), Image.Resampling.LANCZOS)
                self.logo_photo = ImageTk.PhotoImage(logo_img)
                tk.Label(left_panel, image=self.logo_photo, bg="white").pack(anchor="nw", padx=0, pady=0)
            except Exception as e:
                debug_log(f"Error cargando logo: {e}")
                tk.Label(left_panel, text="PAINTFLOW", font=("Segoe UI", 18, "bold"), fg="#1976D2", bg="white").pack(anchor="nw", padx=0, pady=0)
        else:
            tk.Label(left_panel, text="PAINTFLOW", font=("Segoe UI", 18, "bold"), fg="#1976D2", bg="white").pack(anchor="nw", padx=0, pady=0)

        # Divisor vertical
        divider = ttk.Separator(card, orient="vertical")
        divider.grid(row=0, column=1, sticky="ns", pady=8)

        # Panel derecho con formulario
        right_panel = tk.Frame(card, bg="white")
        right_panel.grid(row=0, column=2, sticky="nsew", padx=(10, 12), pady=12)

        form_frame = tk.Frame(right_panel, bg="white")
        form_frame.pack(fill="both", expand=True)

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

        mostrar_row = tk.Frame(controls_frame, bg="white")
        mostrar_row.pack(fill="x")
        chk_mostrar = ttk.Checkbutton(mostrar_row, text="", variable=mostrar_var, bootstyle="primary-round-toggle", command=toggle_password)
        chk_mostrar.pack(side="left")
        tk.Label(mostrar_row, text="Mostrar contraseña", font=("Segoe UI", 9), bg="white", fg="#333333").pack(side="left", padx=(6, 0))

        remember_row = tk.Frame(controls_frame, bg="white")
        remember_row.pack(fill="x", pady=(4, 0))
        chk_recordar_inline = ttk.Checkbutton(remember_row, text="", variable=recordar_acceso_var, bootstyle="primary-round-toggle")
        chk_recordar_inline.pack(side="left")
        tk.Label(remember_row, text="Recordar usuario", font=("Segoe UI", 9), bg="white", fg="#333333").pack(side="left", padx=(6, 0))

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
            
            def limpiar_login():
                try:
                    label_mensaje.configure(text="")
                except:
                    pass  # Ignorar si la ventana ya se cerró
            
            ventana_login.after(tiempo, limpiar_login)
        
        def procesar_login():
            username = entry_usuario.get().strip()
            password = entry_password.get()
            
            if not username or not password:
                mostrar_mensaje("Por favor ingresa usuario y contraseña")
                return
            
            # Deshabilitar botón mientras se verifica
            try:
                btn_login.configure(state="disabled", text="Verificando…")
                # Mostrar y arrancar la barra de progreso
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
                    # Ocultar barra de progreso
                    try:
                        pb_login.stop()
                        pb_login.pack_forget()
                    except Exception:
                        pass
                except Exception:
                    pass
                return
            
            # Login exitoso
            self.usuario_actual = resultado
            self.sucursal = resultado.get('sucursal_nombre', 'SUCURSAL PRINCIPAL')
            
            # Guardar preferencias si corresponde (recordar acceso combinado)
            try:
                cfg_path_local = obtener_ruta_absoluta("paintflow_login.json")
                if recordar_acceso_var.get():
                    try:
                        enc_pwd = base64.b64encode(password.encode('utf-8')).decode('utf-8')
                    except Exception:
                        enc_pwd = ""
                    payload = {
                        "usuario": username,
                        "recordar": True,
                        "recordar_pass": True,
                        "password": enc_pwd
                    }
                    with open(cfg_path_local, 'w', encoding='utf-8') as f:
                        json.dump(payload, f, ensure_ascii=False)
                else:
                    # Si no desea recordar nada, borrar archivo si existe
                    if os.path.exists(cfg_path_local):
                        os.remove(cfg_path_local)
            except Exception:
                pass
            
            mostrar_mensaje(f"¡Bienvenido {resultado['nombre_completo']}!", "exito")
            
            def cerrar_ventana():
                try:
                    ventana_login.destroy()
                except:
                    pass  # Ignorar si la ventana ya se cerró
            
            ventana_login.after(1500, cerrar_ventana)
            try:
                btn_login.configure(state="normal", text="INICIAR SESIÓN")
                # Ocultar barra de progreso al finalizar
                try:
                    pb_login.stop()
                    pb_login.pack_forget()
                except Exception:
                    pass
            except Exception:
                pass
        
        # Botón con estilo similar al gestor (más estrecho, tema info)
        btn_login = ttk.Button(fields_frame, text="INICIAR SESIÓN", bootstyle="info", padding=(10, 8, 10, 14), command=procesar_login, width=26)
        btn_login.pack(anchor="w", ipady=3, pady=(2, 6))

        # (Recordar usuario ya fue colocado debajo de "Mostrar contraseña")
        
        # (Recordar usuario ya se muestra debajo de 'Mostrar contraseña')
        
        # UX: atajos y enfoque
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
            # Esperar cierre del login sin crear otro mainloop
            try:
                master.wait_window(ventana_login)
            except Exception:
                pass
        else:
            ventana_login.mainloop()
        
        return self.usuario_actual is not None
    
    def debug_verificar_bd(self):
        """Método de debug para verificar la base de datos"""
        # Verificación silenciosa de base de datos
        conn = self.conectar_bd()
        if not conn:
            return
        
        try:
            cur = conn.cursor()
            # Verificación rápida sin mensajes
            cur.execute("SELECT COUNT(*) FROM usuarios WHERE activo = true")
            usuarios_activos = cur.fetchone()[0]
            cur.close()
            conn.close()
                
        except Exception as e:
            pass  # Verificación silenciosa
        finally:
            if conn:
                conn.close()

# Función para ejecutar el login
def ejecutar_login(master=None):
    """Ejecuta el sistema de login y retorna la información del usuario.

    Si se pasa `master`, el login se muestra como Toplevel sobre ese root.
    """
    sistema_login = SistemaLoginIntegrado()
    ok = sistema_login.mostrar_login(master=master)
    if ok:
        return sistema_login.usuario_actual, sistema_login.sucursal
    return None, None

   # Intentar importar win32print con manejo de errores
try:
    import win32print
    import win32api
    WIN32_AVAILABLE = True
except ImportError as e:
    WIN32_AVAILABLE = False


# === DB: carga productos desde ProductSW ===
def obtener_productos_desde_db():
    try:
        conn = psycopg2.connect(
            host="dpg-d1b18u8dl3ps73e68v1g-a.oregon-postgres.render.com",
            port=5432,
            database="labels_app_db",
            user="admin",
            password="KCFjzM4KYzSQx63ArufESIXq03EFXHz3",
            sslmode="require"
        )
        cur = conn.cursor()
        # Solo obtener productos activos
        cur.execute("SELECT codigo, nombre, base, ubicacion FROM ProductSW WHERE activo = TRUE;")
        datos = cur.fetchall()
        cur.close()
        conn.close()
        return datos
    except Exception as e:
        return []
    

# === Consulta a la base de datos ===
def obtener_datos_por_pintura(pintura_id):
    try:
        conn = psycopg2.connect(
            host="dpg-d1b18u8dl3ps73e68v1g-a.oregon-postgres.render.com",
            port=5432,
            database="labels_app_db",
            user="admin",
            password="KCFjzM4KYzSQx63ArufESIXq03EFXHz3",
            sslmode="require"
        )
        with conn.cursor() as cur:
            cur.execute("""
                SELECT 
                    p.id AS codigo,
                    p.base,
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
            """, (pintura_id,))
            return cur.fetchall()
    except Exception as e:
        return []

def obtener_datos_por_tinte(tinte_id):
    try:
        conn = psycopg2.connect(
            host="dpg-d1b18u8dl3ps73e68v1g-a.oregon-postgres.render.com",
            port=5432,
            database="labels_app_db",
            user="admin",
            password="KCFjzM4KYzSQx63ArufESIXq03EFXHz3",
            sslmode="require"
        )
        with conn.cursor() as cur:
            cur.execute("""
                SELECT 
                    t.id AS codigo,
                    t.nombre_color,
                    c.nombre AS colorante,
                    p.tipo,
                    p.cantidad
                FROM presentacion_tintes p
                JOIN tintes t ON p.id_tinte = t.id
                JOIN colorantes_tintes c ON p.id_colorante_tinte = c.codigo
                WHERE t.id = %s;
            """, (tinte_id,))
            return cur.fetchall()
    except Exception as e:
        return []




def obtener_sufijo_presentacion(presentacion):
    """Devuelve el sufijo correspondiente a la presentación seleccionada"""
    sufijos = {
        "Medio Galón": "1/2",
        "Cuarto": "QT",
        "Galón": "1",
        "Cubeta": "5",
        "1/8": "1/8"
    }
    return sufijos.get(presentacion, "")



def mostrar_codigo_base():
    base = base_var.get()
    producto = producto_var.get()
    terminacion = terminacion_var.get()
    presentacion = presentacion_var.get()

    if not base or not producto or not terminacion:
        aviso_var.set("Completa todos los campos")
        return

    # Obtener código base
    resultado = obtener_codigo_base(base, producto, terminacion)
    
    # Agregar sufijo de presentación si está seleccionada
    if presentacion and resultado != "No encontrado" and resultado != "No Aplica":
        sufijo_presentacion = obtener_sufijo_presentacion(presentacion)
        if sufijo_presentacion:
            resultado += sufijo_presentacion

    # Copiar al portapeles y mostrar avisos
    app.clipboard_clear()
    app.clipboard_append(resultado)
    codigo_base_var.set(resultado)
    aviso_var.set("Código facturación copiado en el portapapeles")
    aviso_var.set("Copiado al portapapeles")
    actualizar_vista()

    # Borrar el mensaje después de 3 segundos
    limpiar_mensaje_despues(3000)


# === Gestión de temporizadores ===
timer_id = None  # Variable global para almacenar el ID del temporizador actual

def limpiar_mensaje_despues(milisegundos):
    """Función centralizada para limpiar mensajes después de un tiempo"""
    global timer_id
    
    # Cancelar temporizador anterior si existe
    if timer_id is not None:
        try:
            app.after_cancel(timer_id)
        except:
            pass  # Ignorar errores si el temporizador ya no existe
    
    # Crear nuevo temporizador
    def limpiar():
        global timer_id
        aviso_var.set("")
        timer_id = None
    
    timer_id = app.after(milisegundos, limpiar)

# === Rutas y configuración ===
def obtener_ruta_absoluta(rel_path):
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

def cargar_sucursal():
    """Carga la sucursal desde parámetros del sistema de login o archivo local"""
    try:
        # Primero verificar si se pasó información desde el sistema de login
        if len(sys.argv) > 1:
            try:
                # Formato esperado: "usuario_id|username|sucursal_nombre"
                params = sys.argv[1].split('|')
                if len(params) >= 3 and params[2].strip():
                    sucursal_desde_login = params[2].strip()
                    return sucursal_desde_login
                else:
                    pass
            except Exception as e:
                pass
        
        # Fallback: usar archivo local para la sucursal
        config_path = obtener_ruta_absoluta("sucursal.txt")
        if os.path.exists(config_path):
            with open(config_path, "r", encoding="utf-8") as f:
                s = f.read().strip()
                if s:
                    return s
        
        # Último recurso: usar sucursal por defecto para pruebas
        # Cuando se ejecuta directamente sin login, no solicitar sucursal
        return "SUCURSAL PRINCIPAL"
                
    except Exception as e:
        # Fallback final al archivo local
        config_path = obtener_ruta_absoluta("sucursal.txt")
        if os.path.exists(config_path):
            with open(config_path, "r", encoding="utf-8") as f:
                s = f.read().strip()
                if s:
                    return s
        # Usar sucursal por defecto sin solicitar al usuario
        return "SUCURSAL PRINCIPAL"

def obtener_icono_path():
    if getattr(sys, 'frozen', False):
        return os.path.join(sys._MEIPASS, "icono.ico")
    else:
        return os.path.abspath("icono.ico")



# === Configuración inicial ===
CSV_PATH = obtener_ruta_absoluta("etiquetas_guardadas.csv")
IMPRESORA_CONF_PATH = obtener_ruta_absoluta("config_impresora.txt")
PERSONALIZADOS_PATH = obtener_ruta_absoluta("productos_personalizados.csv")
LOGO_PATH = obtener_ruta_absoluta("logo.png")
SUCURSAL = cargar_sucursal()

# Variables globales para información de usuario desde login integrado
USUARIO_ID = None
USUARIO_USERNAME = None
SUCURSAL_USUARIO = None
USUARIO_ROL = None

# Crear root base y ejecutar login como Toplevel para mantener un único mainloop
app = ttk.Window(themename='flatly')
try:
    app.withdraw()
except Exception:
    pass

# Ejecutar sistema de login integrado
usuario_info, sucursal_info = ejecutar_login(master=app)

# Ya estamos usando 'flatly' también en el login, no es necesario restaurar

if usuario_info:
    USUARIO_ID = str(usuario_info['id'])
    USUARIO_USERNAME = usuario_info['username']
    SUCURSAL_USUARIO = sucursal_info
    USUARIO_ROL = usuario_info['rol']
    # Sobreescribir SUCURSAL con la del usuario autenticado
    SUCURSAL = sucursal_info
else:
    try:
        app.destroy()
    except Exception:
        pass
    sys.exit(1)

def cargar_productos_personalizados():
    """Carga productos personalizados desde archivo CSV local"""
    try:
        if os.path.exists(PERSONALIZADOS_PATH):
            df = pd.read_csv(PERSONALIZADOS_PATH)
            productos = []
            for _, row in df.iterrows():
                # Convertir todos los valores a string y filtrar NaN
                codigo = str(row['codigo']) if pd.notna(row['codigo']) else ""
                nombre = str(row['nombre']) if pd.notna(row['nombre']) else ""
                base = str(row['base']) if pd.notna(row['base']) else ""
                ubicacion = str(row['ubicacion']) if pd.notna(row['ubicacion']) else ""
                
                if codigo and nombre:  # Solo agregar si tienen código y nombre válidos
                    productos.append((codigo, nombre, base, ubicacion))
            
            return productos
        return []
    except Exception as e:
        return []

# === Carga de datos ===
datos = obtener_productos_desde_db()
# Filtrar valores None, NaN y convertir a string
codigos = [str(r[0]) for r in datos if r[0] is not None]
nombres = [str(r[1]) for r in datos if r[1] is not None and str(r[1]) != 'nan']

data_por_codigo = {}
data_por_nombre = {}

for r in datos:
    if r[0] is not None and r[1] is not None and str(r[1]) != 'nan':
        codigo = str(r[0])
        nombre = str(r[1])
        base = str(r[2]) if r[2] is not None else ""
        ubicacion = str(r[3]) if r[3] is not None else ""
        
        data_por_codigo[codigo] = {"nombre": nombre, "base": base, "ubicacion": ubicacion}
        data_por_nombre[nombre] = {"codigo": codigo, "base": base, "ubicacion": ubicacion}

# Cargar productos personalizados inmediatamente y combinarlos
productos_personalizados = cargar_productos_personalizados()
for producto in productos_personalizados:
    codigo, nombre, base, ubicacion = producto
    codigo = str(codigo)
    nombre = str(nombre)
    
    if codigo not in data_por_codigo:  # Evitar duplicados
        codigos.append(codigo)
        nombres.append(nombre)
        data_por_codigo[codigo] = {"nombre": nombre, "base": base, "ubicacion": ubicacion}
        data_por_nombre[nombre] = {"codigo": codigo, "base": base, "ubicacion": ubicacion}

def recargar_productos():
    """Recarga los productos activos desde la base de datos"""
    global datos, codigos, nombres, data_por_codigo, data_por_nombre
    
    datos = obtener_productos_desde_db()
    # Filtrar valores None, NaN y convertir a string
    codigos = [str(r[0]) for r in datos if r[0] is not None]
    nombres = [str(r[1]) for r in datos if r[1] is not None and str(r[1]) != 'nan']
    
    data_por_codigo = {}
    data_por_nombre = {}
    
    for r in datos:
        if r[0] is not None and r[1] is not None and str(r[1]) != 'nan':
            codigo = str(r[0])
            nombre = str(r[1])
            base = str(r[2]) if r[2] is not None else ""
            ubicacion = str(r[3]) if r[3] is not None else ""
            
            data_por_codigo[codigo] = {"nombre": nombre, "base": base, "ubicacion": ubicacion}
            data_por_nombre[nombre] = {"codigo": codigo, "base": base, "ubicacion": ubicacion}
    
    # Cargar productos personalizados y combinarlos
    productos_personalizados = cargar_productos_personalizados()
    for producto in productos_personalizados:
        codigo, nombre, base, ubicacion = producto
        codigo = str(codigo)
        nombre = str(nombre)
        
        if codigo not in data_por_codigo:  # Evitar duplicados
            codigos.append(codigo)
            nombres.append(nombre)
            data_por_codigo[codigo] = {"nombre": nombre, "base": base, "ubicacion": ubicacion}
            data_por_nombre[nombre] = {"codigo": codigo, "base": base, "ubicacion": ubicacion}
    
    # Actualizar las listas de autocompletado con filtrado seguro solo si existen
    if 'codigo_entry' in globals() and hasattr(codigo_entry, 'lista'):
        codigo_entry.lista = sorted(set([c for c in codigos if c and str(c) != 'nan']))
    if 'descripcion_entry' in globals() and hasattr(descripcion_entry, 'lista'):
        descripcion_entry.lista = sorted(set([n for n in nombres if n and str(n) != 'nan']))

def guardar_producto_personalizado(codigo, nombre, base, ubicacion):
    """Guarda un nuevo producto personalizado"""
    try:
        # Cargar datos existentes o crear DataFrame vacío
        if os.path.exists(PERSONALIZADOS_PATH):
            df = pd.read_csv(PERSONALIZADOS_PATH)
        else:
            df = pd.DataFrame(columns=['codigo', 'nombre', 'base', 'ubicacion'])
        
        # Verificar si el código ya existe
        if codigo in df['codigo'].values:
            return False, "El código ya existe en productos personalizados"
        
        # Agregar nuevo producto
        nuevo_producto = pd.DataFrame([{
            'codigo': codigo,
            'nombre': nombre,
            'base': base,
            'ubicacion': ubicacion
        }])
        
        df = pd.concat([df, nuevo_producto], ignore_index=True)
        df.to_csv(PERSONALIZADOS_PATH, index=False)
        
        return True, "Producto personalizado guardado exitosamente"
        
    except Exception as e:
        return False, f"Error al guardar producto: {e}"

def eliminar_producto_personalizado(codigo):
    """Elimina un producto personalizado"""
    try:
        if not os.path.exists(PERSONALIZADOS_PATH):
            return False, "No hay productos personalizados"
        
        df = pd.read_csv(PERSONALIZADOS_PATH)
        if codigo not in df['codigo'].values:
            return False, "Código no encontrado en productos personalizados"
        
        df = df[df['codigo'] != codigo]
        df.to_csv(PERSONALIZADOS_PATH, index=False)
        
        return True, "Producto personalizado eliminado"
        
    except Exception as e:
        return False, f"Error al eliminar producto: {e}"

def abrir_vista_gestor():
    """Abre una vista previa del gestor con el MISMO diseño visual pero en modo solo lectura."""
    try:
        conn = psycopg2.connect(
            host="dpg-d1b18u8dl3ps73e68v1g-a.oregon-postgres.render.com",
            port=5432,
            database="labels_app_db",
            user="admin",
            password="KCFjzM4KYzSQx63ArufESIXq03EFXHz3",
            sslmode="require"
        )

        ventana = tk.Toplevel(app)
        ventana.title("PaintFlow — Vista Previa Gestor")
        ventana.geometry("1150x680")
        ventana.minsize(950, 560)
        try:
            if ICONO_PATH and os.path.exists(ICONO_PATH):
                ventana.iconbitmap(ICONO_PATH)
        except Exception:
            pass
        ventana.update_idletasks()
        vx = (ventana.winfo_screenwidth() // 2) - (1150 // 2)
        vy = (ventana.winfo_screenheight() // 2) - (680 // 2)
        ventana.geometry(f"1150x680+{vx}+{vy}")

        # ===== Estilos similares =====
        try:
            style = ttk.Style()
            style.configure("Modern.Treeview", rowheight=30, font=("Segoe UI", 11))
            style.configure("Treeview.Heading", font=("Segoe UI", 12, "bold"))
            style.map("Modern.Treeview", background=[("selected", "#1976D2")], foreground=[("selected", "white")])
        except Exception:
            pass

        # ===== Header =====
        header = ttk.Frame(ventana, style="Card.TFrame")
        header.pack(fill="x", padx=10, pady=8)
        ttk.Label(header, text="🔍 Vista Previa del Gestor (Solo Lectura)", font=("Segoe UI", 18, "bold"), style="Card.TLabel").pack(side="left", padx=8)

        sucursal_actual = obtener_sucursal_usuario(USUARIO_USERNAME) if USUARIO_USERNAME else 'principal'
        right_header = ttk.Frame(header, style="Card.TFrame")
        right_header.pack(side="right")
        ttk.Label(right_header, text=f"🏢 {sucursal_actual}", font=("Segoe UI", 13), style="Card.TLabel", bootstyle="info").pack(side="right", padx=5)
        ttk.Label(right_header, text=f"👤 {USUARIO_USERNAME}", font=("Segoe UI", 11), style="Card.TLabel", bootstyle="success").pack(side="right", padx=5)

        # ===== Barra Filtros =====
        filtros_bar = ttk.Frame(ventana, style="Card.TFrame")
        filtros_bar.pack(fill="x", padx=10, pady=4)
        ttk.Label(filtros_bar, text="Estado:", style="Card.TLabel", font=("Segoe UI", 11)).pack(side="left", padx=5)
        filtro_estado = ttk.Combobox(filtros_bar, values=["Todos", "Pendiente", "En Espera", "En Proceso", "Finalizados"], state="readonly", width=12, font=("Segoe UI", 10))
        filtro_estado.set("Todos")
        filtro_estado.pack(side="left", padx=5)
        ttk.Label(filtros_bar, text="Prioridad:", style="Card.TLabel", font=("Segoe UI", 11)).pack(side="left", padx=5)
        filtro_prioridad = ttk.Combobox(filtros_bar, values=["Todas", "Alta", "Media", "Baja"], state="readonly", width=12, font=("Segoe UI", 10))
        filtro_prioridad.set("Todas")
        filtro_prioridad.pack(side="left", padx=5)
        label_ultima = ttk.Label(filtros_bar, text="🕐 Esperando...", font=("Segoe UI", 10), style="Card.TLabel", bootstyle="secondary")
        label_ultima.pack(side="right", padx=12)
        btn_actualizar = ttk.Button(filtros_bar, text="🔄 Actualizar", bootstyle="primary")
        btn_actualizar.pack(side="left", padx=16)

        # ===== Mensaje inferior temporal =====
        label_msg = ttk.Label(ventana, text="", font=("Segoe UI", 11), style="Card.TLabel", bootstyle="info")
        label_msg.pack(pady=2)

        # ===== Tabla =====
        tree_frame = ttk.Frame(ventana, style="Card.TFrame")
        tree_frame.pack(fill="both", expand=True, padx=10, pady=6)
        tree = ttk.Treeview(tree_frame, style="Modern.Treeview")
        columnas = ("ID Prof.", "Factura", "Código", "Producto", "Terminación", "Código Base", "Prioridad", "Cantidad", "Estado", "Tiempo Est.", "Operador")
        tree["columns"] = columnas
        tree["show"] = "headings"
        anchos = {"ID Prof.": 100, "Factura": 100, "Código": 100, "Producto": 180, "Terminación": 105, "Código Base": 120, "Prioridad": 90, "Cantidad": 80, "Estado": 110, "Tiempo Est.": 105, "Operador": 120}
        anchors = {"ID Prof.": "center", "Factura": "center", "Código": "center", "Producto": "w", "Terminación": "center", "Código Base": "center", "Prioridad": "center", "Cantidad": "center", "Estado": "center", "Tiempo Est.": "center", "Operador": "w"}
        for col in columnas:
            tree.heading(col, text=col, anchor=anchors.get(col, "w"))
            tree.column(col, width=anchos.get(col, 100), anchor=anchors.get(col, "w"))
        sv = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        sh = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=sv.set, xscrollcommand=sh.set)
        tree.grid(row=0, column=0, sticky="nsew")
        sv.grid(row=0, column=1, sticky="ns")
        sh.grid(row=1, column=0, sticky="ew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        tree.tag_configure('alta', background='#ffebee')
        tree.tag_configure('media', background='#fff3e0')
        tree.tag_configure('baja', background='#e8f5e8')
        tree.tag_configure('proceso', background='#e3f2fd')
        tree.tag_configure('finalizado', background='#f3e5f5')

        # ===== Lógica de carga =====
        def cargar_datos():
            for i in tree.get_children():
                tree.delete(i)
            inicio = time.time()
            sucursal = obtener_sucursal_usuario(USUARIO_USERNAME) if USUARIO_USERNAME else 'principal'
            tabla = f"pedidos_pendientes_{sucursal}"
            try:
                cur = conn.cursor()
                query_base = f"""
                    SELECT id_orden_profesional, id_factura, codigo, producto, terminacion,
                           codigo_base, prioridad, cantidad, estado, tiempo_estimado, operador
                    FROM {tabla}
                    WHERE TRIM(COALESCE(estado,'')) <> 'Cancelado'
                      AND (estado IS NULL OR TRIM(estado) IN ('Pendiente','En Proceso'))
                """
                # Filtros
                estado_sel = filtro_estado.get()
                prioridad_sel = filtro_prioridad.get()
                filtros_extra = []
                if estado_sel != "Todos":
                    if estado_sel == "Finalizados":
                        filtros_extra.append("TRIM(estado) IN ('Finalizado','Completado')")
                    else:
                        filtros_extra.append("TRIM(estado) = %s")
                if prioridad_sel != "Todas":
                    filtros_extra.append("TRIM(prioridad) = %s")
                if filtros_extra:
                    query_base += " AND " + " AND ".join(filtros_extra)
                query_base += " ORDER BY CASE TRIM(prioridad) WHEN 'Alta' THEN 1 WHEN 'Media' THEN 2 WHEN 'Baja' THEN 3 ELSE 4 END, fecha_creacion DESC LIMIT 250"
                params = []
                if estado_sel not in ["Todos", "Finalizados"]:
                    params.append(estado_sel)
                if prioridad_sel != "Todas":
                    params.append(prioridad_sel)
                cur.execute(query_base, params)
                rows = cur.fetchall()
                if not rows:
                    tree.insert("", "end", values=("—", "—", "—", f"Sin pedidos en {tabla}", "—", "—", "—", "—", "—", "—", "—"))
                else:
                    for r in rows:
                        (id_prof, factura, codigo, producto, terminacion, codigo_base, prioridad, cantidad, estado, tiempo_est, operador) = r
                        prioridad_txt = (prioridad or '').strip()
                        tag = None
                        if prioridad_txt == 'Alta':
                            tag = 'alta'
                        elif prioridad_txt == 'Media':
                            tag = 'media'
                        elif prioridad_txt == 'Baja':
                            tag = 'baja'
                        estado_txt = (estado or 'Pendiente').strip()
                        if estado_txt in ('En Proceso','Producción'):
                            tag = 'proceso'
                        elif estado_txt in ('Finalizado','Completado'):
                            tag = 'finalizado'
                        tree.insert("", "end", values=(
                            id_prof or '—', factura or '—', codigo or '—', (producto[:40] + '…') if producto and len(producto) > 40 else (producto or '—'),
                            terminacion or '—', codigo_base or '—', prioridad_txt or '—', cantidad or 0, estado_txt or '—', tiempo_est or 0, operador or '—'
                        ), tags=(tag,) if tag else ())
                cur.close()
                dur = (time.time() - inicio)
                label_ultima.configure(text=f"✅ {len(rows)} filas • {dur:.2f}s", bootstyle="success")
                label_msg.configure(text=f"Tabla: {tabla}")
            except Exception as e:
                label_ultima.configure(text="❌ Error", bootstyle="danger")
                tree.insert("", "end", values=("ERROR", "—", "—", str(e)[:60], "—", "—", "—", "—", "—", "—", "—"))
                debug_log(f"Error cargando vista previa gestor: {e}")

        def refrescar_programado():
            try:
                cargar_datos()
            finally:
                ventana.after(15000, refrescar_programado)  # cada 15s

        btn_actualizar.configure(command=cargar_datos)
        filtro_estado.bind('<<ComboboxSelected>>', lambda e: cargar_datos())
        filtro_prioridad.bind('<<ComboboxSelected>>', lambda e: cargar_datos())

        cargar_datos()
        ventana.after(15000, refrescar_programado)

        def on_close():
            try:
                conn.close()
            except Exception:
                pass
            ventana.destroy()
        ventana.protocol("WM_DELETE_WINDOW", on_close)
    except Exception as e:
        debug_log(f"Error abriendo vista gestor (preview): {e}")
        messagebox.showerror("Error", f"No se pudo abrir vista previa del gestor:\n{e}")

def abrir_ventana_personalizados():
    """Abre la ventana para gestionar productos personalizados"""
    ventana_pers = tk.Toplevel(app)
    ventana_pers.title("PaintFlow — Productos Personalizados")
    ventana_pers.geometry("650x450")  # Aumentado de 600 a 650 para dar más espacio
    ventana_pers.resizable(False, False)
    ventana_pers.configure(bg="#f5f5f5")
    
    # Agregar ícono a la ventana
    try:
        ventana_pers.iconbitmap(ICONO_PATH)
    except Exception as e:
        pass
    
    # Centrar ventana
    ventana_pers.update_idletasks()
    x = (ventana_pers.winfo_screenwidth() // 2) - (650 // 2)  # Actualizado para nuevo ancho
    y = (ventana_pers.winfo_screenheight() // 2) - (450 // 2)
    ventana_pers.geometry(f"650x450+{x}+{y}")
    
    # Frame principal
    main_frame = ttk.Frame(ventana_pers, padding=15)  # Reducido de 20 a 15
    main_frame.pack(fill="both", expand=True)
    
    # Logo opcional
    try:
        if LOGO_PATH and os.path.exists(LOGO_PATH):
            from PIL import Image, ImageTk
            _img = Image.open(LOGO_PATH)
            _img.thumbnail((160, 60))
            ventana_pers.logo_photo = ImageTk.PhotoImage(_img)
            ttk.Label(main_frame, image=ventana_pers.logo_photo).pack(pady=(0, 6))
    except Exception:
        pass

    # Título
    ttk.Label(main_frame, text="Gestión de Productos Personalizados", 
              font=("Segoe UI", 16, "bold")).pack(pady=(0, 15))  # Reducido de 20 a 15
    
    # Frame para agregar nuevo producto
    add_frame = ttk.LabelFrame(main_frame, text="Agregar Nuevo Producto", padding=12)  # Reducido de 15 a 12
    add_frame.pack(fill="x", pady=(0, 15))  # Reducido de 20 a 15
    
    # Variables para los campos
    codigo_pers_var = tk.StringVar()
    nombre_pers_var = tk.StringVar()
    
    # Campos de entrada en una sola fila
    ttk.Label(add_frame, text="Código:").grid(row=0, column=0, sticky="w", pady=10, padx=(0, 5))
    ttk.Entry(add_frame, textvariable=codigo_pers_var, width=25).grid(row=0, column=1, pady=10, padx=5)
    
    ttk.Label(add_frame, text="Descripción:").grid(row=0, column=2, sticky="w", pady=10, padx=(15, 5))
    ttk.Entry(add_frame, textvariable=nombre_pers_var, width=35).grid(row=0, column=3, pady=10, padx=5)
    
    def agregar_producto():
        codigo = codigo_pers_var.get().strip().upper()
        nombre = nombre_pers_var.get().strip()
        base = "custom"  # Siempre 'custom' para productos personalizados
        # Calcular ubicación incremental
        if os.path.exists(PERSONALIZADOS_PATH):
            df = pd.read_csv(PERSONALIZADOS_PATH)
            ubicacion = str(len(df) + 1)
        else:
            ubicacion = "1"
        
        if not all([codigo, nombre]):
            messagebox.showwarning("Campos incompletos", "Por favor complete todos los campos")
            return
        
        # Verificar si el código ya existe en la base de datos principal
        if codigo in data_por_codigo:
            messagebox.showwarning("Código existente", "Este código ya existe en la base de datos principal")
            return
        
        exito, mensaje = guardar_producto_personalizado(codigo, nombre, base, ubicacion)
        
        if exito:
            messagebox.showinfo("Éxito", mensaje)
            # Limpiar campos
            codigo_pers_var.set("")
            nombre_pers_var.set("")
            # Actualizar lista
            actualizar_lista_personalizados()
            # Recargar productos en la aplicación principal
            recargar_productos()
        else:
            messagebox.showerror("Error", mensaje)
    
    # Botón agregar centrado
    btn_frame_form = ttk.Frame(add_frame)
    btn_frame_form.grid(row=1, column=0, columnspan=4, pady=15)
    
    ttk.Button(btn_frame_form, text="Agregar Producto", command=agregar_producto,
               style="BotonImprimir.TButton", width=20).pack()
    
    # Frame para lista de productos existentes
    list_frame = ttk.LabelFrame(main_frame, text="Productos Personalizados Existentes", padding=12)  # Reducido de 15 a 12
    list_frame.pack(fill="both", expand=True)
    
    # Treeview para mostrar productos
    columns = ("Código", "Descripción", "Base", "Ubicación")
    tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=14)  # Aumentado de 12 a 14
    
    # Ajustar el ancho de las columnas para mejor distribución
    tree.heading("Código", text="Código")
    tree.column("Código", width=100)
    
    tree.heading("Descripción", text="Descripción")
    tree.column("Descripción", width=200)
    
    tree.heading("Base", text="Base")
    tree.column("Base", width=80)
    
    tree.heading("Ubicación", text="Ubicación")
    tree.column("Ubicación", width=100)
    
    tree.pack(side="left", fill="both", expand=True)
    
    # Scrollbar para el treeview
    scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=tree.yview)
    scrollbar.pack(side="right", fill="y")
    tree.configure(yscrollcommand=scrollbar.set)
    
    def actualizar_lista_personalizados():
        """Actualiza la lista de productos personalizados"""
        for item in tree.get_children():
            tree.delete(item)
        
        productos = cargar_productos_personalizados()
        for producto in productos:
            tree.insert("", "end", values=producto)
    
    def eliminar_seleccionado():
        """Elimina el producto seleccionado"""
        seleccion = tree.selection()
        if not seleccion:
            messagebox.showwarning("Sin selección", "Por favor seleccione un producto para eliminar")
            return
        
        item = tree.item(seleccion[0])
        codigo = item['values'][0]
        
        if messagebox.askyesno("Confirmar eliminación", 
                              f"¿Está seguro de eliminar el producto {codigo}?"):
            exito, mensaje = eliminar_producto_personalizado(codigo)
            
            if exito:
                messagebox.showinfo("Éxito", mensaje)
                # Limpiar campos del formulario
                codigo_pers_var.set("")
                nombre_pers_var.set("")
                # Actualizar lista
                actualizar_lista_personalizados()
                # Recargar productos en la aplicación principal
                recargar_productos()
            else:
                messagebox.showerror("Error", mensaje)
    
    # Frame para botones
    btn_frame = ttk.Frame(list_frame)
    btn_frame.pack(fill="x", pady=(10, 0))
    
    ttk.Button(btn_frame, text="Eliminar", command=eliminar_seleccionado,
               style="BotonGrande.TButton").pack(side="left", padx=5)
    
    # Cargar lista inicial
    actualizar_lista_personalizados()

# === Autocomplete personalizado ===
class AutoCompleteEntry(tk.Entry):
    def __init__(self, master, lista, callback=None, **kwargs):
        super().__init__(master, **kwargs)
        self.lista = sorted(set(lista))
        self.callback = callback
        self.listbox = None
        self.bind('<KeyRelease>', self.check_input)
        self.bind('<Down>', self.focus_listbox)
        self.bind('<Return>', self.select_listbox)

    def check_input(self, event=None):
        txt = self.get().lower()
        if not txt:
            self.close_listbox(); return
        matches = [i for i in self.lista if txt in i.lower()]
        if matches:
            self.show_listbox(matches)
        else:
            self.close_listbox()

    def show_listbox(self, matches):
        if self.listbox: self.listbox.destroy()
        lb = tk.Listbox(self.master, height=5)
        lb.place(x=self.winfo_x(), y=self.winfo_y()+self.winfo_height())
        for m in matches: lb.insert('end', m)
        lb.bind('<<ListboxSelect>>', self.on_select)
        lb.bind('<Return>', self.on_select_keyboard)
        lb.bind('<Escape>', lambda e: self.close_listbox())
        self.listbox = lb

    def on_select(self, e=None):
        sel = self.listbox.get(self.listbox.curselection())
        self.delete(0, 'end')
        self.insert(0, sel)
        self.close_listbox()
        if self.callback:
            self.callback()
        self.focus_set()

    def on_select_keyboard(self, e=None):
        self.on_select(e)

    def focus_listbox(self, event=None):
        if self.listbox:
            self.listbox.focus()
            self.listbox.select_set(0)

    def select_listbox(self, event=None):
        if self.listbox:
            self.on_select()
        else:
            self.event_generate('<Tab>')

    def close_listbox(self):
        if self.listbox:
            self.listbox.destroy()
            self.listbox = None
            

# === ZPL + impresión ===
def generar_zpl(codigo, descripcion, producto, terminacion, presentacion, cantidad=1):
    w, h = 406, 203  # 2x1 pulgadas a 203 dpi
    
    # Ajuste dinámico de fuentes según longitud del contenido
    font_codigo = 70 if len(codigo) <= 6 else 70 if len(codigo) <= 8 else 30
    font_desc = 24 if len(descripcion) > 25 else 28
    
    # Construir texto del producto sin presentación
    producto_completo = '/'.join([x for x in [producto, terminacion] if x])
    font_producto = 20 if len(producto_completo) > 25 else 22 if len(producto_completo) > 20 else 26
    
    # Posiciones optimizadas - movido más a la derecha
    margin = 65  # Incrementado de 55 a 65
    y_cod = 25  # Bajado de 15 a 25
    y_desc = y_cod + font_codigo + 5  # Reducido de 8 a 5
    
    # Calcular posición de producto/terminación dinámicamente
    desc_lines = 1 if len(descripcion) <= 32 else 2
    y_producto = y_desc + (font_desc * desc_lines) + 12  # Reducido de 18 a 12
    
    # === Borde decorativo ===
    border_thickness = 2
    
    # === Sucursal lateral vertical optimizada ===
    sucursal_font_size = 16  # Reducido de 20 a 16
    x_sucursal = 18  # Movido de 8 a 18
    y_sucursal_start = 30
    
    # === Base/Ubicación en la parte inferior ===
    base = base_var.get() if base_var.get() else ""
    ubicacion = ubicacion_var.get() if ubicacion_var.get() else ""
    
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

    zpl = (
        "^XA\n"
        "^CI28\n"  # Codificación UTF-8
        f"^PW{w}\n^LL{h}\n^LH0,0\n"
        
        # === BORDE DECORATIVO ===
        f"^FO0,0^GB{w},{border_thickness},B^FS\n"  # Borde superior
        f"^FO0,{h-border_thickness}^GB{w},{border_thickness},B^FS\n"  # Borde inferior
        f"^FO{w-border_thickness},0^GB{border_thickness},{h},B^FS\n"  # Borde derecho
        f"^FO15,0^GB{border_thickness},{h},B^FS\n"  # Borde izquierdo movido hacia la izquierda
        
        # === LÍNEA DECORATIVA SUPERIOR ===
        f"^FO15,15^GB{w-30},1,B^FS\n"  # Línea arriba del código que toca los bordes
        
        # === CÓDIGO PRINCIPAL (Destacado y centrado) ===
        f"^CF0,{font_codigo}\n"
        f"^FO{margin},{y_cod}^FB{w-margin*2-5},1,0,C,0^FD{codigo}^FS\n"
        
        # === DESCRIPCIÓN (Centrada, máximo 2 líneas) ===
        f"^CF0,{font_desc}\n"
        f"^FO{margin},{y_desc}^FB{w-margin*2-5},{desc_lines},0,C,0^FD{descripcion}^FS\n"
        
        # === PRODUCTO/TERMINACIÓN/PRESENTACIÓN (Destacado y centrado) ===
        f"^CF0,{font_producto}\n"
        f"^FO{margin-10},{y_producto}^FB{w-margin*2+15},1,0,C,0^FD{'/'.join([x.upper() for x in [producto, terminacion, presentacion] if x])}^FS\n"
    )
    
    # === INFORMACIÓN ADICIONAL (Base/Ubicación) ===
    if info_adicional:
        zpl += (
            f"^CF0,{font_info}\n"
            f"^FO{margin},{y_info}^FB{w-margin*2-5},1,0,C,0^FD{info_adicional}^FS\n"
        )
    
    # === LÍNEA SEPARADORA ENTRE PRODUCTO Y BASE ===
    y_linea_separadora = y_producto + font_producto + 5  # Subido 5 píxeles
    zpl += f"^FO{margin+20},{y_linea_separadora}^GB{w-margin*2-50},1,B^FS\n"  # Línea más pequeña
    
    # === SUCURSAL LATERAL (Rotada 90°) ===
    if SUCURSAL:
        # Calcular el centro vertical real de la etiqueta
        centro_etiqueta = h // 2  # Centro absoluto de la etiqueta (203/2 = 101.5)
        
        # Calcular la longitud del texto para centrarlo perfectamente
        longitud_texto = len(SUCURSAL) * (sucursal_font_size * 0.5)  # Ajustado de 0.6 a 0.5
        y_inicio_centrado = centro_etiqueta - (longitud_texto // 2)  # Cambiado + por -
        
        zpl += (
            f"^A0R,{sucursal_font_size},{sucursal_font_size}\n"
            f"^FO{x_sucursal},{y_inicio_centrado}^FD{SUCURSAL.upper()}^FS\n"
        )
    
    # === LÍNEA SEPARADORA DECORATIVA ===
    y_linea = y_producto + font_producto + 8
    # Esta línea ya no es necesaria porque agregamos la línea separadora arriba
    # if not info_adicional or y_linea < y_info - 10:
    #     zpl += f"^FO{margin + 10},{y_linea}^GB{w-margin*2-55},1,B^FS\n"
    
    zpl += "^XZ\n"
    
    return zpl * int(cantidad)


# === Generar PDF ===
def generar_pdf_ficha(data, filename="ficha_pintura.pdf"):
    if not data:
        return

    codigo = data[0][0]
    base = data[0][1]

    c = canvas.Canvas(filename, pagesize=landscape(A4))
    width, height = landscape(A4)

    # Título
    c.setFont("Helvetica-Bold", 20)
    c.drawString(40, height - 40, f"Fórmula - Código: {codigo}")
    c.setFont("Helvetica", 14)
    c.drawString(40, height - 65, f"Base: {base}")

    # Encabezado de tabla
    encabezado = [
        ["COLORANTE", "CUARTOS", "", "", "", "GALONES", "", "", "", "CUBETAS", "", "", ""],
        ["", "oz", "32s", "64s", "128s", "oz", "32s", "64s", "128s", "oz", "32s", "64s", "128s"]
    ]

    # Datos organizados por colorante y tipo
    filas = {}
    for _, _, colorante, tipo, oz, _32s, _64s, _128s in data:
        if colorante not in filas:
            filas[colorante] = {
                "cuarto": ["", "", "", ""],
                "galon": ["", "", "", ""],
                "cubeta": ["", "", "", ""]
            }
        filas[colorante][tipo] = [oz, _32s, _64s, _128s]

    # Cuerpo de la tabla
    cuerpo = []
    for colorante, valores in filas.items():
        fila = [colorante]
        for tipo in ["cuarto", "galon", "cubeta"]:
            for i in range(4):
                val = valores[tipo][i]
                if val is None or str(val).lower() == "nan" or (isinstance(val, float) and math.isnan(val)):
                    fila.append("")
                else:
                    try:
                        num = float(val)
                        if num.is_integer():
                            fila.append(str(int(num)))
                        else:
                            fila.append(str(num))
                    except:
                        fila.append(str(val))
        cuerpo.append(fila)

    tabla = encabezado + cuerpo

    # Estilos
    t = Table(tabla, colWidths=[80] + [40]*12)
    t.setStyle(TableStyle([
     ("BACKGROUND", (0, 0), (-1, 0), colors.darkblue),
     ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
     ("BACKGROUND", (1, 1), (4, 1), colors.orange),
     ("BACKGROUND", (5, 1), (8, 1), colors.lightblue),
     ("BACKGROUND", (9, 1), (12, 1), colors.gold),
     ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
     ("ALIGN", (0, 0), (-1, -1), "CENTER"),
     ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
     ("FONTNAME", (0, 0), (-1, 1), "Helvetica-Bold"),
 
       # Span y centrado para CUARTOS, GALONES, CUBETAS
      ("SPAN", (1, 0), (4, 0)),
      ("SPAN", (5, 0), (8, 0)),
      ("SPAN", (9, 0), (12, 0)),
      ("ALIGN", (1, 0), (12, 0), "CENTER"),
      ("VALIGN", (1, 0), (12, 0), "MIDDLE"),
    ]))


    t.wrapOn(c, width, height)
    t.drawOn(c, 40, height - 150 - 25 * len(cuerpo))

    c.save()

def generar_pdf_tinte(data, filename="ficha_tinte.pdf"):
    if not data:
        return

    codigo = data[0][0]
    nombre_color = data[0][1]

    c = canvas.Canvas(filename, pagesize=landscape(A4))
    width, height = landscape(A4)

    # Título
    c.setFont("Helvetica-Bold", 20)
    c.drawString(40, height - 40, f"Tinte - Código: {codigo}")
    c.setFont("Helvetica", 14)
    c.drawString(40, height - 65, f"Nombre del color: {nombre_color}")

    # Orden deseado de las unidades
    orden_tipos = ["1/8", "QT", "1/2", "GALON"]

    # Construimos estructura: {colorante: {tipo: cantidad}}
    estructura = defaultdict(dict)
    tipos_encontrados = set()
    for _, _, colorante, tipo, cantidad in data:
        tipos_encontrados.add(tipo)
        try:
            num = float(cantidad)
            cantidad_str = str(int(num)) if num.is_integer() else str(num)
        except:
            cantidad_str = str(cantidad)
        estructura[colorante][tipo] = cantidad_str

    # Usar solo los tipos en el orden deseado que existan en los datos
    tipos = [t for t in orden_tipos if t in tipos_encontrados]
    colorantes = sorted(estructura.keys())

    # Encabezado: COLORANTE | 1/8 | QT | 1/2 | GALON
    encabezado = ["COLORANTE"] + tipos
    cuerpo = []

    for colorante in colorantes:
        fila = [colorante]
        for tipo in tipos:
            fila.append(estructura[colorante].get(tipo, ""))
        cuerpo.append(fila)

    tabla = [encabezado] + cuerpo

    # Tamaño de columnas dinámico   
    col_widths = [130] + [80] * (len(encabezado) - 1)

    t = Table(tabla, colWidths=col_widths)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.darkblue),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("ALIGN", (1, 1), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
    ]))

    t.wrapOn(c, width, height)
    t.drawOn(c, 40, height - 150 - 25 * len(cuerpo))

    c.save()

def generar_pdf_por_cada_tinte():
    try:
        conn = psycopg2.connect(
            host="dpg-d1b18u8dl3ps73e68v1g-a.oregon-postgres.render.com",
            port=5432,
            database="labels_app_db",
            user="admin",
            password="KCFjzM4KYzSQx63ArufESIXq03EFXHz3",
            sslmode="require"
        )
        with conn.cursor() as cur:
            cur.execute("SELECT id FROM tintes;")
            ids = cur.fetchall()

        for (tinte_id,) in ids:
            data = obtener_datos_por_tinte(tinte_id)
            generar_pdf_tinte(data, filename=f"tinte_{tinte_id}.pdf")

    except Exception as e:
        pass


def imprimir_zebra_zpl(zpl_code):
    if not WIN32_AVAILABLE:
        messagebox.showerror("Error de impresión",
                           "Módulos de impresión no disponibles.\n"
                           "Instala pywin32: pip install pywin32")
        return

    try:
        pr = printer_var.get()
        guardar_impresora(pr)

        # Verificar lista de impresoras disponibles
        try:
            available = [p[2] for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
        except Exception:
            available = []

        if not pr:
            messagebox.showwarning("Impresora no seleccionada", "Por favor selecciona una impresora antes de imprimir.")
            return

        if available and pr not in available:
            messagebox.showwarning("Impresora no encontrada", f"La impresora seleccionada ('{pr}') no está entre las impresoras detectadas.\nLista detectada: {available}")

        # Intentar enviar ZPL en RAW
        h = win32print.OpenPrinter(pr)
        try:
            win32print.StartDocPrinter(h, 1, ("Etiqueta", None, "RAW"))
            win32print.StartPagePrinter(h)
            win32print.WritePrinter(h, zpl_code.encode())
            win32print.EndPagePrinter(h)
            win32print.EndDocPrinter(h)
        finally:
            try:
                win32print.ClosePrinter(h)
            except:
                pass
    except Exception as e:
        # Mostrar error detallado y ofrecer escribir ZPL a archivo como fallback
        try:
            # Guardar ZPL en archivo temporal para envío manual
            import tempfile
            tmp = tempfile.mktemp(suffix='.zpl')
            with open(tmp, 'w', encoding='utf-8') as f:
                f.write(zpl_code)
            messagebox.showerror("Error impresión", f"No se pudo imprimir:\n{e}\nSe guardó el ZPL en: {tmp}")
        except Exception as e2:
            messagebox.showerror("Error impresión", f"No se pudo imprimir:\n{e}\nAdemás, no se pudo crear archivo de fallback: {e2}")

def guardar_impresora(nombre):
    try:
        with open(IMPRESORA_CONF_PATH,'w',encoding='utf-8') as f:
            f.write(nombre)
    except: pass

def cargar_impresora_guardada():
    if not WIN32_AVAILABLE:
        return ''
        
    if os.path.exists(IMPRESORA_CONF_PATH):
        try:
            n = open(IMPRESORA_CONF_PATH,'r',encoding='utf-8').read().strip()
            printers = [p[2] for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL|win32print.PRINTER_ENUM_CONNECTIONS)]
            if n in printers:
                return n
        except: pass
    return ''
def on_btn_imprimir_click():
    codigo = codigo_entry.get()
    if codigo:
        imprimir_ficha_pintura(codigo)
    else:
        messagebox.showinfo("Campo vacío", "Por favor ingrese un código de pintura.")


# === UI ===
# Reutilizar el root existente creado antes del login (app)
app.title(f"PaintFlow {APP_VERSION} - {SUCURSAL}")
app.geometry("900x540")  # Ventana ligeramente más compacta
app.resizable(False, False)

# Configurar icono de la aplicación
ICONO_PATH = obtener_icono_path()
try:
    app.iconbitmap(ICONO_PATH)
except Exception as e:
    pass
    
aviso_var = tk.StringVar()

printer_var = tk.StringVar(value=cargar_impresora_guardada())


# Obtener lista de impresoras con manejo de errores
if WIN32_AVAILABLE:
    try:
        printers = [p[2] for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
    except Exception as e:
        printers = []
else:
    printers = ["Sin impresoras disponibles (instalar pywin32)"]

if not printers:
    messagebox.showwarning("Sin impresoras", "No se detectaron impresoras")

# Variables

descripcion_var = tk.StringVar()
producto_var = tk.StringVar()
terminacion_var = tk.StringVar()
presentacion_var = tk.StringVar()
spin = tk.IntVar(value=1)
base_var = tk.StringVar()
ubicacion_var = tk.StringVar()
codigo_base_var = tk.StringVar()

# Lista temporal de productos para factura múltiple
lista_productos_factura = []

# Actualiza la vista previa al cambiar producto o terminación
producto_var.trace_add('write', lambda *args: actualizar_vista())
terminacion_var.trace_add('write', lambda *args: actualizar_vista())

# Diccionario de terminaciones válidas por producto
TERMINACIONES_POR_PRODUCTO = {
    'laca': ['Mate', 'Semimate', 'Brillo'],
    'esmalte multiuso': ['Mate', 'Satin', 'Gloss'],
    'excello premium': ['Mate', 'Satin', 'Semigloss', 'Semisatin'],
    'excello voc': ['Mate', 'Satin'],
    'master paint': ['Mate'],
    'tinte al thinner': ['Claro', 'Intermedio', 'Especial'],
    'super paint': ['Mate', 'Satin', 'Gloss'],
    'esmalte kem': ['Mate', 'Semimate', 'Brillo'],
    'excello pastel': ['Mate'],
    'water blocking': ['Mate'],
    'kem aqua': ['Satin'],
    'emerald': ['Satin', 'SemiGloss'],
    'monocapa': ['Mate', 'Semimate', 'Brillo'],
    'uretano': ['Mate', 'Semimate', 'Brillo'],
    'airpuretec': ['Mate', 'Satin'],
    'kem pro': ['Mate'],
    'sanitizing': ['Satin'],
    'scuff tuff-wb' : ['Mate', 'Satin',],
    'armoseal t-p' : ['Semigloss'],
    'armoseal 1000hs' : ['Gloss'],
    'pro industrial dtm' : ['Gloss'],
    'promar® 400 voc' : ['Satin'],
    'h&c heavy-shield' : ['Gloss'],
    'h&c silicone-acrylic' : ['Mate'],
    'promar® 200 voc' : ['Satin', 'Mate'],

    
    
}

def actualizar_terminaciones(*args):
    """Actualiza las terminaciones disponibles según el producto seleccionado"""
    producto = producto_var.get().lower()
    base = base_var.get().lower()
    
    # Buscar terminaciones válidas para el producto
    terminaciones_validas = []
    for key, terminaciones in TERMINACIONES_POR_PRODUCTO.items():
        if key in producto:
            terminaciones_validas = terminaciones
            break
    
    # Si no se encuentra el producto, usar todas las terminaciones
    if not terminaciones_validas:
        terminaciones_validas = ['Mate', 'Satin', 'Semigloss', 'Semimate', 'Gloss', 'Brillo', 
                                "N/A", "ESPECIAL", "CLARO", "INTERMEDIO", "MADERA", "PERLADO", "METALICO", "SEMISATIN"]
    
    # Lógica especial para Excello Premium con bases Ultra Deep / Ultra Deep II
    if 'excello premium' in producto:
        # Detectar Ultra Deep y Ultra Deep II (distintas variantes)
        es_ultra_deep_ii = any(k in base for k in ['ultra deep ii', 'ultradeep ii', 'ultra-deep ii', 'ultra deep 2'])
        es_ultra_deep = ('ultra deep' in base)
        if es_ultra_deep or es_ultra_deep_ii:
            # Solo permitir Semisatin para ambas bases
            terminaciones_validas = ['Semisatin']

    # Actualizar el combobox
    terminaciones_combobox['values'] = terminaciones_validas
    
    # Limpiar la selección actual si no es válida
    terminacion_actual = terminacion_var.get()
    if terminacion_actual and terminacion_actual not in terminaciones_validas:
        terminacion_var.set('')
    
    # Si solo hay una terminación válida, seleccionarla automáticamente
    if len(terminaciones_validas) == 1:
        terminacion_var.set(terminaciones_validas[0])
    
    # Actualizar vista previa
    actualizar_vista()

def actualizar_presentaciones(*args):
    """Actualiza las presentaciones disponibles según el producto seleccionado"""
    producto = producto_var.get().lower()
    
    # Presentaciones por defecto (incluye Medio Galón)
    presentaciones_disponibles = ['Cuarto', 'Medio Galón', 'Galón', 'Cubeta']
    
    # Si es laca, agregar octavos (1/8)
    if 'laca' in producto:
        presentaciones_disponibles = ['1/8', 'Cuarto', 'Medio Galón', 'Galón']

    # Regla solicitada: para Esmalte Kem quitar 'Cubeta' y permitir '1/8'
    if 'esmalte kem' in producto:
        presentaciones_disponibles = ['1/8', 'Cuarto', 'Medio Galón', 'Galón']
    
    # Actualizar el combobox
    presentacion_combobox['values'] = presentaciones_disponibles
    
    # Limpiar la selección actual si no es válida
    presentacion_actual = presentacion_var.get()
    if presentacion_actual and presentacion_actual not in presentaciones_disponibles:
        presentacion_var.set('')

# Actualiza la vista previa al cambiar producto o terminación
producto_var.trace_add('write', actualizar_terminaciones)
producto_var.trace_add('write', actualizar_presentaciones)
terminacion_var.trace_add('write', lambda *args: actualizar_vista())
presentacion_var.trace_add('write', lambda *args: actualizar_vista())
# Agregar listener para cuando cambie la base (importante para Excello Premium)
base_var.trace_add('write', actualizar_terminaciones)

# Layout
labels = ["Código", "Descripción", "Producto", "Terminación", "Presentación", "Ubicación", "Base", "Codigo Base", "Cantidad"]
for i, l in enumerate(labels):
    ttk.Label(app, text=f"{l}:").place(x=30, y=20 + i * 50)

codigo_entry = AutoCompleteEntry(app, sorted(set([c for c in codigos if c and str(c) != 'nan'])), callback=lambda: completar_datos())
codigo_entry.place(x=150, y=20, width=200)

descripcion_entry = AutoCompleteEntry(app, sorted(set([n for n in nombres if n and str(n) != 'nan'])), callback=lambda: completar_datos(), textvariable=descripcion_var)
descripcion_entry.place(x=150, y=70, width=200)

# Guarda las referencias de los combobox
producto_combobox = ttk.Combobox(app, textvariable=producto_var,
    values=['Excello Premium','Laca', 'Esmalte Kem', "Excello VOC","Master Paint","Tinte al Thinner", 'Super Paint', 'Esmalte Multiuso','Excello Pastel', 'Water Blocking', 'Kem Aqua', 'Emerald', 'Monocapa', 'Uretano', 'Airpuretec', 'Kem Pro', 'Sanitizing', 'h&c silicone-acrylic', 'h&c heavy-shield', 'promar® 200 voc', 'promar® 400 voc', 'pro industrial dtm', 'armoseal 1000hs','armoseal t-p', 'scuff tuff-wb' ],
    state='readonly')
producto_combobox.place(x=150, y=120, width=200)

terminaciones_combobox = ttk.Combobox(app, textvariable=terminacion_var,
    values=['Mate', 'Satin', 'Semigloss', 'Semimate', 'Gloss', 'Brillo', "N/A", "ESPECIAL", "CLARO", "INTERMEDIO", "ESPECIAL", "MADERA", "PERLADO", "METALICO"],
    state='readonly')
terminaciones_combobox.place(x=150, y=170, width=200)

# Combobox de presentación
presentacion_combobox = ttk.Combobox(app, textvariable=presentacion_var,
    values=['1/8', 'Cuarto', 'Medio Galón', 'Galón', 'Cubeta'],
    state='readonly')
presentacion_combobox.place(x=150, y=220, width=200)

# Eliminado el selector de impresora (se conserva la lectura/guardado silencioso)


# Función para soporte de teclado en combobox
def combobox_keydown(event, combobox):
    if event.keysym == 'Down':
        combobox.event_generate('<Button-1>')
    elif event.keysym == 'Return':
        combobox.event_generate('<Tab>')

# Bind a todos los combobox
producto_combobox.bind('<Key>', lambda e: combobox_keydown(e, producto_combobox))
terminaciones_combobox.bind('<Key>', lambda e: combobox_keydown(e, terminaciones_combobox))
# Eliminado el bind del selector de impresora

ttk.Entry(app, textvariable=ubicacion_var, state='readonly').place(x=150, y=270, width=200)
ttk.Entry(app, textvariable=base_var, state='readonly').place(x=150, y=320, width=200)
ttk.Entry(app, textvariable=codigo_base_var, state='readonly').place(x=150, y=370, width=200)
ttk.Spinbox(app, from_=1, to=100, textvariable=spin).place(x=150, y=420, width=60)

# Botón minimalista para ver gestor (debajo de cantidad, más pequeño)
btn_gestor = ttk.Button(app, text="👁️", command=abrir_vista_gestor, 
                       width=3)
btn_gestor.place(x=220, y=420)

# Tooltip para el botón gestor
def crear_tooltip(widget, texto):
    def mostrar_tooltip(event):
        tooltip = tk.Toplevel()
        tooltip.wm_overrideredirect(True)
        tooltip.wm_geometry(f"+{event.x_root+10}+{event.y_root+10}")
        tooltip.configure(bg="#ffffcc")
        ttk.Label(tooltip, text=texto, background="#ffffcc", 
                 font=("Segoe UI", 8)).pack()
        widget.tooltip_window = tooltip
    
    def ocultar_tooltip(event):
        if hasattr(widget, 'tooltip_window'):
            widget.tooltip_window.destroy()
            del widget.tooltip_window
    
    widget.bind('<Enter>', mostrar_tooltip)
    widget.bind('<Leave>', ocultar_tooltip)

crear_tooltip(btn_gestor, "Ver estado del gestor")





# Vista Previa dentro de un LabelFrame igual que Acciones
frame_vista = ttk.LabelFrame(app, text="Vista Previa", padding=14, style="Acciones.TLabelframe")
frame_vista.place(x=420, y=15, width=440, height=260)
# Canvas para dibujar una vista previa que imita la ZPL del gestor (406x203)
vista_canvas = tk.Canvas(frame_vista, width=406, height=203, bg="#ffffff", highlightthickness=0)
vista_canvas.pack(anchor='center')
_vista_imgs_cache = {}

# Label para avisos debajo de la vista previa
aviso_label = ttk.Label(app, textvariable=aviso_var, font=('Segoe UI', 10), foreground="#1976d2", background="#fff")
aviso_label.place(x=500, y=270, width=380)  # Subido más para dejar espacio al cuadro de acciones

codigo_base_actual = ""

def actualizar_vista():
    """Dibuja una vista previa alineada a la etiqueta del gestor (layout ZPL simulado)."""
    # Limpiar canvas
    vista_canvas.delete("all")

    # Medidas de etiqueta ZPL
    w, h = 406, 203
    margin = 65
    border = 2

    # Datos actuales
    c = codigo_entry.get().strip().upper()
    d = descripcion_var.get().strip()
    p = (producto_var.get() or '').strip()
    t = (terminacion_var.get() or '').strip()
    pr = (presentacion_var.get() or '').strip()
    b = (base_var.get() or '').strip()
    u = (ubicacion_var.get() or '').strip()

    # Regla del gestor: mostrar PRODUCTO/TERMINACIÓN (sin presentación)
    producto_linea = '/'.join([x for x in [p, t] if x])

    # Productos que no deben mostrar la base (coincide con lógica ZPL)
    productos_sin_base = ['laca', 'uretano', 'esmalte kem', 'esmalte multiuso', 'monocapa']
    mostrar_base = not any(prod.lower() in p.lower() for prod in productos_sin_base)
    info_adicional = f"{b} | {u}" if mostrar_base and b and u else (b if mostrar_base else (u or ""))

    # Cálculo de líneas de descripción (máx 2)
    desc_lines = 1 if len(d) <= 32 else 2
    desc1 = d[:32]
    desc2 = d[32:64] if desc_lines == 2 else ''

    # Tipografías aproximadas
    from tkinter import font as tkfont
    font_codigo = tkfont.Font(family='Consolas', size=28, weight='bold')
    font_desc = tkfont.Font(family='Segoe UI', size=14, weight='bold')
    font_prod = tkfont.Font(family='Segoe UI', size=16, weight='bold')
    font_info = tkfont.Font(family='Segoe UI', size=12)
    font_sucursal = tkfont.Font(family='Segoe UI', size=12, weight='bold')

    # Fondo y bordes
    vista_canvas.create_rectangle(0, 0, w, h, outline='#000000', width=border)
    # Línea decorativa superior interna
    vista_canvas.create_line(15, 15, w-15, 15, fill='#000000')

    # Código principal (centrado)
    y_cod = 25
    vista_canvas.create_text(w//2, y_cod, text=c or '', font=font_codigo, fill='#000000', anchor='n')

    # Descripción (1-2 líneas centradas)
    y_desc = y_cod + 28 + 5
    vista_canvas.create_text(w//2, y_desc, text=desc1, font=font_desc, fill='#000000', anchor='n')
    if desc2:
        vista_canvas.create_text(w//2, y_desc + font_desc.metrics('linespace'), text=desc2, font=font_desc, fill='#000000', anchor='n')

    # Producto/Terminación (centrado, sin presentación)
    y_prod = y_desc + (font_desc.metrics('linespace') * desc_lines) + 12
    vista_canvas.create_text(w//2, y_prod, text=(producto_linea.upper()), font=font_prod, fill='#000000', anchor='n')

    # Info adicional (Base/Ubicación) centrado al pie si aplica
    if info_adicional:
        y_info = h - 25
        vista_canvas.create_text(w//2, y_info, text=info_adicional, font=font_info, fill='#000000', anchor='s')
        # Línea separadora sobre info
        vista_canvas.create_line(margin+20, y_prod + font_prod.metrics('linespace') + 5, w - (margin+20), y_prod + font_prod.metrics('linespace') + 5, fill='#000000')

    # Sin texto lateral de sucursal para una vista previa más limpia


def completar_datos():
    c = codigo_entry.get().strip()
    d = descripcion_entry.get().strip()

    if c in data_por_codigo:
        info = data_por_codigo[c]
        descripcion_var.set(info["nombre"])
        base_var.set(info["base"])
        ubicacion_var.set(info["ubicacion"])
    elif d in data_por_nombre:
        info = data_por_nombre[d]
        codigo_entry.delete(0, 'end')
        codigo_entry.insert(0, info["codigo"])
        descripcion_var.set(d)
        base_var.set(info["base"])
        ubicacion_var.set(info["ubicacion"])
    
    # Limpiar código base para que solo aparezca al presionar el botón
    codigo_base_var.set("")
    
    # Actualizar terminaciones después de completar datos
    actualizar_terminaciones()
    actualizar_vista()

def mostrar_ventana_factura():
    """Muestra ventana emergente para ingresar ID de factura y prioridad"""
    ventana_factura = tk.Toplevel()
    ventana_factura.title("PaintFlow — Información de Factura")
    ventana_factura.geometry("400x300")
    ventana_factura.resizable(False, False)
    ventana_factura.grab_set()  # Hacer modal
    
    # Centrar la ventana respecto a la ventana principal
    ventana_factura.transient(app)
    # Icono
    try:
        if ICONO_PATH and os.path.exists(ICONO_PATH):
            ventana_factura.iconbitmap(ICONO_PATH)
    except Exception:
        pass
    
    # Centrar la ventana en pantalla
    x = (ventana_factura.winfo_screenwidth() // 2) - (400 // 2)
    y = (ventana_factura.winfo_screenheight() // 2) - (250 // 2)
    ventana_factura.geometry(f"400x250+{x}+{y}")
    
    # Variables para almacenar los valores
    id_factura_var = tk.StringVar()
    prioridad_var = tk.StringVar(value="Media")
    resultado = {"continuar": False, "id_factura": "", "prioridad": ""}
    
    # Logo (opcional)
    try:
        if LOGO_PATH and os.path.exists(LOGO_PATH):
            from PIL import Image, ImageTk
            _img = Image.open(LOGO_PATH)
            _img.thumbnail((150, 56))
            ventana_factura.logo_photo = ImageTk.PhotoImage(_img)
            ttk.Label(ventana_factura, image=ventana_factura.logo_photo).pack(pady=(12, 0))
    except Exception:
        pass

    # Título
    ttk.Label(ventana_factura, text="Información del Pedido", 
             font=("Segoe UI", 14, "bold")).pack(pady=15)
    
    # Frame para los campos
    frame_campos = ttk.Frame(ventana_factura)
    frame_campos.pack(pady=10, padx=20, fill="x")
    
    # Campo ID Factura
    ttk.Label(frame_campos, text="ID Factura:", font=("Segoe UI", 10)).grid(row=0, column=0, sticky="w", pady=5)
    entry_factura = ttk.Entry(frame_campos, textvariable=id_factura_var, font=("Segoe UI", 10), width=25)
    entry_factura.grid(row=0, column=1, pady=5, padx=(10, 0))
    entry_factura.focus()
    
    # Campo Prioridad
    ttk.Label(frame_campos, text="Prioridad:", font=("Segoe UI", 10)).grid(row=1, column=0, sticky="w", pady=5)
    combo_prioridad = ttk.Combobox(frame_campos, textvariable=prioridad_var, 
                                  values=["Alta", "Media", "Baja"], 
                                  state="readonly", font=("Segoe UI", 10), width=22)
    combo_prioridad.grid(row=1, column=1, pady=5, padx=(10, 0))
    
    # Frame para botones
    frame_botones = ttk.Frame(ventana_factura)
    frame_botones.pack(pady=20)
    
    def aceptar():
        if not id_factura_var.get().strip():
            messagebox.showwarning("Campo requerido", "Debe ingresar el ID de la factura.")
            entry_factura.focus()
            return
        
        resultado["continuar"] = True
        resultado["id_factura"] = id_factura_var.get().strip()
        resultado["prioridad"] = prioridad_var.get()
        ventana_factura.destroy()
    
    def cancelar():
        resultado["continuar"] = False
        ventana_factura.destroy()
    
    # Botones
    ttk.Button(frame_botones, text="Aceptar", command=aceptar, 
              bootstyle="success").pack(side="left", padx=10)
    ttk.Button(frame_botones, text="Cancelar", command=cancelar, 
              bootstyle="secondary").pack(side="left", padx=10)
    
    # Bind Enter para aceptar
    ventana_factura.bind('<Return>', lambda e: aceptar())
    # Bind Escape para cancelar
    ventana_factura.bind('<Escape>', lambda e: cancelar())
    
    # Esperar hasta que se cierre la ventana
    ventana_factura.wait_window()
    
    return resultado

def imprimir_guardar():
    c = codigo_entry.get()
    d = descripcion_var.get()
    p = producto_var.get()
    t = terminacion_var.get()
    pr = presentacion_var.get()
    q = spin.get()
    
    if not c:
        return  # Salir silenciosamente
    
    # VALIDACIÓN OBLIGATORIA: Presentación debe estar seleccionada
    if not pr:
        messagebox.showwarning("Presentación requerida", 
                              "Debe seleccionar una presentación antes de enviar el producto.\n\n" +
                              "Las presentaciones disponibles son:\n" +
                              "• 1/8 (para lacas)\n" +
                              "• Cuarto\n" +
                              "• Medio Galón\n" +
                              "• Galón\n" +
                              "• Cubeta")
        return
    
    # Mostrar ventana para ID factura y prioridad
    datos_factura = mostrar_ventana_factura()
    
    # Si el usuario canceló, no continuar
    if not datos_factura["continuar"]:
        return
    
    # Asegurar ID de factura único en la cola de la sucursal, con sufijo incremental si ya existe
    usuario_para_deteccion = USUARIO_USERNAME if 'USUARIO_USERNAME' in globals() and USUARIO_USERNAME else USUARIO_ID
    id_factura_input = datos_factura["id_factura"]
    try:
        id_factura = generar_id_factura_unico(id_factura_input, usuario_para_deteccion)
    except Exception:
        # Si algo falla, usar el ingresado por el usuario
        id_factura = id_factura_input
    prioridad = datos_factura["prioridad"]
    
    # Guardar en CSV (para compatibilidad)
    reg = {'Codigo': c, 'Descripcion': d, 'Producto': p, 'Terminacion': t, 'Presentacion': pr, 'ID_Factura': id_factura, 'Prioridad': prioridad}
    df = pd.read_csv(CSV_PATH) if os.path.exists(CSV_PATH) else pd.DataFrame()
    df = pd.concat([df, pd.DataFrame([reg])], ignore_index=True)
    df.to_csv(CSV_PATH, index=False)
    
    # Limpiar campos inmediatamente para velocidad
    limpiar_campos()
    
    # Operaciones de BD en paralelo sin bloquear interfaz
    import threading
    def operaciones_bd():
        try:
            # Registrar en base de datos para estadísticas
            registrar_impresion(c, d, p, t, pr, q, SUCURSAL, USUARIO_ID, id_factura, prioridad)
            # Agregar a lista de espera con presentación correcta
            agregar_a_lista_espera(c, p, t, id_factura, prioridad, q, base=base_var.get(), presentacion=pr, ubicacion=ubicacion_var.get())
        except:
            pass  # Operación silenciosa
    
    threading.Thread(target=operaciones_bd, daemon=True).start()

def generar_id_factura_unico(id_factura: str, usuario_id: str = None) -> str:
    """Devuelve un ID de factura único para la tabla de pedidos pendientes de la sucursal.
    Si ya existe ese id_factura, agrega un sufijo incremental "-2", "-3", ... hasta encontrar uno libre.
    No modifica registros existentes; sólo propone un nombre disponible.
    """
    try:
        base = str(id_factura).strip()
        if not base:
            return id_factura

        # Quitar sufijo numérico final si ya fue provisto (p.ej., FAC-123-2 -> FAC-123)
        try:
            import re
            m = re.match(r"^(.*?)(?:-(\d+))?$", base)
            base_name = m.group(1).strip() if m else base
        except Exception:
            base_name = base

        usuario_para_deteccion = usuario_id or (USUARIO_USERNAME if 'USUARIO_USERNAME' in globals() and USUARIO_USERNAME else USUARIO_ID)
        sucursal = obtener_sucursal_usuario(usuario_para_deteccion)
        tabla_pendientes = f"pedidos_pendientes_{sucursal}"

        def existe_factura(nombre: str) -> bool:
            try:
                conn = psycopg2.connect(
                    host="dpg-d1b18u8dl3ps73e68v1g-a.oregon-postgres.render.com",
                    port=5432,
                    database="labels_app_db",
                    user="admin",
                    password="KCFjzM4KYzSQx63ArufESIXq03EFXHz3",
                    sslmode="require"
                )
                cur = conn.cursor()
                cur.execute(f"SELECT 1 FROM {tabla_pendientes} WHERE id_factura = %s LIMIT 1", (nombre,))
                row = cur.fetchone()
                cur.close(); conn.close()
                return row is not None
            except Exception:
                # Si no podemos verificar, consideramos que no existe para no bloquear el flujo
                return False

        # Si el nombre base no existe, usarlo tal cual
        if not existe_factura(base_name):
            return base_name

        # Buscar el siguiente sufijo disponible
        for n in range(2, 501):  # límite razonable
            candidato = f"{base_name}-{n}"
            if not existe_factura(candidato):
                return candidato

        # Fallback si excede el límite
        return f"{base_name}-{int(time.time()) % 1000}"
    except Exception:
        return id_factura

def imprimir_pdf(path_pdf):
    try:
        if WIN32_AVAILABLE:
            # Intentar imprimir usando ShellExecute (funciona con visores asociados)
            try:
                # Preferir ShellExecute desde win32api
                win32api.ShellExecute(0, "print", path_pdf, None, ".", 0)
                return
            except Exception:
                # Fallback a os.startfile con verbo 'print'
                try:
                    os.startfile(path_pdf, 'print')
                    return
                except Exception as e:
                    messagebox.showerror("Error impresión PDF", f"No se pudo imprimir el PDF: {e}")
                    return
        else:
            # Abrir el PDF si no hay soporte Win32 (solo mostrar)
            try:
                os.startfile(path_pdf)
                return
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo abrir el PDF: {e}")
                return
    except Exception as e:
        messagebox.showerror("Error impresión", f"Error al intentar imprimir/abrir el PDF: {e}")

def limpiar_campos():
    global codigo_base_actual
    codigo_entry.delete(0, 'end')
    descripcion_var.set('')
    producto_var.set('')
    terminacion_var.set('')
    presentacion_var.set('')
    spin.set(1)
    base_var.set('')
    ubicacion_var.set('')
    vista_canvas.delete('all')
    codigo_base_actual = ''
    codigo_base_var.set('')
    
def calcular_tiempo_compromiso(producto, cantidad):
    """Calcula el tiempo de compromiso basado en el tipo de producto y cantidad"""
    producto_lower = producto.lower()
    
    # Productos que requieren 10 minutos por unidad
    productos_10_min = ['laca', 'esmalte', 'industrial']
    
    # Verificar si el producto es de 10 minutos
    tiempo_por_unidad = 10 if any(tipo in producto_lower for tipo in productos_10_min) else 6
    
    return tiempo_por_unidad * cantidad

def obtener_tiempo_acumulado_sucursal(sucursal):
    """Obtiene el tiempo acumulado de todas las órdenes pendientes y en proceso de una sucursal"""
    try:
        conn = psycopg2.connect(
            host="dpg-d1b18u8dl3ps73e68v1g-a.oregon-postgres.render.com",
            port=5432,
            database="labels_app_db",
            user="admin",
            password="KCFjzM4KYzSQx63ArufESIXq03EFXHz3",
            sslmode="require"
        )
        cur = conn.cursor()
        
        # Obtener la suma de tiempos estimados de órdenes pendientes y en proceso
        cur.execute("""
            SELECT COALESCE(SUM(tiempo_estimado), 0)
            FROM lista_espera 
            WHERE sucursal = %s AND estado IN ('Pendiente', 'En Proceso')
        """, (sucursal,))
        
        tiempo_acumulado = cur.fetchone()[0]
        cur.close()
        conn.close()
        return tiempo_acumulado
        
    except Exception as e:
        print(f"Error al obtener tiempo acumulado: {e}")
        return 0

def calcular_fecha_compromiso(tiempo_total_minutos):
    """Calcula la fecha de compromiso basada en el tiempo total en minutos"""
    from datetime import datetime, timedelta
    
    # Obtener fecha y hora actual
    ahora = datetime.now()
    
    # Agregar los minutos al tiempo actual
    fecha_compromiso = ahora + timedelta(minutes=tiempo_total_minutos)
    
    return fecha_compromiso

def generar_id_profesional(id_factura=None):
    """Genera un ID profesional basado en factura: todos los productos de la misma factura tendrán el mismo ID"""
    from datetime import datetime
    import random
    import time
    
    # Si no hay ID de factura, usar método anterior como fallback
    if not id_factura:
        # Obtener fecha actual
        ahora = datetime.now()
        
        # Iniciales de días de la semana
        dias_semana = {
            0: 'L',  # Lunes
            1: 'M',  # Martes  
            2: 'X',  # Miércoles
            3: 'J',  # Jueves
            4: 'V',  # Viernes
            5: 'S',  # Sábado
            6: 'D'   # Domingo
        }
        
        inicial_dia = dias_semana[ahora.weekday()]
    else:
        # Usar ID de factura como base
        # Extraer los últimos 4 dígitos de la factura
        import re
        numeros = re.findall(r'\d+', str(id_factura))
        if numeros:
            factura_num = int(numeros[-1])  # Tomar el último número encontrado
            # Usar F + últimos 4 dígitos de la factura
            inicial_dia = 'F'
            # Si el número es muy grande, tomar solo los últimos 4 dígitos
            base_numero = factura_num % 10000
        else:
            # Fallback si no hay números en la factura
            inicial_dia = 'F'
            base_numero = abs(hash(str(id_factura))) % 10000
    
    try:
        # Conectar a base de datos
        conn = psycopg2.connect(
            host="dpg-d1b18u8dl3ps73e68v1g-a.oregon-postgres.render.com",
            port=5432,
            database="labels_app_db",
            user="admin",
            password="KCFjzM4KYzSQx63ArufESIXq03EFXHz3",
            sslmode="require"
        )
        cur = conn.cursor()
        
        # Obtener todas las tablas de pedidos pendientes
        cur.execute("""
            SELECT table_name FROM information_schema.tables 
            WHERE table_name LIKE 'pedidos_pendientes_%' AND table_schema = 'public'
        """)
        tablas = [row[0] for row in cur.fetchall()]
        
        # También incluir lista_espera si existe
        cur.execute("""
            SELECT table_name FROM information_schema.tables 
            WHERE table_name = 'lista_espera' AND table_schema = 'public'
        """)
        if cur.fetchone():
            tablas.append('lista_espera')
        
        # Si tenemos ID de factura, generar ID secuencial basado en la factura
        if id_factura:
            # Obtener el prefijo base para esta factura
            id_base = f"F{base_numero:03d}"
            
            # Buscar cuántos productos ya existen para esta factura
            total_productos_factura = 0
            for tabla in tablas:
                cur.execute(f"""
                    SELECT COUNT(*) FROM {tabla}
                    WHERE id_factura = %s
                """, (id_factura,))
                
                count = cur.fetchone()[0]
                total_productos_factura += count
            
            # Generar el siguiente número secuencial para esta factura
            siguiente_secuencia = total_productos_factura + 1
            
            # Intentar generar IDs secuenciales hasta encontrar uno libre
            max_intentos = 10
            for intento in range(max_intentos):
                # Formato: F001A, F001B, F001C... (base + letra secuencial)
                if siguiente_secuencia <= 26:
                    sufijo = chr(64 + siguiente_secuencia)  # A=65, B=66, etc.
                else:
                    # Si hay más de 26 productos, usar números
                    sufijo = str(siguiente_secuencia - 26).zfill(2)
                
                id_propuesto = f"{id_base}{sufijo}"
                
                # Verificar que no exceda 5 caracteres
                if len(id_propuesto) > 5:
                    # Reducir el número base si es muy largo
                    base_reducido = base_numero % 100  # Solo 2 dígitos
                    id_base = f"F{base_reducido:02d}"
                    id_propuesto = f"{id_base}{sufijo}"
                
                # Verificar si ya existe este ID
                existe = False
                for tabla in tablas:
                    cur.execute(f"""
                        SELECT COUNT(*) FROM {tabla}
                        WHERE id_orden_profesional = %s
                    """, (id_propuesto,))
                    
                    count = cur.fetchone()[0]
                    if count > 0:
                        existe = True
                        break
                
                if not existe:
                    cur.close()
                    conn.close()
                    print(f"✅ Nuevo ID generado para factura {id_factura}: {id_propuesto}")
                    return id_propuesto
                
                # Ya existe, probar con el siguiente
                siguiente_secuencia += 1
        
        # Fallback: usar el método anterior (sin factura o si hay conflicto)
        patron_busqueda = f"{inicial_dia}%"
        ultimo_numero_encontrado = 0
        
        # Buscar el mayor número usado en todas las tablas
        for tabla in tablas:
            cur.execute(f"""
                SELECT id_orden_profesional FROM {tabla}
                WHERE id_orden_profesional LIKE %s
                AND LENGTH(id_orden_profesional) <= 5
                ORDER BY CAST(SUBSTRING(id_orden_profesional FROM '\\d+') AS INTEGER) DESC 
                LIMIT 1
            """, (patron_busqueda,))
            
            resultado = cur.fetchone()
            if resultado and resultado[0]:
                try:
                    import re
                    numeros = re.findall(r'\d+', resultado[0])
                    if numeros:
                        numero = int(numeros[0])
                        ultimo_numero_encontrado = max(ultimo_numero_encontrado, numero)
                except (ValueError, IndexError):
                    pass
        
        if ultimo_numero_encontrado > 0:
            siguiente_numero = ultimo_numero_encontrado + 1
        else:
            siguiente_numero = 1001  # Empezar desde 1001 para tener 4 dígitos
        
        # Intentar generar un ID único ultra corto (máximo 5 caracteres)
        max_intentos = 50
        for intento in range(max_intentos):
            # Formato: L1234 (1 letra + 4 números = 5 caracteres máximo)
            id_propuesto = f"{inicial_dia}{siguiente_numero}"
            
            # Verificar que no exceda 5 caracteres
            if len(id_propuesto) > 5:
                siguiente_numero = 1001  # Reset si se hace muy largo
                id_propuesto = f"{inicial_dia}{siguiente_numero}"
            
            # Verificar si ya existe en cualquier tabla
            existe = False
            for tabla in tablas:
                cur.execute(f"""
                    SELECT COUNT(*) FROM {tabla}
                    WHERE id_orden_profesional = %s
                """, (id_propuesto,))
                
                count = cur.fetchone()[0]
                if count > 0:
                    existe = True
                    break
            
            if not existe:
                # No existe, podemos usarlo
                cur.close()
                conn.close()
                return id_propuesto
            
            # Ya existe, incrementar
            siguiente_numero += 1
        
        # Si llegamos aquí, usar timestamp como fallback ultra corto
        cur.close()
        conn.close()
        timestamp = int(time.time())
        return f"{inicial_dia}{timestamp % 9999}"
        
    except Exception as e:
        debug_log(f"Error generando ID profesional: {e}")
        # Fallback con timestamp ultra corto (máximo 5 caracteres)
        timestamp = int(time.time())
        return f"{inicial_dia}{timestamp % 9999}"

def agregar_a_lista_espera(codigo, producto, terminacion, id_factura, prioridad, cantidad, base=None, presentacion=None, ubicacion=None):
    """Agrega el pedido a la tabla específica de la sucursal del usuario"""
    try:
        # Detectar sucursal automáticamente basándose en el USERNAME (no ID numérico)
        usuario_para_deteccion = USUARIO_USERNAME if 'USUARIO_USERNAME' in globals() and USUARIO_USERNAME else USUARIO_ID
        sucursal = obtener_sucursal_usuario(usuario_para_deteccion)
        tabla_pendientes = f'pedidos_pendientes_{sucursal}'

        # Normalizar prioridad
        prioridad_limpia = (prioridad or 'Media').strip().title()
        if prioridad_limpia not in ['Alta', 'Media', 'Baja']:
            prioridad_limpia = 'Media'

        debug_log(f"🏢 Usuario: {usuario_para_deteccion} → Sucursal: {sucursal} → Tabla: {tabla_pendientes} → Prioridad: {prioridad_limpia}")

        # Preparar datos previos (indentación corregida)
        base_a_usar = (base if base not in [None, ""] else base_var.get())
        # Si viene vacía desde el diccionario del producto, tomar la selección actual de la UI
        _pres_tmp = presentacion if presentacion not in [None, ""] else presentacion_var.get()
        presentacion_a_usar = (_pres_tmp or "").strip()
        ubicacion_a_usar = (ubicacion if ubicacion not in [None, ""] else ubicacion_var.get())
        terminacion_a_usar = (terminacion or "").strip()

        # Calcular código base completo (con sufijo presentación si aplica)
        codigo_base_calculado = ""
        if base_a_usar and producto and terminacion:
            codigo_base_calculado = obtener_codigo_base(base_a_usar, producto, terminacion)
            if presentacion_a_usar and codigo_base_calculado not in ["No encontrado", "No Aplica"]:
                sufijo = obtener_sufijo_presentacion(presentacion_a_usar)
                if sufijo:
                    codigo_base_calculado += sufijo

        # Generar ID profesional único usando la función actualizada (basado en factura)
        id_profesional = generar_id_profesional(id_factura)

        # Calcular tiempo de compromiso
        tiempo_esta_orden = calcular_tiempo_compromiso(producto, cantidad)
        _ = calcular_fecha_compromiso(tiempo_esta_orden)

        # INTENTO CON REINTENTOS Y VERIFICACIÓN
        backoffs = [0.15, 0.4, 0.9]
        last_error = None

        for intento, espera in enumerate(backoffs, start=1):
            conn = None
            try:
                conn = psycopg2.connect(
                    host="dpg-d1b18u8dl3ps73e68v1g-a.oregon-postgres.render.com",
                    port=5432,
                    database="labels_app_db",
                    user="admin",
                    password="KCFjzM4KYzSQx63ArufESIXq03EFXHz3",
                    sslmode="require"
                )
                cur = conn.cursor()

                # Descubrir columnas disponibles de la tabla (antes de consultas)
                cur.execute("""
                    SELECT column_name FROM information_schema.columns
                    WHERE table_schema='public' AND table_name=%s
                """, (tabla_pendientes,))
                cols_disponibles = {r[0] for r in cur.fetchall()}

                # verificar si ya existe misma línea según columnas disponibles
                where_parts = ["id_factura = %s", "codigo = %s"]
                where_params = [id_factura, codigo]
                if 'presentacion' in cols_disponibles:
                    where_parts.append("TRIM(COALESCE(presentacion,'')) = %s")
                    where_params.append(presentacion_a_usar)
                if 'terminacion' in cols_disponibles:
                    where_parts.append("TRIM(COALESCE(terminacion,'')) = %s")
                    where_params.append(terminacion_a_usar)
                where_sql = " AND ".join(where_parts)
                cur.execute(f"SELECT id FROM {tabla_pendientes} WHERE {where_sql}", tuple(where_params))
                row = cur.fetchone()
                if row:
                    # Ya existe una línea con mismo código+factura+presentación
                    # En lugar de descartar, acumulamos cantidad y tiempo estimado
                    try:
                        # Sumar cantidad y actualizar prioridad si es superior
                        # Además, si existe 'estado' y la línea estaba Finalizada/Completada/Cancelada, reabrir a 'Pendiente'.
                        set_partes = [
                            "cantidad = COALESCE(cantidad,0) + %s",
                            "tiempo_estimado = COALESCE(tiempo_estimado,0) + %s",
                            "prioridad = CASE WHEN %s = 'Alta' OR prioridad = 'Alta' THEN 'Alta' WHEN %s = 'Media' AND prioridad = 'Baja' THEN 'Media' ELSE prioridad END"
                        ]
                        params_update = [cantidad, tiempo_esta_orden, prioridad_limpia, prioridad_limpia]
                        if 'estado' in cols_disponibles:
                            set_partes.append("estado = CASE WHEN TRIM(COALESCE(estado,'')) IN ('Finalizado','Completado','Cancelado') THEN 'Pendiente' ELSE estado END")
                        if 'fecha_asignacion' in cols_disponibles:
                            set_partes.append("fecha_asignacion = CASE WHEN TRIM(COALESCE(estado,'')) IN ('Finalizado','Completado','Cancelado') THEN NULL ELSE fecha_asignacion END")
                        if 'fecha_completado' in cols_disponibles:
                            set_partes.append("fecha_completado = CASE WHEN TRIM(COALESCE(estado,'')) IN ('Finalizado','Completado','Cancelado') THEN NULL ELSE fecha_completado END")
                        sql_update = f"UPDATE {tabla_pendientes} SET " + ", ".join(set_partes) + f" WHERE {where_sql}"
                        cur.execute(sql_update, tuple(params_update + where_params))
                        conn.commit()
                        cur.close(); conn.close()
                        debug_log(f"✅ Acumulado en {tabla_pendientes}: {codigo} x+{cantidad} ({presentacion_a_usar or ''})")
                        return True
                    except Exception as e2:
                        try:
                            conn.rollback()
                        except Exception:
                            pass
                        # Si falla el UPDATE por cualquier razón, seguimos al flujo de inserción como fallback
                        debug_log(f"⚠️ Falló acumulación; intentaremos insertar nueva línea. Motivo: {e2}")

                # cols_disponibles ya obtenido arriba

                # columnas base (siempre esperadas)
                cols = [
                    'id_orden_profesional','codigo','producto','terminacion','id_factura',
                    'prioridad','cantidad','tiempo_estimado','base','ubicacion','sucursal'
                ]
                vals = [
                    id_profesional, codigo, producto, terminacion_a_usar, id_factura,
                    prioridad_limpia, cantidad, tiempo_esta_orden, base_a_usar, ubicacion_a_usar, sucursal.title()
                ]

                # opcionales
                if 'presentacion' in cols_disponibles:
                    cols.append('presentacion'); vals.append(presentacion_a_usar)
                if 'codigo_base' in cols_disponibles:
                    cols.append('codigo_base'); vals.append(codigo_base_calculado)
                if 'estado' in cols_disponibles:
                    cols.append('estado'); vals.append('Pendiente')

                placeholders = ", ".join(["%s"] * len(cols))
                sql = f"INSERT INTO {tabla_pendientes} (" + ", ".join(cols) + ") VALUES (" + placeholders + ")"

                cur.execute(sql, tuple(vals))
                conn.commit()

                # verificación inmediata de inserción
                # verificación inmediata de inserción
                where_verify = ["id_orden_profesional = %s"]
                params_verify = [id_profesional]
                where_dup = ["id_factura=%s", "codigo=%s"]
                params_dup = [id_factura, codigo]
                if 'presentacion' in cols_disponibles:
                    where_dup.append("TRIM(COALESCE(presentacion,''))=%s")
                    params_dup.append(presentacion_a_usar)
                if 'terminacion' in cols_disponibles:
                    where_dup.append("TRIM(COALESCE(terminacion,''))=%s")
                    params_dup.append(terminacion_a_usar)
                verify_sql = f"SELECT id_orden_profesional FROM {tabla_pendientes} WHERE (" + " AND ".join(where_verify) + ") OR (" + " AND ".join(where_dup) + ") LIMIT 1"
                cur.execute(verify_sql, tuple(params_verify + params_dup))
                ok = cur.fetchone()
                cur.close()
                conn.close()
                if ok:
                    debug_log(f"✅ Insert verificado en {tabla_pendientes}: {id_profesional} [{prioridad_limpia}]")
                    return True
                else:
                    last_error = Exception("Verificación post-inserción no encontró el registro")
                    debug_log(f"⚠️ No se encontró tras inserción; reintento {intento}")
                    time.sleep(espera)
                    continue

            except Exception as e:
                last_error = e
                try:
                    if conn:
                        conn.rollback()
                        conn.close()
                except Exception:
                    pass
                debug_log(f"❌ Error intento {intento} insertando en {tabla_pendientes}: {e}")
                time.sleep(espera)
                continue

        # Si llegamos aquí, todos los intentos fallaron
        debug_log(f"🚫 Falló el envío tras varios intentos para factura {id_factura}, código {codigo} ({prioridad_limpia}). Último error: {last_error}")
        return False

    except Exception as e:
        debug_log(f"❌ Error general en agregar_a_lista_espera: {e}")
        return False

def registrar_impresion(codigo, descripcion, producto, terminacion, presentacion, cantidad, sucursal, usuario_id='sistema', id_factura='', prioridad='Media'):
    """Registra una impresión en la base de datos para estadísticas.
    Inserta únicamente las columnas que existan en la tabla para evitar fallos por esquema.
    """
    try:
        conn = psycopg2.connect(
            host="dpg-d1b18u8dl3ps73e68v1g-a.oregon-postgres.render.com",
            port=5432,
            database="labels_app_db",
            user="admin",
            password="KCFjzM4KYzSQx63ArufESIXq03EFXHz3",
            sslmode="require"
        )
        cur = conn.cursor()

        # Detectar columnas disponibles en la tabla
        cur.execute("""
            SELECT column_name
            FROM information_schema.columns
            WHERE table_schema = 'public' AND table_name = 'historial_impresiones'
        """)
        cols_db = {r[0] for r in cur.fetchall()}

        # Mapeo de posibles columnas
        valores = {
            'codigo': codigo,
            'descripcion': descripcion,
            'producto': producto,
            'terminacion': terminacion,
            'presentacion': presentacion,
            'cantidad': cantidad,
            'sucursal': sucursal,
            'usuario_id': usuario_id,
            'prioridad': prioridad,
            # Soportar ambas variantes para factura
            'id_factura': id_factura,
            'factura': id_factura,
        }

        columnas_insert = [c for c in valores.keys() if c in cols_db]
        if not columnas_insert:
            cur.close(); conn.close()
            return False
        valores_insert = [valores[c] for c in columnas_insert]

        placeholders = ', '.join(['%s'] * len(columnas_insert))
        sql = f"INSERT INTO historial_impresiones ({', '.join(columnas_insert)}) VALUES ({placeholders});"
        cur.execute(sql, valores_insert)
        conn.commit()
        cur.close(); conn.close()
        return True
    except Exception:
        try:
            conn.rollback()
        except Exception:
            pass
        try:
            cur.close(); conn.close()
        except Exception:
            pass
        return False

# === HISTORIAL DE ENVÍOS ===
# Mantener referencia única para evitar múltiples ventanas abiertas
ventana_historial = None

def asegurar_esquema_historial():
    """Garantiza que la tabla historial_impresiones tenga columnas básicas para filtros.
    - id_factura: para filtrar por factura
    - created_at: para limitar a últimas 24h
    """
    try:
        conn = psycopg2.connect(
            host="dpg-d1b18u8dl3ps73e68v1g-a.oregon-postgres.render.com",
            port=5432,
            database="labels_app_db",
            user="admin",
            password="KCFjzM4KYzSQx63ArufESIXq03EFXHz3",
            sslmode="require"
        )
        cur = conn.cursor()
        # Agregar columnas si no existen (no rompe si ya están)
        cur.execute("""
            ALTER TABLE IF EXISTS historial_impresiones
            ADD COLUMN IF NOT EXISTS id_factura TEXT,
            ADD COLUMN IF NOT EXISTS created_at TIMESTAMPTZ DEFAULT NOW();
        """)
        conn.commit()
        cur.close(); conn.close()
    except Exception:
        # Silencioso: no interrumpir UI si el esquema no se puede ajustar
        pass
# Mantener referencia única para evitar múltiples ventanas abiertas
ventana_historial = None
def obtener_historial_impresiones(limit=100, sucursal=None, factura=None, solo_ultimas_24h=True):
    """Obtiene los últimos envíos desde la tabla del Gestor (lista_espera).
    Devuelve (columnas, filas) con una única columna de fecha como 'fecha' y 'factura' desde id_factura.
    """
    try:
        conn = psycopg2.connect(
            host="dpg-d1b18u8dl3ps73e68v1g-a.oregon-postgres.render.com",
            port=5432,
            database="labels_app_db",
            user="admin",
            password="KCFjzM4KYzSQx63ArufESIXq03EFXHz3",
            sslmode="require"
        )
        cur = conn.cursor()

        # Descubrir columnas disponibles en lista_espera
        cur.execute(
            """
            SELECT column_name
            FROM information_schema.columns
            WHERE table_schema = 'public' AND table_name = 'lista_espera'
            """
        )
        cols_db = [r[0] for r in cur.fetchall()]
        if not cols_db:
            cur.close(); conn.close()
            return [], []

        # Preferencia de columna de fecha (una sola para evitar duplicados)
        preferencia_fechas = ['fecha_creacion', 'fecha_asignacion', 'fecha_completado', 'fecha']
        col_fecha = next((c for c in preferencia_fechas if c in cols_db), None)

        # Partes del SELECT con alias estables
        select_parts = []
        columnas = []
        if col_fecha:
            select_parts.append(f"{col_fecha} AS fecha")
            columnas.append('fecha')

        def add_if(colname, alias=None):
            if colname in cols_db:
                select_parts.append(f"{colname} AS {alias or colname}")
                columnas.append(alias or colname)

        add_if('id_factura', alias='factura')
        add_if('codigo')
        add_if('producto')
        add_if('terminacion')
        add_if('presentacion')
        add_if('cantidad')
        add_if('prioridad')
        add_if('operador')
        add_if('sucursal')

        if not select_parts:
            cur.close(); conn.close()
            return [], []

        base_q = f"SELECT {', '.join(select_parts)} FROM lista_espera"
        conds = []
        params = []

        if sucursal and 'sucursal' in cols_db:
            conds.append("TRIM(LOWER(sucursal)) = TRIM(LOWER(%s))")
            params.append(sucursal)

        if factura and 'id_factura' in cols_db:
            conds.append("id_factura = %s")
            params.append(factura)

        if solo_ultimas_24h and col_fecha:
            conds.append(f"{col_fecha} >= NOW() - INTERVAL '24 HOURS'")
        elif col_fecha:
            # Cuando no es sólo 24h, limitar a últimos 7 días para evitar consultas pesadas
            conds.append(f"{col_fecha} >= NOW() - INTERVAL '7 DAYS'")

        if conds:
            base_q += " WHERE " + " AND ".join(conds)

        if col_fecha:
            base_q += f" ORDER BY {col_fecha} DESC"
        else:
            base_q += " ORDER BY 1 DESC"

        base_q += " LIMIT %s"
        params.append(limit)

        cur.execute(base_q, params)
        filas = cur.fetchall()
        cur.close(); conn.close()
        return columnas, filas
    except Exception as e:
        debug_log(f"❌ Error obteniendo historial desde lista_espera: {e}")
        return [], []


def abrir_historial_impresiones():
    """Abre una ventana modal con un listado de las últimas impresiones enviadas."""
    global ventana_historial
    try:
        # Si ya existe, traer al frente y enfocar
        if ventana_historial is not None:
            try:
                if ventana_historial.winfo_exists():
                    ventana_historial.deiconify()
                    ventana_historial.lift()
                    ventana_historial.focus_force()
                    return
            except Exception:
                pass

        ventana = tk.Toplevel(app)
        ventana.title("PaintFlow — Historial de Envíos")
        ventana.geometry("1000x600")
        ventana.resizable(True, True)
        ventana.grab_set()
        ventana.transient(app)
        # Icono de la ventana
        try:
            if ICONO_PATH and os.path.exists(ICONO_PATH):
                ventana.iconbitmap(ICONO_PATH)
        except Exception:
            pass

        # Maximizar a pantalla completa (Windows)
        try:
            ventana.state('zoomed')
        except Exception:
            # Fallback: usar tamaño de pantalla
            try:
                sw = ventana.winfo_screenwidth(); sh = ventana.winfo_screenheight()
                ventana.geometry(f"{sw}x{sh}+0+0")
            except Exception:
                pass

        # Logo (opcional) y encabezado
        try:
            if LOGO_PATH and os.path.exists(LOGO_PATH):
                from PIL import Image, ImageTk
                _img = Image.open(LOGO_PATH)
                # Redimensionar manteniendo aspecto
                _img.thumbnail((160, 60))
                ventana.logo_photo = ImageTk.PhotoImage(_img)
                ttk.Label(ventana, image=ventana.logo_photo).pack(pady=(8,0))
        except Exception:
            pass

        # Encabezado
        header_label = ttk.Label(ventana, text=f"Últimos envíos (últimas 24 horas) - {SUCURSAL}", font=("Segoe UI", 14, "bold"))
        header_label.pack(pady=8)

        # Barra superior con búsqueda y filtro por factura (el historial está limitado a 24 horas)
        topbar = ttk.Frame(ventana)
        topbar.pack(fill='x', padx=10, pady=4)
        ttk.Label(topbar, text="Buscar:").pack(side='left')
        filtro_var = tk.StringVar()
        entry_filtro = ttk.Entry(topbar, textvariable=filtro_var, width=30)
        entry_filtro.pack(side='left', padx=6)
    # Filtro por factura
        ttk.Label(topbar, text="Factura:").pack(side='left', padx=(12,4))
        factura_var = tk.StringVar()
        entry_factura = ttk.Entry(topbar, textvariable=factura_var, width=14)
        entry_factura.pack(side='left')

        # Toggle: Solo 24h
        solo24_var = tk.BooleanVar(value=True)
        chk_24h = ttk.Checkbutton(topbar, text="Sólo 24h", variable=solo24_var)
        chk_24h.pack(side='left', padx=(12,4))

        # Botón Exportar CSV (a la derecha)
        def exportar_csv():
            try:
                import csv, os
                from tkinter import filedialog
                # Encabezados visibles
                headers = [alias.get(c, c).title() for c in columnas]
                # Filas visibles en la tabla (respetar filtro actual)
                items = tree.get_children()
                filas_exp = [tree.item(i, 'values') for i in items]
                if not filas_exp:
                    messagebox.showinfo('Exportar CSV', 'No hay filas para exportar.')
                    return
                filename = filedialog.asksaveasfilename(
                    title='Guardar historial como CSV',
                    defaultextension='.csv',
                    filetypes=[('CSV', '*.csv')],
                    initialfile='historial.csv'
                )
                if not filename:
                    return
                with open(filename, 'w', newline='', encoding='utf-8') as f:
                    writer = csv.writer(f)
                    writer.writerow(headers)
                    writer.writerows(filas_exp)
                messagebox.showinfo('Exportar CSV', f'Se exportó correctamente a:\n{filename}')
            except Exception as e:
                messagebox.showerror('Exportar CSV', f'Error al exportar: {e}')

        btn_exportar = ttk.Button(topbar, text='Exportar CSV', command=exportar_csv)
        btn_exportar.pack(side='right', padx=(6,0))
        btn_refrescar = ttk.Button(topbar, text="Aplicar")
        btn_refrescar.pack(side='right', padx=(6,0))

        # Tabla
        frame_tabla = ttk.Frame(ventana)
        frame_tabla.pack(fill='both', expand=True, padx=10, pady=8)
        tree = ttk.Treeview(frame_tabla, show='headings')
        tree.pack(fill='both', expand=True, side='left')
        scroll_y = ttk.Scrollbar(frame_tabla, orient='vertical', command=tree.yview)
        scroll_y.pack(side='right', fill='y')
        tree.configure(yscrollcommand=scroll_y.set)

        # Datos (estado en memoria para filtrar)
        columnas, filas = obtener_historial_impresiones(limit=200, sucursal=SUCURSAL, solo_ultimas_24h=solo24_var.get())
        datos = []

        # Configurar columnas
        if not columnas:
            ttk.Label(frame_tabla, text="No hay datos de historial disponibles.").pack()
            return

        tree["columns"] = columnas
        alias = {
            'created_at': 'Fecha', 'fecha': 'Fecha', 'fecha_impresion': 'Fecha', 'timestamp': 'Fecha',
            'id': 'ID', 'codigo': 'Código', 'descripcion': 'Descripción',
            'producto': 'Producto', 'terminacion': 'Terminación', 'presentacion': 'Presentación',
            'cantidad': 'Cant.', 'id_factura': 'Factura', 'factura': 'Factura', 'prioridad': 'Prioridad',
            'sucursal': 'Sucursal', 'usuario_id': 'Operador', 'operador': 'Operador'
        }
        for c in columnas:
            tree.heading(c, text=alias.get(c, c).title())
            # Anchos aproximados
            ancho = 120
            if c in ('descripcion',): ancho = 250
            if c in ('codigo','producto','terminacion','presentacion'): ancho = 120
            if c in ('cantidad',): ancho = 70
            if c in ('id_factura','factura','prioridad','usuario_id','operador'): ancho = 120
            if c in ('created_at','fecha','timestamp','sucursal'): ancho = 140
            tree.column(c, width=ancho, anchor='center')

        # Normalizar celdas a string
        def normalizar(val):
            try:
                import datetime
                if isinstance(val, (datetime.date, datetime.datetime)):
                    return val.strftime("%Y-%m-%d %H:%M")
            except Exception:
                pass
            return "" if val is None else str(val)

        def cargar_tabla(filas_origen):
            nonlocal datos
            tree.delete(*tree.get_children())
            datos = [[normalizar(v) for v in fila] for fila in filas_origen]
            for fila in datos:
                tree.insert('', 'end', values=fila)

        cargar_tabla(filas)

        # Filtrado básico por texto
        def aplicar_filtro(*_):
            txt = filtro_var.get().strip().lower()
            if not txt:
                cargar_tabla(filas)
                return
            filtradas = []
            for fila in filas:
                if any((str(v).lower().find(txt) >= 0) for v in fila):
                    filtradas.append(fila)
            cargar_tabla(filtradas)

        entry_filtro.bind('<KeyRelease>', aplicar_filtro)

        # Refrescar desde BD
        def refrescar():
            nonlocal columnas, filas
            filtros = {
                'limit': 200,
                'sucursal': SUCURSAL,
                'factura': factura_var.get().strip() or None,
                'solo_ultimas_24h': bool(solo24_var.get()),
            }
            columnas, filas = obtener_historial_impresiones(**filtros)
            # Reconfigurar columnas si cambiaron
            if columnas and list(tree["columns"]) != columnas:
                tree["columns"] = columnas
                for c in columnas:
                    tree.heading(c, text=alias.get(c, c).title())
            cargar_tabla(filas)

            # Actualizar encabezado
            try:
                suffix = " (últimas 24 horas)" if solo24_var.get() else " (últimos 7 días)"
                header_label.configure(text=f"Últimos envíos{suffix} - {SUCURSAL}")
            except Exception:
                pass

        btn_refrescar.configure(command=refrescar)
        # Activar toggle 24h
        try:
            chk_24h.configure(command=refrescar)
        except Exception:
            pass

        # Al cerrar, limpiar referencia
        def on_close():
            global ventana_historial
            try:
                ventana.destroy()
            finally:
                ventana_historial = None

        ventana.protocol("WM_DELETE_WINDOW", on_close)

        # Guardar referencia global
        ventana_historial = ventana

    except Exception as e:
        try:
            messagebox.showerror("Historial", f"No se pudo abrir el historial: {e}")
        except Exception:
            pass

# === FUNCIONES PARA LISTA DE PRODUCTOS DE FACTURA ===

def agregar_producto_a_lista():
    """Agrega el producto actual a la lista temporal de la factura"""
    global lista_productos_factura
    
    # Validar campos necesarios
    c = codigo_entry.get().strip()
    d = descripcion_var.get().strip()
    p = producto_var.get().strip()
    t = terminacion_var.get().strip()
    pr = presentacion_var.get().strip()
    q = spin.get()
    
    if not c:
        messagebox.showwarning("Campo requerido", "Debe ingresar un código.")
        return
    
    if not p:
        messagebox.showwarning("Campo requerido", "Debe seleccionar un producto.")
        return
    
    if not pr:
        messagebox.showwarning("Presentación requerida", 
                              "Debe seleccionar una presentación antes de agregar el producto a la lista.\n\n" +
                              "Las presentaciones disponibles son:\n" +
                              "• 1/8 (para lacas)\n" +
                              "• Cuarto\n" +
                              "• Medio Galón\n" +
                              "• Galón\n" +
                              "• Cubeta")
        return
    
    # Crear producto para la lista
    producto_item = {
        'codigo': c,
        'descripcion': d,
        'producto': p,
        'terminacion': t,
        'presentacion': pr,
        'cantidad': q,
        'base': base_var.get(),
        'ubicacion': ubicacion_var.get()
    }
    
    # Agregar a la lista directamente (sin validación de duplicados)
    
    # Agregar a la lista
    lista_productos_factura.append(producto_item)
    
    # Mostrar confirmación
    messagebox.showinfo("Producto agregado", f"Código {c} agregado a la lista.\nTotal productos: {len(lista_productos_factura)}")
    
    # Limpiar campos para el siguiente producto
    limpiar_campos()

def abrir_gestor_lista_factura():
    """Abre la ventana para gestionar la lista de productos de la factura"""
    global lista_productos_factura
    
    if not lista_productos_factura:
        messagebox.showinfo("Lista vacía", "No hay productos en la lista.\nAgregue productos usando 'Agregar a Lista'.")
        return
    
    # Crear ventana
    ventana_lista = tk.Toplevel(app)
    ventana_lista.title("PaintFlow — Gestionar Lista de Factura")
    # Ampliar ventana para que quepan botones cómodamente
    ventana_lista.geometry("1040x660")
    ventana_lista.resizable(True, True)
    ventana_lista.iconbitmap(ICONO_PATH)
    ventana_lista.grab_set()
    ventana_lista.transient(app)
    
    # Centrar ventana
    x = (ventana_lista.winfo_screenwidth() // 2) - (1040 // 2)
    y = (ventana_lista.winfo_screenheight() // 2) - (660 // 2)
    ventana_lista.geometry(f"1040x660+{x}+{y}")
    
    # Logo opcional
    try:
        if LOGO_PATH and os.path.exists(LOGO_PATH):
            from PIL import Image, ImageTk
            _img = Image.open(LOGO_PATH)
            _img.thumbnail((180, 68))
            ventana_lista.logo_photo = ImageTk.PhotoImage(_img)
            ttk.Label(ventana_lista, image=ventana_lista.logo_photo).pack(pady=(8, 2))
    except Exception:
        pass

    # Título
    ttk.Label(ventana_lista, text="Lista de Productos para Factura", 
             font=("Segoe UI", 14, "bold")).pack(pady=10)
    
    # Frame para información de factura
    frame_factura = ttk.LabelFrame(ventana_lista, text="Información de Factura", padding=10)
    frame_factura.pack(fill="x", padx=20, pady=5)
    
    # Variables para factura
    id_factura_var = tk.StringVar()
    prioridad_var = tk.StringVar(value="Media")
    
    # Campos de factura
    ttk.Label(frame_factura, text="ID Factura:").grid(row=0, column=0, sticky="w", padx=5)
    entry_factura = ttk.Entry(frame_factura, textvariable=id_factura_var, width=25)
    entry_factura.grid(row=0, column=1, padx=5, sticky="w")
    
    ttk.Label(frame_factura, text="Prioridad:").grid(row=0, column=2, sticky="w", padx=5)
    combo_prioridad = ttk.Combobox(frame_factura, textvariable=prioridad_var, 
                                  values=["Alta", "Media", "Baja"], state="readonly", width=15)
    combo_prioridad.grid(row=0, column=3, padx=5, sticky="w")
    
    # Frame para lista de productos
    frame_lista = ttk.LabelFrame(ventana_lista, text="Productos en la Lista", padding=10)
    frame_lista.pack(fill="both", expand=True, padx=20, pady=10)
    
    # Treeview para mostrar productos
    columns = ("Código", "Descripción", "Producto", "Terminación", "Presentación", "Cantidad")
    tree_productos = ttk.Treeview(frame_lista, columns=columns, show="headings", height=12)
    
    # Configurar columnas
    for col in columns:
        tree_productos.heading(col, text=col)
        if col == "Código":
            tree_productos.column(col, width=100)
        elif col == "Descripción":
            tree_productos.column(col, width=200)
        elif col == "Cantidad":
            tree_productos.column(col, width=80)
        else:
            tree_productos.column(col, width=120)
    
    # Scrollbar
    scrollbar = ttk.Scrollbar(frame_lista, orient="vertical", command=tree_productos.yview)
    tree_productos.configure(yscrollcommand=scrollbar.set)
    
    tree_productos.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")
    
    def cargar_productos_en_tree():
        """Carga los productos de la lista en el treeview"""
        for item in tree_productos.get_children():
            tree_productos.delete(item)
        
        for producto in lista_productos_factura:
            tree_productos.insert("", "end", values=(
                producto['codigo'],
                producto['descripcion'],
                producto['producto'],
                producto['terminacion'],
                producto['presentacion'],
                producto['cantidad']
            ))
    
    def eliminar_producto():
        """Elimina el producto seleccionado de la lista"""
        global lista_productos_factura
        selection = tree_productos.selection()
        if not selection:
            messagebox.showwarning("Selección", "Seleccione un producto para eliminar.")
            return
        
        item = tree_productos.item(selection[0])
        vals = item['values']
        codigo_sel = vals[0]
        desc_sel = vals[1]
        prod_sel = vals[2]
        term_sel = vals[3]
        pres_sel = vals[4]
        cant_sel = vals[5]
        
        # Confirmar eliminación
        if messagebox.askyesno("Confirmar", f"¿Eliminar el producto {codigo_sel} ({pres_sel})?"):
            # Eliminar solo la primera coincidencia exacta (permitiendo duplicados restantes)
            idx_to_remove = None
            for i, p in enumerate(lista_productos_factura):
                if (
                    p.get('codigo') == codigo_sel and
                    p.get('presentacion') == pres_sel and
                    p.get('terminacion') == term_sel and
                    p.get('descripcion') == desc_sel and
                    p.get('producto') == prod_sel and
                    str(p.get('cantidad')) == str(cant_sel)
                ):
                    idx_to_remove = i
                    break
            if idx_to_remove is None:
                # Fallback: coincidir por código + presentación + terminación
                for i, p in enumerate(lista_productos_factura):
                    if (
                        p.get('codigo') == codigo_sel and
                        p.get('presentacion') == pres_sel and
                        p.get('terminacion') == term_sel
                    ):
                        idx_to_remove = i
                        break
            if idx_to_remove is not None:
                del lista_productos_factura[idx_to_remove]
            # Recargar tree
            cargar_productos_en_tree()
            messagebox.showinfo("Eliminado", f"Producto {codigo_sel} ({pres_sel}) eliminado.")
    
    def enviar_todos_a_lista_espera():
        """Envía todos los productos a la lista de espera"""
        global lista_productos_factura
        
        # Validar ID de factura silenciosamente
        id_factura = id_factura_var.get().strip()
        if not id_factura:
            return  # Salir silenciosamente
        
        prioridad = prioridad_var.get()
        
        # Envío directo sin confirmación para velocidad
        # Tomar un snapshot inmutable de la lista ANTES de limpiarla para evitar carrera
        items_a_enviar = [p.copy() for p in lista_productos_factura]
        
        # Operaciones en paralelo para máxima velocidad (sobre el snapshot)
        import threading
        def procesar_lista_bd(items):
            try:
                for producto in items:
                    try:
                        # Registrar impresión (para estadísticas)
                        registrar_impresion(
                            producto['codigo'], producto['descripcion'], producto['producto'],
                            producto['terminacion'], producto['presentacion'], producto['cantidad'],
                            SUCURSAL, USUARIO_ID, id_factura, prioridad
                        )
                        
                        # Agregar a lista de espera con todos los datos necesarios
                        agregar_a_lista_espera(
                            producto['codigo'], producto['producto'], producto['terminacion'],
                            id_factura, prioridad, producto['cantidad'],
                            base=producto['base'], 
                            presentacion=producto['presentacion'],
                            ubicacion=producto['ubicacion']
                        )
                    except:
                        pass  # Continuar con otros productos
            except:
                pass

        # Limpiar inmediatamente la lista global y la UI; el hilo usa el snapshot
        lista_productos_factura = []
        try:
            cargar_productos_en_tree()
        except Exception:
            pass

        threading.Thread(target=procesar_lista_bd, args=(items_a_enviar,), daemon=True).start()

        # Cerrar sin mensajes para velocidad
        ventana_lista.destroy()
    
    # Cargar productos iniciales
    cargar_productos_en_tree()
    
    # Frame para botones
    frame_botones = ttk.Frame(ventana_lista)
    frame_botones.pack(fill="x", padx=20, pady=15)
    
    ttk.Button(frame_botones, text="❌ Eliminar Seleccionado", 
              command=eliminar_producto, bootstyle="danger").pack(side="left", padx=8, pady=5)
    
    ttk.Button(frame_botones, text="🔄 Limpiar Lista", 
              command=lambda: [setattr(globals(), 'lista_productos_factura', []), 
                              cargar_productos_en_tree(),
                              messagebox.showinfo("Lista limpiada", "Todos los productos han sido eliminados.")], 
              bootstyle="warning").pack(side="left", padx=8, pady=5)
    
    ttk.Button(frame_botones, text="📋 Enviar Todos a Lista de Espera", 
              command=enviar_todos_a_lista_espera, 
              bootstyle="success").pack(side="right", padx=8, pady=5)
    
    ttk.Button(frame_botones, text="❌ Cerrar", 
              command=ventana_lista.destroy, 
              bootstyle="secondary").pack(side="right", padx=8, pady=5)

# === Código Base desde tabla CodigoBase ===
def obtener_codigo_base( base, producto, terminacion):
    try:
        conn = psycopg2.connect(
            host="dpg-d1b18u8dl3ps73e68v1g-a.oregon-postgres.render.com",
            port=5432,
            database="labels_app_db",
            user="admin",
            password="KCFjzM4KYzSQx63ArufESIXq03EFXHz3",
            sslmode="require"
        )
        cur = conn.cursor()
        cur.execute("SELECT base, tath, tath2, tath3 , flat, satin, sgi , flat2, satin3, sg4, satinkq, flatkp, flatmp, flatcov, flatpas, satinem, sgem, flatsp, satinsp, glossp, flatap, satinap, satinsan   FROM CodigoBase WHERE base ILIKE %s", (base,))
        row = cur.fetchone()
        conn.close()

        if not row:
            return "No encontrado"

        _, tath, tath2, tath3 ,flat, satin, sgi , flat2, satin3, sg4, satinkq, flatkp, flatmp, flatcov, flatpas, satinem, sgem, flatsp, satinsp, glossp, flatap, satinap, satinsan = row
        producto = producto.lower()
        terminacion = terminacion.lower()

        es_esmalte = any(p in producto for p in ["esmalte multiuso"])
        es_kempro = any(p in producto for p in ["kem pro"])
        es_kemaqua = any(p in producto for p in ["kem aqua"])
        es_masterpaint = any(p in producto for p in ["master paint"])
        es_pastel = any(p in producto for p in ["excello pastel"])
        es_emerald = any(p in producto for p in ["emerald"])
        es_superpaint = any(p in producto for p in ["super paint"])
        es_superpaintAP = any(p in producto for p in ["airpurtec"])
        es_sanitizing= any(p in producto for p in ["sanitizing"])
        es_laca= any(p in producto for p in ["laca"])
        es_EsmalteIndustrial= any(p in producto for p in ["esmalte kem"])
        es_uretano= any(p in producto for p in ["uretano"])
        es_tintealthinner= any(p in producto for p in ["tinte al thinner"])
        es_monocapa= any(p in producto for p in ["monocapa"])
        es_excellocov= any(p in producto for p in ["excello voc"])
        es_excellopremium= any(p in producto for p in ["excello premium"])
        es_waterblocking= any(p in producto for p in ["water blocking"])
        es_airpuretec= any(p in producto for p in ["airpuretec"])
        es_hcsiloconeacr= any(p in producto for p in ["h&c silicone-acrylic"])
        es_hcheavyshield= any(p in producto for p in ["h&c heavy-shield"]) 
        es_ProMarEgShel= any(p in producto for p in ["promar® 200 voc"])
        es_ProMarEgShel400= any(p in producto for p in ["promar® 400 voc"]) 
        es_proindustrialDTM= any(p in producto for p in ["pro industrial dtm"])                     
        es_armoseal= any(p in producto for p in ["armoseal 1000hs"])                     
        es_armosealtp= any(p in producto for p in ["armoseal t-p"]) 
        es_scufftuff= any(p in producto for p in ["scuff tuff-wb"]) 


        base_color = base_var.get().lower()


        if es_kemaqua:

            if terminacion == "satin":
                return satinkq
            else:
                return "No Aplica"

        if es_airpuretec:

            if  terminacion == "mate":

                if base_color == "extra white":
                 return "A86W00061-" 
                
                elif base_color == "deep":
                    return "A86W00063-"
                    
            elif terminacion == "satin":

                if base_color== "extra white":
                  
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

                return flatcov
            
            elif terminacion == "satin" and base_color == "extra white":

                return "A20DR2651-"
            
            else:
                return "No Aplica"    

        if es_laca:

            if terminacion == "mate":
                return "L15-" 
            elif terminacion == "semimate":
                return "L15-" 
            elif terminacion == "brillo":
                return "L15-" 
            else:
                return "No Aplica"
            

        if es_EsmalteIndustrial:

            if terminacion == "mate":
                return "F300-"
                
            elif terminacion == "semimate":
                return "F300-"
            
            elif terminacion == "brillo":
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
                
             elif terminacion == "satin" :
                 return "S24W00051-" 
             
               
             elif terminacion == "semigloss" :
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
             
        if  es_ProMarEgShel:

             if terminacion == "satin":

                if base_color == "deep":
                 return "B20W02653-" 
                
                elif base_color == "extra white":
                 return "B20W12651-"                
                
                else:
                    return "No aplica"
                                 
             elif terminacion== "mate":

                if base_color == "ultra deep":
                 return "B30T02654-" 
                
                elif base_color == "extra White":
                 return "B30W02651-1"                
                
                elif base_color == "deep":
                 return "B30W02653-"   

             elif terminacion== "semigloss":

                if base_color == "extra White":
                 return "B31W02651-"                
                                                  
                         
             else:

                return "No aplica"

        if  es_ProMarEgShel400:

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

             if terminacion == "mate":

                if base_color == "extra white":
                 return "ASPPA-" 
                
                elif base_color == "deep" or base_color == "ultra deep":
                 return "ASPPB-" 
                
                else:
                    return "ASPPD-"
                
             elif terminacion == "semimate":

                if base_color == "extra white":
                 return "ASPPA-" 
                
                elif base_color == "deep" or base_color == "ultra deep":
                 return "ASPPB-" 
                
                else:
                    return "ASPPD-"
                
             elif terminacion == "brillo":

                if base_color == "extra white":
                 return "ASPPA-" 
                
                elif base_color == "deep" or base_color == "ultra deep":
                 return "ASPPB-" 
                
                else:
                    return "ASPPD-"                    
             else:
                return "No Aplica"
    
        if es_tintealthinner:

            if terminacion == "claro":
                return tath 
            
            elif terminacion == "intermedio":
                return tath2
            
            elif terminacion== "especial":
                return tath3
            else:
                return "No Aplica"
            
        if es_monocapa:

            if terminacion == "mate":

                if base_color == "extra white":
                 return "ASMCA-" 
                
                elif base_color == "deep" or base_color == "ultra deep":
                 return "ASMCB-" 
                
                else:
                    return "ASMCD-" 
                
            elif terminacion == "semimate":

                if base_color == "extra white":
                 return "ASMCA-" 
                
                elif base_color == "deep" or base_color == "ultra deep":
                 return "ASMCB-" 
                
                else:
                    return "ASMCD-" 
                
            elif terminacion == "brillo":

                if base_color == "extra white":
                 return "ASMCA-" 
                
                elif base_color == "deep" or base_color == "ultra deep":
                 return "ASMCB-" 
                
                else:
                    return "ASMCD-" 
            else:
                return "No Aplica"

        if es_esmalte:

            if terminacion == "mate":
                return flat2
            elif terminacion == "satin":
                return satin3
            elif terminacion == "gloss":
                return sg4
            else:
                return "No Aplica"
            
        elif es_kempro:
                
            if terminacion == "mate":
                 return flatkp
            else:
                return "No Aplica"
            
        elif es_masterpaint:
                
            if terminacion == "mate":
                 return flatmp
            else:
                return "No Aplica"
            
        elif es_pastel:
                
            if terminacion == "mate":
                 return flatpas
            else:
                return "No Aplica"   
        elif es_emerald: 
                    
            if terminacion == "satin":
                return satinem + " K37W02751-"
            
            elif terminacion == "gloss":
                return sgem + "K38W02751-"
            else:
                return "No Aplica"
        elif es_superpaint: 
                    
            if terminacion == "mate":
                return flatsp
            
            elif terminacion == "satin":
                return satinsp    

            elif  terminacion == "gloss":
                return glossp
            else:
                return "No Aplica"           
            
        elif es_superpaintAP: 
                    
            if terminacion == "mate":
                return flatap
            elif terminacion == "satin":
                return satinap
            else:
                return "No Aplica" 
            
        elif es_sanitizing:
                
            if terminacion == "satin":
                 return satinsan
            else:
                return "No Aplica"
            
        elif es_excellopremium: 
            # Reglas especiales Excello Premium
            # 1) Ultra Deep II -> código PP4-
            es_ultra_deep_ii = any(k in base_color for k in ["ultra deep ii", "ultradeep ii", "ultra-deep ii", "ultra deep 2"])
            if es_ultra_deep_ii:
                return "PP4-"

            # 2) Ultra Deep (no II) -> solo Semisatin con código A27WDR03-
            es_ultra_deep = ("ultra deep" in base_color) and not es_ultra_deep_ii
            if es_ultra_deep:
                if terminacion == "semisatin":
                    return "A27WDR03-"
                else:
                    return "No Aplica"

            # 3) Resto de mapeos Excello Premium
            if terminacion == "mate":
                return flat
            elif terminacion == "satin":
                return satin
            elif terminacion == "semigloss":
                return sgi
            else:
                return "No Aplica"   
                                  
        else:
            return "No Aplica"
        
    except Exception as e:
        return "Error"
    
def imprimir_ficha_pintura(codigo_pintura):
    producto = producto_var.get().strip().lower()

    try:
        if producto == "excello premium":
            datos = obtener_datos_por_pintura(codigo_pintura)
            if datos:
                temp_pdf = tempfile.mktemp(".pdf")
                generar_pdf_ficha(datos, temp_pdf)
                os.startfile(temp_pdf)
                imprimir_pdf(temp_pdf)
            else:
                messagebox.showwarning("No encontrado", f"No hay fórmula disponible para el producto: {producto}")

        elif producto == "tinte al thinner":
            datos = obtener_datos_por_tinte(codigo_pintura)
            if datos:
                temp_pdf = tempfile.mktemp(".pdf")
                generar_pdf_tinte(datos, temp_pdf)
                os.startfile(temp_pdf)
                imprimir_pdf(temp_pdf)
            else:
                messagebox.showwarning("No encontrado", f"No hay fórmula disponible para el producto: {producto}")

        else:
            messagebox.showwarning("Producto no soportado", f"Producto no reconocido: {producto}")

    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error al imprimir: {e}")

            

def mostrar_codigo_base():
    global codigo_base_actual

    base = base_var.get()
    producto = producto_var.get()
    terminacion = terminacion_var.get()
    presentacion = presentacion_var.get()

    if not base or not producto or not terminacion:
        aviso_var.set("Completa todos los campos")
        return

    # Obtener código base
    resultado = obtener_codigo_base(base, producto, terminacion)
    
    # Agregar sufijo de presentación si está seleccionada
    if presentacion and resultado != "No encontrado" and resultado != "No Aplica":
        sufijo_presentacion = obtener_sufijo_presentacion(presentacion)
        if sufijo_presentacion:
            resultado += sufijo_presentacion

    app.clipboard_clear()
    app.clipboard_append(resultado)
    aviso_var.set("Código facturación copiado en el portapapeles")

    # Guardamos solo para vista previa
    codigo_base_actual = resultado
    codigo_base_var.set(resultado)

    actualizar_vista()
    limpiar_mensaje_despues(3000)

def actualizar_terminaciones(*args):
    """Actualiza las terminaciones disponibles según el producto seleccionado"""
    producto = producto_var.get().lower()
    base = base_var.get().lower()
    
    # Buscar terminaciones válidas para el producto
    terminaciones_validas = []
    for key, terminaciones in TERMINACIONES_POR_PRODUCTO.items():
        if key in producto:
            terminaciones_validas = terminaciones
            break
    
    # Si no se encuentra el producto, usar todas las terminaciones
    if not terminaciones_validas:
        terminaciones_validas = ['Mate', 'Satin', 'Semigloss', 'Semimate', 'Gloss', 'Brillo', 
                                "N/A", "ESPECIAL", "CLARO", "INTERMEDIO", "MADERA", "PERLADO", "METALICO"]
    
    # Lógica especial para Excello Premium con bases Ultra Deep / Ultra Deep II
    if 'excello premium' in producto:
        es_ultra_deep_ii = any(k in base for k in ['ultra deep ii', 'ultradeep ii', 'ultra-deep ii', 'ultra deep 2'])
        es_ultra_deep = ('ultra deep' in base)
        if es_ultra_deep or es_ultra_deep_ii:
            # Solo permitir Semisatin en ambas bases
            terminaciones_validas = ['Semisatin']

    # Actualizar el combobox
    terminaciones_combobox['values'] = terminaciones_validas
    
    # Limpiar la selección actual si no es válida
    terminacion_actual = terminacion_var.get()
    if terminacion_actual and terminacion_actual not in terminaciones_validas:
        terminacion_var.set('')
    
    # Si solo hay una terminación válida, seleccionarla automáticamente
    if len(terminaciones_validas) == 1:
        terminacion_var.set(terminaciones_validas[0])
    
    # Actualizar vista previa
    actualizar_vista()    


# Agrupa los botones en un LabelFrame moderno


# Crear estilo personalizado para botones y LabelFrame
style = ttk.Style()
# Botones azul oscuro, texto/iconos blancos
style.configure("BotonGrande.TButton", font=("Segoe UI", 11, "bold"), foreground="#fff", background="#222c3c", padding=8, borderwidth=1)
# Botón Imprimir con azul menos claro
style.configure("BotonImprimir.TButton", font=("Segoe UI", 11, "bold"), foreground="#fff", background="#1976d2", padding=8, borderwidth=1)
# Botones especiales para lista (verde)
style.configure("BotonEspecial.TButton", font=("Segoe UI", 10, "bold"), foreground="#fff", background="#28a745", padding=6, borderwidth=1)
# LabelFrame y título fondo blanco, título azul oscuro
style.configure("Acciones.TLabelframe", background="#fff", borderwidth=2, relief="groove")
style.configure("Acciones.TLabelframe.Label", font=("Segoe UI", 12, "bold"), foreground="#222c3c", background="#fff")

frame_botones = ttk.LabelFrame(app, text="Acciones", padding=16, style="Acciones.TLabelframe")
frame_botones.place(x=420, y=295, width=440, height=220)  # Subido un poco más nuevamente (de 310 a 295)

btn_style = {"width": 12, "style": "BotonGrande.TButton"}  # Reducir width para 4 botones


btn_limpiar = ttk.Button(frame_botones, text="Limpiar", command=limpiar_campos, **btn_style)
btn_limpiar.grid(row=0, column=0, padx=8, pady=10, sticky='ew')

btn_codigo = ttk.Button(frame_botones, text="Código", command=mostrar_codigo_base, **btn_style)
btn_codigo.grid(row=0, column=1, padx=8, pady=10, sticky='ew')

btn_ficha = ttk.Button(frame_botones, text="Fórmula", command=on_btn_imprimir_click, **btn_style)
btn_ficha.grid(row=0, column=2, padx=8, pady=10, sticky='ew')

btn_personalizar = ttk.Button(frame_botones, text="Custom", command=abrir_ventana_personalizados, **btn_style)
btn_personalizar.grid(row=0, column=3, padx=8, pady=10, sticky='ew')

# Segunda fila de botones
btn_agregar_lista = ttk.Button(frame_botones, text="📋 Agregar a Lista", command=lambda: agregar_producto_a_lista(), style="BotonEspecial.TButton")
btn_agregar_lista.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky='ew')

btn_gestionar_lista = ttk.Button(frame_botones, text="📑 Gestionar Lista", command=lambda: abrir_gestor_lista_factura(), style="BotonEspecial.TButton")
btn_gestionar_lista.grid(row=1, column=2, columnspan=2, padx=5, pady=5, sticky='ew')

# Botón Enviar ocupa 4 columnas y se centra debajo
btn_print = ttk.Button(frame_botones, text="� Enviar", command=imprimir_guardar, width=14, style="BotonImprimir.TButton")
btn_print.grid(row=2, column=0, columnspan=4, padx=10, pady=10, sticky='ew')

for i in range(4):  # Cambiar a 4 columnas
    frame_botones.columnconfigure(i, weight=1)

# Inicializar vista previa en blanco
actualizar_vista()

# Inicializar presentaciones
actualizar_presentaciones()

# Atajo de teclado: Ctrl+H para abrir historial
try:
    app.bind('<Control-h>', lambda e: abrir_historial_impresiones())
except Exception:
    pass

def main():
    """Función principal de la aplicación"""
    try:
        # Asegurar que el root esté visible tras el login
        try:
            app.deiconify()
        except Exception:
            pass
        app.mainloop()
    except Exception as e:
        messagebox.showerror("Error", f"Error inesperado al ejecutar la aplicación:\n{e}")

if __name__ == "__main__":
    main()

