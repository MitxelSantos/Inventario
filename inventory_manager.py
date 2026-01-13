#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
INVENTORY MANAGER - Sistema de Inventario Tecnológico
=========================================================
Hospital Regional Alfonso Jaramillo Salazar

VERSIÓN: 1.0
FECHA: Enero 2026
"""

import customtkinter as ctk
from tkinter import messagebox, filedialog
import platform
import socket
import subprocess
import os
import re
from datetime import datetime
from pathlib import Path
import threading

# Configurar tema CustomTkinter
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("green")

# Importar configuración
try:
    from config_listas import *
except ImportError:
    messagebox.showerror("Error", "No se encontró config_listas.py\nAsegúrate de tener ambos archivos en la misma carpeta")
    exit(1)

# Librerías opcionales
try:
    import openpyxl
    from openpyxl import load_workbook
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False
    messagebox.showwarning("Advertencia", "openpyxl no instalado. Ejecuta:\npip install openpyxl")

try:
    import psutil
    HAS_PSUTIL = True
except ImportError:
    HAS_PSUTIL = False

try:
    import wmi
    HAS_WMI = True
except ImportError:
    HAS_WMI = False
    print("WMI no disponible - Detección de hardware limitada")

try:
    import winreg
    HAS_WINREG = True
except ImportError:
    HAS_WINREG = False

# PIL para cargar imágenes (logo)
try:
    from PIL import Image
    HAS_PIL = True
except ImportError:
    HAS_PIL = False
    print("PIL/Pillow no disponible - Logo no se mostrará")


# ============================================================================
# COLORES INSTITUCIONALES
# ============================================================================

COLOR_VERDE_HOSPITAL = "#A9FA7B"
COLOR_AZUL_HOSPITAL = "#008ACC"
COLOR_NARANJA = "#F4B183"
COLOR_FONDO = "#F5F5F5"
COLOR_ERROR = "#DC3545"


# ============================================================================
# FUNCIONES DE DETECCIÓN
# ============================================================================

def detect_hardware_wmi():
    """Detectar hardware usando WMI: Serial + Discos primario y secundario."""
    info = {
        'marca': 'No detectado',
        'modelo': 'No detectado',
        'serial': 'No detectado',
        'tipo_disco': 'No detectado',
        # Disco secundario
        'disco_secundario': 'No tiene',
        'tipo_disco_secundario': 'No tiene',
        'serial_disco_secundario': 'No tiene',
        'marca_disco_secundario': 'No tiene',
        'modelo_disco_secundario': 'No tiene'
    }
    
    if not HAS_WMI:
        return info
    
    try:
        # Inicializar COM para evitar errores en threads
        try:
            import pythoncom
            pythoncom.CoInitialize()
        except:
            pass  # Si falla, continuar de todas formas
        
        c = wmi.WMI()
        
        # Información del sistema
        for system in c.Win32_ComputerSystem():
            info['marca'] = system.Manufacturer or 'No detectado'
            info['modelo'] = system.Model or 'No detectado'
        
        # Serial: Buscar en múltiples lugares
        serial_found = False
        serials_invalidos = ['default string', 'to be filled by o.e.m.', 'system serial number', 
                            'base board serial number', 'chassis serial number', '']
        
        # 1. Intentar desde BIOS
        for bios in c.Win32_BIOS():
            serial = (bios.SerialNumber or '').strip()
            if serial and serial.lower() not in serials_invalidos:
                info['serial'] = serial
                serial_found = True
                break
        
        # 2. Si no se encuentra, intentar desde BaseBoard (placa base)
        if not serial_found:
            for board in c.Win32_BaseBoard():
                serial = (board.SerialNumber or '').strip()
                if serial and serial.lower() not in serials_invalidos:
                    info['serial'] = f"MB-{serial}"  # Prefijo para identificar origen
                    serial_found = True
                    break
        
        # 3. Si aún no, intentar desde ComputerSystemProduct
        if not serial_found:
            for product in c.Win32_ComputerSystemProduct():
                serial = (product.IdentifyingNumber or '').strip()
                if serial and serial.lower() not in serials_invalidos:
                    info['serial'] = serial
                    serial_found = True
                    break
        
        # 4. Si aún no hay serial válido, dejar mensaje
        if not serial_found:
            info['serial'] = "No detectado (PC genérico/armado)"
        
        # DETECCIÓN DE DISCOS (Primario y Secundario)
        disks = list(c.Win32_DiskDrive())
        
        if len(disks) > 0:
            # Disco primario
            disk = disks[0]
            media_type = disk.MediaType or ''
            if 'SSD' in media_type.upper() or 'Solid State' in media_type:
                info['tipo_disco'] = 'SSD'
            else:
                info['tipo_disco'] = 'HDD'
        
        if len(disks) > 1:
            # Disco secundario detectado
            disk2 = disks[1]
            
            # Capacidad
            try:
                size_bytes = int(disk2.Size) if disk2.Size else 0
                size_gb = round(size_bytes / (1024**3))
                info['disco_secundario'] = str(size_gb)
            except:
                info['disco_secundario'] = 'Detectado'
            
            # Tipo
            media_type = disk2.MediaType or ''
            if 'SSD' in media_type.upper() or 'Solid State' in media_type:
                info['tipo_disco_secundario'] = 'SSD'
            else:
                info['tipo_disco_secundario'] = 'HDD'
            
            # Serial
            serial_disk = (disk2.SerialNumber or '').strip()
            if serial_disk:
                info['serial_disco_secundario'] = serial_disk
            else:
                info['serial_disco_secundario'] = 'No detectado'
            
            # Marca
            marca_disk = (disk2.Manufacturer or '').strip()
            if marca_disk and marca_disk.lower() not in ['(standard disk drives)', '']:
                info['marca_disco_secundario'] = marca_disk
            else:
                info['marca_disco_secundario'] = 'No detectado'
            
            # Modelo
            modelo_disk = (disk2.Model or '').strip()
            if modelo_disk:
                info['modelo_disco_secundario'] = modelo_disk
            else:
                info['modelo_disco_secundario'] = 'No detectado'
    
    except Exception as e:
        print(f"Error WMI: {e}")
    
    return info


def detect_office_version():
    """Detectar versión de Office: Busca ejecutables incluso sin licencia."""
    if not HAS_WINREG:
        return "No detectado", "No detectado"
    
    try:
        # ESTRATEGIA 1: Buscar en InstallRoot (instalación completa licenciada)
        key_paths = [
            r"SOFTWARE\Microsoft\Office\16.0\Common\InstallRoot",  # Office 2016/2019/365
            r"SOFTWARE\Microsoft\Office\15.0\Common\InstallRoot",  # Office 2013
            r"SOFTWARE\Microsoft\Office\14.0\Common\InstallRoot",  # Office 2010
        ]
        
        for key_path in key_paths:
            try:
                key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path)
                path = winreg.QueryValueEx(key, "Path")[0]
                winreg.CloseKey(key)
                
                if "16.0" in key_path:
                    version = "Office 2016/2019/365"
                elif "15.0" in key_path:
                    version = "Office 2013"
                elif "14.0" in key_path:
                    version = "Office 2010"
                else:
                    version = "Detectado"
                
                licencia = "Retail/Volume"
                return version, licencia
            except:
                continue
        
        # ESTRATEGIA 2: Buscar ejecutables de Office (incluso sin licencia completa)
        office_paths = [
            (r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE", "Office 2016/2019/365"),
            (r"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE", "Office 2016/2019/365"),
            (r"C:\Program Files\Microsoft Office\Office16\WINWORD.EXE", "Office 2016/2019/365"),
            (r"C:\Program Files (x86)\Microsoft Office\Office16\WINWORD.EXE", "Office 2016/2019/365"),
            (r"C:\Program Files\Microsoft Office\Office15\WINWORD.EXE", "Office 2013"),
            (r"C:\Program Files (x86)\Microsoft Office\Office15\WINWORD.EXE", "Office 2013"),
            (r"C:\Program Files\Microsoft Office\Office14\WINWORD.EXE", "Office 2010"),
            (r"C:\Program Files (x86)\Microsoft Office\Office14\WINWORD.EXE", "Office 2010"),
        ]
        
        for path, version in office_paths:
            if os.path.exists(path):
                return version, "Instalado (verificar licencia)"
        
        # ESTRATEGIA 3: Buscar en registro de desinstalación
        try:
            for hive in [winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER]:
                try:
                    key = winreg.OpenKey(hive, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
                    for i in range(winreg.QueryInfoKey(key)[0]):
                        try:
                            subkey_name = winreg.EnumKey(key, i)
                            if 'Office' in subkey_name or 'Microsoft 365' in subkey_name:
                                subkey = winreg.OpenKey(key, subkey_name)
                                try:
                                    display_name = winreg.QueryValueEx(subkey, "DisplayName")[0]
                                    if 'Office' in display_name or 'Microsoft 365' in display_name:
                                        winreg.CloseKey(subkey)
                                        winreg.CloseKey(key)
                                        return display_name, "Instalado (verificar licencia)"
                                except:
                                    pass
                                winreg.CloseKey(subkey)
                        except:
                            continue
                    winreg.CloseKey(key)
                except:
                    continue
        except:
            pass
        
        return "No instalado", "N/A"
    
    except Exception as e:
        return "No detectado", "No detectado"


def detect_office_apps():
    """Detectar si Teams y Outlook están instalados."""
    teams = "No"
    outlook = "No"
    
    # Rutas comunes de Teams
    teams_paths = [
        r"C:\Users\{}\AppData\Local\Microsoft\Teams\current\Teams.exe",
        r"C:\Program Files\Microsoft\Teams\current\Teams.exe",
        r"C:\Program Files (x86)\Microsoft\Teams\current\Teams.exe"
    ]
    
    username = os.environ.get('USERNAME', '')
    for path in teams_paths:
        full_path = path.format(username)
        if os.path.exists(full_path):
            teams = "Sí"
            break
    
    # Rutas comunes de Outlook
    outlook_paths = [
        r"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE",
        r"C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE",
        r"C:\Program Files\Microsoft Office\Office16\OUTLOOK.EXE",
        r"C:\Program Files (x86)\Microsoft Office\Office16\OUTLOOK.EXE",
    ]
    
    for path in outlook_paths:
        if os.path.exists(path):
            outlook = "Sí"
            break
    
    return teams, outlook


def detect_windows_license():
    """Detectar información de licencia de Windows."""
    licencia_info = {
        'tipo': 'No detectado',
        'key': 'No detectado',
        'estado': 'No detectado'
    }
    
    try:
        # Ejecutar slmgr para obtener info de licencia
        result = subprocess.run(
            ['cscript', '//nologo', r'C:\Windows\System32\slmgr.vbs', '/dli'],
            capture_output=True,
            text=True,
            timeout=10
        )
        
        output = result.stdout
        
        # Parsear tipo de licencia
        if 'OEM' in output:
            licencia_info['tipo'] = 'OEM'
        elif 'Retail' in output:
            licencia_info['tipo'] = 'Retail'
        elif 'Volume' in output:
            licencia_info['tipo'] = 'Volume'
        else:
            licencia_info['tipo'] = 'Detectado'
        
        # Estado
        if 'Licensed' in output or 'Licenciado' in output:
            licencia_info['estado'] = 'Activado'
        else:
            licencia_info['estado'] = 'No activado'
        
        # Obtener últimos 5 dígitos de la key
        key_result = subprocess.run(
            ['cscript', '//nologo', r'C:\Windows\System32\slmgr.vbs', '/dli'],
            capture_output=True,
            text=True,
            timeout=10
        )
        
        key_output = key_result.stdout
        # Buscar patrón de product key (últimos 5)
        key_match = re.search(r'([A-Z0-9]{5})$', key_output, re.MULTILINE)
        if key_match:
            licencia_info['key'] = key_match.group(1)
        else:
            licencia_info['key'] = 'XXXXX'
    
    except Exception as e:
        print(f"Error detectando licencia Windows: {e}")
    
    return licencia_info


def detect_last_windows_update():
    """Detectar última actualización de Windows."""
    try:
        if not HAS_WINREG:
            return "No detectado"
        
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, 
                             r"SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\Results\Install")
        last_success = winreg.QueryValueEx(key, "LastSuccessTime")[0]
        winreg.CloseKey(key)
        
        # Formatear fecha
        if last_success:
            # Formato: YYYY-MM-DD HH:MM:SS
            try:
                date_obj = datetime.strptime(last_success, "%Y-%m-%d %H:%M:%S")
                return date_obj.strftime("%Y-%m-%d")
            except:
                return last_success[:10]  # Primeros 10 caracteres (fecha)
        
        return "No detectado"
    
    except Exception as e:
        return "No detectado"


# ============================================================================
# CLASE PRINCIPAL - INVENTORY MANAGER
# ============================================================================

class InventoryManagerApp:
    """Aplicación principal con CustomTkinter."""
    
    def __init__(self, root):
        self.root = root
        self.root.title("Sistema de Inventario Tecnológico - HRAJS")
        
        # Configurar tamaño de ventana inicial (1400x900 o 90% de pantalla)
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        # Usar 90% de pantalla o tamaño fijo (el menor)
        window_width = min(int(screen_width * 0.9), 1600)
        window_height = min(int(screen_height * 0.9), 1000)
        
        # Centrar ventana
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        # Variables de estado
        self.excel_path = None
        self.current_row = None
        self.current_sheet = "Equipos de Cómputo"  # Sheet actual
        self.equipment_data = {}
        self.verde_data = {}
        self.azul_data = {}
        
        # Widgets de formulario (para acceso posterior)
        self.manual_widgets = {}
        self.main_container = None  # Contenedor principal para cambiar vistas
        
        # PRIMERO: Crear menú nativo (por encima de todo)
        self.create_native_menu()
        
        # SEGUNDO: Crear header
        self.create_header()
        
        # TERCERO: Contenedor principal para las vistas
        self.main_container = ctk.CTkFrame(self.root, fg_color=COLOR_FONDO)
        self.main_container.pack(fill="both", expand=True, padx=0, pady=0)
        
        # CUARTO: Intentar cargar Excel automáticamente (después de que la ventana esté lista)
        self.root.after(100, self.auto_load_excel)
    
    
    def create_native_menu(self):
        """Crear menú nativo de tkinter (por encima del header)."""
        import tkinter as tk
        
        # Crear barra de menú nativa
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # MENÚ ARCHIVO
        menu_archivo = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Archivo", menu=menu_archivo)
        menu_archivo.add_command(label="Cargar Excel", command=self.browse_excel)
        menu_archivo.add_separator()
        menu_archivo.add_command(label="Salir", command=self.root.quit)
        
        # MENÚ INVENTARIOS
        menu_inventarios = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Inventarios", menu=menu_inventarios)
        menu_inventarios.add_command(
            label="Equipos de Cómputo", 
            command=lambda: self.show_form_directo("Equipos de Cómputo")
        )
        menu_inventarios.add_command(
            label="Impresoras", 
            command=lambda: self.show_form_directo("Impresoras")
        )
        menu_inventarios.add_command(
            label="Periféricos", 
            command=lambda: self.show_form_directo("Periféricos")
        )
        menu_inventarios.add_command(
            label="Equipos de Red", 
            command=lambda: self.show_form_directo("Red")
        )
        
        # MENÚ OPERACIONES
        menu_operaciones = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Operaciones", menu=menu_operaciones)
        menu_operaciones.add_command(
            label="Mantenimiento", 
            command=lambda: self.show_form_directo("Mantenimiento")
        )
        menu_operaciones.add_command(
            label="Dar de Baja", 
            command=lambda: self.show_form_directo("Dados de Baja")
        )
        
        # MENÚ AYUDA
        menu_ayuda = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Ayuda", menu=menu_ayuda)
        menu_ayuda.add_command(label="Guía de Formulario", command=self.show_classification_guide)
    
    def show_form_directo(self, tipo):
        """Mostrar formulario directamente sin tabs."""
        if not self.excel_path:
            messagebox.showwarning("Advertencia", "Primero debes cargar un archivo Excel.\n\nVe a: Archivo → Cargar Excel")
            return
        
        # Limpiar contenedor principal
        for widget in self.main_container.winfo_children():
            widget.destroy()
        
        # Mostrar formulario correspondiente
        if tipo == "Equipos de Cómputo":
            self.show_manual_form_in_container()
        elif tipo == "Impresoras":
            self.create_impresoras_form_directo()
        elif tipo == "Periféricos":
            self.create_perifericos_form_directo()
        elif tipo == "Red":
            self.create_red_form_directo()
        elif tipo == "Mantenimiento":
            self.create_mantenimientos_form_directo()
        elif tipo == "Dados de Baja":
            self.create_baja_form_directo()
    
    def show_manual_form_in_container(self):
        """Mostrar formulario de datos manuales en contenedor principal."""
        # Guardar valores de campos que deben mantenerse antes de limpiar
        campos_a_mantener = ["tipo_equipo",'area_servicio',"macro_proceso", 'proceso',"sihos","office_basico",
                             "software_especializado","horario_uso", 'periodicidad_mtto', 'tecnico_responsable']
        valores_guardados = {}
        
        if hasattr(self, 'manual_widgets'):
            for field_name in campos_a_mantener:
                if field_name in self.manual_widgets:
                    try:
                        widget = self.manual_widgets[field_name]
                        if hasattr(widget, 'winfo_exists') and widget.winfo_exists():
                            if isinstance(widget, (ctk.CTkEntry, ctk.CTkComboBox)):
                                valores_guardados[field_name] = widget.get()
                    except:
                        pass
        
        # Limpiar contenedor
        for widget in self.main_container.winfo_children():
            widget.destroy()
        
        # Frame scrollable para formulario
        form_frame = ctk.CTkScrollableFrame(
            self.main_container,
            fg_color="#FAFAFA",
            label_text=f"📝 DATOS MANUALES - Equipo #{self.current_row-1} (Código EQC-{self.current_row-1:04d})",
            label_fg_color=COLOR_VERDE_HOSPITAL,
            label_text_color="white",
            label_font=("Segoe UI", 15, "bold")
        )
        form_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Guardar referencia para actualizar título después
        self.equipo_form_frame = form_frame
        
        # Campos del formulario
        fields = [
            ("* Tipo de Equipo ", "tipo_equipo", "combobox", TIPO_EQUIPO),
            ("Área *", "area_servicio", "combobox", AREAS_SERVICIOS),
            ("Ubicación Específica *", "ubicacion_especifica", "entry", None),
            ("Responsable / Custodio *", "responsable_custodio", "entry", None),
            ("Macroproceso", "macro_proceso", "combobox", MACRO_PROCESO)
            ("Proceso *", "proceso", "combobox", PROCESOS),
            ("Uso - SIHOS *", "uso_sihos", "combobox", USO_SIHOS),
            ("Uso - SIFAX", "uso_sifax", "combobox", USO_SIFAX),
            ("Uso - Office Básico", "uso_office_basico", "combobox", USO_OFFICE_BASICO),
            ("Software Especializado", "software_especializado", "combobox", SOFTWARE_ESPECIALIZADO_OPCIONES),
            ("Descripción Software Esp.", "descripcion_software", "entry", None),
            ("Función Principal", "funcion_principal", "entry", None),
            ("Nivel de Criticidad", "criticidad", "combobox", CRITICIDAD),
            ("Clasificación Confidencialidad", "confidencialidad", "combobox", CONFIDENCIALIDAD),
            ("Horario de Uso", "horario_uso", "combobox", HORARIO_USO),
            ("Estado Operativo *", "estado_operativo", "combobox", ESTADO_OPERATIVO),
            ("Fecha de Adquisición (YYYY-MM-DD)", "fecha_adquisicion", "entry", None),
            ("Valor de Adquisición (COP)", "valor_adquisicion", "entry", None),
            ("Fecha Venc. Garantía (YYYY-MM-DD)", "fecha_venc_garantia", "entry", None),
            ("Observaciones Técnicas", "observaciones_tecnicas", "entry", None),
            ("Fecha Exp. Antivirus (YYYY-MM-DD)", "fecha_exp_antivirus", "entry", None),
            ("Periodicidad Mtto", "periodicidad_mtto", "combobox", PERIODICIDAD_MTTO),
            ("Responsable Mtto", "responsable_mtto", "combobox", RESPONSABLE_MTTO),
            ("Último Mantenimiento (YYYY-MM-DD)", "ultimo_mantenimiento", "entry", None),
            ("Tipo Último Mtto", "tipo_ultimo_mtto", "combobox", TIPO_MTTO),
        ]
        
        for label_text, field_name, field_type, options in fields:
            self.create_form_field(form_frame, label_text, field_name, field_type, options)
        
        # Restaurar valores guardados
        for field_name, valor in valores_guardados.items():
            if field_name in self.manual_widgets and valor:
                try:
                    widget = self.manual_widgets[field_name]
                    if isinstance(widget, ctk.CTkEntry):
                        widget.insert(0, valor)
                    elif isinstance(widget, ctk.CTkComboBox):
                        widget.set(valor)
                except:
                    pass
        
        # Frame para botones de acción
        btn_action_frame = ctk.CTkFrame(form_frame, fg_color="transparent")
        btn_action_frame.pack(pady=20, padx=20, fill="x")
        
        # Botón GUARDAR NUEVO (solo datos manuales)
        self.btn_save_equipo = ctk.CTkButton(
            btn_action_frame,
            text="💾 GUARDAR NUEVO (Solo Datos Manuales)",
            command=self.save_equipo_manual_only,
            font=("Segoe UI", 14, "bold"),
            fg_color=COLOR_VERDE_HOSPITAL,
            hover_color="#1F5039",
            height=50,
            width=350
        )
        self.btn_save_equipo.pack(side="left", padx=10)
        
        # Botón ACTUALIZAR EXISTENTE
        btn_update = ctk.CTkButton(
            btn_action_frame,
            text="🔄 ACTUALIZAR EXISTENTE",
            command=self.update_equipo_computo,
            font=("Segoe UI", 14, "bold"),
            fg_color="#2196F3",
            hover_color="#1976D2",
            height=50,
            width=350
        )
        btn_update.pack(side="left", padx=10)
        
        # Separador
        separator = ctk.CTkFrame(form_frame, height=2, fg_color="#E0E0E0")
        separator.pack(fill="x", padx=20, pady=15)
        
        # Botón de recopilación automática
        btn_collect = ctk.CTkButton(
            form_frame,
            text="➡️ CONTINUAR: RECOPILACIÓN AUTOMÁTICA COMPLETA",
            command=self.start_automatic_collection,
            font=("Arial", 16, "bold"),
            fg_color="#FF9800",
            hover_color="#F57C00",
            height=50
        )
        btn_collect.pack(pady=20, padx=20, fill="x")



    def create_impresoras_form_directo(self):
        """Crear formulario de impresoras directamente."""
        for widget in self.main_container.winfo_children():
            widget.destroy()
        self.create_impresoras_form(self.main_container)
    
    def create_perifericos_form_directo(self):
        """Crear formulario de periféricos directamente."""
        for widget in self.main_container.winfo_children():
            widget.destroy()
        self.create_perifericos_form(self.main_container)
    
    def create_red_form_directo(self):
        """Crear formulario de red directamente."""
        for widget in self.main_container.winfo_children():
            widget.destroy()
        self.create_red_form(self.main_container)
    
    def create_mantenimientos_form_directo(self):
        """Crear formulario de mantenimientos directamente."""
        for widget in self.main_container.winfo_children():
            widget.destroy()
        self.create_mantenimientos_form(self.main_container)
    
    def create_baja_form_directo(self):
        """Crear formulario de baja directamente."""
        for widget in self.main_container.winfo_children():
            widget.destroy()
        self.create_baja_form(self.main_container)

    def get_next_available_row(self, sheet_name, check_column=1, max_rows=500):
        """
        Función optimizada para buscar siguiente fila disponible en cualquier hoja.
        
        Args:
            sheet_name: Nombre de la hoja Excel
            check_column: Columna a verificar (default 1 = Consecutivo)
            max_rows: Máximo de filas a buscar (default 500)
        
        Returns:
            int: Número de la siguiente fila disponible
        """
        if not self.excel_path or not HAS_OPENPYXL:
            return 2
        
        try:
            wb = load_workbook(self.excel_path, read_only=True)
            ws = wb[sheet_name]
            
            for row in range(2, max_rows + 2):
                if ws.cell(row=row, column=check_column).value is None:
                    wb.close()
                    return row
            
            wb.close()
            return max_rows + 2
            
        except Exception as e:
            print(f"Error buscando siguiente fila: {e}")
            return 2
    
    def create_header(self):
        """Crear encabezado con diseño profesional - VERDE."""
        header_frame = ctk.CTkFrame(self.root, fg_color=COLOR_VERDE_HOSPITAL, corner_radius=0)
        header_frame.pack(fill="x", padx=0, pady=0)
        
        # Frame interno para organizar logo + texto
        content_frame = ctk.CTkFrame(header_frame, fg_color="transparent")
        content_frame.pack(pady=18)
        
        # Intentar cargar logo institucional
        if HAS_PIL:
            logo_paths = [
                "logo_hospital.png",
                "logo.png", 
                "escudo_hospital.png",
                "hospital_logo.png"
            ]
            
            for logo_path in logo_paths:
                if os.path.exists(logo_path):
                    try:
                        logo_image = Image.open(logo_path)
                        # Redimensionar a altura 70px manteniendo proporción
                        aspect_ratio = logo_image.width / logo_image.height
                        new_height = 70
                        new_width = int(new_height * aspect_ratio)
                        logo_image = logo_image.resize((new_width, new_height), Image.Resampling.LANCZOS)
                        
                        logo_ctk = ctk.CTkImage(light_image=logo_image, dark_image=logo_image, 
                                               size=(new_width, new_height))
                        
                        logo_label = ctk.CTkLabel(
                            content_frame,
                            image=logo_ctk,
                            text=""
                        )
                        logo_label.pack(side="left", padx=(0, 25))
                        print(f"✓ Logo cargado: {logo_path}")
                        break
                    except Exception as e:
                        print(f"✗ Error al cargar logo {logo_path}: {e}")
            else:
                print("ℹ No se encontró logo (logo_hospital.png, logo.png, etc.)")
        else:
            print("ℹ PIL/Pillow no instalado - Logo no disponible")
        
        # Frame para texto (derecha del logo o solo si no hay logo)
        text_frame = ctk.CTkFrame(content_frame, fg_color="transparent")
        text_frame.pack(side="left")
        
        title_label = ctk.CTkLabel(
            text_frame,
            text="SISTEMA DE INVENTARIO TECNOLÓGICO",
            font=("Segoe UI", 22, "bold"),
            text_color="white"
        )
        title_label.pack()
        
        subtitle_label = ctk.CTkLabel(
            text_frame,
            text="Hospital Regional Alfonso Jaramillo Salazar - Líbano, Tolima",
            font=("Segoe UI", 12),
            text_color="white"
        )
        subtitle_label.pack(pady=(2, 0))
        
        # Label de estado del archivo cargado (esquina superior derecha)
        self.status_label = ctk.CTkLabel(
            header_frame,
            text="",
            font=("Segoe UI", 10),
            text_color="white",
            fg_color="transparent"
        )
        self.status_label.place(relx=0.98, rely=0.5, anchor="e")
    
    def browse_excel(self):
        """Abrir diálogo para seleccionar Excel."""
        filename = filedialog.askopenfilename(
            title="Seleccionar archivo Excel - inventario_hospital_v1.xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        if filename:
            self.excel_path = filename
            
            # Detectar siguiente fila
            self.current_row = self.get_next_available_row("Equipos de Cómputo", check_column=1)
            
            # Actualizar status
            filename_short = os.path.basename(filename)
            self.status_label.configure(text=f"✅ {filename_short} cargado")
            
            # Mostrar pestañas
            self.show_manual_form_in_container()
            
            messagebox.showinfo("Éxito", f"✅ Archivo cargado correctamente:\n{filename_short}\n\nSiguiente fila disponible: {self.current_row}")
    
    def auto_load_excel(self):
        """Cargar Excel automáticamente si existe en el directorio actual."""
        default_file = "inventario_hospital_v1.xlsx"
        
        if os.path.exists(default_file):
            self.excel_path = default_file
            
            # Detectar siguiente fila automáticamente
            self.current_row = self.get_next_available_row("Equipos de Cómputo", check_column=1)
            
            # Actualizar status
            self.status_label.configure(text=f"✅ {default_file} cargado")
            
            # Mostrar pestañas directamente
            self.show_manual_form_in_container()
            
            print(f"✅ Excel cargado automáticamente: {default_file}")
            print(f"✅ Siguiente fila disponible: {self.current_row}")
        else:
            # No hay archivo, mostrar mensaje en contenedor
            self.show_no_file_message()
    
    def show_no_file_message(self):
        """Mostrar mensaje cuando no hay archivo cargado."""
        for widget in self.main_container.winfo_children():
            widget.destroy()
        
        msg_frame = ctk.CTkFrame(self.main_container, fg_color=COLOR_FONDO)
        msg_frame.pack(fill="both", expand=True)
        
        # Centrar mensaje
        center_frame = ctk.CTkFrame(msg_frame, fg_color="transparent")
        center_frame.place(relx=0.5, rely=0.5, anchor="center")
        
        icon_label = ctk.CTkLabel(
            center_frame,
            text="📂",
            font=("Segoe UI", 80)
        )
        icon_label.pack(pady=(0, 20))
        
        title_label = ctk.CTkLabel(
            center_frame,
            text="No se encontró el archivo Excel",
            font=("Segoe UI", 24, "bold"),
            text_color=COLOR_VERDE_HOSPITAL
        )
        title_label.pack(pady=(0, 10))
        
        subtitle_label = ctk.CTkLabel(
            center_frame,
            text="El sistema busca: inventario_hospital_v1.xlsx\nen el directorio actual",
            font=("Segoe UI", 14),
            text_color="#666666"
        )
        subtitle_label.pack(pady=(0, 30))
        
        btn_cargar = ctk.CTkButton(
            center_frame,
            text="📁 CARGAR ARCHIVO EXCEL",
            command=self.browse_excel,
            font=("Segoe UI", 16, "bold"),
            fg_color=COLOR_VERDE_HOSPITAL,
            hover_color="#1F5A32",
            height=60,
            width=300,
            corner_radius=12
        )
        btn_cargar.pack()
    
    def detect_next_code(self, sheet_name, prefix):
        """Detectar siguiente código disponible basado en el último consecutivo en columna 1."""
        if not self.excel_path or not HAS_OPENPYXL:
            return f"{prefix}-001"
        
        try:
            wb = load_workbook(self.excel_path, read_only=True)
            
            # Verificar que la hoja existe
            if sheet_name not in wb.sheetnames:
                wb.close()
                print(f"⚠️ Advertencia: Hoja '{sheet_name}' no existe. Creándola...")
                return f"{prefix}-001"
            
            ws = wb[sheet_name]
            
            # Buscar el ÚLTIMO consecutivo en columna 1 (no asumir que es next_row - 1)
            last_consecutive = 0
            for row in range(2, 500):
                value = ws.cell(row=row, column=1).value
                if value is not None:
                    try:
                        consecutivo = int(value)
                        if consecutivo > last_consecutive:
                            last_consecutive = consecutivo
                    except:
                        pass
                else:
                    break  # Primera fila vacía, detener
            
            wb.close()
            
            next_consecutive = last_consecutive + 1
            
            # Todos los códigos ahora son de 4 dígitos
            return f"{prefix}-{next_consecutive:04d}"
            
        except Exception as e:
            print(f"❌ Error detectando código: {e}")
            import traceback
            traceback.print_exc()
            return f"{prefix}-001"
    
    def detect_next_consecutive_mantenimiento(self):
        """Detectar siguiente consecutivo para mantenimientos."""
        next_row = self.get_next_available_row("Mantenimientos", check_column=1)
        return next_row - 1
    
    def detect_next_baja(self):
        """Detectar siguiente número de baja."""
        next_row = self.get_next_available_row("Equipos Dados de Baja", check_column=1, max_rows=200)
        return next_row - 1
    
    def create_form_field(self, parent, label_text, field_name, field_type, options):
        """Crear campo del formulario con diseño mejorado."""
        # Frame con fondo blanco y bordes sutiles
        field_frame = ctk.CTkFrame(parent, fg_color="white", corner_radius=8)
        field_frame.pack(fill="x", padx=15, pady=6)
        
        # Frame interno para contenido
        inner_frame = ctk.CTkFrame(field_frame, fg_color="transparent")
        inner_frame.pack(fill="x", padx=15, pady=10)
        
        # Label mejorado
        label = ctk.CTkLabel(
            inner_frame,
            text=label_text,
            font=("Segoe UI", 12, "bold"),
            width=320,
            anchor="w",
            text_color="#333333"
        )
        label.pack(side="left", padx=(0, 20))
        
        # Widget según tipo
        if field_type == "combobox":
            widget = ctk.CTkComboBox(
                inner_frame,
                values=options,
                width=620,
                height=35,
                font=("Segoe UI", 11),
                dropdown_font=("Segoe UI", 10),
                border_color="#CCCCCC",
                button_color=COLOR_VERDE_HOSPITAL,
                button_hover_color="#1F5A32",
                corner_radius=8
            )
        else:  # entry
            widget = ctk.CTkEntry(
                inner_frame,
                width=620,
                height=35,
                font=("Segoe UI", 11),
                border_color="#CCCCCC",
                fg_color="white",
                corner_radius=8
            )
        
        widget.pack(side="left", fill="x", expand=True)
        self.manual_widgets[field_name] = widget
        return widget  # ← RETORNAR el widget creado
    
    def show_classification_guide(self):
        """Mostrar ventana con guía de clasificación normativa - CLARA Y ÚTIL."""
        guide_window = ctk.CTkToplevel(self.root)
        guide_window.title("Guía de Clasificación Normativa")
        guide_window.geometry("1200x750")
        
        # Centrar
        guide_window.update_idletasks()
        x = (guide_window.winfo_screenwidth() // 2) - 600
        y = (guide_window.winfo_screenheight() // 2) - 375
        guide_window.geometry(f"1200x750+{x}+{y}")
        
        # Header verde profesional
        header_frame = ctk.CTkFrame(guide_window, fg_color=COLOR_VERDE_HOSPITAL, corner_radius=0)
        header_frame.pack(fill="x", padx=0, pady=0)
        
        header = ctk.CTkLabel(
            header_frame,
            text="📋 GUÍA DE CLASIFICACIÓN NORMATIVA",
            font=("Segoe UI", 24, "bold"),
            text_color="white"
        )
        header.pack(pady=(18, 5))
        
        subtitle = ctk.CTkLabel(
            header_frame,
            text="Criterios según MinTIC PETI y MinSalud - Resolución 2183 de 2004",
            font=("Segoe UI", 12),
            text_color="white"
        )
        subtitle.pack(pady=(0, 18))
        
        # Crear Tabview
        tabview = ctk.CTkTabview(guide_window, width=1150, height=580)
        tabview.pack(pady=15, padx=25)
        
        # Crear tabs
        tabview.add("🔴 Criticidad")
        tabview.add("🔒 Confidencialidad")
        tabview.add("🏥 Procesos")
        tabview.add("💻 Sistemas")
        tabview.add("⚡ Otros")
        
        # ===== TAB 1: CRITICIDAD =====
        self._create_criticality_tab_clean(tabview.tab("🔴 Criticidad"))
        
        # ===== TAB 2: CONFIDENCIALIDAD =====
        self._create_confidentiality_tab_clean(tabview.tab("🔒 Confidencialidad"))
        
        # ===== TAB 3: PROCESOS =====
        self._create_processes_tab_clean(tabview.tab("🏥 Procesos"))
        
        # ===== TAB 4: SISTEMAS =====
        self._create_systems_tab_clean(tabview.tab("💻 Sistemas"))
        
        # ===== TAB 5: OTROS =====
        self._create_others_tab_clean(tabview.tab("⚡ Otros"))
        
        # Botón cerrar mejorado
        btn_frame = ctk.CTkFrame(guide_window, fg_color="transparent")
        btn_frame.pack(pady=15)
        
        btn_close = ctk.CTkButton(
            btn_frame,
            text="✓ ENTENDIDO",
            command=guide_window.destroy,
            font=("Segoe UI", 14, "bold"),
            fg_color=COLOR_VERDE_HOSPITAL,
            hover_color="#1F5A32",
            height=45,
            width=250,
            corner_radius=10
        )
        btn_close.pack()
    
    def _create_criticality_tab_clean(self, parent):
        """Tab de criticidad con mensajes CLAROS y ÚTILES."""
        scroll = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Título
        title = ctk.CTkLabel(
            scroll,
            text="NIVEL DE CRITICIDAD - ¿Qué tan importante es este equipo?",
            font=("Segoe UI", 17, "bold"),
            text_color=COLOR_VERDE_HOSPITAL
        )
        title.pack(pady=(0, 15))
        
        intro = ctk.CTkLabel(
            scroll,
            text="Pregúntate: ¿Qué pasa si este equipo falla ahora mismo?",
            font=("Segoe UI", 12, "italic"),
            text_color="#666666"
        )
        intro.pack(pady=(0, 20))
        
        levels = [
            ("🔴 CRÍTICO", "#DC3545", 
             "Si falla, se PARALIZA atención de pacientes",
             [
                "✓ Usa en: Equipos de Urgencias, UCI, Quirófanos",
                "✓ Usa en: Equipos que corren SIHOS/SIFAX 24/7",
                "✓ Usa en: Servidor principal, estaciones de enfermería críticas",
                "✗ NO uses en: Equipos administrativos o de oficina",
                "⏱ Falla: Menos de 1 hora de tolerancia",
                "💡 Ejemplo: PC Estación Enfermería UCI, Servidor SIHOS Principal"
            ]),
            ("🟠 ALTO", "#FD7E14",
             "Si falla, afecta operación importante del hospital",
             [
                "✓ Usa en: Laboratorio, Imágenes, Farmacia, Facturación",
                "✓ Usa en: Equipos que procesan pacientes directamente",
                "✓ Usa en: Consulta Externa, Hospitalización",
                "✗ NO uses en: Equipos que solo hacen Office/email",
                "⏱ Falla: Menos de 4 horas de tolerancia",
                "💡 Ejemplo: PC Laboratorio Clínico, PC Facturación Principal"
            ]),
            ("🟡 MEDIO", "#FFC107",
             "Si falla, afecta trabajo pero NO se paraliza nada",
             [
                "✓ Usa en: Contabilidad, Recursos Humanos, Calidad",
                "✓ Usa en: Oficinas administrativas en general",
                "✓ Usa en: Equipos de apoyo que usan Office/email",
                "✗ NO uses en: Áreas que atienden pacientes",
                "⏱ Falla: Puede esperar 1 día",
                "💡 Ejemplo: PC Contador, PC Recursos Humanos, PC Secretaria"
            ]),
            ("🟢 BAJO", "#28A745",
             "Si falla, casi no afecta - uso esporádico",
             [
                "✓ Usa en: Almacén, Servicios Generales, Mantenimiento",
                "✓ Usa en: Equipos usados ocasionalmente",
                "✓ Usa en: Equipos de respaldo o bodega",
                "✗ NO uses en: Áreas operativas diarias",
                "⏱ Falla: Puede esperar varios días",
                "💡 Ejemplo: PC Almacén, PC Mantenimiento Ocasional"
            ])
        ]
        
        for level_name, color, pregunta, items in levels:
            frame = ctk.CTkFrame(scroll, fg_color="#F5F5F5", corner_radius=10)
            frame.pack(fill="x", pady=10, padx=5)
            
            # Título con color
            label_title = ctk.CTkLabel(
                frame,
                text=level_name,
                font=("Segoe UI", 16, "bold"),
                text_color=color
            )
            label_title.pack(anchor="w", padx=20, pady=(15, 5))
            
            # Pregunta clave
            label_pregunta = ctk.CTkLabel(
                frame,
                text=f"➤ {pregunta}",
                font=("Segoe UI", 12, "bold"),
                text_color="#333333",
                anchor="w"
            )
            label_pregunta.pack(anchor="w", padx=20, pady=(5, 10), fill="x")
            
            # Items
            for item in items:
                label = ctk.CTkLabel(
                    frame,
                    text=f"  {item}",
                    font=("Segoe UI", 11),
                    text_color="#333333",
                    anchor="w",
                    justify="left"
                )
                label.pack(anchor="w", padx=25, pady=2, fill="x")
            
            ctk.CTkLabel(frame, text="", height=8).pack()
    
    def _create_confidentiality_tab_clean(self, parent):
        """Tab de confidencialidad con mensajes CLAROS y ÚTILES."""
        scroll = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=10, pady=10)
        
        title = ctk.CTkLabel(
            scroll,
            text="CONFIDENCIALIDAD - ¿Qué tipo de información maneja?",
            font=("Segoe UI", 17, "bold"),
            text_color=COLOR_VERDE_HOSPITAL
        )
        title.pack(pady=(0, 15))
        
        intro = ctk.CTkLabel(
            scroll,
            text="Pregúntate: ¿Qué tan sensible es la información en este equipo?",
            font=("Segoe UI", 12, "italic"),
            text_color="#666666"
        )
        intro.pack(pady=(0, 20))
        
        levels = [
            ("🔒 CLASIFICADA", "#6F42C1",
             "Información médica ultra-sensible - Máxima protección",
             [
                "✓ Usa en: Equipos que manejan historias clínicas completas",
                "✓ Usa en: Resultados VIH, salud mental, genética",
                "✓ Usa en: Datos financieros de pacientes",
                "⚠ Requiere: Cifrado OBLIGATORIO del disco",
                "⚠ Requiere: Auditoría permanente de accesos",
                "💡 Ejemplo: PC Psicología (salud mental), Servidor Historias Clínicas"
            ]),
            ("🔐 RESERVADA", "#DC3545",
             "Información protegida por ley - Alta protección",
             [
                "✓ Usa en: Equipos con identificación de pacientes",
                "✓ Usa en: Resultados de laboratorio, radiología",
                "✓ Usa en: Nómina, contabilidad sensible",
                "⚠ Requiere: Cifrado recomendado",
                "⚠ Requiere: Auditoría regular",
                "💡 Ejemplo: PC Laboratorio, PC Facturación, PC Nómina"
            ]),
            ("🔓 CONFIDENCIAL", "#FD7E14",
             "Información interna del hospital - Protección estándar",
             [
                "✓ Usa en: Procedimientos internos, manuales",
                "✓ Usa en: Estadísticas sin nombres de pacientes",
                "✓ Usa en: Informes de gestión",
                "⚠ Requiere: Protección estándar (usuario/contraseña)",
                "💡 Ejemplo: PC Calidad (informes), PC Planeación (estadísticas)"
            ]),
            ("🔓 INTERNA", "#20C997",
             "Información de trabajo diario - Protección básica",
             [
                "✓ Usa en: Todo el personal puede ver esta información",
                "✓ Usa en: Políticas, directorio, calendario",
                "✓ Usa en: Circulares, comunicados internos",
                "⚠ Requiere: Solo login básico",
                "💡 Ejemplo: PC Secretaria (circulares), PC Recepción (directorio)"
            ]),
            ("🌐 PÚBLICA", "#17A2B8",
             "Información sin restricciones - Sin protección especial",
             [
                "✓ Usa en: Información que puede ver cualquier persona",
                "✓ Usa en: Horarios, servicios, página web",
                "✓ Usa en: Información de contacto general",
                "⚠ No requiere protección especial",
                "💡 Ejemplo: PC Mercadeo (web pública), Kiosco Información"
            ])
        ]
        
        for level_name, color, pregunta, items in levels:
            frame = ctk.CTkFrame(scroll, fg_color="#F5F5F5", corner_radius=10)
            frame.pack(fill="x", pady=10, padx=5)
            
            label_title = ctk.CTkLabel(
                frame,
                text=level_name,
                font=("Segoe UI", 16, "bold"),
                text_color=color
            )
            label_title.pack(anchor="w", padx=20, pady=(15, 5))
            
            label_pregunta = ctk.CTkLabel(
                frame,
                text=f"➤ {pregunta}",
                font=("Segoe UI", 12, "bold"),
                text_color="#333333",
                anchor="w"
            )
            label_pregunta.pack(anchor="w", padx=20, pady=(5, 10), fill="x")
            
            for item in items:
                label = ctk.CTkLabel(
                    frame,
                    text=f"  {item}",
                    font=("Segoe UI", 11),
                    text_color="#333333",
                    anchor="w",
                    justify="left"
                )
                label.pack(anchor="w", padx=25, pady=2, fill="x")
            
            ctk.CTkLabel(frame, text="", height=8).pack()
    
    def _create_processes_tab_clean(self, parent):
        """Tab de procesos con mensajes CLAROS y ÚTILES."""
        scroll = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=10, pady=10)
        
        title = ctk.CTkLabel(
            scroll,
            text="PROCESO DEL EQUIPO - ¿Para qué se usa este equipo?",
            font=("Segoe UI", 17, "bold"),
            text_color=COLOR_VERDE_HOSPITAL
        )
        title.pack(pady=(0, 15))
        
        intro = ctk.CTkLabel(
            scroll,
            text="Pregúntate: ¿Qué tipo de trabajo hacen en este equipo?",
            font=("Segoe UI", 12, "italic"),
            text_color="#666666"
        )
        intro.pack(pady=(0, 20))
        
        processes = [
            ("🏥 MISIONAL", "#DC3545",
             "Atiende pacientes directamente - Razón de ser del hospital",
             [
                "✓ Usa en: Áreas que atienden, diagnostican o tratan pacientes",
                "✓ Usa en: Urgencias, UCI, Hospitalización, Quirófanos",
                "✓ Usa en: Consulta Externa, Laboratorio, Imágenes",
                "✓ Usa en: Farmacia, Bacteriología, Enfermería",
                "✗ NO uses en: Oficinas o áreas que NO atienden pacientes",
                "💡 Si en este equipo se trabaja CON pacientes → es MISIONAL"
            ]),
            ("📊 APOYO", "#17A2B8",
             "Soporta las operaciones - Servicios necesarios",
             [
                "✓ Usa en: Áreas administrativas y de soporte",
                "✓ Usa en: Facturación, Contabilidad, Recursos Humanos",
                "✓ Usa en: Sistemas/IT, Archivo, Almacén",
                "✓ Usa en: Mantenimiento, Servicios Generales, Seguridad",
                "✗ NO uses en: Áreas que atienden pacientes directamente",
                "💡 Si el trabajo es ADMINISTRATIVO u OPERATIVO → es APOYO"
            ]),
            ("🎯 ESTRATÉGICO", "#6F42C1",
             "Dirige el hospital - Toma decisiones",
             [
                "✓ Usa SOLO en: Dirección General, Subdirección",
                "✓ Usa en: Planeación Estratégica",
                "✓ Usa en: Junta Directiva",
                "✗ NO uses en: Personal operativo o coordinadores",
                "💡 Si toma decisiones de ALTO NIVEL → es ESTRATÉGICO"
            ]),
            ("📋 EVALUACIÓN", "#FFC107",
             "Controla y mejora - Mide resultados",
             [
                "✓ Usa en: Auditoría (interna y médica)",
                "✓ Usa en: Calidad, Control Interno",
                "✓ Usa en: Evaluación de Desempeño",
                "✗ NO uses en: Operaciones diarias normales",
                "💡 Si AUDITA o EVALÚA procesos → es EVALUACIÓN"
            ])
        ]
        
        for proc_name, color, pregunta, items in processes:
            frame = ctk.CTkFrame(scroll, fg_color="#F5F5F5", corner_radius=10)
            frame.pack(fill="x", pady=12, padx=5)
            
            label_title = ctk.CTkLabel(
                frame,
                text=proc_name,
                font=("Segoe UI", 16, "bold"),
                text_color=color
            )
            label_title.pack(anchor="w", padx=20, pady=(15, 5))
            
            label_pregunta = ctk.CTkLabel(
                frame,
                text=f"➤ {pregunta}",
                font=("Segoe UI", 12, "bold"),
                text_color="#333333",
                anchor="w"
            )
            label_pregunta.pack(anchor="w", padx=20, pady=(5, 10), fill="x")
            
            for item in items:
                label = ctk.CTkLabel(
                    frame,
                    text=f"  {item}",
                    font=("Segoe UI", 11),
                    text_color="#333333",
                    anchor="w",
                    justify="left"
                )
                label.pack(anchor="w", padx=25, pady=2, fill="x")
            
            ctk.CTkLabel(frame, text="", height=8).pack()
    
    def _create_systems_tab_clean(self, parent):
        """Tab de sistemas con mensajes CLAROS y ÚTILES."""
        scroll = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=10, pady=10)
        
        title = ctk.CTkLabel(
            scroll,
            text="SOFTWARE DEL EQUIPO - ¿Qué programas usa?",
            font=("Segoe UI", 17, "bold"),
            text_color=COLOR_VERDE_HOSPITAL
        )
        title.pack(pady=(0, 15))
        
        intro = ctk.CTkLabel(
            scroll,
            text="Marca lo que aplique para este equipo específico",
            font=("Segoe UI", 12, "italic"),
            text_color="#666666"
        )
        intro.pack(pady=(0, 20))
        
        systems = [
            ("💻 SIHOS - Sistema de Información Hospitalaria", "#007BFF",
             "El HIS principal del hospital - Historia clínica electrónica",
             [
                "🔵 LOCAL → Programa instalado en el equipo (versión completa)",
                "   • Puede hacer TODO: registrar, consultar, reportes, configurar",
                "   • Más rápido, funciona sin internet interno",
                "",
                "🌐 WEB → Entra por navegador (Chrome, Edge)",
                "   • Solo consultas y algunas funciones según usuario",
                "   • Requiere red funcionando",
                "",
                "❌ NO USA → Este equipo no necesita SIHOS",
                "   • Típico en: oficinas administrativas, almacén, mantenimiento"
            ]),
            ("💊 SIFAX - Sistema de Dispensación Farmacéutica", "#28A745",
             "Sistema de farmacia - Control de medicamentos",
             [
                "✓ SÍ → Este equipo tiene acceso a SIFAX",
                "   • Típico en: Farmacia, Enfermería, Urgencias",
                "",
                "✗ NO → Este equipo NO usa SIFAX",
                "   • Mayoría de equipos NO lo usan"
            ]),
            ("📄 Office Básico - Word, Excel, PowerPoint", "#FD7E14",
             "Suite de oficina Microsoft",
             [
                "✓ SÍ → Necesita Office para trabajar",
                "   • Hace documentos, reportes, presentaciones",
                "   • Mayoría de equipos administrativos",
                "",
                "✗ NO → Solo usa sistemas específicos",
                "   • Algunos equipos clínicos solo usan SIHOS"
            ]),
            ("🔧 Software Especializado", "#6F42C1",
             "Programas específicos del área",
             [
                "✓ SÍ → Tiene software especial instalado",
                "   • Ejemplos: PACS (imágenes), RIS (radiología), LIS (laboratorio)",
                "   • Programas contables, nómina, facturación",
                "   • ⚠ IMPORTANTE: Describe cuál en 'Descripción Software Esp.'",
                "",
                "✗ NO → Solo usa programas estándar",
                "   • SIHOS, Office, navegador web"
            ])
        ]
        
        for sys_name, color, pregunta, items in systems:
            frame = ctk.CTkFrame(scroll, fg_color="#F5F5F5", corner_radius=10)
            frame.pack(fill="x", pady=12, padx=5)
            
            label_title = ctk.CTkLabel(
                frame,
                text=sys_name,
                font=("Segoe UI", 15, "bold"),
                text_color=color
            )
            label_title.pack(anchor="w", padx=20, pady=(15, 5))
            
            label_pregunta = ctk.CTkLabel(
                frame,
                text=f"➤ {pregunta}",
                font=("Segoe UI", 11, "bold"),
                text_color="#333333",
                anchor="w"
            )
            label_pregunta.pack(anchor="w", padx=20, pady=(5, 10), fill="x")
            
            for item in items:
                if item == "":  # Línea en blanco
                    ctk.CTkLabel(frame, text="", height=3).pack()
                else:
                    label = ctk.CTkLabel(
                        frame,
                        text=f"  {item}",
                        font=("Segoe UI", 10),
                        text_color="#333333",
                        anchor="w",
                        justify="left"
                    )
                    label.pack(anchor="w", padx=25, pady=1, fill="x")
            
            ctk.CTkLabel(frame, text="", height=8).pack()
    
    def _create_others_tab_clean(self, parent):
        """Tab otros con mensajes CLAROS y ÚTILES."""
        scroll = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Sección 1: Horarios
        frame1 = ctk.CTkFrame(scroll, fg_color="#F5F5F5", corner_radius=10)
        frame1.pack(fill="x", pady=10, padx=5)
        
        ctk.CTkLabel(
            frame1,
            text="⏰ HORARIO DE USO - ¿Cuándo se usa este equipo?",
            font=("Segoe UI", 16, "bold"),
            text_color="#17A2B8"
        ).pack(anchor="w", padx=20, pady=(15, 5))
        
        ctk.CTkLabel(
            frame1,
            text="➤ Selecciona el horario típico de trabajo en este equipo",
            font=("Segoe UI", 11, "bold"),
            text_color="#333333",
            anchor="w"
        ).pack(anchor="w", padx=20, pady=(5, 10), fill="x")
        
        horarios = [
            "🔴 24/7 → TODO EL TIEMPO, sin parar",
            "   • Urgencias, UCI, Hospitalización, Enfermería 24h",
            "",
            "🟠 Lunes-Viernes 7am-7pm → Jornada extendida",
            "   • Consulta Externa, Facturación, Recepción",
            "",
            "🟡 Lunes-Viernes 7am-5pm → Horario administrativo normal",
            "   • Oficinas, Contabilidad, RH, Archivo",
            "",
            "🔵 Turnos rotativos → Personal por turnos 24h",
            "   • Enfermería por turnos, Personal asistencial rotativo",
            "",
            "🟢 Ocasional → Uso esporádico cuando se necesita",
            "   • Almacén, Mantenimiento, Bodega"
        ]
        
        for h in horarios:
            if h == "":
                ctk.CTkLabel(frame1, text="", height=3).pack()
            else:
                ctk.CTkLabel(
                    frame1,
                    text=f"  {h}",
                    font=("Segoe UI", 10),
                    text_color="#333333",
                    anchor="w"
                ).pack(anchor="w", padx=25, pady=1, fill="x")
        
        ctk.CTkLabel(frame1, text="", height=8).pack()
        
        # Sección 2: Estados Operativos
        frame2 = ctk.CTkFrame(scroll, fg_color="#F5F5F5", corner_radius=10)
        frame2.pack(fill="x", pady=10, padx=5)
        
        ctk.CTkLabel(
            frame2,
            text="⚙️ ESTADO OPERATIVO - ¿Cómo está funcionando?",
            font=("Segoe UI", 16, "bold"),
            text_color="#FD7E14"
        ).pack(anchor="w", padx=20, pady=(15, 5))
        
        ctk.CTkLabel(
            frame2,
            text="➤ Describe el estado actual del equipo",
            font=("Segoe UI", 11, "bold"),
            text_color="#333333",
            anchor="w"
        ).pack(anchor="w", padx=20, pady=(5, 10), fill="x")
        
        estados = [
            "✅ Operativo - Óptimo → Funciona perfecto, sin problemas",
            "",
            "⚠ Operativo - Regular → Funciona pero tiene fallas menores",
            "   • A veces lento, se cuelga ocasionalmente, pero sirve",
            "",
            "⚠ Operativo - Deficiente → Funciona mal, necesita reparación pronto",
            "   • Fallas frecuentes, muy lento, usuario se queja",
            "",
            "❌ Fuera de Servicio - Temporal → NO funciona, en reparación",
            "   • Equipo apagado esperando reparación o repuesto",
            "",
            "❌ Fuera de Servicio - Permanente → Dañado sin reparación",
            "   • Irreparable, se debe dar de baja",
            "",
            "🔧 En Reparación → Actualmente en mantenimiento",
            "   • Con técnico, en taller, en proceso de reparación",
            "",
            "📦 En Bodega → Guardado, no en uso actualmente",
            "   • Equipo funcionando pero almacenado, no asignado"
        ]
        
        for e in estados:
            if e == "":
                ctk.CTkLabel(frame2, text="", height=3).pack()
            else:
                ctk.CTkLabel(
                    frame2,
                    text=f"  {e}",
                    font=("Segoe UI", 10),
                    text_color="#333333",
                    anchor="w"
                ).pack(anchor="w", padx=25, pady=1, fill="x")
        
        ctk.CTkLabel(frame2, text="", height=8).pack()
        
        # Sección 3: Referencias Normativas
        frame3 = ctk.CTkFrame(scroll, fg_color="#F5F5F5", corner_radius=10)
        frame3.pack(fill="x", pady=10, padx=5)
        
        ctk.CTkLabel(
            frame3,
            text="📖 REFERENCIAS NORMATIVAS",
            font=("Segoe UI", 16, "bold"),
            text_color=COLOR_VERDE_HOSPITAL
        ).pack(anchor="w", padx=20, pady=(15, 10))
        
        normas = [
            "MinTIC - PETI → Plan Estratégico de TI para entidades públicas",
            "MinSalud Res. 2183/2004 → Estándares de calidad en salud",
            "Ley 1581/2012 → Protección de Datos Personales (Habeas Data)",
            "Decreto 1377/2013 → Reglamentación Ley 1581",
            "MECI → Modelo Estándar de Control Interno para el Estado"
        ]
        
        for n in normas:
            ctk.CTkLabel(
                frame3,
                text=f"  • {n}",
                font=("Segoe UI", 11),
                text_color="#333333",
                anchor="w"
            ).pack(anchor="w", padx=25, pady=2, fill="x")
        
        ctk.CTkLabel(frame3, text="", height=8).pack()
        """Crear contenido del tab de criticidad."""
        scroll = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Título
        title = ctk.CTkLabel(
            scroll,
            text="NIVEL DE CRITICIDAD (MinTIC - PETI)",
            font=("Segoe UI", 18, "bold"),
            text_color=COLOR_VERDE_HOSPITAL
        )
        title.pack(pady=(0, 20))
        
        levels = [
            ("🔴 CRÍTICO", "#DC3545", [
                "Equipos cuya falla DETIENE operaciones vitales del hospital",
                "Sistemas de información críticos: SIHOS, SIFAX",
                "Áreas: Urgencias, UCI, Quirófanos",
                "Disponibilidad: 24/7 sin interrupción",
                "Tiempo máximo inactividad: < 1 hora",
                "Ejemplos: PC Estación Enfermería UCI, Servidor SIHOS"
            ]),
            ("🟠 ALTO", "#FD7E14", [
                "Equipos importantes para operaciones misionales",
                "Afectan atención de pacientes directamente",
                "Áreas: Hospitalización, Laboratorio, Imágenes",
                "Disponibilidad: Horario extendido",
                "Tiempo máximo inactividad: < 4 horas",
                "Ejemplos: PC Laboratorio, PC Facturación, PC Farmacia"
            ]),
            ("🟡 MEDIO", "#FFC107", [
                "Equipos de apoyo administrativo",
                "Afectan eficiencia, NO bloquean operaciones",
                "Áreas: Contabilidad, RH, Calidad",
                "Disponibilidad: Horario laboral",
                "Tiempo máximo inactividad: < 24 horas",
                "Ejemplos: PC Contabilidad, PC Recursos Humanos"
            ]),
            ("🟢 BAJO", "#28A745", [
                "Equipos de uso ocasional o no prioritario",
                "NO afectan operaciones inmediatas",
                "Áreas: Almacén, Servicios Generales",
                "Disponibilidad: Ocasional",
                "Tiempo máximo inactividad: > 24 horas",
                "Ejemplos: PC Almacén, PC Mantenimiento"
            ])
        ]
        
        for level_name, color, items in levels:
            frame = ctk.CTkFrame(scroll, fg_color=color, corner_radius=10)
            frame.pack(fill="x", pady=10, padx=5)
            
            label_title = ctk.CTkLabel(
                frame,
                text=level_name,
                font=("Segoe UI", 16, "bold"),
                text_color="white"
            )
            label_title.pack(anchor="w", padx=15, pady=(15, 10))
            
            for item in items:
                label = ctk.CTkLabel(
                    frame,
                    text=f"  • {item}",
                    font=("Segoe UI", 12),
                    text_color="white",
                    anchor="w",
                    justify="left"
                )
                label.pack(anchor="w", padx=20, pady=2, fill="x")
            
            ctk.CTkLabel(frame, text="", height=10).pack()  # Spacer
    
    
    def start_automatic_collection(self):
        """Iniciar recopilación automática."""
        # Validar campos obligatorios
        required = ['tipo_equipo', 'area_servicio', 'ubicacion_especifica',
                   'responsable_custodio', 'proceso', 'uso_sihos', 'estado_operativo']
        
        missing = []
        for field in required:
            widget = self.manual_widgets.get(field)
            if widget:
                try:
                    # Verificar que widget existe antes de acceder
                    if hasattr(widget, 'winfo_exists') and widget.winfo_exists():
                        value = widget.get().strip()
                        if not value or value == "Seleccionar...":
                            missing.append(field.replace('_', ' ').title())
                    else:
                        # Widget no existe, considerar campo faltante
                        missing.append(field.replace('_', ' ').title())
                except:
                    # Error al acceder al widget, considerar campo faltante
                    missing.append(field.replace('_', ' ').title())
        
        if missing:
            messagebox.showwarning(
                "Campos Incompletos",
                f"Debes completar los siguientes campos:\n\n" + "\n".join(f"• {m}" for m in missing)
            )
            return
        
        # Guardar datos manuales con verificación
        self.equipment_data = {}
        for field_name, widget in self.manual_widgets.items():
            try:
                if hasattr(widget, 'winfo_exists') and widget.winfo_exists():
                    value = widget.get().strip()
                    if value and value != "Seleccionar...":
                        self.equipment_data[field_name] = value
            except:
                pass  # Si falla, simplemente no guarda ese campo
        
        # Mostrar ventana de progreso
        self.show_progress_window()
        
        # Ejecutar recopilación en thread separado
        thread = threading.Thread(target=self.collect_automatic_data)
        thread.daemon = True
        thread.start()
    
    def show_progress_window(self):
        """Mostrar ventana de progreso."""
        self.progress_window = ctk.CTkToplevel(self.root)
        self.progress_window.title("Recopilación Automática")
        self.progress_window.geometry("600x400")
        self.progress_window.transient(self.root)
        self.progress_window.grab_set()
        
        # Centrar
        self.progress_window.update_idletasks()
        x = (self.progress_window.winfo_screenwidth() // 2) - 300
        y = (self.progress_window.winfo_screenheight() // 2) - 200
        self.progress_window.geometry(f"600x400+{x}+{y}")
        
        label = ctk.CTkLabel(
            self.progress_window,
            text="🔄 Recopilando Datos Automáticos...",
            font=("Arial", 16, "bold")
        )
        label.pack(pady=20)
        
        self.progress_bar = ctk.CTkProgressBar(
            self.progress_window,
            mode="indeterminate",
            width=500
        )
        self.progress_bar.pack(pady=10)
        self.progress_bar.start()
        
        self.log_text = ctk.CTkTextbox(
            self.progress_window,
            width=550,
            height=250,
            font=("Consolas", 10)
        )
        self.log_text.pack(pady=10, padx=20)
    
    def log_progress(self, message):
        """Agregar mensaje al log."""
        if hasattr(self, 'log_text'):
            self.log_text.insert("end", message + "\n")
            self.log_text.see("end")
            self.root.update()
    
    def collect_automatic_data(self):
        """Recopilar datos automáticos (VERDES) con detección WMI real."""
        self.verde_data = {}
        
        # 1. Nombre del equipo
        self.log_progress("📋 Identificación del equipo...")
        self.verde_data['nombre_equipo'] = socket.gethostname()
        self.log_progress(f"   ✓ Nombre: {self.verde_data['nombre_equipo']}")
        
        # 2-4. Hardware con WMI
        self.log_progress("\n💻 Detectando hardware con WMI...")
        hw_info = detect_hardware_wmi()
        self.verde_data['marca'] = hw_info['marca']
        self.verde_data['modelo'] = hw_info['modelo']
        self.verde_data['serial'] = hw_info['serial']
        self.verde_data['tipo_disco'] = hw_info['tipo_disco']
        
        self.log_progress(f"   ✓ Marca: {self.verde_data['marca']}")
        self.log_progress(f"   ✓ Modelo: {self.verde_data['modelo']}")
        self.log_progress(f"   ✓ Serial: {self.verde_data['serial']}")
        self.log_progress(f"   ✓ Tipo Disco: {self.verde_data['tipo_disco']}")
        
        # Disco secundario (si existe)
        if hw_info['disco_secundario'] != 'No tiene':
            self.log_progress(f"\n💿 Disco Secundario Detectado:")
            self.log_progress(f"   ✓ Capacidad: {hw_info['disco_secundario']} GB")
            self.log_progress(f"   ✓ Tipo: {hw_info['tipo_disco_secundario']}")
            self.log_progress(f"   ✓ Serial: {hw_info['serial_disco_secundario']}")
            self.log_progress(f"   ✓ Marca: {hw_info['marca_disco_secundario']}")
            self.log_progress(f"   ✓ Modelo: {hw_info['modelo_disco_secundario']}")
        else:
            self.log_progress(f"\n💿 Disco Secundario: No detectado")
        
        # Guardar info disco secundario para validación mixta
        self.disco_secundario_info = {
            'disco_secundario': hw_info['disco_secundario'],
            'tipo_disco_secundario': hw_info['tipo_disco_secundario'],
            'serial_disco_secundario': hw_info['serial_disco_secundario'],
            'marca_disco_secundario': hw_info['marca_disco_secundario'],
            'modelo_disco_secundario': hw_info['modelo_disco_secundario']
        }
        
        # 5-7. Sistema Operativo
        self.log_progress("\n🪟 Sistema Operativo...")
        self.verde_data['sistema_operativo'] = f"{platform.system()} {platform.release()}"
        self.verde_data['arquitectura_so'] = "64 bits" if "64" in platform.machine() else "32 bits"
        self.verde_data['procesador'] = platform.processor() or "No detectado"
        
        self.log_progress(f"   ✓ SO: {self.verde_data['sistema_operativo']}")
        self.log_progress(f"   ✓ Arquitectura: {self.verde_data['arquitectura_so']}")
        self.log_progress(f"   ✓ Procesador: {self.verde_data['procesador'][:50]}...")
        
        # 8-9. RAM y Almacenamiento
        if HAS_PSUTIL:
            ram_gb = round(psutil.virtual_memory().total / (1024**3))
            self.verde_data['ram_gb'] = str(ram_gb)
            self.log_progress(f"   ✓ RAM: {ram_gb} GB")
            
            try:
                disk = psutil.disk_usage('C:\\')
                storage_gb = round(disk.total / (1024**3))
                self.verde_data['almacenamiento_gb'] = str(storage_gb)
                self.log_progress(f"   ✓ Almacenamiento: {storage_gb} GB")
            except:
                self.verde_data['almacenamiento_gb'] = "No detectado"
        else:
            self.verde_data['ram_gb'] = "Requiere psutil"
            self.verde_data['almacenamiento_gb'] = "Requiere psutil"
        
        # 10-15. Software Office
        self.log_progress("\n📦 Detectando Office...")
        office_version, office_licencia = detect_office_version()
        self.verde_data['version_office'] = office_version
        self.verde_data['licencia_office'] = office_licencia
        self.verde_data['uso_navegador_web'] = "Sí"
        
        self.log_progress(f"   ✓ Versión Office: {office_version}")
        self.log_progress(f"   ✓ Licencia Office: {office_licencia}")
        
        # Teams y Outlook
        teams, outlook = detect_office_apps()
        self.verde_data['uso_teams'] = teams
        self.verde_data['uso_outlook'] = outlook
        
        self.log_progress(f"   ✓ Teams: {teams}")
        self.log_progress(f"   ✓ Outlook: {outlook}")
        
        # 16-18. Licencia Windows
        self.log_progress("\n🔑 Detectando licencia Windows...")
        lic_info = detect_windows_license()
        self.verde_data['licencia_windows'] = lic_info['tipo']
        self.verde_data['key_windows'] = lic_info['key']
        self.verde_data['estado_licencia_windows'] = lic_info['estado']
        
        self.log_progress(f"   ✓ Licencia: {lic_info['tipo']}")
        self.log_progress(f"   ✓ Key (últimos 5): {lic_info['key']}")
        self.log_progress(f"   ✓ Estado: {lic_info['estado']}")
        
        # 19-20. Red
        self.log_progress("\n🌐 Red...")
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            s.connect(("8.8.8.8", 80))
            self.verde_data['direccion_ip'] = s.getsockname()[0]
            s.close()
            self.log_progress(f"   ✓ IP: {self.verde_data['direccion_ip']}")
        except:
            self.verde_data['direccion_ip'] = "No detectado"
            self.log_progress(f"   ⚠️  IP no detectada")
        
        self.verde_data['tipo_conexion'] = "Ethernet"  # Default, se puede mejorar
        
        # 21-23. Seguridad
        self.log_progress("\n🔒 Seguridad...")
        self.verde_data['antivirus_instalado'] = "Windows Defender"
        self.verde_data['windows_update_activo'] = "Sí"
        
        last_update = detect_last_windows_update()
        self.verde_data['ultima_act_windows'] = last_update
        
        self.log_progress(f"   ✓ Antivirus: Windows Defender")
        self.log_progress(f"   ✓ Última actualización: {last_update}")
        
        self.log_progress("\n✅ Recopilación automática completada")
        
        # Cerrar ventana de progreso
        self.progress_bar.stop()
        self.root.after(1000, lambda: self.progress_window.destroy())
        
        # Mostrar validación de campos mixtos
        self.root.after(1500, self.show_mixed_validation)
    
    def show_mixed_validation(self):
        """Mostrar ventana de validación de campos mixtos (AZULES) - MEJORADO."""
        validation_window = ctk.CTkToplevel(self.root)
        validation_window.title("Validación de Campos Mixtos")
        validation_window.geometry("900x600")
        validation_window.transient(self.root)
        validation_window.grab_set()
        
        # Centrar
        validation_window.update_idletasks()
        x = (validation_window.winfo_screenwidth() // 2) - 450
        y = (validation_window.winfo_screenheight() // 2) - 300
        validation_window.geometry(f"900x600+{x}+{y}")
        
        header = ctk.CTkLabel(
            validation_window,
            text="🔵 VALIDACIÓN DE CAMPOS MIXTOS (AZULES)",
            font=("Arial", 18, "bold"),
            text_color=COLOR_AZUL_HOSPITAL
        )
        header.pack(pady=20)
        
        info = ctk.CTkLabel(
            validation_window,
            text="Valida o corrige los siguientes campos detectados automáticamente:",
            font=("Arial", 13)
        )
        info.pack(pady=(0, 20))
        
        # Frame scrollable
        scroll_frame = ctk.CTkScrollableFrame(validation_window, width=850, height=380)
        scroll_frame.pack(pady=10, padx=25)
        
        # Campos mixtos
        self.mixed_widgets = {}
        
        mixed_fields = [
            # Disco secundario (si fue detectado)
            ("Almacenamiento Secundario (GB)", "disco_secundario", "entry", 
             self.disco_secundario_info.get('disco_secundario', 'No tiene') if hasattr(self, 'disco_secundario_info') else 'No tiene'),
            ("Tipo Disco Secundario", "tipo_disco_secundario", "combobox", 
             ['No tiene', 'HDD', 'SSD'] if hasattr(self, 'disco_secundario_info') and self.disco_secundario_info.get('disco_secundario') != 'No tiene' else ['No tiene']),
            ("Serial Disco Secundario", "serial_disco_secundario", "entry",
             self.disco_secundario_info.get('serial_disco_secundario', 'No tiene') if hasattr(self, 'disco_secundario_info') else 'No tiene'),
            ("Marca Disco Secundario", "marca_disco_secundario", "entry",
             self.disco_secundario_info.get('marca_disco_secundario', 'No tiene') if hasattr(self, 'disco_secundario_info') else 'No tiene'),
            ("Modelo Disco Secundario", "modelo_disco_secundario", "entry",
             self.disco_secundario_info.get('modelo_disco_secundario', 'No tiene') if hasattr(self, 'disco_secundario_info') else 'No tiene'),
            # Otros campos
            ("Switch / Puerto", "switch_puerto", "entry", "No detectado"),
            ("VLAN Asignada", "vlan_asignada", "entry", "No detectado"),
            ("ID AnyDesk", "id_anydesk", "entry", self.detect_anydesk()),
            ("Otro Acceso Remoto", "otro_acceso_remoto", "entry", "Ninguno"),
            ("Estado Antivirus", "estado_antivirus", "combobox", OPCIONES_ESTADO_ANTIVIRUS),
            ("Cifrado de Disco", "cifrado_disco", "combobox", OPCIONES_CIFRADO_DISCO),
            ("Tipo Usuario Local", "tipo_usuario_local", "combobox", OPCIONES_TIPO_USUARIO),
        ]
        
        for label_text, field_name, field_type, default in mixed_fields:
            field_frame = ctk.CTkFrame(scroll_frame, fg_color="transparent")
            field_frame.pack(fill="x", padx=15, pady=10)
            
            label = ctk.CTkLabel(
                field_frame,
                text=label_text,
                font=("Arial", 13, "bold"),
                width=250,
                anchor="w"
            )
            label.pack(side="left", padx=(0, 15))
            
            # CORRECCIÓN: Crear ComboBox correctamente cuando field_type == "combobox"
            if field_type == "combobox":
                # default es una lista de opciones
                widget = ctk.CTkComboBox(
                    field_frame,
                    values=default if isinstance(default, list) else ["No detectado"],
                    width=500,
                    font=("Arial", 12),
                    dropdown_font=("Arial", 11),
                    height=32
                )
                # Seleccionar primera opción por defecto
                if isinstance(default, list) and len(default) > 0:
                    widget.set(default[0])
            else:
                # Entry normal
                widget = ctk.CTkEntry(
                    field_frame,
                    width=500,
                    font=("Arial", 12),
                    height=32
                )
                widget.insert(0, str(default))
            
            widget.pack(side="left", fill="x", expand=True)
            self.mixed_widgets[field_name] = widget
        
        # Botón continuar
        btn_save = ctk.CTkButton(
            validation_window,
            text="✅ VALIDAR Y GUARDAR EN EXCEL",
            command=lambda: self.save_mixed_and_excel(validation_window),
            font=("Arial", 14, "bold"),
            fg_color=COLOR_VERDE_HOSPITAL,
            hover_color="#1F5A32",
            height=45
        )
        btn_save.pack(pady=20, padx=20, fill="x")
    
    def detect_anydesk(self):
        """Detectar ID de AnyDesk si está instalado."""
        try:
            # Ruta típica de AnyDesk
            anydesk_path = r"C:\Program Files (x86)\AnyDesk\AnyDesk.exe"
            if os.path.exists(anydesk_path):
                # Intentar obtener ID (simplificado)
                return "Instalado - Verificar ID"
            return "No instalado"
        except:
            return "No detectado"
    
    def save_mixed_and_excel(self, validation_window):
        """Guardar datos mixtos y todo en Excel."""
        # Obtener datos de campos mixtos
        self.azul_data = {}
        for field_name, widget in self.mixed_widgets.items():
            value = widget.get().strip()
            if value:
                self.azul_data[field_name] = value
        
        # Cerrar ventana de validación
        validation_window.destroy()
        
        # Guardar en Excel
        self.save_to_excel()
        
        # Mostrar mensaje de completado
        self.show_completion_message()
    
    def save_to_excel(self):
        """Guardar TODOS los datos en Excel (NARANJAS + VERDES + AZULES)."""
        if not HAS_OPENPYXL:
            messagebox.showerror("Error", "Necesitas instalar openpyxl")
            return
        
        try:
            wb = load_workbook(self.excel_path)
            ws = wb["Equipos de Cómputo"]
            
            # Verificar si estamos en modo ACTUALIZACIÓN o GUARDAR NUEVO
            if hasattr(self, 'equipo_update_row') and self.equipo_update_row:
                # MODO ACTUALIZACIÓN
                row = self.equipo_update_row
                codigo = self.equipo_update_code
                consecutive = int(codigo.split('-')[1])  # Extraer número del código EQC-0142
                
                # NO se modifican las columnas 1 y 2 (Consecutivo y Código ya existen)
                
            else:
                # MODO GUARDAR NUEVO
                row = self.current_row
                consecutive = row - 1
                
                # Columna 1: N° Consecutivo
                ws.cell(row=row, column=1, value=consecutive)
                
                # Columna 2: Código Inventario
                ws.cell(row=row, column=2, value=f"EQC-{consecutive:04d}")
            
            # ===== COLUMNA 3: Nombre Equipo (VERDE) =====
            ws.cell(row=row, column=3, value=self.verde_data.get('nombre_equipo', ''))
            
            # ===== COLUMNAS 4-27: NARANJAS (24 campos) =====
            col = 4
            naranja_fields = [
                'tipo_equipo', 'area_servicio', 'ubicacion_especifica',
                'responsable_custodio', 'proceso', 'uso_sihos', 'uso_sifax',
                'uso_office_basico', 'software_especializado', 'descripcion_software',
                'funcion_principal', 'criticidad', 'confidencialidad',
                'horario_uso', 'estado_operativo', 'fecha_adquisicion',
                'valor_adquisicion', 'fecha_venc_garantia', 'observaciones_tecnicas',
                'fecha_exp_antivirus', 'periodicidad_mtto', 'responsable_mtto',
                'ultimo_mantenimiento', 'tipo_ultimo_mtto'
            ]
            
            for field in naranja_fields:
                value = self.equipment_data.get(field, '')
                ws.cell(row=row, column=col, value=value)
                col += 1
            
            # ===== COLUMNAS 28-48: VERDES (21 campos más) =====
            verde_fields = [
                'marca', 'modelo', 'serial', 'sistema_operativo', 'arquitectura_so',
                'procesador', 'ram_gb', 'almacenamiento_gb', 'tipo_disco',
                'uso_navegador_web', 'version_office', 'licencia_office',
                'uso_teams', 'uso_outlook', 'licencia_windows', 'key_windows',
                'estado_licencia_windows', 'direccion_ip', 'tipo_conexion',
                'antivirus_instalado', 'ultima_act_windows', 'windows_update_activo'
            ]
            
            for field in verde_fields:
                value = self.verde_data.get(field, '')
                ws.cell(row=row, column=col, value=value)
                col += 1
            
            # ===== COLUMNAS 49-61: AZULES (12 campos mixtos con disco secundario) =====
            azul_fields = [
                # Disco secundario (5 campos)
                'disco_secundario', 'tipo_disco_secundario', 'serial_disco_secundario',
                'marca_disco_secundario', 'modelo_disco_secundario',
                # Otros campos mixtos (7 campos)
                'switch_puerto', 'vlan_asignada', 'id_anydesk',
                'otro_acceso_remoto', 'estado_antivirus',
                'cifrado_disco', 'tipo_usuario_local'
            ]
            
            for field in azul_fields:
                value = self.azul_data.get(field, '')
                ws.cell(row=row, column=col, value=value)
                col += 1
            
            # ===== COLUMNA 62: Antigüedad (CALCULADA - BLANCA) =====
            # Calcular antigüedad si hay fecha de adquisición
            fecha_adq = self.equipment_data.get('fecha_adquisicion', '')
            if fecha_adq:
                try:
                    fecha = datetime.strptime(fecha_adq, '%Y-%m-%d')
                    hoy = datetime.now()
                    antiguedad = round((hoy - fecha).days / 365.25, 1)
                    ws.cell(row=row, column=col, value=antiguedad)
                except:
                    ws.cell(row=row, column=col, value='')
            
            # Guardar
            wb.save(self.excel_path)
            wb.close()
            
            # Verificar si fue actualización o guardar nuevo
            if hasattr(self, 'equipo_update_row') and self.equipo_update_row:
                # MODO ACTUALIZACIÓN - Mensaje y reseteo completo
                messagebox.showinfo("Éxito", f"✅ Equipo {codigo} actualizado correctamente (datos completos)")
                
                # Reseteo completo usando función unificada
                self.reset_after_update_equipos()
                
            else:
                # MODO GUARDAR NUEVO - Flujo normal
                messagebox.showinfo("Éxito", f"✅ Equipo guardado: EQC-{consecutive:04d}")
                
                # Actualizar para siguiente equipo
                self.current_row += 1
                
                # IMPORTANTE: Recrear formulario completo para que botón automático siempre funcione
                self.root.after(100, self.show_manual_form_in_container)
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al guardar en Excel:\n{e}")
    
    def save_equipo_manual_only(self):
        """Guardar solo datos manuales del equipo (sin detección automática)."""
        if not HAS_OPENPYXL:
            messagebox.showerror("Error", "Necesitas instalar openpyxl")
            return
        
        # Verificar si es actualización o nuevo registro
        if hasattr(self, 'equipo_update_row') and self.equipo_update_row:
            # MODO ACTUALIZACIÓN
            self.save_equipo_update()
            return
        
        # MODO GUARDAR NUEVO
        try:
            # PRIMERO: Leer todos los valores ANTES de hacer cualquier cosa
            datos_guardados = {}
            naranja_fields = [
                'tipo_equipo', 'area_servicio', 'ubicacion_especifica',
                'responsable_custodio', 'proceso', 'uso_sihos', 'uso_sifax',
                'uso_office_basico', 'software_especializado', 'descripcion_software',
                'funcion_principal', 'criticidad', 'confidencialidad',
                'horario_uso', 'estado_operativo', 'fecha_adquisicion',
                'valor_adquisicion', 'fecha_venc_garantia', 'observaciones_tecnicas',
                'fecha_exp_antivirus', 'periodicidad_mtto', 'responsable_mtto',
                'ultimo_mantenimiento', 'tipo_ultimo_mtto'
            ]
            
            for field in naranja_fields:
                try:
                    if field in self.manual_widgets:
                        widget = self.manual_widgets[field]
                        if hasattr(widget, 'winfo_exists') and widget.winfo_exists():
                            if isinstance(widget, ctk.CTkEntry):
                                datos_guardados[field] = widget.get()
                            elif isinstance(widget, ctk.CTkComboBox):
                                datos_guardados[field] = widget.get()
                            else:
                                datos_guardados[field] = ''
                        else:
                            datos_guardados[field] = ''
                    else:
                        datos_guardados[field] = ''
                except:
                    datos_guardados[field] = ''
            
            # SEGUNDO: Guardar en Excel
            wb = load_workbook(self.excel_path)
            ws = wb["Equipos de Cómputo"]
            
            row = self.current_row
            consecutive = row - 1
            
            # Columna 1: N° Consecutivo
            ws.cell(row=row, column=1, value=consecutive)
            
            # Columna 2: Código Inventario
            ws.cell(row=row, column=2, value=f"EQC-{consecutive:04d}")
            
            # Columna 3: Nombre Equipo (vacío en guardado manual)
            ws.cell(row=row, column=3, value='')
            
            # ===== COLUMNAS 4-27: NARANJAS (24 campos) =====
            col = 4
            for field in naranja_fields:
                value = datos_guardados.get(field, '')
                ws.cell(row=row, column=col, value=value)
                col += 1
            
            # Columnas 28-61: vacías (verdes y azules incluyendo disco secundario)
            for i in range(28, 62):
                ws.cell(row=row, column=i, value='')
            
            # Guardar
            wb.save(self.excel_path)
            wb.close()
            
            messagebox.showinfo("Éxito", f"✅ Equipo guardado (solo datos manuales): EQC-{consecutive:04d}")
            
            # Actualizar para siguiente equipo
            self.current_row += 1
            
            # IMPORTANTE: Recrear formulario para que el botón automático funcione
            self.root.after(100, self.show_manual_form_in_container)
            
            # TERCERO: Actualizar título
            if hasattr(self, 'equipo_form_frame'):
                try:
                    if hasattr(self.equipo_form_frame, 'winfo_exists') and self.equipo_form_frame.winfo_exists():
                        self.equipo_form_frame.configure(
                            label_text=f"📝 DATOS MANUALES - Equipo #{self.current_row-1} (Código EQC-{self.current_row-1:04d})"
                        )
                except:
                    pass
            
            # CUARTO: Limpiar campos selectivamente (con verificación)
            campos_a_mantener = ['area_servicio', 'proceso', 'responsable_custodio', 'periodicidad_mtto', 'responsable_mtto']
            
            for key, widget in self.manual_widgets.items():
                if key not in campos_a_mantener:
                    try:
                        if hasattr(widget, 'winfo_exists') and widget.winfo_exists():
                            if isinstance(widget, ctk.CTkEntry):
                                widget.delete(0, "end")
                            elif isinstance(widget, ctk.CTkComboBox):
                                widget.set("")
                    except:
                        pass
                    
        except Exception as e:
            messagebox.showerror("Error", f"Error al guardar:\n{e}")
    
    def update_equipo_computo(self):
        """Actualizar equipo de cómputo existente."""
        if not self.excel_path:
            messagebox.showerror("Error", "No hay Excel cargado")
            return
        
        dialog = ctk.CTkToplevel(self.root)
        dialog.title("Actualizar Equipo de Cómputo")
        dialog.geometry("400x200")
        dialog.transient(self.root)
        dialog.grab_set()
        
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - 200
        y = (dialog.winfo_screenheight() // 2) - 100
        dialog.geometry(f"400x200+{x}+{y}")
        
        ctk.CTkLabel(
            dialog,
            text="Ingresa el código del equipo a actualizar:",
            font=("Segoe UI", 13)
        ).pack(pady=20)
        
        entry_codigo = ctk.CTkEntry(
            dialog,
            width=200,
            height=40,
            font=("Segoe UI", 12),
            placeholder_text="Ej: EQC-0142"
        )
        entry_codigo.pack(pady=10)
        entry_codigo.focus()
        
        def buscar_y_cargar():
            codigo = entry_codigo.get().strip().upper()
            if not codigo:
                messagebox.showerror("Error", "Debes ingresar un código")
                return
            
            try:
                wb = load_workbook(self.excel_path)
                ws = wb["Equipos de Cómputo"]
                
                found = False
                target_row = None
                
                for row in range(2, 500):
                    cell_value = ws.cell(row=row, column=2).value
                    if cell_value and cell_value.upper() == codigo:
                        found = True
                        target_row = row
                        break
                
                if not found:
                    wb.close()
                    messagebox.showerror("Error", f"No se encontró el código {codigo}")
                    return
                
                # Cargar datos NARANJAS (columnas 4-27)
                naranja_fields = [
                    'tipo_equipo', 'area_servicio', 'ubicacion_especifica',
                    'responsable_custodio', 'proceso', 'uso_sihos', 'uso_sifax',
                    'uso_office_basico', 'software_especializado', 'descripcion_software',
                    'funcion_principal', 'criticidad', 'confidencialidad',
                    'horario_uso', 'estado_operativo', 'fecha_adquisicion',
                    'valor_adquisicion', 'fecha_venc_garantia', 'observaciones_tecnicas',
                    'fecha_exp_antivirus', 'periodicidad_mtto', 'responsable_mtto',
                    'ultimo_mantenimiento', 'tipo_ultimo_mtto'
                ]
                
                col = 4
                for field in naranja_fields:
                    value = ws.cell(row=target_row, column=col).value or ''
                    self.equipment_data[field] = value
                    
                    # Cargar en widgets con verificación
                    if field in self.manual_widgets:
                        try:
                            widget = self.manual_widgets[field]
                            if hasattr(widget, 'winfo_exists') and widget.winfo_exists():
                                if isinstance(widget, ctk.CTkEntry):
                                    widget.delete(0, "end")
                                    widget.insert(0, value)
                                elif isinstance(widget, ctk.CTkComboBox):
                                    widget.set(value)
                        except:
                            pass  # Si falla, continuar con el siguiente
                    
                    col += 1
                
                wb.close()
                
                self.equipo_update_code = codigo
                self.equipo_update_row = target_row
                
                # CAMBIAR TÍTULO A MODO ACTUALIZACIÓN (con verificación)
                if hasattr(self, 'equipo_form_frame'):
                    try:
                        if hasattr(self.equipo_form_frame, 'winfo_exists') and self.equipo_form_frame.winfo_exists():
                            self.equipo_form_frame.configure(
                                label_text=f"🔄 ACTUALIZANDO EQUIPO - Código: {codigo}"
                            )
                    except:
                        pass
                
                # CAMBIAR TEXTO DEL BOTÓN (con verificación)
                if hasattr(self, 'btn_save_equipo'):
                    try:
                        if hasattr(self.btn_save_equipo, 'winfo_exists') and self.btn_save_equipo.winfo_exists():
                            self.btn_save_equipo.configure(text="🔄 ACTUALIZAR EQUIPO")
                    except:
                        pass
                
                dialog.destroy()
                
                if messagebox.askyesno(
                    "Confirmar Actualización",
                    f"⚠️ ¿Estás seguro de actualizar {codigo}?\n\n"
                    f"Los datos actuales se han cargado.\n"
                    f"Modifica los campos necesarios y presiona ACTUALIZAR EQUIPO."
                ):
                    messagebox.showinfo("Listo", f"✅ Datos cargados de {codigo}\n\nModifica los campos y presiona ACTUALIZAR EQUIPO.")
                
            except Exception as e:
                messagebox.showerror("Error", f"Error al buscar:\n{e}")
        
        btn_buscar = ctk.CTkButton(
            dialog,
            text="🔍 BUSCAR Y CARGAR",
            command=buscar_y_cargar,
            font=("Segoe UI", 13, "bold"),
            fg_color=COLOR_VERDE_HOSPITAL,
            height=40
        )
        btn_buscar.pack(pady=10)
        entry_codigo.bind("<Return>", lambda e: buscar_y_cargar())
    
    def reset_after_update_equipos(self):
        """Reseteo completo después de actualizar un equipo (volver al estado inicial)."""
        # 1. Limpiar variables de modo actualización
        self.equipo_update_row = None
        self.equipo_update_code = None
        
        # 2. Restaurar título al SIGUIENTE equipo nuevo (con verificación)
        next_code = f"EQC-{self.current_row-1:04d}"
        if hasattr(self, 'equipo_form_frame'):
            try:
                if hasattr(self.equipo_form_frame, 'winfo_exists') and self.equipo_form_frame.winfo_exists():
                    self.equipo_form_frame.configure(
                        label_text=f"📝 DATOS MANUALES - Equipo #{self.current_row-1} (Código {next_code})"
                    )
            except:
                pass
        
        # 3. Restaurar BOTÓN a estado normal (con verificación)
        if hasattr(self, 'btn_save_equipo'):
            try:
                if hasattr(self.btn_save_equipo, 'winfo_exists') and self.btn_save_equipo.winfo_exists():
                    self.btn_save_equipo.configure(text="💾 GUARDAR NUEVO (Solo Datos Manuales)")
            except:
                pass
        
        # 4. Limpiar TODOS los datos
        self.equipment_data = {}
        self.verde_data = {}
        self.azul_data = {}
        
        # 5. Limpiar TODOS los widgets (con verificación)
        for key, widget in self.manual_widgets.items():
            try:
                if hasattr(widget, 'winfo_exists') and widget.winfo_exists():
                    if isinstance(widget, ctk.CTkEntry):
                        widget.delete(0, "end")
                    elif isinstance(widget, ctk.CTkComboBox):
                        widget.set("")
            except:
                pass
    
    def save_equipo_update(self):
        """Guardar actualización de equipo de cómputo (solo datos manuales)."""
        try:
            wb = load_workbook(self.excel_path)
            ws = wb["Equipos de Cómputo"]
            
            row = self.equipo_update_row
            codigo = self.equipo_update_code
            
            # Actualizar NARANJAS (columnas 4-27)
            col = 4
            naranja_fields = [
                'tipo_equipo', 'area_servicio', 'ubicacion_especifica',
                'responsable_custodio', 'proceso', 'uso_sihos', 'uso_sifax',
                'uso_office_basico', 'software_especializado', 'descripcion_software',
                'funcion_principal', 'criticidad', 'confidencialidad',
                'horario_uso', 'estado_operativo', 'fecha_adquisicion',
                'valor_adquisicion', 'fecha_venc_garantia', 'observaciones_tecnicas',
                'fecha_exp_antivirus', 'periodicidad_mtto', 'responsable_mtto',
                'ultimo_mantenimiento', 'tipo_ultimo_mtto'
            ]
            
            # Leer de widgets directamente con verificación (igual que save_equipo_manual_only)
            for field in naranja_fields:
                value = ''
                if field in self.manual_widgets:
                    try:
                        widget = self.manual_widgets[field]
                        if hasattr(widget, 'winfo_exists') and widget.winfo_exists():
                            if isinstance(widget, ctk.CTkEntry):
                                value = widget.get()
                            elif isinstance(widget, ctk.CTkComboBox):
                                value = widget.get()
                    except:
                        pass
                ws.cell(row=row, column=col, value=value)
                col += 1
            
            wb.save(self.excel_path)
            wb.close()
            
            messagebox.showinfo("Éxito", f"✅ Equipo {codigo} actualizado correctamente")
            
            # Reseteo completo usando función unificada
            self.reset_after_update_equipos()
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al actualizar:\n{e}")
    
    def show_completion_message(self):
        """Mostrar mensaje de equipo completado."""
        consecutive = self.current_row - 1
        code = f"EQC-{consecutive:04d}"
        nombre = self.verde_data.get('nombre_equipo', 'N/A')
        area = self.equipment_data.get('area_servicio', 'N/A')
        
        # Ventana de completado
        completion_window = ctk.CTkToplevel(self.root)
        completion_window.title("Equipo Completado")
        completion_window.geometry("500x450")
        completion_window.transient(self.root)
        completion_window.grab_set()
        
        # Centrar
        completion_window.update_idletasks()
        x = (completion_window.winfo_screenwidth() // 2) - 250
        y = (completion_window.winfo_screenheight() // 2) - 225
        completion_window.geometry(f"500x450+{x}+{y}")
        
        # Icono de éxito
        success_label = ctk.CTkLabel(
            completion_window,
            text="✅",
            font=("Arial", 60)
        )
        success_label.pack(pady=20)
        
        title = ctk.CTkLabel(
            completion_window,
            text="EQUIPO COMPLETADO Y GUARDADO",
            font=("Arial", 16, "bold"),
            text_color=COLOR_VERDE_HOSPITAL
        )
        title.pack(pady=10)
        
        info_frame = ctk.CTkFrame(completion_window, fg_color="transparent")
        info_frame.pack(pady=20)
        
        info_text = f"""
N° Consecutivo: {consecutive}
Código: {code}
Nombre: {nombre}
Área: {area}

✓ Datos manuales (24 campos): Guardados
✓ Datos automáticos (22 campos): Guardados
✓ Datos mixtos (7 campos): Guardados
✓ Excel actualizado correctamente

Total: 56 columnas completas
        """
        
        info = ctk.CTkLabel(
            info_frame,
            text=info_text,
            font=("Arial", 11),
            justify="left"
        )
        info.pack()
        
        # Instrucciones
        instructions = ctk.CTkLabel(
            completion_window,
            text="Cierra la aplicación para hacer salida segura del USB\ny proceder al siguiente equipo.",
            font=("Arial", 11),
            text_color="gray"
        )
        instructions.pack(pady=10)
        
        # Botones
        btn_frame = ctk.CTkFrame(completion_window, fg_color="transparent")
        btn_frame.pack(pady=20)
        
        btn_next = ctk.CTkButton(
            btn_frame,
            text="➡️ Siguiente Equipo",
            command=lambda: self.next_equipment(completion_window),
            fg_color=COLOR_VERDE_HOSPITAL,
            width=200
        )
        btn_next.pack(side="left", padx=10)
        
        btn_close = ctk.CTkButton(
            btn_frame,
            text="❌ Cerrar",
            command=self.root.quit,
            fg_color=COLOR_ERROR,
            hover_color="#A02828",
            width=200
        )
        btn_close.pack(side="left", padx=10)
    
    def next_equipment(self, completion_window):
        """Ir a siguiente equipo."""
        completion_window.destroy()
        self.current_row += 1
        
        # Limpiar datos
        self.equipment_data = {}
        self.verde_data = {}
        self.azul_data = {}
        
        # Mostrar nuevo formulario
        self.show_manual_form()


    
    # ========================================================================
    # FORMULARIOS DE OTROS TIPOS DE INVENTARIO
    # ========================================================================
    
    def create_impresoras_form(self, parent_tab):
        """Formulario para Impresoras y Escáneres."""
        scroll = ctk.CTkScrollableFrame(
            parent_tab,
            fg_color="#FAFAFA",
            label_fg_color=COLOR_VERDE_HOSPITAL,
            label_text_color="white",
            label_font=("Segoe UI", 15, "bold")
        )
        scroll.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Detectar siguiente código automáticamente
        next_code = self.detect_next_code("Impresoras y Escáneres", "IMP")
        scroll.configure(label_text=f"🖨️ IMPRESORAS Y ESCÁNERES - Código: {next_code}")
        
        # Widgets para almacenar referencias
        self.imp_widgets = {}
        self.imp_next_code = next_code  # Guardar código para usar al guardar
        self.imp_scroll = scroll  # Guardar referencia al scroll para actualizar título
        
        # Campos
        fields = [
            ("Código Equipo Asignado *", "codigo_asignado", "entry"),
            ("Tipo *", "tipo", "combobox", TIPOS_IMPRESORA),
            ("Marca *", "marca", "combobox", MARCAS_IMPRESORA),
            ("Modelo", "modelo", "entry"),
            ("Serial", "serial", "entry"),
            ("Área / Servicio *", "area", "combobox", AREAS_SERVICIO),
            ("Ubicación Específica", "ubicacion", "entry"),
            ("Función *", "funcion", "combobox", FUNCIONES_IMPRESORA),
            ("Dirección IP", "ip", "entry"),
            ("Estado Operativo *", "estado", "combobox", ESTADOS_IMPRESORA),
            ("Fecha de Adquisición (YYYY-MM-DD)", "fecha_adq", "entry"),
            ("Valor de Adquisición (COP)", "valor", "entry"),
            ("Observaciones", "observaciones", "entry"),
        ]
        
        for field_data in fields:
            if len(field_data) == 4:
                label, key, field_type, options = field_data
                widget = self.create_form_field(scroll, label, key, field_type, options)
            else:
                label, key, field_type = field_data
                widget = self.create_form_field(scroll, label, key, field_type, None)
            self.imp_widgets[key] = widget
        
        # Frame para botones
        btn_frame = ctk.CTkFrame(scroll, fg_color="transparent")
        btn_frame.pack(pady=30)
        
        # Botón guardar nuevo - Referencia global
        self.btn_save_imp = ctk.CTkButton(
            btn_frame,
            text="💾 GUARDAR NUEVO",
            command=self.save_impresora,
            font=("Segoe UI", 14, "bold"),
            fg_color=COLOR_VERDE_HOSPITAL,
            hover_color="#1F5039",
            height=50,
            width=250
        )
        self.btn_save_imp.pack(side="left", padx=10)
        
        # Botón actualizar existente
        btn_update = ctk.CTkButton(
            btn_frame,
            text="🔄 ACTUALIZAR EXISTENTE",
            command=self.update_impresora,
            font=("Segoe UI", 14, "bold"),
            fg_color="#2196F3",
            hover_color="#1976D2",
            height=50,
            width=250
        )
        btn_update.pack(side="left", padx=10)
    
    def save_impresora(self):
        """Guardar impresora en Excel."""
        if not self.excel_path:
            messagebox.showerror("Error", "No hay Excel cargado")
            return
        
        try:
            wb = load_workbook(self.excel_path)
            
            # Verificar que la hoja existe
            if "Impresoras y Escáneres" not in wb.sheetnames:
                wb.close()
                messagebox.showerror("Error", "La hoja 'Impresoras y Escáneres' no existe en el Excel.\n\nCrea esta hoja primero.")
                return
            
            ws = wb["Impresoras y Escáneres"]
            
            # Verificar si es actualización o nuevo registro
            if hasattr(self, 'imp_update_row') and self.imp_update_row:
                # MODO ACTUALIZACIÓN
                row = self.imp_update_row
                codigo = self.imp_update_code
                
                # Actualizar datos en la fila existente (NO modificar columnas 1 y 2)
                ws.cell(row=row, column=3, value=self.imp_widgets["codigo_asignado"].get())
                ws.cell(row=row, column=4, value=self.imp_widgets["tipo"].get())
                ws.cell(row=row, column=5, value=self.imp_widgets["marca"].get())
                ws.cell(row=row, column=6, value=self.imp_widgets["modelo"].get())
                ws.cell(row=row, column=7, value=self.imp_widgets["serial"].get())
                ws.cell(row=row, column=8, value=self.imp_widgets["area"].get())
                ws.cell(row=row, column=9, value=self.imp_widgets["ubicacion"].get())
                ws.cell(row=row, column=10, value=self.imp_widgets["funcion"].get())
                ws.cell(row=row, column=11, value=self.imp_widgets["ip"].get())
                ws.cell(row=row, column=12, value=self.imp_widgets["estado"].get())
                ws.cell(row=row, column=13, value=self.imp_widgets["fecha_adq"].get())
                ws.cell(row=row, column=14, value=self.imp_widgets["valor"].get())
                ws.cell(row=row, column=15, value=self.imp_widgets["observaciones"].get())
                
                wb.save(self.excel_path)
                wb.close()
                
                messagebox.showinfo("Éxito", f"✅ Impresora {codigo} actualizada correctamente")
                
                # Limpiar modo actualización
                self.imp_update_row = None
                self.imp_update_code = None
                
                # Volver a título normal
                next_code = self.detect_next_code("Impresoras y Escáneres", "IMP")
                self.imp_next_code = next_code
                self.imp_scroll.configure(label_text=f"🖨️ IMPRESORAS Y ESCÁNERES - Código: {next_code}")
                
                # Restaurar texto del botón
                if hasattr(self, 'btn_save_imp'):
                    self.btn_save_imp.configure(text="💾 GUARDAR NUEVO")
                
                # Limpiar todos los campos
                for key, widget in self.imp_widgets.items():
                    if isinstance(widget, ctk.CTkEntry):
                        widget.delete(0, "end")
                    elif isinstance(widget, ctk.CTkComboBox):
                        widget.set("")
                
            else:
                # MODO GUARDAR NUEVO
                # Buscar el último consecutivo real
                last_consecutive = 0
                for row in range(2, 200):
                    value = ws.cell(row=row, column=1).value
                    if value is not None:
                        try:
                            consecutivo = int(value)
                            if consecutivo > last_consecutive:
                                last_consecutive = consecutivo
                        except:
                            pass
                
                # Siguiente consecutivo
                next_consecutive = last_consecutive + 1
                
                # Buscar primera fila vacía
                next_row = 2
                for row in range(2, 200):
                    if ws.cell(row=row, column=2).value is None:
                        next_row = row
                        break
                
                # Guardar datos
                ws.cell(row=next_row, column=1, value=next_consecutive)
                ws.cell(row=next_row, column=2, value=f"IMP-{next_consecutive:04d}")
                ws.cell(row=next_row, column=3, value=self.imp_widgets["codigo_asignado"].get())
                ws.cell(row=next_row, column=4, value=self.imp_widgets["tipo"].get())
                ws.cell(row=next_row, column=5, value=self.imp_widgets["marca"].get())
                ws.cell(row=next_row, column=6, value=self.imp_widgets["modelo"].get())
                ws.cell(row=next_row, column=7, value=self.imp_widgets["serial"].get())
                ws.cell(row=next_row, column=8, value=self.imp_widgets["area"].get())
                ws.cell(row=next_row, column=9, value=self.imp_widgets["ubicacion"].get())
                ws.cell(row=next_row, column=10, value=self.imp_widgets["funcion"].get())
                ws.cell(row=next_row, column=11, value=self.imp_widgets["ip"].get())
                ws.cell(row=next_row, column=12, value=self.imp_widgets["estado"].get())
                ws.cell(row=next_row, column=13, value=self.imp_widgets["fecha_adq"].get())
                ws.cell(row=next_row, column=14, value=self.imp_widgets["valor"].get())
                ws.cell(row=next_row, column=15, value=self.imp_widgets["observaciones"].get())
                
                wb.save(self.excel_path)
                wb.close()
                
                messagebox.showinfo("Éxito", f"✅ Impresora guardada: IMP-{next_consecutive:04d}")
                
                # Detectar siguiente código y actualizar título
                next_code = self.detect_next_code("Impresoras y Escáneres", "IMP")
                self.imp_next_code = next_code
                self.imp_scroll.configure(label_text=f"🖨️ IMPRESORAS Y ESCÁNERES - Código: {next_code}")
                
                # Limpiar campos selectivamente (mantener área)
                campos_a_mantener = ['area']
                
                for key, widget in self.imp_widgets.items():
                    if key not in campos_a_mantener:
                        if isinstance(widget, ctk.CTkEntry):
                            widget.delete(0, "end")
                        elif isinstance(widget, ctk.CTkComboBox):
                            widget.set("")
                    
        except Exception as e:
            messagebox.showerror("Error", f"❌ Error al guardar impresora:\n\n{str(e)}\n\nVerifica que la hoja 'Impresoras y Escáneres' existe.")
            import traceback
            traceback.print_exc()
    
    def update_impresora(self):
        """Actualizar impresora existente en Excel."""
        if not self.excel_path:
            messagebox.showerror("Error", "No hay Excel cargado")
            return
        
        # Ventana para pedir código
        dialog = ctk.CTkToplevel(self.root)
        dialog.title("Actualizar Impresora")
        dialog.geometry("400x200")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Centrar
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - 200
        y = (dialog.winfo_screenheight() // 2) - 100
        dialog.geometry(f"400x200+{x}+{y}")
        
        ctk.CTkLabel(
            dialog,
            text="Ingresa el código de la impresora a actualizar:",
            font=("Segoe UI", 13)
        ).pack(pady=20)
        
        entry_codigo = ctk.CTkEntry(
            dialog,
            width=200,
            height=40,
            font=("Segoe UI", 12),
            placeholder_text="Ej: IMP-0026"
        )
        entry_codigo.pack(pady=10)
        entry_codigo.focus()
        
        def buscar_y_cargar():
            codigo = entry_codigo.get().strip().upper()
            if not codigo:
                messagebox.showerror("Error", "Debes ingresar un código")
                return
            
            try:
                wb = load_workbook(self.excel_path)
                ws = wb["Impresoras y Escáneres"]
                
                # Buscar el código en la columna 2
                found = False
                target_row = None
                
                for row in range(2, 200):
                    cell_value = ws.cell(row=row, column=2).value
                    if cell_value and cell_value.upper() == codigo:
                        found = True
                        target_row = row
                        break
                
                if not found:
                    wb.close()
                    messagebox.showerror("Error", f"No se encontró el código {codigo}")
                    return
                
                # Cargar datos en los widgets
                self.imp_widgets["codigo_asignado"].delete(0, "end")
                self.imp_widgets["codigo_asignado"].insert(0, ws.cell(row=target_row, column=3).value or "")
                
                self.imp_widgets["tipo"].set(ws.cell(row=target_row, column=4).value or "")
                self.imp_widgets["marca"].set(ws.cell(row=target_row, column=5).value or "")
                
                self.imp_widgets["modelo"].delete(0, "end")
                self.imp_widgets["modelo"].insert(0, ws.cell(row=target_row, column=6).value or "")
                
                self.imp_widgets["serial"].delete(0, "end")
                self.imp_widgets["serial"].insert(0, ws.cell(row=target_row, column=7).value or "")
                
                self.imp_widgets["area"].set(ws.cell(row=target_row, column=8).value or "")
                
                self.imp_widgets["ubicacion"].delete(0, "end")
                self.imp_widgets["ubicacion"].insert(0, ws.cell(row=target_row, column=9).value or "")
                
                self.imp_widgets["funcion"].set(ws.cell(row=target_row, column=10).value or "")
                
                self.imp_widgets["ip"].delete(0, "end")
                self.imp_widgets["ip"].insert(0, ws.cell(row=target_row, column=11).value or "")
                
                self.imp_widgets["estado"].set(ws.cell(row=target_row, column=12).value or "")
                
                self.imp_widgets["fecha_adq"].delete(0, "end")
                fecha_val = ws.cell(row=target_row, column=13).value
                self.imp_widgets["fecha_adq"].insert(0, str(fecha_val) if fecha_val else "")
                
                self.imp_widgets["valor"].delete(0, "end")
                self.imp_widgets["valor"].insert(0, ws.cell(row=target_row, column=14).value or "")
                
                self.imp_widgets["observaciones"].delete(0, "end")
                self.imp_widgets["observaciones"].insert(0, ws.cell(row=target_row, column=15).value or "")
                
                wb.close()
                
                # Guardar código y fila para actualizar
                self.imp_update_code = codigo
                self.imp_update_row = target_row
                
                # CAMBIAR TÍTULO A MODO ACTUALIZACIÓN
                self.imp_scroll.configure(label_text=f"🔄 ACTUALIZANDO IMPRESORA - Código: {codigo}")
                
                # CAMBIAR TEXTO DEL BOTÓN
                if hasattr(self, 'btn_save_imp'):
                    self.btn_save_imp.configure(text="🔄 ACTUALIZAR IMPRESORA")
                
                dialog.destroy()
                
                # Confirmar
                if messagebox.askyesno(
                    "Confirmar Actualización",
                    f"⚠️ ¿Estás seguro de actualizar {codigo}?\n\n"
                    f"Los datos actuales se han cargado.\n"
                    f"Modifica los campos necesarios y presiona GUARDAR NUEVO."
                ):
                    messagebox.showinfo("Listo", f"✅ Datos cargados de {codigo}\n\nModifica los campos y presiona GUARDAR NUEVO.")
                
            except Exception as e:
                messagebox.showerror("Error", f"Error al buscar:\n{e}")
        
        btn_buscar = ctk.CTkButton(
            dialog,
            text="🔍 BUSCAR Y CARGAR",
            command=buscar_y_cargar,
            font=("Segoe UI", 13, "bold"),
            fg_color=COLOR_VERDE_HOSPITAL,
            height=40
        )
        btn_buscar.pack(pady=10)
        
        # Enter para buscar
        entry_codigo.bind("<Return>", lambda e: buscar_y_cargar())
    
    def create_perifericos_form(self, parent_tab):
        """Formulario para Periféricos."""
        scroll = ctk.CTkScrollableFrame(
            parent_tab,
            fg_color="#FAFAFA",
            label_fg_color=COLOR_VERDE_HOSPITAL,
            label_text_color="white",
            label_font=("Segoe UI", 15, "bold")
        )
        scroll.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Detectar siguiente código automáticamente
        next_code = self.detect_next_code("Periféricos", "PER")
        scroll.configure(label_text=f"🖱️ PERIFÉRICOS - Código: {next_code}")
        
        self.per_widgets = {}
        self.per_next_code = next_code
        self.per_scroll = scroll  # Guardar referencia al scroll
        
        fields = [
            ("Código Equipo Asignado *", "codigo_asignado", "entry"),
            ("Tipo *", "tipo", "combobox", TIPOS_PERIFERICO),
            ("Marca *", "marca", "combobox", MARCAS_PERIFERICO),
            ("Modelo", "modelo", "entry"),
            ("Serial", "serial", "entry"),
            ("Área / Servicio *", "area", "combobox", AREAS_SERVICIO),
            ("Estado Operativo *", "estado", "combobox", ESTADOS_PERIFERICO),
            ("Fecha de Adquisición (YYYY-MM-DD)", "fecha_adq", "entry"),
            ("Observaciones", "observaciones", "entry"),
        ]
        
        for field_data in fields:
            if len(field_data) == 4:
                label, key, field_type, options = field_data
                widget = self.create_form_field(scroll, label, key, field_type, options)
            else:
                label, key, field_type = field_data
                widget = self.create_form_field(scroll, label, key, field_type, None)
            self.per_widgets[key] = widget
        
        # Frame para botones
        btn_frame = ctk.CTkFrame(scroll, fg_color="transparent")
        btn_frame.pack(pady=30)
        
        # Botón guardar nuevo - Referencia global
        self.btn_save_per = ctk.CTkButton(
            btn_frame,
            text="💾 GUARDAR NUEVO",
            command=self.save_periferico,
            font=("Segoe UI", 14, "bold"),
            fg_color=COLOR_VERDE_HOSPITAL,
            hover_color="#1F5039",
            height=50,
            width=250
        )
        self.btn_save_per.pack(side="left", padx=10)
        
        # Botón actualizar existente
        btn_update = ctk.CTkButton(
            btn_frame,
            text="🔄 ACTUALIZAR EXISTENTE",
            command=self.update_periferico,
            font=("Segoe UI", 14, "bold"),
            fg_color="#2196F3",
            hover_color="#1976D2",
            height=50,
            width=250
        )
        btn_update.pack(side="left", padx=10)
    
    def save_periferico(self):
        """Guardar periférico en Excel."""
        if not self.excel_path:
            messagebox.showerror("Error", "No hay Excel cargado")
            return
        
        try:
            wb = load_workbook(self.excel_path)
            
            # Verificar que la hoja existe
            if "Periféricos" not in wb.sheetnames:
                wb.close()
                messagebox.showerror("Error", "La hoja 'Periféricos' no existe en el Excel. Crea esta hoja primero.")
                return
            
            ws = wb["Periféricos"]
            
            # Verificar si es actualización o nuevo registro
            if hasattr(self, 'per_update_row') and self.per_update_row:
                # MODO ACTUALIZACIÓN
                row = self.per_update_row
                codigo = self.per_update_code
                
                # Actualizar datos en la fila existente
                ws.cell(row=row, column=3, value=self.per_widgets["codigo_asignado"].get())
                ws.cell(row=row, column=4, value=self.per_widgets["tipo"].get())
                ws.cell(row=row, column=5, value=self.per_widgets["marca"].get())
                ws.cell(row=row, column=6, value=self.per_widgets["modelo"].get())
                ws.cell(row=row, column=7, value=self.per_widgets["serial"].get())
                ws.cell(row=row, column=8, value=self.per_widgets["area"].get())
                ws.cell(row=row, column=9, value=self.per_widgets["estado"].get())
                ws.cell(row=row, column=10, value=self.per_widgets["fecha_adq"].get())
                ws.cell(row=row, column=11, value=self.per_widgets["observaciones"].get())
                
                wb.save(self.excel_path)
                wb.close()
                
                messagebox.showinfo("Éxito", f"✅ Periférico {codigo} actualizado correctamente")
                
                # Limpiar modo actualización
                self.per_update_row = None
                self.per_update_code = None
                
                # Volver a título normal
                next_code = self.detect_next_code("Periféricos", "PER")
                self.per_next_code = next_code
                self.per_scroll.configure(label_text=f"🖱️ PERIFÉRICOS - Código: {next_code}")
                
                # Restaurar texto del botón
                if hasattr(self, 'btn_save_per'):
                    self.btn_save_per.configure(text="💾 GUARDAR NUEVO")
                
                # Limpiar todos los campos
                for key, widget in self.per_widgets.items():
                    if isinstance(widget, ctk.CTkEntry):
                        widget.delete(0, "end")
                    elif isinstance(widget, ctk.CTkComboBox):
                        widget.set("")
                
            else:
                # MODO GUARDAR NUEVO
                # Buscar el último consecutivo real
                last_consecutive = 0
                for row in range(2, 200):
                    value = ws.cell(row=row, column=1).value
                    if value is not None:
                        try:
                            consecutivo = int(value)
                            if consecutivo > last_consecutive:
                                last_consecutive = consecutivo
                        except:
                            pass
                
                # Siguiente consecutivo
                next_consecutive = last_consecutive + 1
                
                # Buscar primera fila vacía
                next_row = 2
                for row in range(2, 200):
                    if ws.cell(row=row, column=2).value is None:
                        next_row = row
                        break
                
                ws.cell(row=next_row, column=1, value=next_consecutive)
                ws.cell(row=next_row, column=2, value=f"PER-{next_consecutive:04d}")
                ws.cell(row=next_row, column=3, value=self.per_widgets["codigo_asignado"].get())
                ws.cell(row=next_row, column=4, value=self.per_widgets["tipo"].get())
                ws.cell(row=next_row, column=5, value=self.per_widgets["marca"].get())
                ws.cell(row=next_row, column=6, value=self.per_widgets["modelo"].get())
                ws.cell(row=next_row, column=7, value=self.per_widgets["serial"].get())
                ws.cell(row=next_row, column=8, value=self.per_widgets["area"].get())
                ws.cell(row=next_row, column=9, value=self.per_widgets["estado"].get())
                ws.cell(row=next_row, column=10, value=self.per_widgets["fecha_adq"].get())
                ws.cell(row=next_row, column=11, value=self.per_widgets["observaciones"].get())
                
                wb.save(self.excel_path)
                wb.close()
                
                messagebox.showinfo("Éxito", f"✅ Periférico guardado: PER-{next_consecutive:04d}")
                
                # Detectar siguiente código y actualizar título
                next_code = self.detect_next_code("Periféricos", "PER")
                self.per_next_code = next_code
                self.per_scroll.configure(label_text=f"🖱️ PERIFÉRICOS - Código: {next_code}")
                
                # Limpiar campos selectivamente (mantener área)
                campos_a_mantener = ['area']
                
                for key, widget in self.per_widgets.items():
                    if key not in campos_a_mantener:
                        if isinstance(widget, ctk.CTkEntry):
                            widget.delete(0, "end")
                        elif isinstance(widget, ctk.CTkComboBox):
                            widget.set("")
                    
        except Exception as e:
            messagebox.showerror("Error", f"❌ Error al guardar periférico:\n\n{str(e)}\n\nVerifica que la hoja 'Periféricos' existe.")
            import traceback
            traceback.print_exc()
    
    def update_periferico(self):
        """Actualizar periférico existente en Excel."""
        if not self.excel_path:
            messagebox.showerror("Error", "No hay Excel cargado")
            return
        
        # Ventana para pedir código
        dialog = ctk.CTkToplevel(self.root)
        dialog.title("Actualizar Periférico")
        dialog.geometry("400x200")
        dialog.transient(self.root)
        dialog.grab_set()
        
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - 200
        y = (dialog.winfo_screenheight() // 2) - 100
        dialog.geometry(f"400x200+{x}+{y}")
        
        ctk.CTkLabel(
            dialog,
            text="Ingresa el código del periférico a actualizar:",
            font=("Segoe UI", 13)
        ).pack(pady=20)
        
        entry_codigo = ctk.CTkEntry(
            dialog,
            width=200,
            height=40,
            font=("Segoe UI", 12),
            placeholder_text="Ej: PER-0026"
        )
        entry_codigo.pack(pady=10)
        entry_codigo.focus()
        
        def buscar_y_cargar():
            codigo = entry_codigo.get().strip().upper()
            if not codigo:
                messagebox.showerror("Error", "Debes ingresar un código")
                return
            
            try:
                wb = load_workbook(self.excel_path)
                ws = wb["Periféricos"]
                
                found = False
                target_row = None
                
                for row in range(2, 200):
                    cell_value = ws.cell(row=row, column=2).value
                    if cell_value and cell_value.upper() == codigo:
                        found = True
                        target_row = row
                        break
                
                if not found:
                    wb.close()
                    messagebox.showerror("Error", f"No se encontró el código {codigo}")
                    return
                
                # Cargar datos
                self.per_widgets["codigo_asignado"].delete(0, "end")
                self.per_widgets["codigo_asignado"].insert(0, ws.cell(row=target_row, column=3).value or "")
                
                self.per_widgets["tipo"].set(ws.cell(row=target_row, column=4).value or "")
                self.per_widgets["marca"].set(ws.cell(row=target_row, column=5).value or "")
                
                self.per_widgets["modelo"].delete(0, "end")
                self.per_widgets["modelo"].insert(0, ws.cell(row=target_row, column=6).value or "")
                
                self.per_widgets["serial"].delete(0, "end")
                self.per_widgets["serial"].insert(0, ws.cell(row=target_row, column=7).value or "")
                
                self.per_widgets["area"].set(ws.cell(row=target_row, column=8).value or "")
                self.per_widgets["estado"].set(ws.cell(row=target_row, column=9).value or "")
                
                self.per_widgets["fecha_adq"].delete(0, "end")
                fecha_val = ws.cell(row=target_row, column=10).value
                self.per_widgets["fecha_adq"].insert(0, str(fecha_val) if fecha_val else "")
                
                self.per_widgets["observaciones"].delete(0, "end")
                self.per_widgets["observaciones"].insert(0, ws.cell(row=target_row, column=11).value or "")
                
                wb.close()
                
                self.per_update_code = codigo
                self.per_update_row = target_row
                
                # CAMBIAR TÍTULO A MODO ACTUALIZACIÓN
                self.per_scroll.configure(label_text=f"🔄 ACTUALIZANDO PERIFÉRICO - Código: {codigo}")
                
                # CAMBIAR TEXTO DEL BOTÓN
                if hasattr(self, 'btn_save_per'):
                    self.btn_save_per.configure(text="🔄 ACTUALIZAR PERIFÉRICO")
                
                dialog.destroy()
                
                if messagebox.askyesno(
                    "Confirmar Actualización",
                    f"⚠️ ¿Estás seguro de actualizar {codigo}?\n\n"
                    f"Los datos actuales se han cargado.\n"
                    f"Modifica los campos necesarios y presiona GUARDAR NUEVO."
                ):
                    messagebox.showinfo("Listo", f"✅ Datos cargados de {codigo}\n\nModifica los campos y presiona GUARDAR NUEVO.")
                
            except Exception as e:
                messagebox.showerror("Error", f"Error al buscar:\n{e}")
        
        btn_buscar = ctk.CTkButton(
            dialog,
            text="🔍 BUSCAR Y CARGAR",
            command=buscar_y_cargar,
            font=("Segoe UI", 13, "bold"),
            fg_color=COLOR_VERDE_HOSPITAL,
            height=40
        )
        btn_buscar.pack(pady=10)
        entry_codigo.bind("<Return>", lambda e: buscar_y_cargar())
    
    def create_red_form(self, parent_tab):
        """Formulario para Equipos de Red."""
        scroll = ctk.CTkScrollableFrame(
            parent_tab,
            fg_color="#FAFAFA",
            label_fg_color=COLOR_VERDE_HOSPITAL,
            label_text_color="white",
            label_font=("Segoe UI", 15, "bold")
        )
        scroll.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Detectar siguiente código automáticamente
        next_code = self.detect_next_code("Equipos de Red", "RED")
        scroll.configure(label_text=f"🌐 EQUIPOS DE RED - Código: {next_code}")
        
        self.red_widgets = {}
        self.red_next_code = next_code
        self.red_scroll = scroll  # Guardar referencia al scroll
        
        fields = [
            ("Tipo *", "tipo", "combobox", TIPOS_EQUIPO_RED),
            ("Marca *", "marca", "combobox", MARCAS_RED),
            ("Modelo", "modelo", "entry"),
            ("Serial", "serial", "entry"),
            ("Dirección IP *", "ip", "entry"),
            ("Puertos Totales", "puertos", "entry"),
            ("Ubicación *", "ubicacion", "combobox", UBICACIONES_RED),
            ("Área / Servicio", "area", "combobox", AREAS_SERVICIO),
            ("Estado Operativo *", "estado", "combobox", ESTADOS_RED),
            ("Fecha de Adquisición (YYYY-MM-DD)", "fecha_adq", "entry"),
            ("Valor de Adquisición (COP)", "valor", "entry"),
            ("Observaciones", "observaciones", "entry"),
        ]
        
        for field_data in fields:
            if len(field_data) == 4:
                label, key, field_type, options = field_data
                widget = self.create_form_field(scroll, label, key, field_type, options)
            else:
                label, key, field_type = field_data
                widget = self.create_form_field(scroll, label, key, field_type, None)
            self.red_widgets[key] = widget
        
        # Frame para botones
        btn_frame = ctk.CTkFrame(scroll, fg_color="transparent")
        btn_frame.pack(pady=30)
        
        # Botón guardar nuevo - Referencia global
        self.btn_save_red = ctk.CTkButton(
            btn_frame,
            text="💾 GUARDAR NUEVO",
            command=self.save_red,
            font=("Segoe UI", 14, "bold"),
            fg_color=COLOR_VERDE_HOSPITAL,
            hover_color="#1F5039",
            height=50,
            width=250
        )
        self.btn_save_red.pack(side="left", padx=10)
        
        # Botón actualizar existente
        btn_update = ctk.CTkButton(
            btn_frame,
            text="🔄 ACTUALIZAR EXISTENTE",
            command=self.update_red,
            font=("Segoe UI", 14, "bold"),
            fg_color="#2196F3",
            hover_color="#1976D2",
            height=50,
            width=250
        )
        btn_update.pack(side="left", padx=10)
    
    def save_red(self):
        """Guardar equipo de red en Excel."""
        if not self.excel_path:
            messagebox.showerror("Error", "No hay Excel cargado")
            return
        
        try:
            wb = load_workbook(self.excel_path)
            
            # Verificar que la hoja existe
            if "Equipos de Red" not in wb.sheetnames:
                wb.close()
                messagebox.showerror("Error", "La hoja 'Equipos de Red' no existe en el Excel.\n\nCrea esta hoja primero.")
                return
            
            ws = wb["Equipos de Red"]
            
            # Verificar si es actualización o nuevo registro
            if hasattr(self, 'red_update_row') and self.red_update_row:
                # MODO ACTUALIZACIÓN
                row = self.red_update_row
                codigo = self.red_update_code
                
                # Actualizar datos en la fila existente (NO modificar columnas 1 y 2)
                ws.cell(row=row, column=3, value=self.red_widgets["tipo"].get())
                ws.cell(row=row, column=4, value=self.red_widgets["marca"].get())
                ws.cell(row=row, column=5, value=self.red_widgets["modelo"].get())
                ws.cell(row=row, column=6, value=self.red_widgets["serial"].get())
                ws.cell(row=row, column=7, value=self.red_widgets["ip"].get())
                ws.cell(row=row, column=8, value=self.red_widgets["puertos"].get())
                ws.cell(row=row, column=9, value=self.red_widgets["ubicacion"].get())
                ws.cell(row=row, column=10, value=self.red_widgets["area"].get())
                ws.cell(row=row, column=11, value=self.red_widgets["estado"].get())
                ws.cell(row=row, column=12, value=self.red_widgets["fecha_adq"].get())
                ws.cell(row=row, column=13, value=self.red_widgets["valor"].get())
                ws.cell(row=row, column=14, value=self.red_widgets["observaciones"].get())
                
                wb.save(self.excel_path)
                wb.close()
                
                messagebox.showinfo("Éxito", f"✅ Equipo de red {codigo} actualizado correctamente")
                
                # Limpiar modo actualización
                self.red_update_row = None
                self.red_update_code = None
                
                # Volver a título normal
                next_code = self.detect_next_code("Equipos de Red", "RED")
                self.red_next_code = next_code
                self.red_scroll.configure(label_text=f"🌐 EQUIPOS DE RED - Código: {next_code}")
                
                # Restaurar texto del botón
                if hasattr(self, 'btn_save_red'):
                    self.btn_save_red.configure(text="💾 GUARDAR NUEVO")
                
                # Limpiar todos los campos
                for key, widget in self.red_widgets.items():
                    if isinstance(widget, ctk.CTkEntry):
                        widget.delete(0, "end")
                    elif isinstance(widget, ctk.CTkComboBox):
                        widget.set("")
                
            else:
                # MODO GUARDAR NUEVO
                # Buscar el último consecutivo real
                last_consecutive = 0
                for row in range(2, 100):
                    value = ws.cell(row=row, column=1).value
                    if value is not None:
                        try:
                            consecutivo = int(value)
                            if consecutivo > last_consecutive:
                                last_consecutive = consecutivo
                        except:
                            pass
                
                # Siguiente consecutivo
                next_consecutive = last_consecutive + 1
                
                # Buscar primera fila vacía
                next_row = 2
                for row in range(2, 100):
                    if ws.cell(row=row, column=2).value is None:
                        next_row = row
                        break
                
                ws.cell(row=next_row, column=1, value=next_consecutive)
                ws.cell(row=next_row, column=2, value=f"RED-{next_consecutive:04d}")
                ws.cell(row=next_row, column=3, value=self.red_widgets["tipo"].get())
                ws.cell(row=next_row, column=4, value=self.red_widgets["marca"].get())
                ws.cell(row=next_row, column=5, value=self.red_widgets["modelo"].get())
                ws.cell(row=next_row, column=6, value=self.red_widgets["serial"].get())
                ws.cell(row=next_row, column=7, value=self.red_widgets["ip"].get())
                ws.cell(row=next_row, column=8, value=self.red_widgets["puertos"].get())
                ws.cell(row=next_row, column=9, value=self.red_widgets["ubicacion"].get())
                ws.cell(row=next_row, column=10, value=self.red_widgets["area"].get())
                ws.cell(row=next_row, column=11, value=self.red_widgets["estado"].get())
                ws.cell(row=next_row, column=12, value=self.red_widgets["fecha_adq"].get())
                ws.cell(row=next_row, column=13, value=self.red_widgets["valor"].get())
                ws.cell(row=next_row, column=14, value=self.red_widgets["observaciones"].get())
                
                wb.save(self.excel_path)
                wb.close()
                
                messagebox.showinfo("Éxito", f"✅ Equipo de red guardado: RED-{next_consecutive:04d}")
                
                # Detectar siguiente código y actualizar título
                next_code = self.detect_next_code("Equipos de Red", "RED")
                self.red_next_code = next_code
                self.red_scroll.configure(label_text=f"🌐 EQUIPOS DE RED - Código: {next_code}")
                
                # Limpiar campos selectivamente (mantener área y ubicación)
                campos_a_mantener = ['area', 'ubicacion']
                
                for key, widget in self.red_widgets.items():
                    if key not in campos_a_mantener:
                        if isinstance(widget, ctk.CTkEntry):
                            widget.delete(0, "end")
                        elif isinstance(widget, ctk.CTkComboBox):
                            widget.set("")
                    
        except Exception as e:
            messagebox.showerror("Error", f"❌ Error al guardar equipo de red:\n\n{str(e)}\n\nVerifica que la hoja 'Equipos de Red' existe.")
            import traceback
            traceback.print_exc()
    
    def update_red(self):
        """Actualizar equipo de red existente en Excel."""
        if not self.excel_path:
            messagebox.showerror("Error", "No hay Excel cargado")
            return
        
        dialog = ctk.CTkToplevel(self.root)
        dialog.title("Actualizar Equipo de Red")
        dialog.geometry("400x200")
        dialog.transient(self.root)
        dialog.grab_set()
        
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - 200
        y = (dialog.winfo_screenheight() // 2) - 100
        dialog.geometry(f"400x200+{x}+{y}")
        
        ctk.CTkLabel(
            dialog,
            text="Ingresa el código del equipo de red a actualizar:",
            font=("Segoe UI", 13)
        ).pack(pady=20)
        
        entry_codigo = ctk.CTkEntry(
            dialog,
            width=200,
            height=40,
            font=("Segoe UI", 12),
            placeholder_text="Ej: RED-0026"
        )
        entry_codigo.pack(pady=10)
        entry_codigo.focus()
        
        def buscar_y_cargar():
            codigo = entry_codigo.get().strip().upper()
            if not codigo:
                messagebox.showerror("Error", "Debes ingresar un código")
                return
            
            try:
                wb = load_workbook(self.excel_path)
                ws = wb["Equipos de Red"]
                
                found = False
                target_row = None
                
                for row in range(2, 100):
                    cell_value = ws.cell(row=row, column=2).value
                    if cell_value and cell_value.upper() == codigo:
                        found = True
                        target_row = row
                        break
                
                if not found:
                    wb.close()
                    messagebox.showerror("Error", f"No se encontró el código {codigo}")
                    return
                
                # Cargar datos
                self.red_widgets["tipo"].set(ws.cell(row=target_row, column=3).value or "")
                self.red_widgets["marca"].set(ws.cell(row=target_row, column=4).value or "")
                
                self.red_widgets["modelo"].delete(0, "end")
                self.red_widgets["modelo"].insert(0, ws.cell(row=target_row, column=5).value or "")
                
                self.red_widgets["serial"].delete(0, "end")
                self.red_widgets["serial"].insert(0, ws.cell(row=target_row, column=6).value or "")
                
                self.red_widgets["ip"].delete(0, "end")
                self.red_widgets["ip"].insert(0, ws.cell(row=target_row, column=7).value or "")
                
                self.red_widgets["puertos"].delete(0, "end")
                self.red_widgets["puertos"].insert(0, ws.cell(row=target_row, column=8).value or "")
                
                self.red_widgets["ubicacion"].set(ws.cell(row=target_row, column=9).value or "")
                self.red_widgets["area"].set(ws.cell(row=target_row, column=10).value or "")
                self.red_widgets["estado"].set(ws.cell(row=target_row, column=11).value or "")
                
                self.red_widgets["fecha_adq"].delete(0, "end")
                fecha_val = ws.cell(row=target_row, column=12).value
                self.red_widgets["fecha_adq"].insert(0, str(fecha_val) if fecha_val else "")
                
                self.red_widgets["valor"].delete(0, "end")
                self.red_widgets["valor"].insert(0, ws.cell(row=target_row, column=13).value or "")
                
                self.red_widgets["observaciones"].delete(0, "end")
                self.red_widgets["observaciones"].insert(0, ws.cell(row=target_row, column=14).value or "")
                
                wb.close()
                
                self.red_update_code = codigo
                self.red_update_row = target_row
                
                # CAMBIAR TÍTULO A MODO ACTUALIZACIÓN
                self.red_scroll.configure(label_text=f"🔄 ACTUALIZANDO EQUIPO DE RED - Código: {codigo}")
                
                # CAMBIAR TEXTO DEL BOTÓN
                if hasattr(self, 'btn_save_red'):
                    self.btn_save_red.configure(text="🔄 ACTUALIZAR EQUIPO DE RED")
                
                dialog.destroy()
                
                if messagebox.askyesno(
                    "Confirmar Actualización",
                    f"⚠️ ¿Estás seguro de actualizar {codigo}?\n\n"
                    f"Los datos actuales se han cargado.\n"
                    f"Modifica los campos necesarios y presiona GUARDAR NUEVO."
                ):
                    messagebox.showinfo("Listo", f"✅ Datos cargados de {codigo}\n\nModifica los campos y presiona GUARDAR NUEVO.")
                
            except Exception as e:
                messagebox.showerror("Error", f"Error al buscar:\n{e}")
        
        btn_buscar = ctk.CTkButton(
            dialog,
            text="🔍 BUSCAR Y CARGAR",
            command=buscar_y_cargar,
            font=("Segoe UI", 13, "bold"),
            fg_color=COLOR_VERDE_HOSPITAL,
            height=40
        )
        btn_buscar.pack(pady=10)
        entry_codigo.bind("<Return>", lambda e: buscar_y_cargar())
    
    def create_mantenimientos_form(self, parent_tab):
        """Formulario para Mantenimientos."""
        scroll = ctk.CTkScrollableFrame(
            parent_tab,
            fg_color="#FAFAFA",
            label_fg_color=COLOR_VERDE_HOSPITAL,
            label_text_color="white",
            label_font=("Segoe UI", 15, "bold")
        )
        scroll.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Detectar siguiente consecutivo
        next_consecutive = self.detect_next_consecutive_mantenimiento()
        scroll.configure(label_text=f"🔧 MANTENIMIENTOS - Registro #{next_consecutive}")
        
        self.mtt_widgets = {}
        self.mtt_scroll = scroll  # Guardar referencia
        self.mtt_next_consecutive = next_consecutive
        
        fields = [
            ("Código Equipo *", "codigo_equipo", "entry"),
            ("Fecha Mantenimiento (YYYY-MM-DD) *", "fecha_mtto", "entry"),
            ("Tipo Mantenimiento *", "tipo", "combobox", TIPOS_MANTENIMIENTO_MTTO),
            ("Técnico Responsable *", "tecnico", "combobox", TECNICOS_RESPONSABLES),
            ("Descripción Actividades *", "descripcion", "combobox", ACTIVIDADES_MANTENIMIENTO),
            ("Repuestos/Insumos", "repuestos", "entry"),
            ("Estado Post-Mtto *", "estado_post", "combobox", ESTADO_POST_MTTO),
            ("Próximo Mantenimiento (YYYY-MM-DD)", "proximo", "entry"),
            ("Observaciones", "observaciones", "entry"),
        ]
        
        for field_data in fields:
            if len(field_data) == 4:
                label, key, field_type, options = field_data
                widget = self.create_form_field(scroll, label, key, field_type, options)
            else:
                label, key, field_type = field_data
                widget = self.create_form_field(scroll, label, key, field_type, None)
            self.mtt_widgets[key] = widget
        
        btn_save = ctk.CTkButton(
            scroll,
            text="💾 GUARDAR MANTENIMIENTO",
            command=self.save_mantenimiento,
            font=("Segoe UI", 14, "bold"),
            fg_color=COLOR_VERDE_HOSPITAL,
            hover_color="#1F5039",
            height=50
        )
        btn_save.pack(pady=30)
    
    def save_mantenimiento(self):
        """Guardar mantenimiento en Excel."""
        if not self.excel_path:
            messagebox.showerror("Error", "No hay Excel cargado")
            return
        
        try:
            wb = load_workbook(self.excel_path)
            ws = wb["Mantenimientos"]
            
            next_row = 2
            for row in range(2, 500):
                if ws.cell(row=row, column=1).value is None:
                    next_row = row
                    break
            
            consecutive = next_row - 1
            
            ws.cell(row=next_row, column=1, value=consecutive)
            ws.cell(row=next_row, column=2, value=self.mtt_widgets["codigo_equipo"].get())
            ws.cell(row=next_row, column=3, value=self.mtt_widgets["fecha_mtto"].get())
            ws.cell(row=next_row, column=4, value=self.mtt_widgets["tipo"].get())
            ws.cell(row=next_row, column=5, value=self.mtt_widgets["tecnico"].get())
            ws.cell(row=next_row, column=6, value=self.mtt_widgets["descripcion"].get())
            ws.cell(row=next_row, column=7, value=self.mtt_widgets["repuestos"].get())
            ws.cell(row=next_row, column=8, value=self.mtt_widgets["estado_post"].get())
            ws.cell(row=next_row, column=9, value=self.mtt_widgets["proximo"].get())
            ws.cell(row=next_row, column=10, value=self.mtt_widgets["observaciones"].get())
            
            wb.save(self.excel_path)
            wb.close()
            
            messagebox.showinfo("Éxito", f"✅ Mantenimiento registrado #{consecutive}")
            
            # Actualizar título para siguiente registro
            next_consecutive = self.detect_next_consecutive_mantenimiento()
            self.mtt_next_consecutive = next_consecutive
            self.mtt_scroll.configure(label_text=f"🔧 MANTENIMIENTOS - Registro #{next_consecutive}")
            
            # Limpiar campos selectivamente (mantener técnico)
            campos_a_mantener = ['tecnico']
            
            for key, widget in self.mtt_widgets.items():
                if key not in campos_a_mantener:
                    if isinstance(widget, ctk.CTkEntry):
                        widget.delete(0, "end")
                    elif isinstance(widget, ctk.CTkComboBox):
                        widget.set("")
                    
        except Exception as e:
            messagebox.showerror("Error", f"Error al guardar:\n{e}")
    
    def create_baja_form(self, parent_tab):
        """Formulario para Equipos Dados de Baja."""
        scroll = ctk.CTkScrollableFrame(
            parent_tab,
            fg_color="#FAFAFA",
            label_fg_color=COLOR_VERDE_HOSPITAL,
            label_text_color="white",
            label_font=("Segoe UI", 15, "bold")
        )
        scroll.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Detectar siguiente número de baja
        next_baja = self.detect_next_baja()
        scroll.configure(label_text=f"📦 EQUIPOS DADOS DE BAJA - Baja #{next_baja}")
        
        self.baja_widgets = {}
        self.baja_scroll = scroll  # Guardar referencia
        self.baja_next = next_baja
        
        fields = [
            ("Código Original *", "codigo_original", "entry"),
            ("Tipo *", "tipo", "entry"),
            ("Marca", "marca", "entry"),
            ("Modelo", "modelo", "entry"),
            ("Serial", "serial", "entry"),
            ("Fecha de Baja (YYYY-MM-DD) *", "fecha_baja", "entry"),
            ("Motivo Baja *", "motivo", "combobox", MOTIVOS_BAJA),
            ("Destino *", "destino", "combobox", DESTINOS_BAJA),
            ("Responsable Baja *", "responsable", "combobox", RESPONSABLES_BAJA),
            ("Observaciones", "observaciones", "entry"),
        ]
        
        for field_data in fields:
            if len(field_data) == 4:
                label, key, field_type, options = field_data
                widget = self.create_form_field(scroll, label, key, field_type, options)
            else:
                label, key, field_type = field_data
                widget = self.create_form_field(scroll, label, key, field_type, None)
            self.baja_widgets[key] = widget
        
        # Botón para buscar y autocompletar
        btn_search = ctk.CTkButton(
            scroll,
            text="🔍 BUSCAR Y AUTOCOMPLETAR DESDE INVENTARIO",
            command=self.buscar_equipo_baja,
            font=("Segoe UI", 13, "bold"),
            fg_color="#2196F3",
            hover_color="#1976D2",
            height=45
        )
        btn_search.pack(pady=15, padx=20, fill="x")
        
        # Separador
        separator = ctk.CTkFrame(scroll, height=2, fg_color="#E0E0E0")
        separator.pack(fill="x", padx=20, pady=10)
        
        btn_save = ctk.CTkButton(
            scroll,
            text="💾 REGISTRAR BAJA",
            command=self.save_baja,
            font=("Segoe UI", 14, "bold"),
            fg_color="#DC3545",
            hover_color="#A02828",
            height=50
        )
        btn_save.pack(pady=30)
    
    def buscar_equipo_baja(self):
        """Buscar equipo en inventarios y autocompletar datos."""
        codigo = self.baja_widgets["codigo_original"].get().strip().upper()
        
        if not codigo:
            messagebox.showerror("Error", "Primero ingresa el código del equipo en el campo 'Código Original'")
            return
        
        try:
            wb = load_workbook(self.excel_path)
            
            # Determinar en qué hoja buscar según el prefijo
            if codigo.startswith("EQC-"):
                ws_name = "Equipos de Cómputo"
                col_codigo = 2
            elif codigo.startswith("IMP-"):
                ws_name = "Impresoras y Escáneres"
                col_codigo = 2
            elif codigo.startswith("PER-"):
                ws_name = "Periféricos"
                col_codigo = 2
            elif codigo.startswith("RED-"):
                ws_name = "Equipos de Red"
                col_codigo = 2
            else:
                wb.close()
                messagebox.showerror("Error", "Código no válido. Usa: EQC-XXXX, IMP-XXX, PER-XXX, RED-XXX")
                return
            
            ws = wb[ws_name]
            found = False
            target_row = None
            
            # Buscar código
            for row in range(2, 500):
                cell_value = ws.cell(row=row, column=col_codigo).value
                if cell_value and cell_value.upper() == codigo:
                    found = True
                    target_row = row
                    break
            
            if not found:
                wb.close()
                messagebox.showerror("Error", f"No se encontró el código {codigo} en {ws_name}")
                return
            
            # Autocompletar según el tipo
            if codigo.startswith("EQC-"):
                # Equipos de Cómputo
                tipo = ws.cell(row=target_row, column=4).value or "Computador"  # Tipo equipo
                marca = ws.cell(row=target_row, column=28).value or ""  # Marca (verde)
                modelo = ws.cell(row=target_row, column=29).value or ""  # Modelo (verde)
                serial = ws.cell(row=target_row, column=30).value or ""  # Serial (verde)
                
            elif codigo.startswith("IMP-"):
                # Impresoras
                tipo = ws.cell(row=target_row, column=4).value or "Impresora"
                marca = ws.cell(row=target_row, column=5).value or ""
                modelo = ws.cell(row=target_row, column=6).value or ""
                serial = ws.cell(row=target_row, column=7).value or ""
                
            elif codigo.startswith("PER-"):
                # Periféricos
                tipo = ws.cell(row=target_row, column=4).value or "Periférico"
                marca = ws.cell(row=target_row, column=5).value or ""
                modelo = ws.cell(row=target_row, column=6).value or ""
                serial = ws.cell(row=target_row, column=7).value or ""
                
            elif codigo.startswith("RED-"):
                # Equipos de Red
                tipo = ws.cell(row=target_row, column=3).value or "Equipo de Red"
                marca = ws.cell(row=target_row, column=4).value or ""
                modelo = ws.cell(row=target_row, column=5).value or ""
                serial = ws.cell(row=target_row, column=6).value or ""
            
            wb.close()
            
            # Cargar datos en los widgets
            self.baja_widgets["tipo"].delete(0, "end")
            self.baja_widgets["tipo"].insert(0, tipo)
            
            self.baja_widgets["marca"].delete(0, "end")
            self.baja_widgets["marca"].insert(0, marca)
            
            self.baja_widgets["modelo"].delete(0, "end")
            self.baja_widgets["modelo"].insert(0, modelo)
            
            self.baja_widgets["serial"].delete(0, "end")
            self.baja_widgets["serial"].insert(0, serial)
            
            # Guardar información para actualizar después
            self.baja_origen_sheet = ws_name
            self.baja_origen_row = target_row
            
            messagebox.showinfo("Éxito", f"✅ Datos cargados de {codigo}\n\nCompleta los campos de baja y guarda.")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al buscar equipo:\n{e}")
    
    def save_baja(self):
        """Guardar equipo dado de baja en Excel y actualizar estado en inventario original."""
        if not self.excel_path:
            messagebox.showerror("Error", "No hay Excel cargado")
            return
        
        try:
            wb = load_workbook(self.excel_path)
            ws_baja = wb["Equipos Dados de Baja"]
            
            next_row = 2
            for row in range(2, 200):
                if ws_baja.cell(row=row, column=1).value is None:
                    next_row = row
                    break
            
            codigo = self.baja_widgets["codigo_original"].get()
            
            # Guardar en hoja de Dados de Baja
            ws_baja.cell(row=next_row, column=1, value=codigo)
            ws_baja.cell(row=next_row, column=2, value=self.baja_widgets["tipo"].get())
            ws_baja.cell(row=next_row, column=3, value=self.baja_widgets["marca"].get())
            ws_baja.cell(row=next_row, column=4, value=self.baja_widgets["modelo"].get())
            ws_baja.cell(row=next_row, column=5, value=self.baja_widgets["serial"].get())
            ws_baja.cell(row=next_row, column=6, value=self.baja_widgets["fecha_baja"].get())
            ws_baja.cell(row=next_row, column=7, value=self.baja_widgets["motivo"].get())
            ws_baja.cell(row=next_row, column=8, value=self.baja_widgets["destino"].get())
            ws_baja.cell(row=next_row, column=9, value=self.baja_widgets["responsable"].get())
            ws_baja.cell(row=next_row, column=10, value=self.baja_widgets["observaciones"].get())
            
            # Actualizar estado en inventario original (si fue autocompletado)
            if hasattr(self, 'baja_origen_sheet') and hasattr(self, 'baja_origen_row'):
                ws_origen = wb[self.baja_origen_sheet]
                
                # Actualizar estado operativo según el tipo
                if self.baja_origen_sheet == "Equipos de Cómputo":
                    # Columna 15 = Estado Operativo
                    ws_origen.cell(row=self.baja_origen_row, column=15, value="DADO DE BAJA")
                elif self.baja_origen_sheet == "Impresoras y Escáneres":
                    # Columna 12 = Estado
                    ws_origen.cell(row=self.baja_origen_row, column=12, value="DADO DE BAJA")
                elif self.baja_origen_sheet == "Periféricos":
                    # Columna 9 = Estado
                    ws_origen.cell(row=self.baja_origen_row, column=9, value="DADO DE BAJA")
                elif self.baja_origen_sheet == "Equipos de Red":
                    # Columna 11 = Estado
                    ws_origen.cell(row=self.baja_origen_row, column=11, value="DADO DE BAJA")
                
                # Limpiar referencias
                delattr(self, 'baja_origen_sheet')
                delattr(self, 'baja_origen_row')
            
            wb.save(self.excel_path)
            wb.close()
            
            messagebox.showinfo("Éxito", 
                f"✅ Baja registrada: {codigo}\n\n"
                f"• Agregado a 'Equipos Dados de Baja'\n"
                f"• Estado actualizado a 'DADO DE BAJA' en inventario original")
            
            # Actualizar título para siguiente registro
            next_baja = self.detect_next_baja()
            self.baja_next = next_baja
            self.baja_scroll.configure(label_text=f"📦 EQUIPOS DADOS DE BAJA - Baja #{next_baja}")
            
            # Limpiar campos selectivamente (mantener responsable, motivo y destino)
            campos_a_mantener = ['responsable', 'motivo', 'destino']
            
            for key, widget in self.baja_widgets.items():
                if key not in campos_a_mantener:
                    if isinstance(widget, ctk.CTkEntry):
                        widget.delete(0, "end")
                    elif isinstance(widget, ctk.CTkComboBox):
                        widget.set("")
                    
        except Exception as e:
            messagebox.showerror("Error", f"Error al guardar:\n{e}")


# ============================================================================
# MAIN
# ============================================================================

def main():
    """Función principal."""
    root = ctk.CTk()
    app = InventoryManagerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
