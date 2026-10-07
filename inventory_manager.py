#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
INVENTORY MANAGER - Sistema de Inventario Tecnológico
=========================================================
Hospital Regional Alfonso Jaramillo Salazar

VERSIÓN: 1.0
FECHA: Enero 2026
"""

import tkinter as tk
from tkinter import messagebox, filedialog, ttk
from tkcalendar import DateEntry

import customtkinter as ctk
import platform
import socket
import subprocess
import os
import re
import threading
import sys
import json
import time
import shutil
import zipfile
import unicodedata
import urllib.request
import urllib.error
import xml.etree.ElementTree as ET

from datetime import datetime
from dateutil.relativedelta import relativedelta
from pathlib import Path

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
    from PIL import Image, ImageDraw
    HAS_PIL = True
except ImportError:
    HAS_PIL = False
    print("PIL/Pillow no disponible - Logo no se mostrará")


# ============================================================================
# COLORES INSTITUCIONALES
# ============================================================================

COLOR_VERDE_HOSPITAL = "#3C8B0E"
COLOR_AZUL_HOSPITAL = "#3C8B0E"
COLOR_NARANJA = "#F4B183"
COLOR_FONDO = "#F5F5F5"
COLOR_ERROR = "#DC3545"

# ============================================================================
# 1. CLASE TOOLTIP
# ============================================================================

class ToolTip:
    """
    Clase para mostrar tooltips al pasar el mouse sobre un widget.
    """
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tooltip = None
        self.widget.bind("<Enter>", self.show_tooltip)
        self.widget.bind("<Leave>", self.hide_tooltip)
    
    def show_tooltip(self, event=None):
        if self.tooltip:
            return
        
        x, y, _, _ = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 25
        
        self.tooltip = tk.Toplevel(self.widget)
        self.tooltip.wm_overrideredirect(True)
        self.tooltip.wm_geometry(f"+{x}+{y}")
        
        # Frame con borde
        frame = tk.Frame(self.tooltip, bg="#FFFFCC", relief=tk.SOLID, borderwidth=1)
        frame.pack()
        
        # Label con texto (máximo 600px de ancho)
        label = tk.Label(
            frame, 
            text=self.text, 
            bg="#FFFFCC",
            fg="#000000",
            font=("Segoe UI", 10),
            wraplength=600,
            justify=tk.LEFT,
            padx=10,
            pady=8
        )
        label.pack()
    
    def hide_tooltip(self, event=None):
        if self.tooltip:
            self.tooltip.destroy()
            self.tooltip = None


def detect_hardware_wmi():
    """
    Detectar hardware usando WMI.
    """
    info = {
        'marca': 'No detectado',
        'modelo': 'No detectado',
        'serial': 'No detectado',
        'disco1_capacidad': 'No detectado',
        'disco1_tipo': 'No detectado',
        'disco1_serial': 'No detectado',
        'disco1_marca': 'No detectado',
        'disco1_modelo': 'No detectado',
        'disco2_capacidad': 'No tiene',
        'disco2_tipo': 'No tiene',
        'disco2_serial': 'No tiene',
        'disco2_marca': 'No tiene',
        'disco2_modelo': 'No tiene'
    }

    def _normalize_text(value):
        text = str(value or '').strip()
        if not text:
            return ''
        return unicodedata.normalize('NFKD', text).encode('ascii', 'ignore').decode('ascii').lower().strip()

    def _is_placeholder(value):
        normalized = _normalize_text(value)
        if not normalized:
            return True
        if normalized.startswith('no detectado'):
            return True
        return normalized in {'no tiene', 'detectado', 'unknown', 'n/a', 'none', 'null'}

    def _clean_disk_brand(value):
        brand = str(value or '').strip()
        normalized = _normalize_text(brand)
        if not normalized:
            return 'No detectado'
        if 'standard disk' in normalized:
            return 'No detectado'
        if 'unidades de disco' in normalized or ('unidad' in normalized and 'disco' in normalized):
            return 'No detectado'
        return brand

    def _infer_disk_type(media_type, model, interface_type=''):
        data = f"{media_type or ''} {model or ''} {interface_type or ''}".upper()
        if any(token in data for token in ['NVME', 'SSD', 'SOLID STATE']):
            return 'SSD'
        if any(token in data for token in ['HDD', 'ROTATIONAL', 'MAGNETIC']):
            return 'HDD'
        return 'No detectado'

    def _derive_brand_from_model(model):
        model_text = str(model or '').strip()
        if not model_text:
            return 'No detectado'
        tokens = re.split(r"[\s_\-]+", model_text)
        generic_tokens = {
            'SSD', 'HDD', 'NVME', 'SATA', 'ATA', 'DISK', 'DRIVE', 'DEVICE', 'STORAGE', 'SOLID', 'STATE',
            'GENERAL', 'MSFT', 'MICROSOFT', 'VIRTUAL'
        }
        for token in tokens:
            cleaned = re.sub(r"[^A-Za-z0-9]", "", token).upper()
            if len(cleaned) >= 3 and cleaned not in generic_tokens:
                return cleaned
        return 'No detectado'

    def _is_virtual_disk(model, size_bytes, interface_type='', media_type=''):
        model_text = str(model or '').upper()
        interface_text = str(interface_type or '').upper()
        media_text = str(media_type or '').upper()
        virtual_tokens = ['VIRTUAL', 'VMWARE', 'VBOX', 'QEMU', 'MICROSOFT VIRTUAL', 'XENSRC', 'USB', 'UDISK']
        if any(token in model_text for token in virtual_tokens):
            return True
        if 'USB' in interface_text or 'REMOVABLE' in media_text:
            return True
        try:
            return int(size_bytes or 0) <= 0
        except Exception:
            return False

    def _safe_round_gb(size_value):
        try:
            size_bytes = int(size_value or 0)
            if size_bytes <= 0:
                return 'No detectado', 0
            return str(round(size_bytes / (1024 ** 3))), size_bytes
        except Exception:
            return 'No detectado', 0

    def _detect_storage_api_disks():
        cmd = (
            "$d=Get-PhysicalDisk -ErrorAction SilentlyContinue | "
            "Select-Object FriendlyName,Model,SerialNumber,MediaType,BusType,Size;"
            "if($null -eq $d){'[]'} else {$d | ConvertTo-Json -Compress -Depth 4}"
        )

        output = _run_powershell(cmd, timeout=10)
        if not output:
            return []

        try:
            payload = json.loads(output)
        except Exception:
            return []

        if isinstance(payload, dict):
            payload = [payload]
        if not isinstance(payload, list):
            return []

        parsed = []
        for disk in payload:
            if not isinstance(disk, dict):
                continue

            size_label, size_bytes = _safe_round_gb(disk.get('Size'))
            model = str(disk.get('Model') or disk.get('FriendlyName') or '').strip()
            media_type = str(disk.get('MediaType') or '').strip()
            bus_type = str(disk.get('BusType') or '').strip()
            serial = str(disk.get('SerialNumber') or '').strip()

            if size_bytes <= 0:
                continue
            if _is_virtual_disk(model, size_bytes, bus_type, media_type):
                continue

            disk_type = _infer_disk_type(media_type, model, bus_type)
            if disk_type == 'No detectado' and 'HDD' in media_type.upper():
                disk_type = 'HDD'

            brand = _derive_brand_from_model(model)

            parsed.append({
                'size_bytes': size_bytes,
                'capacidad': size_label,
                'tipo': disk_type,
                'serial': serial if serial else 'No detectado',
                'marca': brand,
                'modelo': model if model else 'No detectado'
            })

        parsed.sort(key=lambda item: item['size_bytes'], reverse=True)
        return parsed

    def _detect_hardware_cim_fallback():
        fallback = {
            'marca': 'No detectado',
            'modelo': 'No detectado',
            'serial': 'No detectado',
            'disco1_capacidad': 'No detectado',
            'disco1_tipo': 'No detectado',
            'disco1_serial': 'No detectado',
            'disco1_marca': 'No detectado',
            'disco1_modelo': 'No detectado',
            'disco2_capacidad': 'No tiene',
            'disco2_tipo': 'No tiene',
            'disco2_serial': 'No tiene',
            'disco2_marca': 'No tiene',
            'disco2_modelo': 'No tiene'
        }

        cmd = (
            "$cs=Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue;"
            "$bios=Get-CimInstance Win32_BIOS -ErrorAction SilentlyContinue;"
            "$disks=Get-CimInstance Win32_DiskDrive -ErrorAction SilentlyContinue | Sort-Object Index | "
            "Select-Object Index,Model,SerialNumber,Size,MediaType,InterfaceType,Manufacturer;"
            "[pscustomobject]@{"
            "marca=$cs.Manufacturer;"
            "modelo=$cs.Model;"
            "serial=$bios.SerialNumber;"
            "discos=$disks"
            "}|ConvertTo-Json -Compress -Depth 5"
        )

        output = _run_powershell(cmd, timeout=10)
        if not output:
            return fallback

        try:
            payload = json.loads(output)
        except Exception:
            return fallback

        marca = str(payload.get('marca') or '').strip()
        modelo = str(payload.get('modelo') or '').strip()
        serial = str(payload.get('serial') or '').strip()

        if marca:
            fallback['marca'] = marca
        if modelo:
            fallback['modelo'] = modelo
        if serial and not _is_placeholder(serial):
            fallback['serial'] = serial

        disks = payload.get('discos', [])
        if isinstance(disks, dict):
            disks = [disks]
        if not isinstance(disks, list):
            disks = []

        physical_disks = []
        for disk in disks:
            if not isinstance(disk, dict):
                continue
            size_label, size_bytes = _safe_round_gb(disk.get('Size'))
            if _is_virtual_disk(disk.get('Model'), size_bytes, disk.get('InterfaceType'), disk.get('MediaType')):
                continue
            physical_disks.append((disk, size_label))

        if len(physical_disks) > 0:
            disk1, size_label = physical_disks[0]
            fallback['disco1_capacidad'] = size_label
            fallback['disco1_tipo'] = _infer_disk_type(disk1.get('MediaType'), disk1.get('Model'), disk1.get('InterfaceType'))
            serial_disk1 = str(disk1.get('SerialNumber') or '').strip()
            fallback['disco1_serial'] = serial_disk1 if serial_disk1 else 'No detectado'
            fallback['disco1_marca'] = _clean_disk_brand(disk1.get('Manufacturer'))
            model_disk1 = str(disk1.get('Model') or '').strip()
            fallback['disco1_modelo'] = model_disk1 if model_disk1 else 'No detectado'

        if len(physical_disks) > 1:
            disk2, size_label = physical_disks[1]
            fallback['disco2_capacidad'] = size_label
            fallback['disco2_tipo'] = _infer_disk_type(disk2.get('MediaType'), disk2.get('Model'), disk2.get('InterfaceType'))
            serial_disk2 = str(disk2.get('SerialNumber') or '').strip()
            fallback['disco2_serial'] = serial_disk2 if serial_disk2 else 'No detectado'
            fallback['disco2_marca'] = _clean_disk_brand(disk2.get('Manufacturer'))
            model_disk2 = str(disk2.get('Model') or '').strip()
            fallback['disco2_modelo'] = model_disk2 if model_disk2 else 'No detectado'

        return fallback

    def _merge_missing_fields(primary, secondary):
        for key, value in secondary.items():
            if _is_placeholder(primary.get(key)) and not _is_placeholder(value):
                primary[key] = value
        return primary

    def _apply_storage_disks(target, storage_disks):
        if len(storage_disks) > 0:
            d1 = storage_disks[0]
            target['disco1_capacidad'] = d1['capacidad']
            target['disco1_tipo'] = d1['tipo']
            target['disco1_serial'] = d1['serial']
            target['disco1_marca'] = d1['marca']
            target['disco1_modelo'] = d1['modelo']

        if len(storage_disks) > 1:
            d2 = storage_disks[1]
            target['disco2_capacidad'] = d2['capacidad']
            target['disco2_tipo'] = d2['tipo']
            target['disco2_serial'] = d2['serial']
            target['disco2_marca'] = d2['marca']
            target['disco2_modelo'] = d2['modelo']
        else:
            target['disco2_capacidad'] = 'No tiene'
            target['disco2_tipo'] = 'No tiene'
            target['disco2_serial'] = 'No tiene'
            target['disco2_marca'] = 'No tiene'
            target['disco2_modelo'] = 'No tiene'

        return target

    if not HAS_WMI:
        return _detect_hardware_cim_fallback()

    try:
        try:
            import pythoncom
            pythoncom.CoInitialize()
        except Exception:
            pass

        c = wmi.WMI()

        for system in c.Win32_ComputerSystem():
            info['marca'] = system.Manufacturer or 'No detectado'
            info['modelo'] = system.Model or 'No detectado'

        serial_found = False
        serials_invalidos = ['default string', 'to be filled by o.e.m.',
                            'system serial number', 'base board serial number',
                            'chassis serial number', '']

        try:
            wmic_result = subprocess.run(
                ['wmic', 'bios', 'get', 'serialnumber'],
                capture_output=True,
                text=True,
                timeout=4
            )
            wmic_lines = [line.strip() for line in (wmic_result.stdout or '').splitlines() if line.strip()]
            if len(wmic_lines) >= 2:
                serial_wmic = wmic_lines[1].strip()
                if serial_wmic and serial_wmic.lower() not in serials_invalidos:
                    info['serial'] = serial_wmic
                    serial_found = True
        except Exception:
            pass

        if not serial_found:
            for bios in c.Win32_BIOS():
                serial = (bios.SerialNumber or '').strip()
                if serial and serial.lower() not in serials_invalidos:
                    info['serial'] = serial
                    serial_found = True
                    break

        if not serial_found:
            for board in c.Win32_BaseBoard():
                serial = (board.SerialNumber or '').strip()
                if serial and serial.lower() not in serials_invalidos:
                    info['serial'] = f"MB-{serial}"
                    serial_found = True
                    break

        if not serial_found:
            for product in c.Win32_ComputerSystemProduct():
                serial = (product.IdentifyingNumber or '').strip()
                if serial and serial.lower() not in serials_invalidos:
                    info['serial'] = serial
                    serial_found = True
                    break

        if not serial_found:
            info['serial'] = "No detectado (PC genérico/armado)"

        disks = list(c.Win32_DiskDrive())
        physical_disks = []
        for disk in disks:
            _, size_bytes = _safe_round_gb(getattr(disk, 'Size', 0))
            if _is_virtual_disk(getattr(disk, 'Model', ''), size_bytes, getattr(disk, 'InterfaceType', ''), getattr(disk, 'MediaType', '')):
                continue
            physical_disks.append(disk)

        if len(physical_disks) > 0:
            disk1 = physical_disks[0]

            try:
                size_gb, _ = _safe_round_gb(disk1.Size)
                info['disco1_capacidad'] = size_gb
            except Exception:
                info['disco1_capacidad'] = 'No detectado'

            info['disco1_tipo'] = _infer_disk_type(disk1.MediaType, getattr(disk1, 'Model', ''), getattr(disk1, 'InterfaceType', ''))

            serial_disk = (disk1.SerialNumber or '').strip()
            info['disco1_serial'] = serial_disk if serial_disk else 'No detectado'

            marca_disk = (disk1.Manufacturer or '').strip()
            info['disco1_marca'] = _clean_disk_brand(marca_disk)

            modelo_disk = (disk1.Model or '').strip()
            info['disco1_modelo'] = modelo_disk if modelo_disk else 'No detectado'

        if len(physical_disks) > 1:
            disk2 = physical_disks[1]

            try:
                size_gb, _ = _safe_round_gb(disk2.Size)
                info['disco2_capacidad'] = size_gb
            except Exception:
                info['disco2_capacidad'] = 'Detectado'

            info['disco2_tipo'] = _infer_disk_type(disk2.MediaType, getattr(disk2, 'Model', ''), getattr(disk2, 'InterfaceType', ''))

            serial_disk2 = (disk2.SerialNumber or '').strip()
            info['disco2_serial'] = serial_disk2 if serial_disk2 else 'No detectado'

            marca_disk2 = (disk2.Manufacturer or '').strip()
            info['disco2_marca'] = _clean_disk_brand(marca_disk2)

            modelo_disk2 = (disk2.Model or '').strip()
            info['disco2_modelo'] = modelo_disk2 if modelo_disk2 else 'No detectado'

    except Exception as e:
        print(f"Error WMI: {e}")

    cim_fallback = _detect_hardware_cim_fallback()
    info = _merge_missing_fields(info, cim_fallback)

    storage_disks = _detect_storage_api_disks()
    if storage_disks:
        info = _apply_storage_disks(info, storage_disks)

    return info


def detect_network_speed():
    """Detectar velocidad máxima soportada por la tarjeta de red principal."""
    if not HAS_WMI:
        return "No detectado"
    
    try:
        import pythoncom
        try:
            pythoncom.CoInitialize()
        except:
            pass
        
        c = wmi.WMI()
        
        # Buscar adaptadores físicos activos
        for adapter in c.Win32_NetworkAdapter():
            if adapter.PhysicalAdapter and adapter.NetEnabled:
                speed = adapter.Speed
                if speed:
                    try:
                        speed_bps = int(speed)
                        speed_mbps = speed_bps / 1000000
                        
                        if speed_mbps >= 1000:
                            return f"{speed_mbps/1000:.1f} Gbps"
                        else:
                            return f"{speed_mbps:.0f} Mbps"
                    except:
                        continue
        
        return "No detectado"
    
    except Exception as e:
        print(f"Error detectando velocidad de red: {e}")
        return "No detectado"


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
            timeout=6
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
            timeout=6
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


def _run_powershell(command, timeout=8):
    """Ejecutar comando de PowerShell y devolver salida limpia."""
    try:
        result = subprocess.run(
            ['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', command],
            capture_output=True,
            text=True,
            timeout=timeout
        )
        output = (result.stdout or '').strip()
        return output if output else None
    except Exception:
        return None


def detect_tpm_status():
    """Detectar estado de TPM."""
    output = _run_powershell(
        "$t=Get-Tpm -ErrorAction SilentlyContinue;"
        "if($null -eq $t){'No detectado'}"
        "elseif(-not $t.TpmPresent){'No presente'}"
        "elseif($t.TpmReady){'Presente y listo'}"
        "else{'Presente (no listo)'}",
        timeout=6
    )
    return output or "No detectado"


def detect_bitlocker_status():
    """Detectar estado de BitLocker en unidad del sistema."""
    output = _run_powershell(
        "$v=Get-BitLockerVolume -MountPoint $env:SystemDrive -ErrorAction SilentlyContinue;"
        "if($null -eq $v){'No disponible'}"
        "elseif($v.ProtectionStatus -eq 'On' -or $v.ProtectionStatus -eq 1){"
        "'Activo (' + $v.EncryptionPercentage + '%)'}"
        "else{'Inactivo'}",
        timeout=6
    )

    if output:
        return output

    fallback = _run_powershell(
        "manage-bde -status $env:SystemDrive | Select-String 'Protection Status|Estado de protección' | ForEach-Object { $_.Line }",
        timeout=6
    )
    if fallback:
        text = fallback.lower()
        if 'on' in text or 'activado' in text:
            return 'Activo'
        if 'off' in text or 'desactivado' in text:
            return 'Inactivo'
    return "No detectado"


def detect_defender_status():
    """Detectar estado real de Microsoft Defender."""
    output = _run_powershell(
        "$d=Get-MpComputerStatus -ErrorAction SilentlyContinue;"
        "if($null -eq $d){'No detectado'}"
        "elseif($d.AntivirusEnabled -and $d.RealTimeProtectionEnabled -and -not $d.DefenderSignaturesOutOfDate){'Activo'}"
        "elseif($d.AntivirusEnabled -and $d.RealTimeProtectionEnabled){'Activo (firmas desactualizadas)'}"
        "elseif($d.AntivirusEnabled){'Parcial'}"
        "else{'Desactivado'}",
        timeout=6
    )
    return output or "No detectado"


def detect_cpu_cores():
    """Detectar núcleos físicos de CPU."""
    try:
        if HAS_PSUTIL:
            cores = psutil.cpu_count(logical=False)
            if cores:
                return str(cores)
    except Exception:
        pass

    if HAS_WMI:
        try:
            c = wmi.WMI()
            for cpu in c.Win32_Processor():
                cores = getattr(cpu, 'NumberOfCores', None)
                if cores:
                    return str(cores)
        except Exception:
            pass

    return "No detectado"


def detect_current_cpu_ram_usage():
    """Detectar uso actual de CPU y RAM (%)."""
    if not HAS_PSUTIL:
        return "No detectado", "No detectado"

    try:
        cpu_percent = psutil.cpu_percent(interval=0.7)
        ram_percent = psutil.virtual_memory().percent
        return f"{cpu_percent:.1f}%", f"{ram_percent:.1f}%"
    except Exception:
        return "No detectado", "No detectado"


def detect_last_windows_update():
    """Detectar última actualización de Windows."""
    try:
        hotfix = _run_powershell(
            "Get-HotFix | Sort-Object InstalledOn -Descending | Select-Object -First 1 HotFixID, InstalledOn | ForEach-Object { $_.HotFixID + ' - ' + $_.InstalledOn.ToString('yyyy-MM-dd') }",
            timeout=6
        )
        if hotfix:
            return hotfix

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


def detect_switch_and_vlan():
    """Detectar switch/puerto (si LLDP está disponible) y VLAN de la interfaz principal."""
    try:
        cmd = r"""
        $adaptersUp = Get-NetAdapter -ErrorAction SilentlyContinue |
            Where-Object {
                $_.Status -eq 'Up' -and $_.HardwareInterface -eq $true -and
                $_.InterfaceDescription -notmatch 'Virtual|Hyper-V|VMware|Bluetooth|Loopback|Teredo|VPN|Pseudo'
            }

        $ethernet = $adaptersUp |
            Where-Object {
                $_.Name -match 'Ethernet|LAN' -or
                $_.InterfaceDescription -match 'Ethernet|Gigabit|Realtek|Intel\(R\).*Ethernet|Killer E\d+'
            } |
            Select-Object -First 1

        $adapter = $ethernet

        if (-not $adapter) {
            $ifIndex = (Get-NetRoute -DestinationPrefix '0.0.0.0/0' -ErrorAction SilentlyContinue |
                        Sort-Object { $_.RouteMetric + $_.InterfaceMetric } |
                        Select-Object -First 1 -ExpandProperty ifIndex)

            if ($ifIndex) {
                $adapter = $adaptersUp | Where-Object { $_.ifIndex -eq $ifIndex } | Select-Object -First 1
            }
        }

        if (-not $adapter) {
            $adapter = $adaptersUp | Select-Object -First 1
        }

        if (-not $adapter) {
            [pscustomobject]@{ switch_puerto = 'No detectado'; vlan = 'No detectado' } | ConvertTo-Json -Compress
            return
        }

        $vlan = 'No detectado'
        $advProps = Get-NetAdapterAdvancedProperty -Name $adapter.Name -AllProperties -ErrorAction SilentlyContinue

        # 1) Prioridad: buscar propiedad explícita de VLAN ID
        $vlanIdProp = $advProps |
            Where-Object { $_.DisplayName -match 'VLAN\s*ID|VLAN.*Identifier|ID de VLAN' -or $_.RegistryKeyword -match 'VlanId|VLANID|RegVlanid' } |
            Select-Object -First 1

        if ($vlanIdProp) {
            $vlanCandidates = @()
            if ($vlanIdProp.DisplayValue) { $vlanCandidates += $vlanIdProp.DisplayValue.ToString().Trim() }
            if ($vlanIdProp.RegistryValue) { $vlanCandidates += ($vlanIdProp.RegistryValue -join ' ').Trim() }

            foreach ($candidate in $vlanCandidates) {
                if ($candidate -match '^\d{1,4}$') {
                    $vlan = $candidate
                    break
                }
                if ($candidate -match '(\d{1,4})') {
                    $vlan = $matches[1]
                    break
                }
            }
        }

        # 2) Fallback: estado de VLAN habilitada/deshabilitada
        if ($vlan -eq 'No detectado') {
            $vlanStateProp = $advProps |
                Where-Object {
                    $_.DisplayName -match 'Priority\s*&\s*VLAN|VLAN' -or
                    $_.RegistryKeyword -match 'PriorityVlanTag|VLAN|Vlan'
                } |
                Select-Object -First 1

            if ($vlanStateProp -and $vlanStateProp.DisplayValue) {
                $v = $vlanStateProp.DisplayValue.ToString().Trim()
                if ($v -and $v -notmatch 'Disabled|Deshabilitado|Not Present|No\s*VLAN') {
                    $vlan = $v
                }
            }
        }

        $switchPuerto = 'No detectado'

        if (Get-Command Get-NetLldpAgent -ErrorAction SilentlyContinue) {
            $lldp = Get-NetLldpAgent -Name $adapter.Name -ErrorAction SilentlyContinue
            if ($lldp) {
                $bridge = $null
                $port = $null

                foreach ($prop in @('NearestBridge', 'NearestBridgeChassisId', 'NearestCustomerBridge')) {
                    if ($lldp.PSObject.Properties.Name -contains $prop) {
                        $val = $lldp.$prop
                        if ($val) { $bridge = $val; break }
                    }
                }

                foreach ($prop in @('NearestBridgePortId', 'NearestBridgePortDescription')) {
                    if ($lldp.PSObject.Properties.Name -contains $prop) {
                        $val = $lldp.$prop
                        if ($val) { $port = $val; break }
                    }
                }

                if ($bridge -and $port) {
                    $switchPuerto = "$bridge / $port"
                } elseif ($bridge) {
                    $switchPuerto = "$bridge"
                }
            }
        }

        if ($switchPuerto -eq 'No detectado') {
            $switchPuerto = "Adaptador: $($adapter.Name)"
        }

        if ($vlan -eq 'No detectado' -and ($adapter.Name -match 'Wi-?Fi|WLAN' -or $adapter.InterfaceDescription -match 'Wireless|Wi-?Fi|802\.11')) {
            $vlan = 'N/A (Wi-Fi)'
        }

        [pscustomobject]@{
            switch_puerto = $switchPuerto
            vlan = $vlan
        } | ConvertTo-Json -Compress
        """

        output = _run_powershell(cmd, timeout=10)
        if not output:
            return "No detectado", "No detectado"

        data = json.loads(output)
        switch_puerto = str(data.get("switch_puerto", "No detectado") or "No detectado").strip()
        vlan = str(data.get("vlan", "No detectado") or "No detectado").strip()
        return switch_puerto, vlan

    except Exception:
        return "No detectado", "No detectado"
    
def detect_mac_address():
    """Detectar dirección MAC de la interfaz de red principal."""
    try:
        import uuid
        mac = ':'.join(['{:02x}'.format((uuid.getnode() >> elements) & 0xff)
                       for elements in range(0,2*6,2)][::-1])
        return mac.upper()
    except:
        return "No detectado"


def detect_default_browser():
    """Detectar navegador predeterminado en Windows."""
    if not HAS_WINREG:
        return "No detectado"
    
    try:
        # Leer asociación de protocolo http
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\\Microsoft\\Windows\\Shell\\Associations\\UrlAssociations\\http\\UserChoice"
        )
        prog_id = winreg.QueryValueEx(key, "ProgId")[0]
        winreg.CloseKey(key)
        
        # Mapear ProgId a nombre de navegador
        browser_map = {
            'ChromeHTML': 'Google Chrome',
            'FirefoxURL': 'Mozilla Firefox',
            'MSEdgeHTM': 'Microsoft Edge',
            'IE.HTTP': 'Internet Explorer',
            'BraveHTML': 'Brave',
            'OperaStable': 'Opera'
        }
        
        for key_name, browser_name in browser_map.items():
            if key_name in prog_id:
                return browser_name
        
        return "Otro navegador"
    
    except Exception as e:
        # Fallback: buscar ejecutables comunes
        browsers = [
            (r"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe", "Google Chrome"),
            (r"C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe", "Google Chrome"),
            (r"C:\\Program Files\\Mozilla Firefox\\firefox.exe", "Mozilla Firefox"),
            (r"C:\\Program Files (x86)\\Mozilla Firefox\\firefox.exe", "Mozilla Firefox"),
            (r"C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe", "Microsoft Edge"),
        ]
        
        for path, name in browsers:
            if os.path.exists(path):
                return f"{name} (detectado)"
        
        return "No detectado"


def detect_network_drives():
    """Detectar unidades de red mapeadas (ej: Z:\\, Y:\\)."""
    try:
        import subprocess
        
        # Ejecutar comando "net use" para listar unidades de red
        result = subprocess.run(
            ['net', 'use'],
            capture_output=True,
            text=True,
            timeout=4
        )
        
        output = result.stdout
        drives = []
        
        # Parsear salida de "net use"
        for line in output.split('\\n'):
            # Buscar líneas con unidades (formato: OK   Z:   \\\\servidor\\carpeta)
            if ':' in line and '\\\\\\\\' in line:
                parts = line.split()
                for part in parts:
                    if ':' in part and len(part) == 2:
                        drives.append(part)
        
        if drives:
            return ', '.join(sorted(set(drives)))
        else:
            return "Ninguna"
    
    except Exception as e:
        return "No detectado"


def detect_ip_local():
    """Detectar IP local del equipo."""
    try:
        # Método 1: Conectar a servidor externo (más confiable)
        import socket
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        local_ip = s.getsockname()[0]
        s.close()
        return local_ip
    except:
        try:
            # Método 2: Usar hostname
            hostname = socket.gethostname()
            local_ip = socket.gethostbyname(hostname)
            return local_ip
        except:
            return "No detectado"


def detect_hostname_local():
    """Detectar nombre del equipo local."""
    try:
        hostname = socket.gethostname()
        if hostname and hostname.strip():
            return hostname.strip()
        return "No detectado"
    except:
        return "No detectado"


def extract_json_object(text):
    """Extraer primer objeto JSON válido desde un texto."""
    if not text:
        return None

    text = text.strip()

    try:
        parsed_direct = json.loads(text)
        if isinstance(parsed_direct, dict):
            return parsed_direct
    except Exception:
        pass

    block_match = re.search(r"```(?:json)?\s*(\{[\s\S]*?\})\s*```", text, re.IGNORECASE)
    if block_match:
        candidate = block_match.group(1).strip()
        try:
            parsed_block = json.loads(candidate)
            if isinstance(parsed_block, dict):
                return parsed_block
        except Exception:
            pass

    brace_match = re.search(r"\{[\s\S]*\}", text)
    if brace_match:
        candidate = brace_match.group(0).strip()
        try:
            parsed_brace = json.loads(candidate)
            if isinstance(parsed_brace, dict):
                return parsed_brace
        except Exception:
            return None

    return None


def call_ollama_inventory_assistant(form_data, model_name="llama3.2:3b", timeout_sec=45):
    """Consultar Ollama local para sugerencias de inventario."""
    prompt = (
        "Eres asistente técnico de inventario hospitalario. "
        "Analiza el formulario y responde SOLO JSON válido con este esquema exacto: "
        "{\"campos_faltantes\":[],\"inconsistencias\":[],\"recomendaciones\":[],"
        "\"observaciones_tecnicas_sugeridas\":\"\",\"estado_operativo_sugerido\":\"\"}. "
        "No uses markdown ni texto adicional. "
        "Si no hay hallazgos, devuelve arreglos vacíos. "
        "Los campos del formulario son:\n"
        f"{json.dumps(form_data, ensure_ascii=False, indent=2)}"
    )

    payload = {
        "model": model_name,
        "prompt": prompt,
        "stream": False,
        "options": {
            "temperature": 0.2
        }
    }

    request = urllib.request.Request(
        "http://127.0.0.1:11434/api/generate",
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json"},
        method="POST"
    )

    with urllib.request.urlopen(request, timeout=timeout_sec) as response:
        body = response.read().decode("utf-8", errors="replace")

    response_json = json.loads(body)
    raw_model_text = (response_json.get("response") or "").strip()

    parsed_result = extract_json_object(raw_model_text)
    if not parsed_result:
        raise ValueError("La respuesta del modelo no vino en JSON válido")

    parsed_result.setdefault("campos_faltantes", [])
    parsed_result.setdefault("inconsistencias", [])
    parsed_result.setdefault("recomendaciones", [])
    parsed_result.setdefault("observaciones_tecnicas_sugeridas", "")
    parsed_result.setdefault("estado_operativo_sugerido", "")

    return parsed_result


def detect_best_ollama_model(preferred_model=None):
    """Elegir el mejor modelo disponible en Ollama local."""
    candidates = []
    if preferred_model:
        candidates.append(preferred_model)

    candidates.extend([
        "llama3.1:8b",
        "llama3.2:3b",
        "llama3.2:1b",
        "mistral:7b"
    ])

    request = urllib.request.Request(
        "http://127.0.0.1:11434/api/tags",
        method="GET"
    )

    with urllib.request.urlopen(request, timeout=8) as response:
        body = response.read().decode("utf-8", errors="replace")

    data = json.loads(body)
    installed = [model.get("name", "") for model in data.get("models", []) if model.get("name")]

    if not installed:
        return preferred_model or "llama3.2:3b"

    for candidate in candidates:
        if candidate in installed:
            return candidate

    return installed[0]


def get_ollama_executable_path():
    """Obtener ruta del ejecutable de Ollama en Windows."""
    by_path = shutil.which("ollama")
    if by_path and os.path.exists(by_path):
        return by_path

    local_candidate = os.path.join(
        os.environ.get("LOCALAPPDATA", ""),
        "Programs",
        "Ollama",
        "ollama.exe"
    )

    if os.path.exists(local_candidate):
        return local_candidate

    return None


def is_ollama_api_ready(timeout_sec=3):
    """Verificar si la API local de Ollama responde."""
    try:
        request = urllib.request.Request("http://127.0.0.1:11434/api/tags", method="GET")
        with urllib.request.urlopen(request, timeout=timeout_sec) as response:
            body = response.read().decode("utf-8", errors="replace")
        data = json.loads(body)
        return isinstance(data, dict) and "models" in data
    except Exception:
        return False


def get_installed_ollama_models():
    """Retornar nombres de modelos instalados en Ollama local."""
    request = urllib.request.Request("http://127.0.0.1:11434/api/tags", method="GET")
    with urllib.request.urlopen(request, timeout=8) as response:
        body = response.read().decode("utf-8", errors="replace")

    data = json.loads(body)
    return [model.get("name", "") for model in data.get("models", []) if model.get("name")]


def normalize_text_for_compare(value):
    """Normalizar texto para comparación flexible."""
    raw = str(value or "").strip().lower()
    if not raw:
        return ""

    normalized = unicodedata.normalize("NFD", raw)
    normalized = "".join(ch for ch in normalized if unicodedata.category(ch) != "Mn")
    normalized = re.sub(r"\s+", " ", normalized)
    return normalized.strip()


def normalize_alnum(value):
    """Normalizar valor alfanumérico (útil para serial/MAC)."""
    return re.sub(r"[^a-zA-Z0-9]", "", str(value or "")).lower()


def extract_first_number(value):
    """Extraer primer número de un texto."""
    match = re.search(r"\d+(?:[\.,]\d+)?", str(value or ""))
    if not match:
        return ""
    return match.group(0).replace(",", ".")


def read_docx_text(docx_path):
    """Leer texto plano de un archivo .docx (párrafos y tablas)."""
    if not docx_path.lower().endswith(".docx"):
        raise ValueError("Solo se admite formato .docx")

    with zipfile.ZipFile(docx_path, "r") as docx_zip:
        with docx_zip.open("word/document.xml") as xml_file:
            xml_content = xml_file.read()

    root = ET.fromstring(xml_content)
    namespaces = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

    lines = []
    for paragraph in root.findall(".//w:p", namespaces):
        text_parts = []
        for node in paragraph.findall(".//w:t", namespaces):
            if node.text:
                text_parts.append(node.text)

        paragraph_text = " ".join(text_parts).strip()
        if paragraph_text:
            lines.append(paragraph_text)

    full_text = "\n".join(lines)
    return full_text, lines


def extract_value_from_doc(full_text, lines, regex_patterns, line_labels=None):
    """Extraer un valor desde texto de hoja de vida usando regex y fallback por líneas."""
    for pattern in regex_patterns:
        match = re.search(pattern, full_text, flags=re.IGNORECASE | re.MULTILINE)
        if match:
            value = (match.group(1) or "").strip()
            if value:
                return value

    if line_labels:
        normalized_labels = [normalize_text_for_compare(label) for label in line_labels]
        for idx, line in enumerate(lines):
            normalized_line = normalize_text_for_compare(line)
            if normalized_line in normalized_labels:
                if idx + 1 < len(lines):
                    candidate = (lines[idx + 1] or "").strip()
                    if candidate:
                        return candidate

    return ""


def extract_value_by_label(lines, labels):
    """Extraer valor por etiqueta en la misma línea o en la línea siguiente (útil para tablas)."""
    if not lines:
        return ""

    normalized_labels = [normalize_text_for_compare(label) for label in labels]

    for idx, raw_line in enumerate(lines):
        line = (raw_line or "").strip()
        if not line:
            continue

        normalized_line = normalize_text_for_compare(line)

        for normalized_label in normalized_labels:
            if not normalized_label:
                continue

            if normalized_line.startswith(normalized_label):
                same_line_value = ""
                if ":" in line:
                    same_line_value = (line.split(":", 1)[1] or "").strip()
                elif "-" in line:
                    same_line_value = (line.split("-", 1)[1] or "").strip()
                else:
                    remainder = normalized_line[len(normalized_label):].strip(" .:-")
                    if remainder:
                        same_line_value = remainder

                if same_line_value:
                    return same_line_value

                if idx + 1 < len(lines):
                    next_line = (lines[idx + 1] or "").strip()
                    if next_line:
                        next_normalized = normalize_text_for_compare(next_line)
                        if next_normalized not in normalized_labels:
                            return next_line

    return ""


def extract_cpu_component_fields(lines):
    """Extraer Marca/Modelo/Serial desde tabla de componentes en fila CPU."""
    if not lines:
        return "", "", ""

    normalized_lines = [normalize_text_for_compare(line) for line in lines]

    for index, normalized in enumerate(normalized_lines):
        if normalized != "cpu":
            continue

        context_start = max(0, index - 8)
        context = normalized_lines[context_start:index]
        has_headers = all(header in context for header in ["componente", "marca", "modelo", "serial"])
        if not has_headers:
            continue

        marca = (lines[index + 1].strip() if index + 1 < len(lines) else "")
        modelo = (lines[index + 2].strip() if index + 2 < len(lines) else "")
        serial = (lines[index + 3].strip() if index + 3 < len(lines) else "")

        forbidden = {
            "", "marca", "modelo", "serial", "especif.", "cod. inv.",
            "componente", "monitor", "teclado", "mouse"
        }

        marca_norm = normalize_text_for_compare(marca)
        modelo_norm = normalize_text_for_compare(modelo)
        serial_norm = normalize_text_for_compare(serial)

        if marca_norm in forbidden:
            marca = ""
        if modelo_norm in forbidden:
            modelo = ""
        if serial_norm in forbidden:
            serial = ""

        return marca, modelo, serial

    return "", "", ""


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
        
        # Variables para guardar consecutivo correcto (FIX BUG #1)
        self.last_saved_consecutive = None
        self.last_saved_code = None
        
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
        """Crear menú nativo de tkinter."""
        
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
        menu_operaciones.add_command(
            label="Comparar Hoja de Vida",
            command=lambda: self.show_form_directo("Comparar Hoja de Vida")
        )
        menu_operaciones.add_separator()
        menu_operaciones.add_command(
            label="Generar Hoja de Vida (formato oficial .docx)",
            command=self.generar_hoja_vida_oficial
        )
        menu_operaciones.add_command(
            label="Generar Reporte de Mantenimiento (formato oficial .pdf)",
            command=self.abrir_dialogo_mantenimiento_pdf
        )
        menu_operaciones.add_separator()
        menu_operaciones.add_command(
            label="Inventario Rápido (Externo)",
            command=self.launch_inventario_rapido
        )
        
        # MENÚ AYUDA
        menu_ayuda = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Ayuda", menu=menu_ayuda)
        menu_ayuda.add_command(label="Guía de Formulario", command=self.show_classification_guide)

    def launch_inventario_rapido(self):
        """Lanzar proyecto externo Inventario Rápido como módulo integrado."""
        try:
            if getattr(sys, "frozen", False):
                base_dir = Path(sys.executable).resolve().parent
            else:
                base_dir = Path(__file__).resolve().parent

            search_roots = [base_dir, *list(base_dir.parents)[:4], Path(__file__).resolve().parent, Path(__file__).resolve().parent.parent]
            candidate_dirs = []
            for root in search_roots:
                candidate = root / "Script"
                if candidate not in candidate_dirs:
                    candidate_dirs.append(candidate)

            external_project_dir = next(
                (candidate for candidate in candidate_dirs if (candidate / "main.py").exists()),
                candidate_dirs[0]
            )
            external_main = external_project_dir / "main.py"

            if not external_main.exists():
                messagebox.showerror(
                    "Proyecto no encontrado",
                    "No se encontró el proyecto externo 'Inventario Rápido'.\n"
                    f"Ruta intentada: {external_main}\n\n"
                    "Coloca la carpeta 'Script' junto al .exe "
                    "(o junto a la carpeta del proyecto en modo desarrollo)."
                )
                return

            external_venv_python = external_project_dir / ".venv" / "Scripts" / "python.exe"
            python_executable = str(external_venv_python) if external_venv_python.exists() else sys.executable

            subprocess.Popen(
                [python_executable, str(external_main)],
                cwd=str(external_project_dir)
            )

            messagebox.showinfo(
                "Módulo abierto",
                "Inventario Rápido se abrió en una ventana separada."
            )

        except Exception as e:
            messagebox.showerror("Error", f"No fue posible abrir Inventario Rápido:\n{e}")
    
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
            lambda: self.show_form_in_container(self.create_impresoras_form)()
        elif tipo == "Periféricos":
            lambda: self.show_form_in_container(self.create_perifericos_form)()
        elif tipo == "Red":
            lambda: self.show_form_in_container(self.create_red_form)()
        elif tipo == "Mantenimiento":
            lambda: self.show_form_in_container(self.create_mantenimientos_form)()
        elif tipo == "Dados de Baja":
            lambda: self.show_form_in_container(self.create_baja_form)()
        elif tipo == "Comparar Hoja de Vida":
            lambda: self.show_form_in_container(self.create_hoja_vida_compare_form)()
    
    def show_manual_form_in_container(self):
        """Mostrar formulario manual:"""
        # Limpiar contenedor
        for widget in self.main_container.winfo_children():
            widget.destroy()
        
        self.manual_widgets = {}
        
        # Frame scrollable con label_text
        try:
            next_code = self.get_next_codigo()
        except:
            next_code = "EQC-0001"
        
        form_frame = ctk.CTkScrollableFrame(
            self.main_container, 
            fg_color="#F5F5F5",
            corner_radius=0,
            label_text=f"📝 EQUIPOS DE CÓMPUTO - Código: {next_code}",
            label_fg_color="#CCCCCC",
            label_text_color="#333333",
            label_font=("Segoe UI", 12)
        )
        form_frame.pack(fill="both", expand=True)
        
        # Guardar referencia para poder actualizar el título después
        self.equipo_form_frame = form_frame
        
        # ===== TÍTULO =====
        title_frame = ctk.CTkFrame(form_frame, fg_color=COLOR_VERDE_HOSPITAL, corner_radius=12)
        title_frame.pack(fill="x", padx=20, pady=(20, 10))
        
        # Obtener código siguiente
        try:
            next_code = self.get_next_codigo()
            codigo_text = f"Equipo: {next_code}"
        except:
            codigo_text = "Equipo: EQC-0001"
        
        # GUARDAR REFERENCIA AL TÍTULO PARA ACTUALIZARLO DESPUÉS
        self.form_title_label = ctk.CTkLabel(
            title_frame,
            text=f"{codigo_text}",
            font=("Segoe UI", 16, "bold"),
            text_color="white"
        )
        self.form_title_label.pack(pady=15)
        
        # ===== CAMPOS BÁSICOS =====
        self.create_form_field_centered(form_frame, "Tipo de Equipo", "tipo_equipo", 
                                        "combobox", TIPOS_EQUIPO)
        self.create_form_field_centered(form_frame, "Área / Servicio", "area_servicio", 
                                        "combobox", AREAS_SERVICIO)
        self.manual_widgets['area_servicio'].configure(
            command=lambda choice: self.on_area_servicio_change(choice)
        )
        self.create_form_field_centered(form_frame, "Ubicación Específica", "ubicacion_especifica", 
                                        "entry")
        self.create_form_field_centered(form_frame, "Responsable / Custodio", "responsable_custodio", 
                                        "entry")
        
        # ===== SECCIÓN: CLASIFICACIÓN POR PROCESOS =====
        separator1 = ctk.CTkFrame(form_frame, height=2, fg_color="#CCCCCC")
        separator1.pack(fill="x", padx=20, pady=15)
        
        label_procesos = ctk.CTkLabel(
            form_frame,
            text="🏥 CLASIFICACIÓN POR PROCESOS",
            font=("Segoe UI", 14, "bold"),
            text_color=COLOR_AZUL_HOSPITAL
        )
        label_procesos.pack(pady=(10, 15))
        
        self.create_form_field_centered(form_frame, "Macroproceso", "macroproceso", 
                                        "combobox", list(MACROPROCESOS.keys()))
        self.create_form_field_centered(form_frame, "Proceso", "proceso", 
                                        "combobox", ["Selecciona primero Macroproceso"])
        self.create_form_field_centered(form_frame, "Subproceso", "subproceso", 
                                        "combobox", ["Selecciona primero Proceso"])
        
        # Configurar eventos condicionales
        self.manual_widgets['macroproceso'].configure(
            command=lambda choice: self.on_macroproceso_change(choice)
        )
        self.manual_widgets['proceso'].configure(
            command=lambda choice: self.on_proceso_change(choice)
        )
        
        # ===== SECCIÓN: SOFTWARE =====
        separator2 = ctk.CTkFrame(form_frame, height=2, fg_color="#CCCCCC")
        separator2.pack(fill="x", padx=20, pady=15)
        
        label_software = ctk.CTkLabel(
            form_frame,
            text="💻 SOFTWARE UTILIZADO",
            font=("Segoe UI", 14, "bold"),
            text_color=COLOR_AZUL_HOSPITAL
        )
        label_software.pack(pady=(10, 15))
        
        self.create_form_field_centered(form_frame, "Uso - SIHOS", "uso_sihos", 
                                        "combobox", SI_NO)
        self.create_form_field_centered(form_frame, "Uso - Office Básico", "uso_office_basico", 
                                        "combobox", SI_NO)
        self.create_form_field_centered(form_frame, "Software Especializado", "software_especializado", 
                                        "combobox", SI_NO)
        self.create_form_field_centered(form_frame, "Descripción Software", "descripcion_software", 
                                        "entry")
        self.create_form_field_centered(form_frame, "Función Principal", "funcion_principal", 
                                        "entry")
        
        # ===== SECCIÓN: INFORMACIÓN OPERATIVA =====
        separator4 = ctk.CTkFrame(form_frame, height=2, fg_color="#CCCCCC")
        separator4.pack(fill="x", padx=20, pady=15)
        
        label_operativo = ctk.CTkLabel(
            form_frame,
            text="⚙️ INFORMACIÓN OPERATIVA",
            font=("Segoe UI", 14, "bold"),
            text_color=COLOR_AZUL_HOSPITAL
        )
        label_operativo.pack(pady=(10, 15))
        
        self.create_form_field_centered(form_frame, "Horario de Uso", "horario_uso", 
                                        "combobox", HORARIOS_USO)
        self.create_form_field_centered(form_frame, "Estado Operativo", "estado_operativo", 
                                        "combobox", ESTADOS_OPERATIVOS)
        self.create_form_field_centered(form_frame, "Periodicidad Mantenimiento", "periodicidad_mtto", 
                                        "combobox", PERIODICIDADES_MTTO)
        self.create_form_field_centered(form_frame, "Responsable Mantenimiento", "responsable_mtto", 
                                        "combobox", TECNICOS_RESPONSABLES)
        self.create_form_field_centered(form_frame, "Observaciones Técnicas", "observaciones_tecnicas", 
                                        "entry")
        
        # ===== SECCIÓN: CUESTIONARIO CON RADIOBUTTONS =====
        from textos_tooltips import (
            CONF_LABELS, CONF_TOOLTIPS,
            INT_LABELS, INT_TOOLTIPS,
            CRIT_LABELS, CRIT_TOOLTIPS
        )
        
        separator3 = ctk.CTkFrame(form_frame, height=2, fg_color="#CCCCCC")
        separator3.pack(fill="x", padx=20, pady=15)
        
        label_cuestionario = ctk.CTkLabel(
            form_frame,
            text="🔒 CUESTIONARIO DE CLASIFICACIÓN (18 Preguntas Sí/No)",
            font=("Segoe UI", 14, "bold"),
            text_color=COLOR_AZUL_HOSPITAL
        )
        label_cuestionario.pack(pady=(10, 5))
        
        info_cuestionario = ctk.CTkLabel(
            form_frame,
            text="Selecciona Sí o No • Pasa el mouse sobre el texto para ver la pregunta completa",
            font=("Segoe UI", 11, "italic"),
            text_color="gray"
        )
        info_cuestionario.pack(pady=(0, 15))
        
        # CONFIDENCIALIDAD (9 preguntas - RadioButtons)
        label_conf = ctk.CTkLabel(
            form_frame,
            text="📋 CONFIDENCIALIDAD (Tipo de información):",
            font=("Segoe UI", 12, "bold"),
            text_color="#6F42C1"
        )
        label_conf.pack(anchor="w", padx=40, pady=(10, 5))
        
        for i, (label_text, tooltip_text) in enumerate(zip(CONF_LABELS, CONF_TOOLTIPS), 1):
            field_name = f"conf_{i}"
            self.create_radio_field_centered(form_frame, label_text, field_name, tooltip_text)
        
        # INTEGRIDAD (3 preguntas - RadioButtons)
        label_int = ctk.CTkLabel(
            form_frame,
            text="🔐 INTEGRIDAD (Compromiso de información):",
            font=("Segoe UI", 12, "bold"),
            text_color="#FD7E14"
        )
        label_int.pack(anchor="w", padx=40, pady=(15, 5))
        
        for i, (label_text, tooltip_text) in enumerate(zip(INT_LABELS, INT_TOOLTIPS), 1):
            field_name = f"int_{i}"
            self.create_radio_field_centered(form_frame, label_text, field_name, tooltip_text)
        
        # CRITICIDAD (6 preguntas - RadioButtons)
        label_crit = ctk.CTkLabel(
            form_frame,
            text="⚠️ CRITICIDAD (Impacto operacional):",
            font=("Segoe UI", 12, "bold"),
            text_color="#DC3545"
        )
        label_crit.pack(anchor="w", padx=40, pady=(15, 5))
        
        for i, (label_text, tooltip_text) in enumerate(zip(CRIT_LABELS, CRIT_TOOLTIPS), 1):
            field_name = f"crit_{i}"
            self.create_radio_field_centered(form_frame, label_text, field_name, tooltip_text)
        
        # ===== BOTONES (3 HORIZONTALES IGUALES) =====
        separator5 = ctk.CTkFrame(form_frame, height=2, fg_color="#CCCCCC")
        separator5.pack(fill="x", padx=20, pady=15)
        
        btn_action_frame = ctk.CTkFrame(form_frame, fg_color="transparent")
        btn_action_frame.pack(pady=20)
        
        BTN_WIDTH = 350
        BTN_HEIGHT = 50
        
        # Botón 1: GUARDAR
        self.btn_save_equipo = ctk.CTkButton(
            btn_action_frame,
            text="💾 GUARDAR",
            command=self.handle_save_equipo,
            font=("Segoe UI", 13, "bold"),
            fg_color=COLOR_VERDE_HOSPITAL,
            hover_color="#1F5039",
            height=BTN_HEIGHT,
            width=BTN_WIDTH
        )
        self.btn_save_equipo.pack(side="left", padx=8)
        
        # Botón 2: ACTUALIZAR
        btn_update = ctk.CTkButton(
            btn_action_frame,
            text="🔄 ACTUALIZAR",
            command=self.update_equipo_computo,
            font=("Segoe UI", 13, "bold"),
            fg_color=COLOR_VERDE_HOSPITAL,
            hover_color="#1F5039",
            height=BTN_HEIGHT,
            width=BTN_WIDTH
        )
        btn_update.pack(side="left", padx=8)
        
        # Botón 3: RECOPILACIÓN AUTOMÁTICA
        btn_collect = ctk.CTkButton(
            btn_action_frame,
            text="➡️ RECOPILACIÓN AUTO",
            command=self.start_automatic_collection,
            font=("Segoe UI", 13, "bold"),
            fg_color=COLOR_VERDE_HOSPITAL,
            hover_color="#1F5039",
            height=BTN_HEIGHT,
            width=BTN_WIDTH
        )
        btn_collect.pack(side="left", padx=8)

        # Botón 4: ASISTENTE IA LOCAL (Ollama)
        btn_ai_frame = ctk.CTkFrame(form_frame, fg_color="transparent")
        btn_ai_frame.pack(pady=(0, 15))

        btn_ai = ctk.CTkButton(
            btn_ai_frame,
            text="🤖 ASISTENTE IA (LOCAL)",
            command=self.run_local_ai_assistant,
            font=("Segoe UI", 13, "bold"),
            fg_color=COLOR_VERDE_HOSPITAL,
            hover_color="#1F5039",
            height=BTN_HEIGHT,
            width=BTN_WIDTH
        )
        btn_ai.pack()

    def on_area_servicio_change(self, selected_area):
        """Permitir capturar un área personalizada cuando se selecciona 'Otra Área / Servicio'."""
        if selected_area != "Otra Área / Servicio":
            return

        area_widget = self.manual_widgets.get("area_servicio")
        if not area_widget:
            return

        dialog = ctk.CTkInputDialog(
            text="Escribe el nombre del área / servicio:",
            title="Nueva Área / Servicio"
        )
        nueva_area = (dialog.get_input() or "").strip()

        if not nueva_area:
            messagebox.showwarning(
                "Área / Servicio",
                "Debes escribir un nombre de área para usar la opción personalizada."
            )
            area_widget.set("")
            return

        valores_actuales = list(area_widget.cget("values"))
        if nueva_area not in valores_actuales:
            if "Otra Área / Servicio" in valores_actuales:
                posicion = valores_actuales.index("Otra Área / Servicio")
                valores_actuales.insert(posicion, nueva_area)
            else:
                valores_actuales.append(nueva_area)
            area_widget.configure(values=valores_actuales)

        area_widget.set(nueva_area)

    def on_macroproceso_change(self, selected_macroproceso):
        """Actualizar lista de Procesos cuando cambia el Macroproceso."""
        # Limpiar proceso y subproceso
        self.manual_widgets["proceso"].set("")
        self.manual_widgets["subproceso"].set("")
        self.manual_widgets["subproceso"].configure(values=[])
        
        # Obtener procesos del macroproceso seleccionado
        from config_listas import get_procesos_por_macroproceso
        procesos = get_procesos_por_macroproceso(selected_macroproceso)
        
        # Actualizar lista de procesos
        self.manual_widgets["proceso"].configure(values=procesos)


    def on_proceso_change(self, selected_proceso):
        """Actualizar lista de Subprocesos cuando cambia el Proceso."""
        # Limpiar subproceso
        self.manual_widgets["subproceso"].set("")
        
        # Obtener macroproceso actual
        macroproceso = self.manual_widgets["macroproceso"].get()
        
        if not macroproceso:
            return
        
        # Obtener subprocesos
        from config_listas import get_subprocesos_por_proceso
        subprocesos = get_subprocesos_por_proceso(macroproceso, selected_proceso)
        
        # Actualizar lista de subprocesos
        self.manual_widgets["subproceso"].configure(values=subprocesos)

    def get_manual_form_data_snapshot(self):
        """Capturar estado actual de los campos del formulario manual."""
        snapshot = {}
        for field_name, widget in self.manual_widgets.items():
            try:
                if isinstance(widget, tk.StringVar):
                    snapshot[field_name] = (widget.get() or "").strip()
                elif hasattr(widget, 'winfo_exists') and widget.winfo_exists():
                    if isinstance(widget, (ctk.CTkEntry, ctk.CTkComboBox)):
                        snapshot[field_name] = (widget.get() or "").strip()
            except Exception:
                snapshot[field_name] = ""
        return snapshot

    def run_local_ai_assistant(self):
        """Ejecutar asistente IA local (Ollama) para revisar formulario."""
        if not self.manual_widgets:
            messagebox.showwarning("Asistente IA", "No hay formulario activo para analizar.")
            return

        form_data = self.get_manual_form_data_snapshot()
        preferred_model = os.getenv("OLLAMA_MODEL", "llama3.2:3b")
        self.ai_model_name = preferred_model

        self.ai_wait_window = ctk.CTkToplevel(self.root)
        self.ai_wait_window.title("Asistente IA Local")
        self.ai_wait_window.geometry("430x180")
        self.ai_wait_window.transient(self.root)
        self.ai_wait_window.grab_set()

        self.ai_wait_label = ctk.CTkLabel(
            self.ai_wait_window,
            text="Preparando IA local...",
            font=("Segoe UI", 12, "bold")
        )
        self.ai_wait_label.pack(pady=(25, 10))

        wait_info = ctk.CTkLabel(
            self.ai_wait_window,
            text="Si es el primer uso en este PC, puede tardar varios minutos.",
            font=("Segoe UI", 11)
        )
        wait_info.pack(pady=(0, 10))

        wait_bar = ctk.CTkProgressBar(self.ai_wait_window, mode="indeterminate", width=320)
        wait_bar.pack(pady=10)
        wait_bar.start()

        def worker():
            try:
                selected_model = self.ensure_ollama_ready_for_ai(preferred_model)
                self.ai_model_name = selected_model
                self.update_ai_wait_status(f"Analizando con IA local ({self.ai_model_name})...")

                result = call_ollama_inventory_assistant(
                    form_data=form_data,
                    model_name=self.ai_model_name,
                    timeout_sec=45
                )
                self.root.after(0, lambda: self.show_ai_result_window(result))
            except urllib.error.URLError:
                self.root.after(0, self.show_ollama_connection_error)
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("Asistente IA", f"Error consultando IA local:\n{e}"))
            finally:
                self.root.after(0, lambda: self.safe_close_ai_wait_window())

        thread = threading.Thread(target=worker, daemon=True)
        thread.start()

    def update_ai_wait_status(self, text):
        """Actualizar texto de estado en la ventana de espera de IA."""
        def _set_text():
            try:
                if hasattr(self, 'ai_wait_label') and self.ai_wait_label and self.ai_wait_label.winfo_exists():
                    self.ai_wait_label.configure(text=text)
            except Exception:
                pass

        self.root.after(0, _set_text)

    def ensure_ollama_ready_for_ai(self, preferred_model):
        """Garantizar Ollama + modelo en primer uso (instalación automática)."""
        if is_ollama_api_ready(timeout_sec=3):
            try:
                return detect_best_ollama_model(preferred_model)
            except Exception:
                return preferred_model or "llama3.2:3b"

        ollama_exe = get_ollama_executable_path()

        if not ollama_exe:
            self.update_ai_wait_status("Instalando Ollama automáticamente...")
            winget_path = shutil.which("winget")

            if not winget_path:
                raise RuntimeError("No se encontró winget para instalación automática de Ollama.")

            install_proc = subprocess.run(
                [
                    winget_path,
                    "install",
                    "--id",
                    "Ollama.Ollama",
                    "-e",
                    "--accept-source-agreements",
                    "--accept-package-agreements"
                ],
                capture_output=True,
                text=True
            )

            if install_proc.returncode != 0:
                raise RuntimeError("Falló instalación automática de Ollama con winget.")

            ollama_exe = get_ollama_executable_path()
            if not ollama_exe:
                raise RuntimeError("Ollama se instaló, pero no se encontró el ejecutable.")

        if not is_ollama_api_ready(timeout_sec=3):
            self.update_ai_wait_status("Iniciando servicio Ollama...")

            creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
            subprocess.Popen(
                [ollama_exe, "serve"],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                creationflags=creationflags
            )

            started = False
            for _ in range(30):
                if is_ollama_api_ready(timeout_sec=2):
                    started = True
                    break
                time.sleep(1)

            if not started:
                raise RuntimeError("No fue posible iniciar el servicio local de Ollama.")

        installed_models = []
        try:
            installed_models = get_installed_ollama_models()
        except Exception:
            installed_models = []

        target_model = preferred_model or "llama3.2:3b"

        if target_model not in installed_models:
            self.update_ai_wait_status(f"Descargando modelo {target_model}...")
            pull_proc = subprocess.run([ollama_exe, "pull", target_model], capture_output=True, text=True)

            if pull_proc.returncode != 0:
                if not installed_models:
                    fallback_model = "llama3.2:3b"
                    self.update_ai_wait_status(f"Intentando modelo alterno {fallback_model}...")
                    fallback_proc = subprocess.run([ollama_exe, "pull", fallback_model], capture_output=True, text=True)
                    if fallback_proc.returncode != 0:
                        raise RuntimeError("No se pudo descargar ningún modelo de IA local.")
                    target_model = fallback_model

        try:
            return detect_best_ollama_model(target_model)
        except Exception:
            return target_model

    def safe_close_ai_wait_window(self):
        """Cerrar ventana de espera del asistente IA de forma segura."""
        try:
            if hasattr(self, 'ai_wait_window') and self.ai_wait_window and self.ai_wait_window.winfo_exists():
                self.ai_wait_window.destroy()
        except Exception:
            pass

    def show_ollama_connection_error(self):
        """Mostrar instrucciones cuando Ollama local no está disponible."""
        messagebox.showerror(
            "Asistente IA - Sin conexión",
            "No se pudo conectar con Ollama local (la app intentó configurarlo automáticamente).\n\n"
            "Verifica en PowerShell:\n"
            "1) ollama serve\n"
            "2) ollama pull llama3.2:3b\n"
            "3) (Opcional) setx OLLAMA_MODEL \"llama3.2:3b\""
        )

    def show_ai_result_window(self, ai_result):
        """Mostrar resultados del asistente IA y permitir aplicar sugerencias."""
        self.ai_last_result = ai_result

        campos_faltantes = ai_result.get("campos_faltantes", [])
        inconsistencias = ai_result.get("inconsistencias", [])
        recomendaciones = ai_result.get("recomendaciones", [])
        observaciones = (ai_result.get("observaciones_tecnicas_sugeridas") or "").strip()
        estado_sugerido = (ai_result.get("estado_operativo_sugerido") or "").strip()

        result_window = ctk.CTkToplevel(self.root)
        result_window.title("Resultado Asistente IA")
        result_window.geometry("760x520")
        result_window.transient(self.root)
        result_window.grab_set()

        header = ctk.CTkLabel(
            result_window,
            text="🤖 Asistente IA Local - Resultado",
            font=("Segoe UI", 16, "bold"),
            text_color=COLOR_VERDE_HOSPITAL
        )
        header.pack(pady=(18, 10))

        text_box = ctk.CTkTextbox(result_window, width=700, height=360, font=("Consolas", 11))
        text_box.pack(padx=20, pady=10, fill="both", expand=True)

        resumen = []
        resumen.append("=== CAMPOS FALTANTES ===")
        resumen.extend([f"- {item}" for item in campos_faltantes] or ["- Ninguno"])
        resumen.append("\n=== INCONSISTENCIAS ===")
        resumen.extend([f"- {item}" for item in inconsistencias] or ["- Ninguna"])
        resumen.append("\n=== RECOMENDACIONES ===")
        resumen.extend([f"- {item}" for item in recomendaciones] or ["- Ninguna"])
        resumen.append("\n=== OBSERVACIONES TÉCNICAS SUGERIDAS ===")
        resumen.append(observaciones or "Sin sugerencia")
        resumen.append("\n=== ESTADO OPERATIVO SUGERIDO ===")
        resumen.append(estado_sugerido or "Sin sugerencia")

        text_box.insert("1.0", "\n".join(resumen))
        text_box.configure(state="disabled")

        btn_frame = ctk.CTkFrame(result_window, fg_color="transparent")
        btn_frame.pack(pady=(8, 16))

        btn_apply = ctk.CTkButton(
            btn_frame,
            text="✅ Aplicar sugerencias",
            command=lambda: self.apply_ai_suggestions(result_window),
            fg_color=COLOR_VERDE_HOSPITAL,
            hover_color="#1F5039",
            width=210,
            height=40,
            font=("Segoe UI", 12, "bold")
        )
        btn_apply.pack(side="left", padx=10)

        btn_close = ctk.CTkButton(
            btn_frame,
            text="Cerrar",
            command=result_window.destroy,
            fg_color="#808080",
            hover_color="#666666",
            width=140,
            height=40,
            font=("Segoe UI", 12, "bold")
        )
        btn_close.pack(side="left", padx=10)

    def apply_ai_suggestions(self, result_window):
        """Aplicar sugerencias de IA a campos del formulario manual."""
        if not hasattr(self, 'ai_last_result'):
            return

        ai_result = self.ai_last_result
        applied = []

        obs_sugerida = (ai_result.get("observaciones_tecnicas_sugeridas") or "").strip()
        if obs_sugerida and 'observaciones_tecnicas' in self.manual_widgets:
            widget = self.manual_widgets['observaciones_tecnicas']
            if isinstance(widget, ctk.CTkEntry):
                current = (widget.get() or "").strip()
                if not current or messagebox.askyesno(
                    "Asistente IA",
                    "El campo 'Observaciones Técnicas' ya tiene texto.\n¿Deseas reemplazarlo con la sugerencia de IA?"
                ):
                    widget.delete(0, 'end')
                    widget.insert(0, obs_sugerida)
                    applied.append("Observaciones Técnicas")

        estado_sugerido = (ai_result.get("estado_operativo_sugerido") or "").strip()
        if estado_sugerido and 'estado_operativo' in self.manual_widgets:
            widget = self.manual_widgets['estado_operativo']
            if isinstance(widget, ctk.CTkComboBox):
                widget.set(estado_sugerido)
                applied.append("Estado Operativo")

        if applied:
            messagebox.showinfo("Asistente IA", f"Sugerencias aplicadas:\n- " + "\n- ".join(applied))
        else:
            messagebox.showinfo("Asistente IA", "No había sugerencias aplicables automáticamente.")

        try:
            if result_window and result_window.winfo_exists():
                result_window.destroy()
        except Exception:
            pass

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
        
    # ============================================================================
    # MÉTODO 1: get_next_consecutivo()
    # ============================================================================

    def get_next_consecutivo(self):
        """
        Obtener el siguiente número consecutivo para Equipos de Cómputo.
        
        Busca el último consecutivo en la columna 1 de "Equipos de Cómputo"
        y retorna el siguiente número.
        
        Returns:
            int: Siguiente consecutivo (ej: si el último es 5, retorna 6)
        """
        if not self.excel_path or not HAS_OPENPYXL:
            return 1
        
        try:
            wb = load_workbook(self.excel_path, read_only=True)
            
            # Verificar que la hoja existe
            if "Equipos de Cómputo" not in wb.sheetnames:
                wb.close()
                return 1
            
            ws = wb["Equipos de Cómputo"]
            
            # Buscar el ÚLTIMO consecutivo en columna 1
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
            
            # Retornar siguiente consecutivo
            return last_consecutive + 1
            
        except Exception as e:
            print(f"❌ Error obteniendo consecutivo: {e}")
            return 1


    # ============================================================================
    # MÉTODO 2: get_next_codigo()
    # ============================================================================

    def get_next_codigo(self):
        """
        Obtener el siguiente código para Equipos de Cómputo.
        
        Usa get_next_consecutivo() y retorna el código formateado.
        
        Returns:
            str: Código formateado (ej: "EQC-0001", "EQC-0142")
        """
        try:
            next_consecutive = self.get_next_consecutivo()
            return f"EQC-{next_consecutive:04d}"
        except:
            return "EQC-0001"
    
    def show_form_in_container(self, form_creator):
        """Limpiar contenedor y mostrar formulario (reemplaza todas las _directo)."""
        for widget in self.main_container.winfo_children():
            widget.destroy()
        form_creator(self.main_container)
    
    def create_header(self):
        """Crear encabezado con logo en círculo blanco."""
        header_frame = ctk.CTkFrame(self.root, fg_color=COLOR_VERDE_HOSPITAL, corner_radius=0)
        header_frame.pack(fill="x", padx=0, pady=0)
        
        # Frame interno
        content_frame = ctk.CTkFrame(header_frame, fg_color="transparent")
        content_frame.pack(pady=18)
        
        # Intentar cargar logo
        if HAS_PIL:
            logo_paths = [
                "logo_hospital.png",
                "logo.png", 
                "hospital_logo.png"
            ]
            
            for logo_path in logo_paths:
                if os.path.exists(logo_path):
                    try:
                        from PIL import Image, ImageDraw
                        
                        # 1. Cargar logo original
                        logo_original = Image.open(logo_path)
                        
                        # 2. Convertir a RGBA si no lo es
                        if logo_original.mode != 'RGBA':
                            logo_original = logo_original.convert('RGBA')
                        
                        # 3. Redimensionar logo
                        aspect_ratio = logo_original.width / logo_original.height
                        new_height = 65
                        new_width = int(new_height * aspect_ratio)
                        logo_resized = logo_original.resize((new_width, new_height), Image.Resampling.LANCZOS)
                        
                        # 4. Crear círculo blanco de fondo
                        circle_size = 90
                        background = Image.new('RGBA', (circle_size, circle_size), (0, 0, 0, 0))
                        
                        # Dibujar círculo blanco sólido
                        draw = ImageDraw.Draw(background)
                        draw.ellipse([0, 0, circle_size-1, circle_size-1], 
                                    fill=(255, 255, 255, 255))
                        
                        # 5. Centrar logo sobre círculo blanco
                        x_offset = (circle_size - new_width) // 2
                        y_offset = (circle_size - new_height) // 2
                        
                        # Pegar logo
                        background.paste(logo_resized, (x_offset, y_offset), logo_resized)
                        
                        # 6. Convertir a CTkImage
                        logo_ctk = ctk.CTkImage(
                            light_image=background, 
                            dark_image=background, 
                            size=(circle_size, circle_size)
                        )
                        
                        logo_label = ctk.CTkLabel(
                            content_frame,
                            image=logo_ctk,
                            text=""
                        )
                        logo_label.pack(side="left", padx=(0, 25))
                        print(f"✓ Logo cargado en círculo blanco: {logo_path}")
                        break
                        
                    except Exception as e:
                        print(f"✗ Error al cargar logo {logo_path}: {e}")
        
        # Texto del header
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
        
        # Status label
        self.status_label = ctk.CTkLabel(
            header_frame,
            text="",
            font=("Segoe UI", 10),
            text_color="white",
            fg_color="transparent"
        )
        self.status_label.place(relx=0.98, rely=0.5, anchor="e")
    
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
            self.current_row = self.current_row-1
            
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
            self.current_row = self.current_row-1
            
            # Actualizar status
            self.status_label.configure(text=f"✅ {default_file} cargado")
            
            # Mostrar pestañas directamente
            self.show_manual_form_in_container()
            
            print(f"✅ Excel cargado automáticamente: {default_file}")
            print(f"✅ Siguiente fila disponible: {self.current_row}")
        else:
            # No hay archivo, mostrar mensaje en contenedor
            self.show_no_file_message()
    

    
    def show_form_directo(self, tipo):
        """Mostrar formulario directamente desde el menú."""
        if not self.excel_path:
            messagebox.showwarning("Advertencia", "Primero debes cargar un archivo Excel.\n\nVe a: Archivo → Cargar Excel")
            return
        
        # Mostrar formulario correspondiente usando la función consolidada
        if tipo == "Equipos de Cómputo":
            self.show_manual_form_in_container()
        elif tipo == "Impresoras":
            self.show_form_in_container(self.create_impresoras_form)
        elif tipo == "Periféricos":
            self.show_form_in_container(self.create_perifericos_form)
        elif tipo == "Red":
            self.show_form_in_container(self.create_red_form)
        elif tipo == "Mantenimiento":
            self.show_form_in_container(self.create_mantenimientos_form)
        elif tipo == "Dados de Baja":
            self.show_form_in_container(self.create_baja_form)
        elif tipo == "Comparar Hoja de Vida":
            self.show_form_in_container(self.create_hoja_vida_compare_form)

    # ------------------------------------------------------------------
    # FORMATOS OFICIALES (plantillas en carpeta "plantillas/")
    # ------------------------------------------------------------------
    def _abrir_archivo(self, ruta):
        try:
            if platform.system() == "Windows":
                os.startfile(str(ruta))
            elif platform.system() == "Darwin":
                subprocess.Popen(["open", str(ruta)])
            else:
                subprocess.Popen(["xdg-open", str(ruta)])
        except Exception:
            pass

    def _pedir_codigo_equipo(self, titulo):
        """Pide el código EQC-xxxx; propone el del equipo cargado en el formulario si existe."""
        from tkinter import simpledialog
        if not self.excel_path or not os.path.exists(self.excel_path):
            messagebox.showwarning(titulo, "Primero carga o crea el archivo Excel del inventario.")
            return ""
        sugerido = getattr(self, "equipo_update_code", "") or ""
        codigo = simpledialog.askstring(titulo, "Código del equipo (ej: EQC-0004):",
                                        initialvalue=sugerido, parent=self)
        return (codigo or "").strip().upper()

    def generar_hoja_vida_oficial(self):
        """Genera la hoja de vida (.docx) con la plantilla oficial."""
        import generador_formatos as gf
        codigo = self._pedir_codigo_equipo("Generar Hoja de Vida")
        if not codigo:
            return
        try:
            salida = gf.generar_hoja_vida(self.excel_path, codigo,
                                          Path(self.excel_path).parent / "formatos_generados")
        except Exception as e:
            messagebox.showerror("Hoja de Vida", f"No se pudo generar la hoja de vida:\n\n{e}")
            return
        messagebox.showinfo("Hoja de Vida", f"Hoja de vida generada:\n{salida}")
        self._abrir_archivo(salida)

    def abrir_dialogo_mantenimiento_pdf(self):
        """Diálogo para completar el reporte de mantenimiento preventivo (PDF oficial)."""
        import generador_formatos as gf
        codigo = self._pedir_codigo_equipo("Reporte de Mantenimiento")
        if not codigo:
            return
        try:
            gf.leer_equipo(self.excel_path, codigo)
        except Exception as e:
            messagebox.showerror("Reporte de Mantenimiento", str(e))
            return

        win = ctk.CTkToplevel(self)
        win.title(f"Reporte de Mantenimiento Preventivo - {codigo}")
        win.geometry("640x700")
        win.transient(self)
        frame = ctk.CTkScrollableFrame(win)
        frame.pack(fill="both", expand=True, padx=10, pady=10)

        def campo(etiqueta, valor=""):
            fila = ctk.CTkFrame(frame, fg_color="transparent")
            fila.pack(fill="x", pady=2)
            ctk.CTkLabel(fila, text=etiqueta, width=170, anchor="w").pack(side="left")
            e = ctk.CTkEntry(fila)
            e.pack(side="left", fill="x", expand=True)
            e.insert(0, valor)
            return e

        hoy = datetime.now().strftime("%d/%m/%Y")
        e_orden = campo("No. Orden")
        e_prog = campo("Fecha programada (dd/mm/aaaa)", hoy)
        e_ejec = campo("Fecha ejecución (dd/mm/aaaa)", hoy)
        e_ini = campo("Hora inicio (HH:MM)", "08:00")
        e_fin = campo("Hora finalización (HH:MM)", "08:30")

        ctk.CTkLabel(frame, text="Checklist (por defecto: Cumple)",
                     font=ctk.CTkFont(weight="bold")).pack(anchor="w", pady=(10, 2))
        actividades = [
            "1. Condiciones ambientales", "2. Conexiones eléctricas y cables",
            "3. Limpieza externa integral", "4. Limpieza interna del CPU",
            "5. Componentes completos", "6. Funcionamiento del sistema operativo",
            "7. Actualizaciones SO y software", "8. Estado del disco duro",
            "9. Antivirus institucional", "10. Conectividad de red e impresoras",
        ]
        opciones = {"Cumple": "cumple", "No cumple": "no_cumple", "N/A": "na"}
        combos = []
        for act in actividades:
            fila = ctk.CTkFrame(frame, fg_color="transparent")
            fila.pack(fill="x", pady=1)
            ctk.CTkLabel(fila, text=act, width=300, anchor="w").pack(side="left")
            cb = ctk.CTkComboBox(fila, values=list(opciones.keys()), width=130, state="readonly")
            cb.set("Cumple")
            cb.pack(side="left")
            combos.append(cb)

        ctk.CTkLabel(frame, text="Materiales / repuestos (opcional)",
                     font=ctk.CTkFont(weight="bold")).pack(anchor="w", pady=(10, 2))
        e_mat = campo("Material")
        e_desc = campo("Descripción")
        e_cant = campo("Cantidad")
        ctk.CTkLabel(frame, text="Observaciones y hallazgos").pack(anchor="w", pady=(10, 0))
        t_obs = ctk.CTkTextbox(frame, height=70)
        t_obs.pack(fill="x")
        ctk.CTkLabel(frame, text="Seguimiento y acciones requeridas").pack(anchor="w", pady=(8, 0))
        t_seg = ctk.CTkTextbox(frame, height=70)
        t_seg.pack(fill="x")

        def generar():
            materiales = []
            if e_mat.get().strip():
                materiales.append((e_mat.get().strip(), e_desc.get().strip(), e_cant.get().strip()))
            try:
                salida = gf.generar_mantenimiento_pdf(
                    self.excel_path, codigo, Path(self.excel_path).parent / "formatos_generados",
                    numero_orden=e_orden.get().strip(),
                    fecha_programada=e_prog.get().strip(),
                    fecha_ejecucion=e_ejec.get().strip(),
                    hora_inicio=e_ini.get().strip(), hora_fin=e_fin.get().strip(),
                    checklist=[opciones[c.get()] for c in combos],
                    materiales=materiales,
                    observaciones=t_obs.get("1.0", "end").strip(),
                    seguimiento=t_seg.get("1.0", "end").strip(),
                )
            except Exception as e:
                messagebox.showerror("Reporte de Mantenimiento", f"No se pudo generar el PDF:\n\n{e}", parent=win)
                return
            messagebox.showinfo("Reporte de Mantenimiento", f"PDF generado:\n{salida}", parent=win)
            self._abrir_archivo(salida)
            win.destroy()

        ctk.CTkButton(win, text="📄 Generar PDF", command=generar, height=38).pack(pady=8)

    def create_hoja_vida_compare_form(self, parent):
        """Crear sección para comparar hoja de vida (.docx) contra registro en Excel."""
        frame = ctk.CTkScrollableFrame(
            parent,
            fg_color=COLOR_FONDO,
            corner_radius=0
        )
        frame.pack(fill="both", expand=True)

        title_frame = ctk.CTkFrame(frame, fg_color=COLOR_VERDE_HOSPITAL, corner_radius=12)
        title_frame.pack(fill="x", padx=20, pady=(20, 12))

        ctk.CTkLabel(
            title_frame,
            text="📄 COMPARADOR HOJA DE VIDA VS EXCEL",
            font=("Segoe UI", 17, "bold"),
            text_color="white"
        ).pack(pady=14)

        help_text = (
            "1) Selecciona la hoja de vida (.docx).\n"
            "2) El sistema detecta automáticamente el código EQC-xxxx desde el documento.\n"
            "3) Presiona COMPARAR para completar vacíos en Excel y cargar el equipo en inicio."
        )
        ctk.CTkLabel(
            frame,
            text=help_text,
            justify="left",
            font=("Segoe UI", 12)
        ).pack(anchor="w", padx=25, pady=(5, 12))

        top_controls = ctk.CTkFrame(frame, fg_color="transparent")
        top_controls.pack(fill="x", padx=20)

        self.hoja_vida_docx_path = getattr(self, "hoja_vida_docx_path", "")

        def seleccionar_docx():
            filename = filedialog.askopenfilename(
                title="Seleccionar Hoja de Vida del Equipo",
                filetypes=[("Word", "*.docx"), ("Todos los archivos", "*.*")]
            )
            if filename:
                self.hoja_vida_docx_path = filename
                self.hoja_vida_selected_label.configure(text=f"📄 {os.path.basename(filename)}")

        btn_select = ctk.CTkButton(
            top_controls,
            text="📁 Seleccionar Hoja de Vida (.docx)",
            command=seleccionar_docx,
            font=("Segoe UI", 12, "bold"),
            fg_color=COLOR_VERDE_HOSPITAL,
            hover_color="#1F5039",
            width=320,
            height=42
        )
        btn_select.pack(side="left", padx=(0, 12), pady=(0, 10))

        btn_preview_doc = ctk.CTkButton(
            top_controls,
            text="👁 Visualizar datos útiles",
            command=self.on_preview_docx_data,
            font=("Segoe UI", 12, "bold"),
            fg_color=COLOR_VERDE_HOSPITAL,
            hover_color="#1F5039",
            width=230,
            height=42
        )
        btn_preview_doc.pack(side="left", padx=(0, 12), pady=(0, 10))

        self.hoja_vida_selected_label = ctk.CTkLabel(
            top_controls,
            text=(f"📄 {os.path.basename(self.hoja_vida_docx_path)}" if self.hoja_vida_docx_path else "📄 Ningún archivo seleccionado"),
            font=("Segoe UI", 12)
        )
        self.hoja_vida_selected_label.pack(side="left", pady=(0, 10))

        btn_compare = ctk.CTkButton(
            frame,
            text="🔎 COMPARAR CON EXCEL",
            command=self.run_hoja_vida_excel_comparison,
            font=("Segoe UI", 13, "bold"),
            fg_color=COLOR_VERDE_HOSPITAL,
            hover_color="#1F5039",
            height=44,
            width=280
        )
        btn_compare.pack(anchor="w", padx=20, pady=(6, 14))

        self.compare_result_text = ctk.CTkTextbox(
            frame,
            width=1300,
            height=520,
            font=("Consolas", 11)
        )
        self.compare_result_text.pack(fill="both", expand=True, padx=20, pady=(0, 16))

        self.compare_result_text.insert(
            "end",
            "Resultado de comparación aparecerá aquí...\n"
            "Tip: el sistema detecta automáticamente el código EQC-xxxx desde la hoja de vida."
        )
        self.compare_result_text.configure(state="disabled")

    def set_compare_result(self, content):
        """Actualizar caja de resultados del comparador."""
        if not hasattr(self, "compare_result_text"):
            return
        self.compare_result_text.configure(state="normal")
        self.compare_result_text.delete("1.0", "end")
        self.compare_result_text.insert("end", content)
        self.compare_result_text.configure(state="disabled")

    def get_excel_equipo_by_codigo(self, codigo):
        """Obtener campos principales de un equipo desde Excel por código."""
        if not self.excel_path or not HAS_OPENPYXL:
            return None

        column_map = {
            2: "codigo",
            3: "nombre_equipo",
            4: "tipo_equipo",
            5: "area_servicio",
            6: "ubicacion_especifica",
            7: "responsable_custodio",
            39: "marca",
            40: "modelo",
            41: "serial",
            42: "sistema_operativo",
            44: "procesador",
            45: "ram_gb",
            46: "disco1_capacidad",
            64: "direccion_ip",
            65: "mac_address",
        }

        wb = None
        try:
            wb = load_workbook(self.excel_path, read_only=True)
            ws = wb["Equipos de Cómputo"]

            target_row = None
            for row in range(2, 5000):
                value = ws.cell(row=row, column=2).value
                if value and str(value).strip().upper() == codigo:
                    target_row = row
                    break

            if not target_row:
                return None

            data = {}
            for col, field in column_map.items():
                value = ws.cell(row=target_row, column=col).value
                data[field] = "" if value is None else str(value).strip()

            data["__row"] = target_row

            return data
        finally:
            if wb:
                wb.close()

    def fill_missing_excel_fields_from_doc(self, codigo, doc_data):
        """Completar en Excel los campos vacíos usando datos extraídos de la hoja de vida."""
        if not self.excel_path or not HAS_OPENPYXL:
            return 0, []

        field_to_column = {
            "nombre_equipo": 3,
            "tipo_equipo": 4,
            "area_servicio": 5,
            "ubicacion_especifica": 6,
            "responsable_custodio": 7,
            "marca": 39,
            "modelo": 40,
            "serial": 41,
            "sistema_operativo": 42,
            "procesador": 44,
            "ram_gb": 45,
            "disco1_capacidad": 46,
            "direccion_ip": 64,
            "mac_address": 65,
        }

        wb = None
        updated_fields = []
        try:
            wb = load_workbook(self.excel_path)
            ws = wb["Equipos de Cómputo"]

            target_row = None
            for row in range(2, 5000):
                value = ws.cell(row=row, column=2).value
                if value and str(value).strip().upper() == codigo:
                    target_row = row
                    break

            if not target_row:
                return 0, []

            for field_name, col in field_to_column.items():
                current_value = ws.cell(row=target_row, column=col).value
                current_raw = str(current_value or "").strip()
                doc_raw = str(doc_data.get(field_name, "") or "").strip()

                if not current_raw and doc_raw:
                    ws.cell(row=target_row, column=col, value=doc_raw)
                    updated_fields.append(field_name)

            if updated_fields:
                wb.save(self.excel_path)

            return len(updated_fields), updated_fields
        finally:
            if wb:
                wb.close()

    def create_new_equipo_from_doc_data(self, doc_data, forced_codigo=None):
        """Crear nuevo equipo en Excel usando datos de hoja de vida y dejar faltantes para completar."""
        if not self.excel_path or not HAS_OPENPYXL:
            return None, 0

        wb = None
        try:
            wb = load_workbook(self.excel_path)
            ws = wb["Equipos de Cómputo"]

            last_consecutive = 0
            first_empty_row = 2
            for row in range(2, 5000):
                value = ws.cell(row=row, column=1).value
                if value is not None:
                    try:
                        consecutive = int(value)
                        if consecutive > last_consecutive:
                            last_consecutive = consecutive
                    except Exception:
                        pass
                else:
                    first_empty_row = row
                    break

            next_consecutive = last_consecutive + 1
            codigo = (forced_codigo or f"EQC-{next_consecutive:04d}").strip().upper()

            # Si el código forzado ya existe, usar código automático para evitar duplicado
            for row in range(2, 5000):
                current_code = ws.cell(row=row, column=2).value
                if current_code and str(current_code).strip().upper() == codigo:
                    codigo = f"EQC-{next_consecutive:04d}"
                    break

            ws.cell(row=first_empty_row, column=1, value=next_consecutive)
            ws.cell(row=first_empty_row, column=2, value=codigo)

            field_to_column = {
                "nombre_equipo": 3,
                "tipo_equipo": 4,
                "area_servicio": 5,
                "ubicacion_especifica": 6,
                "responsable_custodio": 7,
                "marca": 39,
                "modelo": 40,
                "serial": 41,
                "sistema_operativo": 42,
                "procesador": 44,
                "ram_gb": 45,
                "disco1_capacidad": 46,
                "direccion_ip": 64,
                "mac_address": 65,
            }

            filled_count = 0
            for field_name, col in field_to_column.items():
                value = str(doc_data.get(field_name, "") or "").strip()
                if value:
                    ws.cell(row=first_empty_row, column=col, value=value)
                    filled_count += 1

            wb.save(self.excel_path)
            return codigo, filled_count
        finally:
            if wb:
                wb.close()

    def get_doc_fill_field_meta(self):
        """Metadatos de campos que se llenan desde hoja de vida."""
        return [
            ("nombre_equipo", "Nombre Equipo"),
            ("tipo_equipo", "Tipo de Equipo"),
            ("area_servicio", "Área / Servicio"),
            ("ubicacion_especifica", "Ubicación Específica"),
            ("responsable_custodio", "Responsable / Custodio"),
            ("marca", "Marca"),
            ("modelo", "Modelo"),
            ("serial", "Serial"),
            ("sistema_operativo", "Sistema Operativo"),
            ("procesador", "Procesador"),
            ("ram_gb", "RAM (GB)"),
            ("disco1_capacidad", "Disco 1 Capacidad"),
            ("direccion_ip", "Dirección IP"),
            ("mac_address", "MAC Address"),
        ]

    def build_fill_preview_rows_existing(self, excel_data, doc_data):
        """Construir filas a llenar para un equipo existente (solo vacíos)."""
        rows = []
        for field_key, field_label in self.get_doc_fill_field_meta():
            excel_value = str(excel_data.get(field_key, "") or "").strip()
            doc_value = str(doc_data.get(field_key, "") or "").strip()
            if (not excel_value) and doc_value:
                rows.append((field_label, "(vacío)", doc_value))
        return rows

    def build_fill_preview_rows_new(self, doc_data):
        """Construir filas a llenar para un equipo nuevo."""
        rows = []
        for field_key, field_label in self.get_doc_fill_field_meta():
            doc_value = str(doc_data.get(field_key, "") or "").strip()
            if doc_value:
                rows.append((field_label, "(nuevo)", doc_value))
        return rows

    def show_fill_preview_window(self, title, rows):
        """Mostrar mini excel con campos que se llenarán y confirmar acción."""
        if not rows:
            return messagebox.askyesno(
                title,
                "No se detectaron campos nuevos para llenar en Excel.\n\n"
                "¿Deseas continuar de todos modos?"
            )

        preview = ctk.CTkToplevel(self.root)
        preview.title(title)
        preview.geometry("900x430")
        preview.transient(self.root)
        preview.grab_set()

        preview.update_idletasks()
        x = (preview.winfo_screenwidth() // 2) - 450
        y = (preview.winfo_screenheight() // 2) - 215
        preview.geometry(f"900x430+{x}+{y}")

        ctk.CTkLabel(
            preview,
            text="🧾 Vista previa (mini excel) de campos a llenar",
            font=("Segoe UI", 15, "bold")
        ).pack(pady=(14, 8))

        container = ctk.CTkFrame(preview, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=16, pady=8)

        cols = ("campo", "excel", "nuevo")
        tree = ttk.Treeview(container, columns=cols, show="headings", height=12)
        tree.heading("campo", text="Campo")
        tree.heading("excel", text="Excel actual")
        tree.heading("nuevo", text="Valor a llenar")
        tree.column("campo", width=250, anchor="w")
        tree.column("excel", width=180, anchor="w")
        tree.column("nuevo", width=430, anchor="w")

        scrollbar = ttk.Scrollbar(container, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        for row in rows:
            tree.insert("", "end", values=row)

        decision = {"ok": False}

        btns = ctk.CTkFrame(preview, fg_color="transparent")
        btns.pack(fill="x", padx=16, pady=(6, 14))

        def on_continue():
            decision["ok"] = True
            preview.destroy()

        def on_cancel():
            decision["ok"] = False
            preview.destroy()

        ctk.CTkButton(
            btns,
            text="✅ Continuar y llenar",
            command=on_continue,
            fg_color=COLOR_VERDE_HOSPITAL,
            hover_color="#1F5039",
            width=200
        ).pack(side="left")

        ctk.CTkButton(
            btns,
            text="❌ Cancelar",
            command=on_cancel,
            fg_color="#6C757D",
            hover_color="#5A6268",
            width=140
        ).pack(side="right")

        preview.wait_window()
        return decision["ok"]

    def on_preview_docx_data(self):
        """Mostrar vista previa manual de datos útiles dentro del panel actual."""
        docx_path = (getattr(self, "hoja_vida_docx_path", "") or "").strip()
        if not docx_path:
            messagebox.showwarning("Vista previa", "Primero selecciona una hoja de vida (.docx).")
            return

        if not os.path.exists(docx_path):
            messagebox.showerror("Vista previa", "El archivo seleccionado ya no existe.")
            return

        try:
            full_text, lines = read_docx_text(docx_path)
            doc_data = self.extract_hoja_vida_data(full_text, lines)

            detected_codigo = ""
            match_codigo = re.search(r"\bEQC-\d{4,5}\b", full_text, flags=re.IGNORECASE)
            if match_codigo:
                detected_codigo = match_codigo.group(0).upper()

            codigo = detected_codigo or (doc_data.get("codigo", "") or "").strip().upper()

            report = []
            report.append("=" * 92)
            report.append("VISTA PREVIA DESDE HOJA DE VIDA (.DOCX)")
            report.append("=" * 92)
            report.append(f"Archivo: {os.path.basename(docx_path)}")
            report.append(f"Código detectado: {codigo if codigo else '(sin código)'}")
            report.append("")
            report.append("DATOS DETECTADOS")

            for field_key, field_label in self.get_doc_fill_field_meta():
                value = str(doc_data.get(field_key, "") or "").strip()
                report.append(f"- {field_label}: {value if value else '(vacío/no detectado)'}")

            report.append("")
            report.append("CAMPOS QUE SE LLENARÍAN EN EXCEL")

            if codigo:
                excel_data = self.get_excel_equipo_by_codigo(codigo)
                if excel_data:
                    rows = self.build_fill_preview_rows_existing(excel_data, doc_data)
                    report.append("Modo: actualizar registro existente")
                else:
                    rows = self.build_fill_preview_rows_new(doc_data)
                    report.append("Modo: crear nuevo registro (código detectado no existe)")
            else:
                rows = self.build_fill_preview_rows_new(doc_data)
                report.append("Modo: crear nuevo registro con código automático")

            if rows:
                for field_label, excel_current, new_value in rows:
                    report.append(f"- {field_label}: {excel_current} -> {new_value}")
            else:
                report.append("- No hay campos nuevos para llenar.")

            self.set_compare_result("\n".join(report))

        except Exception as exc:
            messagebox.showwarning(
                "Vista previa",
                f"No se pudo generar vista previa:\n{exc}"
            )

    def show_preview_only_window(self, title, rows, info_text=""):
        """Mostrar mini excel solo informativo al cargar hoja de vida."""
        preview = ctk.CTkToplevel(self.root)
        preview.title(title)
        preview.geometry("900x430")
        preview.transient(self.root)
        preview.grab_set()

        preview.update_idletasks()
        x = (preview.winfo_screenwidth() // 2) - 450
        y = (preview.winfo_screenheight() // 2) - 215
        preview.geometry(f"900x430+{x}+{y}")

        ctk.CTkLabel(
            preview,
            text="🧾 Vista previa automática (mini excel)",
            font=("Segoe UI", 15, "bold")
        ).pack(pady=(14, 8))

        if info_text:
            ctk.CTkLabel(
                preview,
                text=info_text,
                font=("Segoe UI", 11)
            ).pack(pady=(0, 8))

        container = ctk.CTkFrame(preview, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=16, pady=8)

        cols = ("campo", "excel", "nuevo")
        tree = ttk.Treeview(container, columns=cols, show="headings", height=12)
        tree.heading("campo", text="Campo")
        tree.heading("excel", text="Excel actual")
        tree.heading("nuevo", text="Valor detectado")
        tree.column("campo", width=250, anchor="w")
        tree.column("excel", width=180, anchor="w")
        tree.column("nuevo", width=430, anchor="w")

        scrollbar = ttk.Scrollbar(container, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        if rows:
            for row in rows:
                tree.insert("", "end", values=row)
        else:
            tree.insert("", "end", values=("(sin cambios detectados)", "", ""))

        ctk.CTkButton(
            preview,
            text="Cerrar",
            command=preview.destroy,
            fg_color=COLOR_VERDE_HOSPITAL,
            hover_color="#1F5039",
            width=120
        ).pack(pady=(6, 14))

        preview.wait_window()

    def preview_docx_fill_plan(self, docx_path):
        """Mostrar mini excel automáticamente al cargar hoja de vida."""
        try:
            full_text, lines = read_docx_text(docx_path)
            doc_data = self.extract_hoja_vida_data(full_text, lines)

            detected_codigo = ""
            match_codigo = re.search(r"\bEQC-\d{4,5}\b", full_text, flags=re.IGNORECASE)
            if match_codigo:
                detected_codigo = match_codigo.group(0).upper()

            codigo = detected_codigo or (doc_data.get("codigo", "") or "").strip().upper()

            if codigo:
                excel_data = self.get_excel_equipo_by_codigo(codigo)
                if excel_data:
                    rows = self.build_fill_preview_rows_existing(excel_data, doc_data)
                    info = f"Código detectado: {codigo} (registro existente)"
                    self.show_preview_only_window("Mini Excel - Campos a completar", rows, info)
                else:
                    rows = self.build_fill_preview_rows_new(doc_data)
                    info = f"Código detectado: {codigo} (se creará nuevo registro)"
                    self.show_preview_only_window("Mini Excel - Campos detectados", rows, info)
            else:
                rows = self.build_fill_preview_rows_new(doc_data)
                info = "Sin código detectado (se usará código automático al crear)."
                self.show_preview_only_window("Mini Excel - Campos detectados", rows, info)

        except Exception as exc:
            messagebox.showwarning(
                "Vista previa",
                f"No se pudo generar vista previa automática:\n{exc}"
            )

    def load_equipo_into_form_by_codigo(self, codigo, show_ready_message=True):
        """Cargar equipo por código en el formulario principal en modo actualización."""
        if not self.excel_path:
            return False, "No hay Excel cargado"

        wb = None
        try:
            wb = load_workbook(self.excel_path)
            ws = wb["Equipos de Cómputo"]

            target_row = None
            for row in range(2, 5000):
                cell_value = ws.cell(row=row, column=2).value
                if cell_value and str(cell_value).strip().upper() == codigo:
                    target_row = row
                    break

            if not target_row:
                return False, f"No se encontró el código {codigo}"

            naranja_map = {
                4: 'tipo_equipo', 5: 'area_servicio', 6: 'ubicacion_especifica',
                7: 'responsable_custodio', 8: 'macroproceso', 9: 'proceso',
                10: 'subproceso', 11: 'uso_sihos', 12: 'uso_office_basico',
                13: 'software_especializado', 14: 'descripcion_software',
                15: 'funcion_principal', 34: 'horario_uso', 35: 'estado_operativo',
                36: 'observaciones_tecnicas', 37: 'periodicidad_mtto', 38: 'responsable_mtto'
            }

            for i in range(1, 10):
                naranja_map[15 + i] = f'conf_{i}'
            for i in range(1, 4):
                naranja_map[24 + i] = f'int_{i}'
            for i in range(1, 7):
                naranja_map[27 + i] = f'crit_{i}'

            verde_map = {
                39: 'marca', 40: 'modelo', 41: 'serial',
                42: 'sistema_operativo', 43: 'arquitectura_so', 44: 'procesador', 45: 'ram_gb',
                46: 'disco1_capacidad', 47: 'disco1_tipo', 48: 'disco1_serial',
                49: 'disco1_marca', 50: 'disco1_modelo',
                51: 'disco2_capacidad', 52: 'disco2_tipo', 53: 'disco2_serial',
                54: 'disco2_marca', 55: 'disco2_modelo',
                56: 'uso_navegador_web', 57: 'version_office', 58: 'licencia_office',
                59: 'uso_teams', 60: 'uso_outlook',
                61: 'licencia_windows', 62: 'key_windows', 63: 'estado_licencia_windows',
                64: 'direccion_ip', 65: 'mac_address', 66: 'tipo_conexion',
                67: 'navegador_predeterminado', 68: 'unidades_red_mapeadas',
                69: 'antivirus_instalado', 70: 'ultima_act_windows', 71: 'windows_update_activo',
                79: 'velocidad_red',
                80: 'tpm_status', 81: 'bitlocker_status', 82: 'cpu_nucleos',
                83: 'cpu_uso_actual', 84: 'ram_uso_actual'
            }

            azul_map = {
                72: 'switch_puerto', 73: 'vlan_asignada', 74: 'id_anydesk',
                75: 'otro_acceso_remoto', 76: 'estado_antivirus',
                77: 'cifrado_disco', 78: 'tipo_usuario_local'
            }

            self.equipment_data = {}
            self.verde_data = {}
            self.azul_data = {}

            for col, field_name in naranja_map.items():
                value = ws.cell(row=target_row, column=col).value or ''
                self.equipment_data[field_name] = value

            for col, field_name in verde_map.items():
                value = ws.cell(row=target_row, column=col).value or ''
                self.verde_data[field_name] = value

            for col, field_name in azul_map.items():
                value = ws.cell(row=target_row, column=col).value or ''
                self.azul_data[field_name] = value

            for field_name, value in self.equipment_data.items():
                if field_name in self.manual_widgets:
                    try:
                        widget = self.manual_widgets[field_name]
                        if hasattr(widget, 'winfo_exists') and widget.winfo_exists():
                            if isinstance(widget, ctk.CTkEntry):
                                widget.delete(0, "end")
                                widget.insert(0, value)
                            elif isinstance(widget, ctk.CTkComboBox):
                                widget.set(value)
                        elif isinstance(widget, tk.StringVar):
                            widget.set(value)
                    except Exception:
                        pass

            self.equipo_update_code = codigo
            self.equipo_update_row = target_row

            if hasattr(self, 'equipo_form_frame'):
                try:
                    if hasattr(self.equipo_form_frame, 'winfo_exists') and self.equipo_form_frame.winfo_exists():
                        self.equipo_form_frame.configure(label_text=f"🔄 ACTUALIZANDO - Código: {codigo}")
                except Exception:
                    pass

            if hasattr(self, 'form_title_label'):
                try:
                    if hasattr(self.form_title_label, 'winfo_exists') and self.form_title_label.winfo_exists():
                        self.form_title_label.configure(text=f"🔄 ACTUALIZANDO EQUIPO - Código: {codigo}")
                except Exception:
                    pass

            if hasattr(self, 'btn_save_equipo'):
                try:
                    if hasattr(self.btn_save_equipo, 'winfo_exists') and self.btn_save_equipo.winfo_exists():
                        self.btn_save_equipo.configure(text="🔄 ACTUALIZAR EQUIPO")
                except Exception:
                    pass

            if show_ready_message:
                messagebox.showinfo("Listo", f"✅ Datos cargados de {codigo}\n\nModifica los campos y presiona ACTUALIZAR EQUIPO.")

            return True, "ok"
        except Exception as exc:
            return False, str(exc)
        finally:
            if wb:
                wb.close()

    def extract_hoja_vida_data(self, full_text, lines):
        """Extraer valores de hoja de vida usando patrones comunes."""
        cpu_marca, cpu_modelo, cpu_serial = extract_cpu_component_fields(lines)

        return {
            "codigo": extract_value_from_doc(
                full_text,
                lines,
                [r"(?im)\bcodigo\s*(?:del\s*equipo)?\s*[:\-]\s*(EQC-\d{4,5})\b", r"(?im)\b(EQC-\d{4,5})\b"],
                ["codigo", "codigo equipo", "codigo del equipo"],
            ),
            "nombre_equipo": extract_value_from_doc(
                full_text,
                lines,
                [r"(?im)\bnombre\s*(?:del\s*equipo)?\s*[:\-]\s*([^\n\r]+)", r"(?im)\bhostname\s*[:\-]\s*([^\n\r]+)"],
                ["nombre equipo", "nombre del equipo", "hostname"],
            ),
            "tipo_equipo": extract_value_from_doc(
                full_text,
                lines,
                [r"(?im)\btipo\s*de\s*equipo\s*[:\-]\s*([^\n\r]+)"],
                ["tipo de equipo"],
            ),
            "area_servicio": extract_value_from_doc(
                full_text,
                lines,
                [
                    r"(?im)\b[aá]rea\s*asignada\s*(?::|\-|\s+)\s*([^\n\r]+)",
                    r"(?im)\b[aá]rea\s*(?:\/|y)?\s*servicio\s*(?::|\-|\s+)\s*([^\n\r]+)",
                    r"(?im)\bservicio\s*(?::|\-|\s+)\s*([^\n\r]+)",
                    r"(?im)\bdependencia\s*(?::|\-)\s*([^\n\r]+)",
                ],
                ["area / servicio", "area servicio", "servicio"],
            ) or extract_value_by_label(
                lines,
                ["Área Asignada", "Area Asignada", "Área / Servicio", "Area / Servicio", "Área Servicio", "Area Servicio", "Servicio", "Dependencia"],
            ),
            "ubicacion_especifica": extract_value_from_doc(
                full_text,
                lines,
                [
                    r"(?im)\bubicaci[oó]n\s*espec[ií]fica\s*(?::|\-)\s*([^\n\r]+)",
                    r"(?im)\bubicaci[oó]n\s*f[ií]sica\s*(?::|\-)\s*([^\n\r]+)",
                ],
                ["ubicacion", "ubicacion especifica"],
            ) or extract_value_by_label(
                lines,
                ["Ubicación Específica", "Ubicacion Especifica", "Ubicación", "Ubicacion", "Ubicación Física", "Ubicacion Fisica"],
            ),
            "responsable_custodio": extract_value_from_doc(
                full_text,
                lines,
                [
                    r"(?im)\bcoordinador\s*[aá]rea\s*(?::|\-|\s+)\s*([^\n\r]+)",
                    r"(?im)\b(?:responsable|custodio)\s*(?::|\-|\s+)\s*([^\n\r]+)",
                    r"(?im)\bresponsable\s*del\s*equipo\s*(?::|\-|\s+)\s*([^\n\r]+)",
                    r"(?im)\bfuncionario\s*responsable\s*(?::|\-|\s+)\s*([^\n\r]+)",
                ],
                ["responsable", "custodio", "responsable / custodio"],
            ) or extract_value_by_label(
                lines,
                ["Coordinador Área", "Coordinador Area", "Responsable", "Custodio", "Responsable / Custodio", "Responsable del Equipo", "Funcionario Responsable"],
            ),
            "marca": extract_value_from_doc(
                full_text,
                lines,
                [r"(?im)\bmarca\s*[:\-]\s*([^\n\r]+)"],
            ) if not cpu_marca else cpu_marca,
            "modelo": extract_value_from_doc(
                full_text,
                lines,
                [r"(?im)\bmodelo\s*[:\-]\s*([^\n\r]+)"],
            ) if not cpu_modelo else cpu_modelo,
            "serial": extract_value_from_doc(
                full_text,
                lines,
                [r"(?im)\bserial\s*(?:number|no\.?|nro\.?|#)?\s*[:\-]\s*([^\n\r]+)"],
            ) if not cpu_serial else cpu_serial,
            "sistema_operativo": extract_value_from_doc(
                full_text,
                lines,
                [r"(?im)\b(?:sistema\s*operativo|so)\s*[:\-]\s*([^\n\r]+)"],
                ["sistema operativo", "so"],
            ),
            "procesador": extract_value_from_doc(
                full_text,
                lines,
                [r"(?im)\b(?:procesador|cpu)\s*[:\-]\s*([^\n\r]+)"],
                ["procesador", "cpu"],
            ),
            "ram_gb": extract_value_from_doc(
                full_text,
                lines,
                [r"(?im)\bram\s*(?:instalada)?\s*[:\-]\s*([^\n\r]+)", r"(?im)\bmemoria\s*ram\s*[:\-]\s*([^\n\r]+)"],
                ["ram", "ram instalada", "memoria ram"],
            ),
            "disco1_capacidad": extract_value_from_doc(
                full_text,
                lines,
                [r"(?im)\b(?:disco\s*1\s*capacidad|capacidad\s*disco\s*1|almacenamiento)\s*[:\-]\s*([^\n\r]+)"],
                ["disco 1 capacidad", "capacidad disco 1", "almacenamiento"],
            ),
            "direccion_ip": extract_value_from_doc(
                full_text,
                lines,
                [r"(?im)\b(?:direccion\s*ip|ip)\s*[:\-]\s*([0-9]{1,3}(?:\.[0-9]{1,3}){3})"],
                ["direccion ip", "ip"],
            ),
            "mac_address": extract_value_from_doc(
                full_text,
                lines,
                [r"(?im)\b(?:mac\s*(?:address)?|direccion\s*mac)\s*[:\-]\s*([0-9A-Fa-f:\-]{12,17})"],
                ["mac", "mac address", "direccion mac"],
            ),
        }

    def values_match_for_field(self, field_name, excel_value, doc_value):
        """Comparar valores de Excel vs hoja de vida según tipo de campo."""
        excel_raw = str(excel_value or "").strip()
        doc_raw = str(doc_value or "").strip()

        if not doc_raw:
            return False
        if not excel_raw:
            return False

        if field_name in {"serial", "mac_address"}:
            return normalize_alnum(excel_raw) == normalize_alnum(doc_raw)

        if field_name in {"ram_gb", "disco1_capacidad"}:
            return extract_first_number(excel_raw) == extract_first_number(doc_raw)

        return normalize_text_for_compare(excel_raw) == normalize_text_for_compare(doc_raw)

    def run_hoja_vida_excel_comparison(self):
        """Ejecutar comparación entre hoja de vida (.docx) y fila del Excel."""
        docx_path = (getattr(self, "hoja_vida_docx_path", "") or "").strip()
        if not docx_path:
            messagebox.showwarning("Comparador", "Primero selecciona la hoja de vida (.docx).")
            return

        if not os.path.exists(docx_path):
            messagebox.showerror("Comparador", "El archivo seleccionado ya no existe.")
            return

        try:
            full_text, lines = read_docx_text(docx_path)
        except Exception as exc:
            messagebox.showerror("Comparador", f"No se pudo leer el archivo .docx:\n{exc}")
            return

        detected_codigo = ""
        match_codigo = re.search(r"\bEQC-\d{4,5}\b", full_text, flags=re.IGNORECASE)
        if match_codigo:
            detected_codigo = match_codigo.group(0).upper()

        doc_data = self.extract_hoja_vida_data(full_text, lines)
        codigo = detected_codigo or (doc_data.get("codigo", "") or "").strip().upper()

        if not codigo:
            create_without_code = messagebox.askyesno(
                "Comparador",
                "No se detectó código EQC-xxxx en la hoja de vida.\n\n"
                "¿Quieres crear el equipo con código automático y completar faltantes en inicio?"
            )
            if not create_without_code:
                return

            preview_rows = self.build_fill_preview_rows_new(doc_data)
            if not self.show_fill_preview_window("Campos a crear en Excel", preview_rows):
                return

            new_code, filled_count = self.create_new_equipo_from_doc_data(doc_data)
            if not new_code:
                messagebox.showerror("Comparador", "No fue posible crear el equipo en Excel.")
                return

            self.show_manual_form_in_container()
            loaded_ok, load_msg = self.load_equipo_into_form_by_codigo(new_code, show_ready_message=False)
            if not loaded_ok:
                messagebox.showwarning(
                    "Equipo creado",
                    f"Se creó el equipo {new_code} con {filled_count} campos iniciales,\n"
                    f"pero no se pudo cargar en el formulario:\n{load_msg}"
                )
                return

            messagebox.showinfo(
                "Equipo creado",
                f"✅ Equipo {new_code} creado desde hoja de vida.\n"
                f"Campos iniciales cargados: {filled_count}.\n\n"
                "Ya estás en inicio para completar los campos faltantes y guardar actualización."
            )
            return

        excel_data = self.get_excel_equipo_by_codigo(codigo)
        if not excel_data:
            create_with_detected_code = messagebox.askyesno(
                "Comparador",
                f"El código {codigo} no existe en Excel.\n\n"
                "¿Deseas crear el equipo con los datos detectados y luego completar faltantes?"
            )
            if not create_with_detected_code:
                return

            preview_rows = self.build_fill_preview_rows_new(doc_data)
            if not self.show_fill_preview_window("Campos a crear en Excel", preview_rows):
                return

            new_code, filled_count = self.create_new_equipo_from_doc_data(doc_data, forced_codigo=codigo)
            if not new_code:
                messagebox.showerror("Comparador", "No fue posible crear el equipo en Excel.")
                return

            self.show_manual_form_in_container()
            loaded_ok, load_msg = self.load_equipo_into_form_by_codigo(new_code, show_ready_message=False)
            if not loaded_ok:
                messagebox.showwarning(
                    "Equipo creado",
                    f"Se creó el equipo {new_code} con {filled_count} campos iniciales,\n"
                    f"pero no se pudo cargar en el formulario:\n{load_msg}"
                )
                return

            messagebox.showinfo(
                "Equipo creado",
                f"✅ Equipo {new_code} creado desde hoja de vida.\n"
                f"Campos iniciales cargados: {filled_count}.\n\n"
                "Ya estás en inicio para completar los campos faltantes y guardar actualización."
            )
            return

        if not doc_data.get("codigo"):
            doc_data["codigo"] = codigo

        fields_to_compare = [
            ("codigo", "Código"),
            ("nombre_equipo", "Nombre Equipo"),
            ("tipo_equipo", "Tipo de Equipo"),
            ("area_servicio", "Área / Servicio"),
            ("ubicacion_especifica", "Ubicación Específica"),
            ("responsable_custodio", "Responsable / Custodio"),
            ("marca", "Marca"),
            ("modelo", "Modelo"),
            ("serial", "Serial"),
            ("sistema_operativo", "Sistema Operativo"),
            ("procesador", "Procesador"),
            ("ram_gb", "RAM (GB)"),
            ("disco1_capacidad", "Disco 1 Capacidad"),
            ("direccion_ip", "Dirección IP"),
            ("mac_address", "MAC Address"),
        ]

        matches = 0
        differences = 0
        missing = 0
        report_lines = []

        report_lines.append("=" * 92)
        report_lines.append("COMPARACIÓN HOJA DE VIDA vs EXCEL")
        report_lines.append("=" * 92)
        report_lines.append(f"Archivo: {os.path.basename(docx_path)}")
        report_lines.append(f"Código comparado: {codigo}")
        report_lines.append("")

        for field_key, field_label in fields_to_compare:
            excel_value = (excel_data.get(field_key) or "").strip()
            doc_value = (doc_data.get(field_key) or "").strip()

            if not doc_value:
                status = "❓ NO ENCONTRADO EN HOJA DE VIDA"
                missing += 1
            elif self.values_match_for_field(field_key, excel_value, doc_value):
                status = "✅ COINCIDE"
                matches += 1
            else:
                status = "⚠️ DIFERENTE"
                differences += 1

            report_lines.append(f"{field_label}: {status}")
            report_lines.append(f"  Excel: {excel_value if excel_value else '(vacío)'}")
            report_lines.append(f"  Hoja : {doc_value if doc_value else '(no encontrado)'}")
            report_lines.append("-" * 92)

        report_lines.append("")
        report_lines.append("RESUMEN")
        report_lines.append(f"✅ Coinciden: {matches}")
        report_lines.append(f"⚠️ Diferentes: {differences}")
        report_lines.append(f"❓ No encontrados en hoja de vida: {missing}")

        preview_rows = self.build_fill_preview_rows_existing(excel_data, doc_data)
        if not self.show_fill_preview_window("Campos a completar en Excel", preview_rows):
            self.set_compare_result("\n".join(report_lines + ["", "Acción cancelada por usuario antes de llenar Excel."]))
            return

        updated_count, updated_fields = self.fill_missing_excel_fields_from_doc(codigo, doc_data)
        report_lines.append(f"🧩 Campos vacíos completados en Excel: {updated_count}")
        if updated_fields:
            report_lines.append("   " + ", ".join(updated_fields))

        self.set_compare_result("\n".join(report_lines))

        self.show_manual_form_in_container()
        loaded_ok, load_msg = self.load_equipo_into_form_by_codigo(codigo, show_ready_message=False)
        if loaded_ok:
            messagebox.showinfo(
                "Comparación finalizada",
                f"✅ Equipo {codigo} procesado.\n"
                f"Campos completados en Excel: {updated_count}\n\n"
                "Ya estás en inicio con el equipo cargado para revisar y actualizar."
            )
        else:
            messagebox.showwarning(
                "Comparación finalizada",
                f"Se completaron campos faltantes en Excel para {codigo},\n"
                f"pero no se pudo cargar el formulario de actualización:\n{load_msg}"
            )
    
    def show_classification_guide(self):
        """Mostrar guía de clasificación."""
        messagebox.showinfo(
            "Guía de Formulario",
            "📋 GUÍA DE CAMPOS\n\n"
            "🟧 NARANJAS: Campos manuales (obligatorios)\n"
            "   • Información operativa y de clasificación\n"
            "   • Debes completarlos manualmente\n\n"
            "🟩 VERDES: Campos automáticos\n"
            "   • Hardware, software, red\n"
            "   • Se detectan automáticamente\n\n"
            "🟦 AZULES: Campos mixtos\n"
            "   • Validación manual de datos\n"
            "   • Aparecen después de recopilación automática"
        )
    
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
            
            # Todos los códigos son de 4 dígitos
            return f"{prefix}-{next_consecutive:04d}"
            
        except Exception as e:
            print(f"❌ Error detectando código: {e}")
            import traceback
            traceback.print_exc()
            return f"{prefix}-001"
        
    def get_next_codigo(self):
        """
        Obtener siguiente código EQC para equipos de cómputo.
        Wrapper de detect_next_code() para compatibilidad.
        """
        return self.detect_next_code("Equipos de Cómputo", "EQC")
    
    def get_next_consecutivo(self):
        """
        Obtener siguiente número consecutivo para equipos de cómputo.
        Busca el último consecutivo en la columna 1 de "Equipos de Cómputo".
        """
        if not self.excel_path or not HAS_OPENPYXL:
            return 1
        
        try:
            wb = load_workbook(self.excel_path, read_only=True)
            
            # Verificar que la hoja existe
            if "Equipos de Cómputo" not in wb.sheetnames:
                wb.close()
                return 1
            
            ws = wb["Equipos de Cómputo"]
            
            # Buscar el ÚLTIMO consecutivo en columna 1
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
            
            # Siguiente consecutivo
            return last_consecutive + 1
            
        except Exception as e:
            print(f"❌ Error obteniendo consecutivo: {e}")
            return 1
    
    def detect_next_consecutive_mantenimiento(self):
        """Detectar siguiente consecutivo para mantenimientos."""
        next_row = self.get_next_available_row("Mantenimientos", check_column=1)
        return next_row - 1
    
    def detect_next_baja(self):
        """Detectar siguiente número de baja."""
        next_row = self.get_next_available_row("Equipos Dados de Baja", check_column=1, max_rows=200)
        return next_row - 1
    
    def create_form_field_centered(self, parent, label_text, field_name, field_type, 
                               options=None, tooltip_text=None):
        """
        Crear campo centrado.
        
        Args:
            parent: Frame padre donde se creará el campo
            label_text: Texto del label (izquierda)
            field_name: Nombre del campo (key en self.manual_widgets)
            field_type: "entry" o "combobox"
            options: Lista de opciones para combobox (opcional)
            tooltip_text: Texto del tooltip al pasar mouse (opcional)
        
        Returns:
            widget: El widget creado (Entry o ComboBox)
        
        Grid:
            - Columna 0 (Label): 60% peso, alineado izquierda
            - Columna 1 (Widget): 40% peso, ancho fijo 300px
            - Espacio entre columnas: 30px
        """
        # Frame principal - CENTRADO con padding lateral
        field_frame = ctk.CTkFrame(parent, fg_color="white", corner_radius=8)
        field_frame.pack(fill="x", padx=40, pady=6)
        
        # Frame interno - Grid layout con espacio
        inner_frame = ctk.CTkFrame(field_frame, fg_color="transparent")
        inner_frame.pack(fill="x", padx=20, pady=10)
        
        # Configurar grid (2 columnas con espacio)
        inner_frame.grid_columnconfigure(0, weight=6)  # Label: flexible
        inner_frame.grid_columnconfigure(1, weight=0, minsize=300)  # Widget: fijo 300px
        
        # ===== LABEL (COLUMNA 0) =====
        label = ctk.CTkLabel(
            inner_frame,
            text=label_text,
            font=("Segoe UI", 12, "bold"),
            anchor="w",  # Alineado a la izquierda
            text_color="#333333"
        )
        label.grid(row=0, column=0, sticky="w", padx=(0, 30))  # 30px espacio a la derecha
        
        # Tooltip si existe
        if tooltip_text:
            ToolTip(label, tooltip_text)
        
        # ===== WIDGET (COLUMNA 1) =====
        if field_type == "combobox":
            widget = ctk.CTkComboBox(
                inner_frame,
                values=options if options else [],
                height=35,
                width=300, 
                font=("Segoe UI", 11),
                dropdown_font=("Segoe UI", 10),
                border_color="#CCCCCC",
                button_color=COLOR_VERDE_HOSPITAL,
                button_hover_color="#1F5A32",
                corner_radius=8
            )
            widget.grid(row=0, column=1, sticky="e")  # Alineado a la derecha de su columna
            
        elif field_type == "entry":
            widget = ctk.CTkEntry(
                inner_frame,
                height=35,
                width=300,
                font=("Segoe UI", 11),
                border_color="#CCCCCC",
                fg_color="white",
                corner_radius=8
            )
            widget.grid(row=0, column=1, sticky="e")  # Alineado a la derecha de su columna
        
        # Guardar widget
        self.manual_widgets[field_name] = widget
        return widget
    
    def create_date_field_centered(self, parent, label_text, field_name, tooltip_text=None):
        """
        Crear campo de FECHA con calendario (DateEntry).
        Similar a create_form_field_centered pero con calendario.
        """
        # Frame principal - CENTRADO con padding lateral
        field_frame = ctk.CTkFrame(parent, fg_color="white", corner_radius=8)
        field_frame.pack(fill="x", padx=40, pady=6)
        
        # Frame interno - Grid layout
        inner_frame = ctk.CTkFrame(field_frame, fg_color="transparent")
        inner_frame.pack(fill="x", padx=20, pady=10)
        
        # Configurar grid (2 columnas)
        inner_frame.grid_columnconfigure(0, weight=6)  # Label: flexible
        inner_frame.grid_columnconfigure(1, weight=0, minsize=300)  # Widget: fijo 300px
        
        # ===== LABEL (COLUMNA 0) =====
        label = ctk.CTkLabel(
            inner_frame,
            text=label_text,
            font=("Segoe UI", 12, "bold"),
            anchor="w",
            text_color="#333333"
        )
        label.grid(row=0, column=0, sticky="w", padx=(0, 30))
        
        # Tooltip si existe
        if tooltip_text:
            ToolTip(label, tooltip_text)
        
        # ===== WIDGET FECHA (COLUMNA 1) =====
        # Usar DateEntry (calendario visual)
        widget = DateEntry(
            inner_frame,
            width=28,
            background=COLOR_VERDE_HOSPITAL,
            foreground='white',
            borderwidth=2,
            font=("Segoe UI", 11),
            date_pattern='yyyy-mm-dd',  # Formato ISO
            showweeknumbers=False,
            showothermonthdays=False,
            selectbackground=COLOR_VERDE_HOSPITAL,
            selectforeground='white',
            normalbackground='white',
            normalforeground='black',
            weekendbackground='#F0F0F0',
            weekendforeground='black',
            othermonthbackground='white',
            othermonthweforeground='gray',
            othermonthwebackground='#F0F0F0'
        )
        widget.grid(row=0, column=1, sticky="e")
        
        # Guardar widget
        self.manual_widgets[field_name] = widget
        return widget
        
    def get_date_value_safe(self, widget):
        """
        Obtener valor de fecha de forma segura desde DateEntry o Entry.
        
        Args:
            widget: Widget DateEntry o Entry con fecha
        
        Returns:
            str: Fecha en formato YYYY-MM-DD o cadena vacía
        """
        try:
            # DateEntry de tkcalendar - tiene get_date()
            if hasattr(widget, 'get_date'):
                date_obj = widget.get_date()
                return date_obj.strftime('%Y-%m-%d')
            
            # Entry normal - tiene get()
            elif hasattr(widget, 'get'):
                value = widget.get()
                if isinstance(value, str):
                    return value.strip()
                else:
                    return str(value)
            
            else:
                return ''
        
        except Exception as e:
            print(f"Error obteniendo fecha: {e}")
            return ''
    
    def create_radio_field_centered(self, parent, label_text, field_name, tooltip_text=None):
        """
        Crear campo con RadioButtons (para preguntas Sí/No).
        """
        # Frame principal - CENTRADO
        field_frame = ctk.CTkFrame(parent, fg_color="white", corner_radius=8)
        field_frame.pack(fill="x", padx=40, pady=6)
        
        # Frame interno - Grid layout
        inner_frame = ctk.CTkFrame(field_frame, fg_color="transparent")
        inner_frame.pack(fill="x", padx=20, pady=10)
        
        # Configurar grid con columna fija para RadioButtons
        inner_frame.grid_columnconfigure(0, weight=6)  # Label: flexible
        inner_frame.grid_columnconfigure(1, weight=0, minsize=150)  # RadioButtons: fijo 150px
        
        # ===== LABEL (COLUMNA 0) =====
        label = ctk.CTkLabel(
            inner_frame,
            text=label_text,
            font=("Segoe UI", 12, "bold"),
            anchor="w",  # Alineado a la izquierda
            text_color="#333333"
        )
        label.grid(row=0, column=0, sticky="w", padx=(0, 30))  # 30px espacio a la derecha
        
        # Tooltip si existe
        if tooltip_text:
            ToolTip(label, tooltip_text)
        
        # ===== FRAME PARA RADIOBUTTONS (COLUMNA 1) =====
        radio_frame = ctk.CTkFrame(inner_frame, fg_color="transparent", width=150)
        radio_frame.grid(row=0, column=1, sticky="e")  # Alineado a la derecha
        radio_frame.grid_propagate(False)  # Mantener ancho fijo
        
        # Variable para almacenar selección
        var = tk.StringVar(value="")
        
        # RadioButton SÍ
        radio_si = ctk.CTkRadioButton(
            radio_frame,
            text="Sí",
            variable=var,
            value="Sí",
            font=("Segoe UI", 12),
            fg_color=COLOR_VERDE_HOSPITAL,
            hover_color="#1F5A32",
            border_width_checked=8,
            border_width_unchecked=2,
            width=60  # Ancho fijo para consistencia
        )
        radio_si.pack(side="left", padx=(0, 10))  # 10px entre Sí y No
        
        # RadioButton NO
        radio_no = ctk.CTkRadioButton(
            radio_frame,
            text="No",
            variable=var,
            value="No",
            font=("Segoe UI", 12),
            fg_color=COLOR_VERDE_HOSPITAL,
            hover_color="#1F5A32",
            border_width_checked=8,
            border_width_unchecked=2,
            width=60  # Ancho fijo para consistencia
        )
        radio_no.pack(side="left")
        
        # Guardar variable
        self.manual_widgets[field_name] = var
        
        return var
    
    def _create_flujo_tab(self, parent):
        """Tab de flujo del proceso."""
        scroll = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=10, pady=10)
        
        content = """
    🔄 FLUJO DEL PROCESO DE INVENTARIO

    1️⃣ DATOS MANUALES
    • Información básica del equipo
    • Área, ubicación, responsable
    • Software utilizado (SIFAX, Office, Especializado)
    • Información operativa (horarios, estado)
    • Macroproceso → Proceso → Subproceso (condicional)
    • Cuestionario de clasificación (18 preguntas)
    
    
    → Al terminar, puedes:
        💾 GUARDAR: Solo datos manuales (rápido)
        🔄 ACTUALIZAR: Modificar equipo existente
        ➡️ CONTINUAR: Pasar a detección automática

    2️⃣ RECOPILACIÓN AUTOMÁTICA
    • Detección de hardware (WMI)
        - Marca, Modelo, Serial del equipo, Discos
    • Sistema operativo y arquitectura
    • RAM y procesador
    • Software instalado (Office, Teams, Outlook)
    • Licencias de Windows y Office
    • Red (IP, conexión)
    • Seguridad (Antivirus, actualizaciones)
    
    ⚠️ Este proceso tarda 10-30 segundos

    3️⃣ VALIDACIÓN MIXTA
    • Revisar y corregir datos detectados
    • Agregar información de red (Switch, VLAN)
    • Configurar acceso remoto (AnyDesk)
    • Validar seguridad (Antivirus, Cifrado)

    4️⃣ GUARDADO FINAL
    • Se guardan TODAS las columnas en Excel
    • Archivo: inventario_hospital_v1.xlsx

    🎯 RECOMENDACIONES

    ✓ Usa GUARDAR si no necesitas detección automática
    ✓ Usa RECOPILACIÓN AUTO para inventario completo
    ✓ Puedes ACTUALIZAR equipos en cualquier momento
    ✓ Pasa el mouse sobre preguntas para ver texto completo
        """
        
        label = ctk.CTkLabel(
            scroll,
            text=content,
            font=("Consolas", 11),
            justify="left",
            anchor="w"
        )
        label.pack(fill="both", padx=20, pady=10)

    def _create_macroproceso_tab(self, parent):
        """Tab de explicación de Macroproceso."""
        scroll = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=10, pady=10)
        
        content = """
    🏥 MACROPROCESO → PROCESO → SUBPROCESO

    Esta estructura clasifica el equipo según el proceso hospitalario.
    Las listas son CONDICIONALES (se actualizan automáticamente).

    ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    📊 ESTRATÉGICO (Dirección del hospital)
    → GERENCIA
        • Direccionamiento Estratégico
        • Asignación de Recursos
        • Evaluación y Desempeño
        • Rendición de Cuentas
        • Asesoría Jurídica
    
    → PLANEACIÓN Y CALIDAD
        • Seguridad del Paciente
        • Epidemiología
        • Estadística
        • Formulación de Planes
        • Revisión Documental

    ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    🏥 MISIONAL (Atención de pacientes)
    → AMBULATORIA
        • Consulta Externa
        • Optometría
        • Odontología
        • Atención Hospitalizada
        • Atención Quirúrgica
    
    → SOPORTE DIAGNÓSTICO
        • Laboratorio Clínico
        • Imágenes Diagnósticas
        • Farmacia
        • Fisioterapia
        • Trabajo Social
    
    → URGENCIAS
        • Atención de Urgencias
        • Referencia y Contrarreferencia

    ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    🔧 APOYO (Soporte operativo)
    → FINANCIERA
        • Facturación, Contabilidad, Cartera
    
    → TALENTO HUMANO
        • Recursos Humanos, Nómina
    
    → INFORMACIÓN Y COMUNICACIONES
        • Sistemas, SIAU, Archivo
    
    → AMBIENTE FÍSICO Y TECNOLOGÍA
        • Mantenimiento, Almacén, Servicios

    ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    📋 EVALUACIÓN Y CONTROL
    → CONTROL INTERNO
    → AUDITORÍA MÉDICA
        """
        
        label = ctk.CTkLabel(
            scroll,
            text=content,
            font=("Consolas", 10),
            justify="left",
            anchor="w"
        )
        label.pack(fill="both", padx=20, pady=10)

    def _create_tips_tab(self, parent):
        """Tab de tips para campos."""
        scroll = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=10, pady=10)
        
        content = """
    💡 TIPS PARA LLENAR CAMPOS ESPECÍFICOS

    📅 FECHAS (Formato: YYYY-MM-DD)
    Ejemplo: 2024-01-15
    • Último Mantenimiento: Fecha del último mtto

    📋 CAMPOS OBLIGATORIOS (*)
    • Tipo de Equipo, Área, Ubicación
    • Responsable/Custodio
    • Macroproceso, Proceso, Subproceso
    • Uso SIHOS
    • Estado Operativo

    🔒 CUESTIONARIO (18 PREGUNTAS)
    • Responde Sí o No según corresponda
    • Pasa el mouse sobre cada pregunta para ver texto completo
    • Si tienes dudas, consulta con supervisor
    • Las respuestas se usan para clasificación posterior

    🏥 SOFTWARE ESPECIALIZADO
    • Marca "Sí" si usa programas especiales
    • En "Descripción", especifica cuáles:
        - PACS (Imágenes)
        - RIS (Radiología)
        - LIS (Laboratorio)
        - Contable específico
        - Otros sistemas propietarios

    ⏰ HORARIO DE USO
    • 24/7: Equipos que NUNCA se apagan (UCI, Urgencias)
    • Lun-Vie 7am-7pm: Atención extendida
    • Lun-Vie 7am-5pm: Horario administrativo

    🔧 MANTENIMIENTO
    • Periodicidad: Con qué frecuencia debe hacerse
    • Responsable: Técnico asignado
    • Último: Fecha del último realizado
    • Tipo: Preventivo, Correctivo o Predictivo

    ⚠️ ESTADO OPERATIVO
    • Operativo - Óptimo: Funciona perfecto
    • Operativo - Regular: Funciona con fallas menores
    • Operativo - Deficiente: Funciona mal, necesita atención
    • Fuera de Servicio: No funciona
    • En Reparación: En proceso de arreglo
        """
        
        label = ctk.CTkLabel(
            scroll,
            text=content,
            font=("Consolas", 10),
            justify="left",
            anchor="w"
        )
        label.pack(fill="both", padx=20, pady=10)        
    
    def log_progress(self, message):
        """Agregar mensaje al log."""
        if hasattr(self, 'log_text'):
            self.log_text.insert("end", message + "\n")
            self.log_text.see("end")
            self.root.update()
    
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
    
    def detect_anydesk(self):
        """Detectar ID de AnyDesk si está instalado."""
        try:
            rutas_posibles = []

            pf = os.environ.get("ProgramFiles")
            pf86 = os.environ.get("ProgramFiles(x86)")
            appdata = os.environ.get("APPDATA")
            localappdata = os.environ.get("LOCALAPPDATA")

            if pf:
                rutas_posibles.append(os.path.join(pf, "AnyDesk", "AnyDesk.exe"))
            if pf86:
                rutas_posibles.append(os.path.join(pf86, "AnyDesk", "AnyDesk.exe"))
            if appdata:
                rutas_posibles.append(os.path.join(appdata, "AnyDesk", "AnyDesk.exe"))
            if localappdata:
                rutas_posibles.append(os.path.join(localappdata, "AnyDesk", "AnyDesk.exe"))

            path_anydesk = shutil.which("AnyDesk.exe")
            if path_anydesk:
                rutas_posibles.append(path_anydesk)

            instalado = any(os.path.exists(ruta) for ruta in rutas_posibles)

            # Intentar extraer ID desde archivos de configuración típicos
            archivos_config = []
            programdata = os.environ.get("PROGRAMDATA")

            if programdata:
                archivos_config.extend([
                    os.path.join(programdata, "AnyDesk", "service.conf"),
                    os.path.join(programdata, "AnyDesk", "system.conf"),
                ])
            if appdata:
                archivos_config.extend([
                    os.path.join(appdata, "AnyDesk", "service.conf"),
                    os.path.join(appdata, "AnyDesk", "system.conf"),
                ])
            if localappdata:
                archivos_config.extend([
                    os.path.join(localappdata, "AnyDesk", "service.conf"),
                    os.path.join(localappdata, "AnyDesk", "system.conf"),
                ])

            patron_clave = re.compile(
                r"(?i)(?:ad\.anynet\.id|anynet\.id|client_id)\s*=\s*\"?([0-9]{6,12})\"?"
            )
            patron_numerico = re.compile(r"\b([0-9]{6,12})\b")

            for archivo in archivos_config:
                if not os.path.exists(archivo):
                    continue
                try:
                    with open(archivo, "r", encoding="utf-8", errors="ignore") as f:
                        contenido = f.read()

                    match = patron_clave.search(contenido)
                    if match:
                        return match.group(1)

                    match_num = patron_numerico.search(contenido)
                    if match_num:
                        return match_num.group(1)
                except:
                    continue

            if instalado:
                return "Instalado - Verificar ID"
            return "No instalado"
        except Exception:
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
    


    def log_progress(self, message):
        """Agregar mensaje al log."""
        if hasattr(self, 'log_text'):
            self.log_text.insert("end", message + "\n")
            self.log_text.see("end")
            self.root.update()
    
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

    def start_automatic_collection(self):
        """Iniciar recopilación automática."""
        # Validar campos obligatorios
        required = ['tipo_equipo', 'area_servicio', 'ubicacion_especifica',
           'responsable_custodio', 'macroproceso', 'proceso', 'subproceso',
           'uso_sihos', 'estado_operativo']
        
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
                # ⚠️ CRÍTICO: StringVar NO tiene winfo_exists() - verificar PRIMERO
                if isinstance(widget, tk.StringVar):
                    value = widget.get().strip()
                    if value and value != "":
                        self.equipment_data[field_name] = value
                
                # Widgets de CTk - SÍ tienen winfo_exists()
                elif hasattr(widget, 'winfo_exists') and widget.winfo_exists():
                    if isinstance(widget, ctk.CTkEntry):
                        value = widget.get().strip()
                        if value:
                            self.equipment_data[field_name] = value
                    elif isinstance(widget, ctk.CTkComboBox):
                        value = widget.get().strip()
                        if value and value != "Seleccionar...":
                            self.equipment_data[field_name] = value
            except Exception as e:
                print(f"Error capturando {field_name}: {e}")
                pass
        
        # Mostrar ventana de progreso
        self.show_progress_window()
        
        # Ejecutar recopilación en thread separado
        thread = threading.Thread(target=self.collect_automatic_data)
        thread.daemon = True
        thread.start()
    


    def collect_automatic_data(self):
        """Recopilar datos automáticos."""
        self.verde_data = {}
        
        # 1. Nombre del equipo
        self.log_progress("📋 Identificación del equipo...")
        self.verde_data['nombre_equipo'] = detect_hostname_local()
        self.log_progress(f"   ✓ Nombre: {self.verde_data['nombre_equipo']}")
        
        # 2. Hardware con WMI
        self.log_progress("\n💻 Detectando hardware con WMI...")
        hw_info = detect_hardware_wmi()
        
        # Equipo
        self.verde_data['marca'] = hw_info['marca']
        self.verde_data['modelo'] = hw_info['modelo']
        self.verde_data['serial'] = hw_info['serial']
        
        self.log_progress(f"   ✓ Marca: {self.verde_data['marca']}")
        self.log_progress(f"   ✓ Modelo: {self.verde_data['modelo']}")
        self.log_progress(f"   ✓ Serial: {self.verde_data['serial']}")
        
        # ===== DISCO 1 (PRIMARIO) =====
        self.log_progress(f"\n💿 Disco 1 (Primario):")
        self.verde_data['disco1_capacidad'] = hw_info['disco1_capacidad']
        self.verde_data['disco1_tipo'] = hw_info['disco1_tipo']
        self.verde_data['disco1_serial'] = hw_info['disco1_serial']
        self.verde_data['disco1_marca'] = hw_info['disco1_marca']
        self.verde_data['disco1_modelo'] = hw_info['disco1_modelo']
        
        self.log_progress(f"   ✓ Capacidad: {hw_info['disco1_capacidad']} GB")
        self.log_progress(f"   ✓ Tipo: {hw_info['disco1_tipo']}")
        self.log_progress(f"   ✓ Serial: {hw_info['disco1_serial']}")
        self.log_progress(f"   ✓ Marca: {hw_info['disco1_marca']}")
        self.log_progress(f"   ✓ Modelo: {hw_info['disco1_modelo']}")
        
        # ===== DISCO 2 (SECUNDARIO) =====
        if hw_info['disco2_capacidad'] != 'No tiene':
            self.log_progress(f"\n💿 Disco 2 (Secundario) Detectado:")
            self.log_progress(f"   ✓ Capacidad: {hw_info['disco2_capacidad']} GB")
            self.log_progress(f"   ✓ Tipo: {hw_info['disco2_tipo']}")
            self.log_progress(f"   ✓ Serial: {hw_info['disco2_serial']}")
            self.log_progress(f"   ✓ Marca: {hw_info['disco2_marca']}")
            self.log_progress(f"   ✓ Modelo: {hw_info['disco2_modelo']}")
        else:
            self.log_progress(f"\n💿 Disco 2 (Secundario): No detectado")
        
        # 5-7. Sistema Operativo
        self.log_progress("\n🪟 Sistema Operativo...")
        self.verde_data['sistema_operativo'] = f"{platform.system()} {platform.release()}"
        self.verde_data['arquitectura_so'] = "64 bits" if "64" in platform.machine() else "32 bits"
        self.verde_data['procesador'] = platform.processor() or "No detectado"
        self.verde_data['cpu_nucleos'] = detect_cpu_cores()
        
        self.log_progress(f"   ✓ SO: {self.verde_data['sistema_operativo']}")
        self.log_progress(f"   ✓ Arquitectura: {self.verde_data['arquitectura_so']}")
        self.log_progress(f"   ✓ Procesador: {self.verde_data['procesador'][:50]}...")
        self.log_progress(f"   ✓ Núcleos CPU: {self.verde_data['cpu_nucleos']}")
        
        # 8-9. RAM y Almacenamiento
        if HAS_PSUTIL:
            # Obtener RAM utilizable
            ram_bytes = psutil.virtual_memory().total
            ram_gib_usable = ram_bytes / (1024**3)  # GiB utilizables
            
            # Tamaños comerciales estándar
            common_sizes = [2, 4, 6, 8, 12, 16, 24, 32, 48, 64, 128]
            
            # Redondeo inteligente con margen del 15%
            # Busca el tamaño comercial más probable
            # considerando que puede haber RAM reservada
            best_match = None
            min_diff = float('inf')
            
            for size in common_sizes:
                # Considerar margen de -20% (por GPU integrada, BIOS, etc)
                expected_usable = size * 0.80  # 80% del tamaño comercial
                diff = abs(ram_gib_usable - expected_usable)
                
                # También considerar coincidencia directa
                direct_diff = abs(ram_gib_usable - size)
                
                # Usar la mejor coincidencia
                actual_diff = min(diff, direct_diff)
                
                if actual_diff < min_diff:
                    min_diff = actual_diff
                    best_match = size
            
            ram_gb = best_match
            
            self.verde_data['ram_gb'] = str(ram_gb)
            self.log_progress(f"   ✓ RAM: {ram_gb} GB (utilizable: {ram_gib_usable:.2f} GiB)")
            
            # Almacenamiento
            try:
                disk = psutil.disk_usage('C:\\\\')
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
        
        # 19-23. Red
        self.log_progress("\\n🌐 Red...")

        # IP Local
        ip_local = detect_ip_local()
        self.verde_data['direccion_ip'] = ip_local
        self.log_progress(f"   ✓ IP Local: {ip_local}")

        # MAC Address
        mac_address = detect_mac_address()
        self.verde_data['mac_address'] = mac_address
        self.log_progress(f"   ✓ MAC Address: {mac_address}")

        self.verde_data['tipo_conexion'] = "Ethernet"  # Default

        # Navegador predeterminado
        navegador = detect_default_browser()
        self.verde_data['navegador_predeterminado'] = navegador
        self.log_progress(f"   ✓ Navegador: {navegador}")

        # Unidades de red mapeadas (nuevo)
        unidades_red = detect_network_drives()
        self.verde_data['unidades_red_mapeadas'] = unidades_red
        self.log_progress(f"   ✓ Unidades de red: {unidades_red}")
        
        # Velocidad de tarjeta de red (NUEVA)
        velocidad_red = detect_network_speed()
        self.verde_data['velocidad_red'] = velocidad_red
        self.log_progress(f"   ✓ Velocidad de red: {velocidad_red}")
        
        # Uso actual de recursos
        cpu_uso, ram_uso = detect_current_cpu_ram_usage()
        self.verde_data['cpu_uso_actual'] = cpu_uso
        self.verde_data['ram_uso_actual'] = ram_uso
        self.log_progress(f"   ✓ Uso actual CPU: {cpu_uso}")
        self.log_progress(f"   ✓ Uso actual RAM: {ram_uso}")

        # 24-26. Seguridad
        self.log_progress("\n🔒 Seguridad...")
        defender_status = detect_defender_status()
        self.verde_data['antivirus_instalado'] = defender_status
        self.verde_data['tpm_status'] = detect_tpm_status()
        self.verde_data['bitlocker_status'] = detect_bitlocker_status()
        self.verde_data['windows_update_activo'] = "Sí"
        
        last_update = detect_last_windows_update()
        self.verde_data['ultima_act_windows'] = last_update
        
        self.log_progress(f"   ✓ Defender: {defender_status}")
        self.log_progress(f"   ✓ TPM: {self.verde_data['tpm_status']}")
        self.log_progress(f"   ✓ BitLocker: {self.verde_data['bitlocker_status']}")
        self.log_progress(f"   ✓ Última actualización: {last_update}")
        
        self.log_progress("\n✅ Recopilación automática completada")
        
        # Cerrar ventana de progreso
        self.progress_bar.stop()
        self.root.after(800, lambda: self.progress_window.destroy())
        
        # Mostrar validación de campos mixtos
        self.root.after(1200, self.show_mixed_validation)

    def show_mixed_validation(self):
        """Mostrar ventana de validación de campos mixtos (AZULES)."""
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
            text="Valida o corrige los siguientes campos:",
            font=("Arial", 13)
        )
        info.pack(pady=(0, 20))
        
        # Frame scrollable
        scroll_frame = ctk.CTkScrollableFrame(validation_window, width=850, height=380)
        scroll_frame.pack(pady=10, padx=25)
        
        # Campos mixtos (SIN DISCO 2)
        self.mixed_widgets = {}
        switch_puerto_detectado, vlan_detectada = detect_switch_and_vlan()
        
        mixed_fields = [
            # RED Y ACCESO REMOTO
            ("Switch / Puerto", "switch_puerto", "entry", switch_puerto_detectado),
            ("VLAN Asignada", "vlan_asignada", "entry", vlan_detectada),
            ("ID AnyDesk", "id_anydesk", "entry", self.detect_anydesk()),
            ("Otro Acceso Remoto", "otro_acceso_remoto", "entry", "Ninguno"),
            
            # SEGURIDAD
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
            
            if field_type == "combobox":
                widget = ctk.CTkComboBox(
                    field_frame,
                    values=default if isinstance(default, list) else ["No detectado"],
                    width=500,
                    font=("Arial", 12),
                    dropdown_font=("Arial", 11),
                    height=32
                )
                if isinstance(default, list) and len(default) > 0:
                    widget.set(default[0])
            else:
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
    


    def save_equipo_automatic_update(self):
        """Actualizar equipo con datos automáticos (recopilación automática en modo actualización)."""
        try:
            wb = load_workbook(self.excel_path)
            ws = wb["Equipos de Cómputo"]
            
            row = self.equipo_update_row
            codigo = self.equipo_update_code
            
            print(f"DEBUG: Actualizando {codigo} en fila {row} con datos automáticos")
            
            # ===== ACTUALIZAR COLUMNAS 4-38: NARANJAS (datos manuales) =====
            col = 4
            naranja_fields = [
                'tipo_equipo', 'area_servicio', 'ubicacion_especifica',
                'responsable_custodio', 'macroproceso', 'proceso', 'subproceso',
                'uso_sihos', 'uso_office_basico', 'software_especializado', 
                'descripcion_software', 'funcion_principal'
            ]
            
            for field in naranja_fields:
                value = self.equipment_data.get(field, '')
                ws.cell(row=row, column=col, value=value)
                col += 1
            
            # Cuestionario (cols 16-33)
            for i in range(1, 10):  # CONFIDENCIALIDAD
                ws.cell(row=row, column=15 + i, value=self.equipment_data.get(f'conf_{i}', ''))
            for i in range(1, 4):   # INTEGRIDAD
                ws.cell(row=row, column=24 + i, value=self.equipment_data.get(f'int_{i}', ''))
            for i in range(1, 7):   # CRITICIDAD
                ws.cell(row=row, column=27 + i, value=self.equipment_data.get(f'crit_{i}', ''))
            
            # Operativos (cols 34-38)
            ws.cell(row=row, column=34, value=self.equipment_data.get('horario_uso', ''))
            ws.cell(row=row, column=35, value=self.equipment_data.get('estado_operativo', ''))
            ws.cell(row=row, column=36, value=self.equipment_data.get('observaciones_tecnicas', ''))
            ws.cell(row=row, column=37, value=self.equipment_data.get('periodicidad_mtto', ''))
            ws.cell(row=row, column=38, value=self.equipment_data.get('responsable_mtto', ''))
            
            # ===== ACTUALIZAR COLUMNAS 39-79: VERDES (datos automáticos) =====
            ws.cell(row=row, column=3, value=self.verde_data.get('nombre_equipo', ''))
            ws.cell(row=row, column=39, value=self.verde_data.get('marca', ''))
            ws.cell(row=row, column=40, value=self.verde_data.get('modelo', ''))
            ws.cell(row=row, column=41, value=self.verde_data.get('serial', ''))
            ws.cell(row=row, column=42, value=self.verde_data.get('sistema_operativo', ''))
            ws.cell(row=row, column=43, value=self.verde_data.get('arquitectura_so', ''))
            ws.cell(row=row, column=44, value=self.verde_data.get('procesador', ''))
            ws.cell(row=row, column=45, value=self.verde_data.get('ram_gb', ''))
            
            # Disco 1
            ws.cell(row=row, column=46, value=self.verde_data.get('disco1_capacidad', ''))
            ws.cell(row=row, column=47, value=self.verde_data.get('disco1_tipo', ''))
            ws.cell(row=row, column=48, value=self.verde_data.get('disco1_serial', ''))
            ws.cell(row=row, column=49, value=self.verde_data.get('disco1_marca', ''))
            ws.cell(row=row, column=50, value=self.verde_data.get('disco1_modelo', ''))
            
            # Disco 2
            ws.cell(row=row, column=51, value=self.verde_data.get('disco2_capacidad', 'No tiene'))
            ws.cell(row=row, column=52, value=self.verde_data.get('disco2_tipo', 'No tiene'))
            ws.cell(row=row, column=53, value=self.verde_data.get('disco2_serial', 'No tiene'))
            ws.cell(row=row, column=54, value=self.verde_data.get('disco2_marca', 'No tiene'))
            ws.cell(row=row, column=55, value=self.verde_data.get('disco2_modelo', 'No tiene'))
            
            # Software
            ws.cell(row=row, column=56, value=self.verde_data.get('uso_navegador_web', ''))
            ws.cell(row=row, column=57, value=self.verde_data.get('version_office', ''))
            ws.cell(row=row, column=58, value=self.verde_data.get('licencia_office', ''))
            ws.cell(row=row, column=59, value=self.verde_data.get('uso_teams', ''))
            ws.cell(row=row, column=60, value=self.verde_data.get('uso_outlook', ''))
            
            # Licencia Windows
            ws.cell(row=row, column=61, value=self.verde_data.get('licencia_windows', ''))
            ws.cell(row=row, column=62, value=self.verde_data.get('key_windows', ''))
            ws.cell(row=row, column=63, value=self.verde_data.get('estado_licencia_windows', ''))
            
            # Red
            ws.cell(row=row, column=64, value=self.verde_data.get('direccion_ip', ''))
            ws.cell(row=row, column=65, value=self.verde_data.get('mac_address', ''))
            ws.cell(row=row, column=66, value=self.verde_data.get('tipo_conexion', ''))
            ws.cell(row=row, column=67, value=self.verde_data.get('navegador_predeterminado', ''))
            ws.cell(row=row, column=68, value=self.verde_data.get('unidades_red_mapeadas', ''))
            ws.cell(row=row, column=69, value=self.verde_data.get('antivirus_instalado', ''))
            ws.cell(row=row, column=70, value=self.verde_data.get('ultima_act_windows', ''))
            ws.cell(row=row, column=71, value=self.verde_data.get('windows_update_activo', ''))
            ws.cell(row=row, column=79, value=self.verde_data.get('velocidad_red', 'No detectado'))
            ws.cell(row=row, column=80, value=self.verde_data.get('tpm_status', ''))
            ws.cell(row=row, column=81, value=self.verde_data.get('bitlocker_status', ''))
            ws.cell(row=row, column=82, value=self.verde_data.get('cpu_nucleos', ''))
            ws.cell(row=row, column=83, value=self.verde_data.get('cpu_uso_actual', ''))
            ws.cell(row=row, column=84, value=self.verde_data.get('ram_uso_actual', ''))
            
            # ===== ACTUALIZAR COLUMNAS 72-78: AZULES (datos mixtos) =====
            ws.cell(row=row, column=72, value=self.azul_data.get('switch_puerto', ''))
            ws.cell(row=row, column=73, value=self.azul_data.get('vlan_asignada', ''))
            ws.cell(row=row, column=74, value=self.azul_data.get('id_anydesk', ''))
            ws.cell(row=row, column=75, value=self.azul_data.get('otro_acceso_remoto', ''))
            ws.cell(row=row, column=76, value=self.azul_data.get('estado_antivirus', ''))
            ws.cell(row=row, column=77, value=self.azul_data.get('cifrado_disco', ''))
            ws.cell(row=row, column=78, value=self.azul_data.get('tipo_usuario_local', ''))
            
            wb.save(self.excel_path)
            wb.close()

            messagebox.showinfo("Éxito", f"✅ Equipo {codigo} actualizado con datos automáticos")
            
            # Limpiar modo actualización
            self.reset_after_update_equipos()
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al actualizar con datos automáticos:\n{e}")
            import traceback
            traceback.print_exc()

    def save_to_excel(self):
        """Guardar TODOS los datos en Excel (después de recopilación automática)."""
        if not HAS_OPENPYXL:
            messagebox.showerror("Error", "Necesitas instalar openpyxl")
            return
        
        # ===== DETECTAR SI ESTAMOS EN MODO ACTUALIZACIÓN =====
        if hasattr(self, 'equipo_update_code') and self.equipo_update_code:
            # MODO ACTUALIZACIÓN - llamar función específica
            self.save_equipo_automatic_update()
            return
        
        # ===== MODO NUEVO EQUIPO =====
        try:
            wb = load_workbook(self.excel_path)
            ws = wb["Equipos de Cómputo"]
            
            # ===== BUSCAR ÚLTIMO CONSECUTIVO Y PRIMERA FILA VACÍA =====
            last_consecutive = 0
            primera_fila_vacia = 2
            
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
                    primera_fila_vacia = row
                    break
            
            # Siguiente consecutivo
            next_consecutivo = last_consecutive + 1
            next_codigo = f"EQC-{next_consecutivo:04d}"
            
            # Guardar para mensaje de completado (FIX BUG #3)
            self.last_saved_consecutive = next_consecutivo
            self.last_saved_code = next_codigo
            nueva_fila = primera_fila_vacia
            
            print(f"DEBUG AUTO: last_consecutive={last_consecutive}, next={next_consecutivo}, fila={nueva_fila}")
            
            # ===== COLUMNAS 1-3: IDENTIFICACIÓN =====
            ws.cell(row=nueva_fila, column=1, value=next_consecutivo)
            ws.cell(row=nueva_fila, column=2, value=next_codigo)
            ws.cell(row=nueva_fila, column=3, value=self.verde_data.get('nombre_equipo', ''))
            
            # ===== COLUMNAS 4-10: BÁSICOS (NARANJAS) =====
            ws.cell(row=nueva_fila, column=4, value=self.equipment_data.get('tipo_equipo', ''))
            ws.cell(row=nueva_fila, column=5, value=self.equipment_data.get('area_servicio', ''))
            ws.cell(row=nueva_fila, column=6, value=self.equipment_data.get('ubicacion_especifica', ''))
            ws.cell(row=nueva_fila, column=7, value=self.equipment_data.get('responsable_custodio', ''))
            ws.cell(row=nueva_fila, column=8, value=self.equipment_data.get('macroproceso', ''))
            ws.cell(row=nueva_fila, column=9, value=self.equipment_data.get('proceso', ''))
            ws.cell(row=nueva_fila, column=10, value=self.equipment_data.get('subproceso', ''))
            
            # ===== COLUMNAS 11-15: SOFTWARE (NARANJAS) =====
            ws.cell(row=nueva_fila, column=11, value=self.equipment_data.get('uso_sihos', ''))
            ws.cell(row=nueva_fila, column=12, value=self.equipment_data.get('uso_office_basico', ''))
            ws.cell(row=nueva_fila, column=13, value=self.equipment_data.get('software_especializado', ''))
            ws.cell(row=nueva_fila, column=14, value=self.equipment_data.get('descripcion_software', ''))
            ws.cell(row=nueva_fila, column=15, value=self.equipment_data.get('funcion_principal', ''))
            
            # ===== COLUMNAS 16-24: CONFIDENCIALIDAD (NARANJAS) =====
            for i in range(1, 10):
                ws.cell(row=nueva_fila, column=15 + i, value=self.equipment_data.get(f'conf_{i}', ''))
            
            # ===== COLUMNAS 25-27: INTEGRIDAD (NARANJAS) =====
            for i in range(1, 4):
                ws.cell(row=nueva_fila, column=24 + i, value=self.equipment_data.get(f'int_{i}', ''))
            
            # ===== COLUMNAS 28-33: CRITICIDAD (NARANJAS) =====
            for i in range(1, 7):
                ws.cell(row=nueva_fila, column=27 + i, value=self.equipment_data.get(f'crit_{i}', ''))
            
            # ===== COLUMNAS 34-38: OPERATIVOS (NARANJAS) =====
            ws.cell(row=nueva_fila, column=34, value=self.equipment_data.get('horario_uso', ''))
            ws.cell(row=nueva_fila, column=35, value=self.equipment_data.get('estado_operativo', ''))
            ws.cell(row=nueva_fila, column=36, value=self.equipment_data.get('observaciones_tecnicas', ''))
            ws.cell(row=nueva_fila, column=37, value=self.equipment_data.get('periodicidad_mtto', ''))
            ws.cell(row=nueva_fila, column=38, value=self.equipment_data.get('responsable_mtto', ''))
            
            # ===== COLUMNAS 39-45: HARDWARE (VERDES) =====
            ws.cell(row=nueva_fila, column=39, value=self.verde_data.get('marca', ''))
            ws.cell(row=nueva_fila, column=40, value=self.verde_data.get('modelo', ''))
            ws.cell(row=nueva_fila, column=41, value=self.verde_data.get('serial', ''))
            ws.cell(row=nueva_fila, column=42, value=self.verde_data.get('sistema_operativo', ''))
            ws.cell(row=nueva_fila, column=43, value=self.verde_data.get('arquitectura_so', ''))
            ws.cell(row=nueva_fila, column=44, value=self.verde_data.get('procesador', ''))
            ws.cell(row=nueva_fila, column=45, value=self.verde_data.get('ram_gb', ''))
            
            # ===== COLUMNAS 46-50: DISCO 1 (VERDES) =====
            ws.cell(row=nueva_fila, column=46, value=self.verde_data.get('disco1_capacidad', ''))
            ws.cell(row=nueva_fila, column=47, value=self.verde_data.get('disco1_tipo', ''))
            ws.cell(row=nueva_fila, column=48, value=self.verde_data.get('disco1_serial', ''))
            ws.cell(row=nueva_fila, column=49, value=self.verde_data.get('disco1_marca', ''))
            ws.cell(row=nueva_fila, column=50, value=self.verde_data.get('disco1_modelo', ''))
            
            # ===== COLUMNAS 51-55: DISCO 2 (VERDES) =====
            ws.cell(row=nueva_fila, column=51, value=self.verde_data.get('disco2_capacidad', 'No tiene'))
            ws.cell(row=nueva_fila, column=52, value=self.verde_data.get('disco2_tipo', 'No tiene'))
            ws.cell(row=nueva_fila, column=53, value=self.verde_data.get('disco2_serial', 'No tiene'))
            ws.cell(row=nueva_fila, column=54, value=self.verde_data.get('disco2_marca', 'No tiene'))
            ws.cell(row=nueva_fila, column=55, value=self.verde_data.get('disco2_modelo', 'No tiene'))
            
            # ===== COLUMNAS 56-63: SOFTWARE (VERDES) =====
            ws.cell(row=nueva_fila, column=56, value=self.verde_data.get('uso_navegador_web', ''))
            ws.cell(row=nueva_fila, column=57, value=self.verde_data.get('version_office', ''))
            ws.cell(row=nueva_fila, column=58, value=self.verde_data.get('licencia_office', ''))
            ws.cell(row=nueva_fila, column=59, value=self.verde_data.get('uso_teams', ''))
            ws.cell(row=nueva_fila, column=60, value=self.verde_data.get('uso_outlook', ''))
            ws.cell(row=nueva_fila, column=61, value=self.verde_data.get('licencia_windows', ''))
            ws.cell(row=nueva_fila, column=62, value=self.verde_data.get('key_windows', ''))
            ws.cell(row=nueva_fila, column=63, value=self.verde_data.get('estado_licencia_windows', ''))
            
            # ===== COLUMNAS 64-68: RED (VERDES) =====
            ws.cell(row=nueva_fila, column=64, value=self.verde_data.get('direccion_ip', ''))
            ws.cell(row=nueva_fila, column=65, value=self.verde_data.get('mac_address', ''))
            ws.cell(row=nueva_fila, column=66, value=self.verde_data.get('tipo_conexion', ''))
            ws.cell(row=nueva_fila, column=67, value=self.verde_data.get('navegador_predeterminado', ''))
            ws.cell(row=nueva_fila, column=68, value=self.verde_data.get('unidades_red_mapeadas', ''))
            
            # ===== COLUMNAS 69-71: SEGURIDAD (VERDES) =====
            ws.cell(row=nueva_fila, column=69, value=self.verde_data.get('antivirus_instalado', ''))
            ws.cell(row=nueva_fila, column=70, value=self.verde_data.get('ultima_act_windows', ''))
            ws.cell(row=nueva_fila, column=71, value=self.verde_data.get('windows_update_activo', ''))
            
            # ===== COLUMNAS 72-78: MIXTOS (AZULES) =====
            ws.cell(row=nueva_fila, column=72, value=self.azul_data.get('switch_puerto', ''))
            ws.cell(row=nueva_fila, column=73, value=self.azul_data.get('vlan_asignada', ''))
            ws.cell(row=nueva_fila, column=74, value=self.azul_data.get('id_anydesk', ''))
            ws.cell(row=nueva_fila, column=75, value=self.azul_data.get('otro_acceso_remoto', ''))
            ws.cell(row=nueva_fila, column=76, value=self.azul_data.get('estado_antivirus', ''))
            ws.cell(row=nueva_fila, column=77, value=self.azul_data.get('cifrado_disco', ''))
            ws.cell(row=nueva_fila, column=78, value=self.azul_data.get('tipo_usuario_local', ''))

            # ===== COLUMNA 79: VELOCIDAD DE RED (VERDE - NUEVA) =====
            ws.cell(row=nueva_fila, column=79, value=self.verde_data.get('velocidad_red', ''))
            ws.cell(row=nueva_fila, column=80, value=self.verde_data.get('tpm_status', ''))
            ws.cell(row=nueva_fila, column=81, value=self.verde_data.get('bitlocker_status', ''))
            ws.cell(row=nueva_fila, column=82, value=self.verde_data.get('cpu_nucleos', ''))
            ws.cell(row=nueva_fila, column=83, value=self.verde_data.get('cpu_uso_actual', ''))
            ws.cell(row=nueva_fila, column=84, value=self.verde_data.get('ram_uso_actual', ''))
            
            # Guardar
            wb.save(self.excel_path)
            wb.close()

            print(f"✅ Guardado exitoso: {next_codigo} en fila {nueva_fila}")
            
            messagebox.showinfo("Éxito", f"✅ Equipo guardado: {next_codigo}\n\nDatos completos: 78 columnas")
            
            # Actualizar current_row para próximo equipo
            self.current_row = nueva_fila + 1
            
            # Volver al formulario
            self.root.after(100, self.show_manual_form_in_container)
            
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"Error al guardar en Excel:\n{e}")
    
    def save_equipo_manual_only(self):
        """Guardar solo datos manuales (sin recopilación automática)."""
        
        # ===== VALIDACIÓN DE CAMPOS OBLIGATORIOS =====
        required_fields = {
            'tipo_equipo': 'Tipo de Equipo',
            'area_servicio': 'Área / Servicio',
            'ubicacion_especifica': 'Ubicación Específica',
            'responsable_custodio': 'Responsable / Custodio',
            'macroproceso': 'Macroproceso',
            'proceso': 'Proceso',
            'subproceso': 'Subproceso',
            'uso_sihos': 'Uso SIHOS',
            'estado_operativo': 'Estado Operativo'
        }
        
        missing_fields = []
        
        for field_name, field_label in required_fields.items():
            if field_name in self.manual_widgets:
                widget = self.manual_widgets[field_name]
                try:
                    if hasattr(widget, 'winfo_exists') and widget.winfo_exists():
                        if isinstance(widget, (ctk.CTkEntry, ctk.CTkComboBox)):
                            value = widget.get().strip()
                        elif isinstance(widget, tk.StringVar):
                            value = widget.get().strip()
                        else:
                            value = ''
                        
                        if not value:
                            missing_fields.append(field_label)
                    else:
                        missing_fields.append(field_label)
                except:
                    missing_fields.append(field_label)
            else:
                missing_fields.append(field_label)
        
        if missing_fields:
            messagebox.showwarning(
                "Campos Requeridos",
                f"Por favor completa los siguientes campos obligatorios:\n\n" + 
                "\n".join(f"• {field}" for field in missing_fields)
            )
            return
        
        # ===== RECOPILAR DATOS MANUALES =====
        datos_guardados = {}
        
        for field_name, widget in self.manual_widgets.items():
            try:
                # ⚠️ CRÍTICO: StringVar NO tiene winfo_exists() - verificar PRIMERO
                if isinstance(widget, tk.StringVar):
                    datos_guardados[field_name] = widget.get()
                
                # Widgets de CTk - SÍ tienen winfo_exists()
                elif hasattr(widget, 'winfo_exists') and widget.winfo_exists():
                    # Entry (campos de texto)
                    if isinstance(widget, ctk.CTkEntry):
                        datos_guardados[field_name] = widget.get()
                    
                    # ComboBox (listas desplegables)
                    elif isinstance(widget, ctk.CTkComboBox):
                        datos_guardados[field_name] = widget.get()
                    
                    else:
                        datos_guardados[field_name] = ''
                else:
                    datos_guardados[field_name] = ''
            except Exception as e:
                print(f"Error al obtener valor de {field_name}: {e}")
                datos_guardados[field_name] = ''
        
        # ===== GUARDAR EN EXCEL =====
        try:
            # Abrir Excel
            if not self.excel_path:
                messagebox.showerror("Error", "No hay Excel cargado")
                return
            
            if not os.path.exists(self.excel_path):
                messagebox.showerror("Error", f"No se encontró el archivo: {self.excel_path}")
                return
            
            wb = openpyxl.load_workbook(self.excel_path)
            
            if "Equipos de Cómputo" not in wb.sheetnames:
                messagebox.showerror("Error", "No se encontró la hoja: Equipos de Cómputo")
                return
            
            ws = wb["Equipos de Cómputo"]
            
            # ===== BUSCAR ÚLTIMO CONSECUTIVO Y PRIMERA FILA VACÍA =====
            # Buscar ÚLTIMA fila con datos (para consecutivo correcto)
            last_consecutive = 0
            primera_fila_vacia = 2  # Por defecto

            for row in range(2, 500):
                value = ws.cell(row=row, column=1).value
                
                if value is not None:
                    # Hay datos - actualizar consecutivo
                    try:
                        consecutivo = int(value)
                        if consecutivo > last_consecutive:
                            last_consecutive = consecutivo
                    except:
                        pass
                else:
                    # Primera fila vacía encontrada
                    primera_fila_vacia = row
                    break

            # Siguiente consecutivo (si no hay datos, empieza en 1)
            next_consecutivo = last_consecutive + 1
            next_codigo = f"EQC-{next_consecutivo:04d}"
            nueva_fila = primera_fila_vacia

            print(f"DEBUG: last_consecutive={last_consecutive}, next_consecutivo={next_consecutivo}, nueva_fila={nueva_fila}")
            
            # ===== MAPEO A 80 COLUMNAS (VERSION ACTUAL) =====
            
            # Cols 1-2: Identificación
            ws.cell(row=nueva_fila, column=1).value = next_consecutivo  # N° Consecutivo
            ws.cell(row=nueva_fila, column=2).value = next_codigo       # Código
            
            # Col 3: Nombre Equipo (VERDE - auto detección básica)
            ws.cell(row=nueva_fila, column=3).value = detect_hostname_local()
            
            # Cols 4-10: Básicos (NARANJA)
            ws.cell(row=nueva_fila, column=4).value = datos_guardados.get('tipo_equipo', '')
            ws.cell(row=nueva_fila, column=5).value = datos_guardados.get('area_servicio', '')
            ws.cell(row=nueva_fila, column=6).value = datos_guardados.get('ubicacion_especifica', '')
            ws.cell(row=nueva_fila, column=7).value = datos_guardados.get('responsable_custodio', '')
            ws.cell(row=nueva_fila, column=8).value = datos_guardados.get('macroproceso', '')
            ws.cell(row=nueva_fila, column=9).value = datos_guardados.get('proceso', '')
            ws.cell(row=nueva_fila, column=10).value = datos_guardados.get('subproceso', '')
            
            # Cols 11-15: Software (NARANJA)
            ws.cell(row=nueva_fila, column=11).value = datos_guardados.get('uso_sihos', '')
            ws.cell(row=nueva_fila, column=12).value = datos_guardados.get('uso_office_basico', '')
            ws.cell(row=nueva_fila, column=13).value = datos_guardados.get('software_especializado', '')
            ws.cell(row=nueva_fila, column=14).value = datos_guardados.get('descripcion_software', '')
            ws.cell(row=nueva_fila, column=15).value = datos_guardados.get('funcion_principal', '')
            
            # Cols 16-33: Cuestionario 18 preguntas (NARANJA)
            # CONFIDENCIALIDAD (9)
            for i in range(1, 10):
                ws.cell(row=nueva_fila, column=15 + i).value = datos_guardados.get(f'conf_{i}', '')

            # INTEGRIDAD (3)
            for i in range(1, 4):
                ws.cell(row=nueva_fila, column=24 + i).value = datos_guardados.get(f'int_{i}', '')

            # CRITICIDAD (6)
            for i in range(1, 7):
                ws.cell(row=nueva_fila, column=27 + i).value = datos_guardados.get(f'crit_{i}', '')
            
            # Cols 34-38: Operativos (NARANJA) - Solo 5 campos
            ws.cell(row=nueva_fila, column=34).value = datos_guardados.get('horario_uso', '')
            ws.cell(row=nueva_fila, column=35).value = datos_guardados.get('estado_operativo', '')
            ws.cell(row=nueva_fila, column=36).value = datos_guardados.get('observaciones_tecnicas', '')
            ws.cell(row=nueva_fila, column=37).value = datos_guardados.get('periodicidad_mtto', '')
            ws.cell(row=nueva_fila, column=38).value = datos_guardados.get('responsable_mtto', '')
            
            # Cols 39-71: Hardware/Software (VERDE) - Vacíos por ahora
            for col in range(39, 72):
                ws.cell(row=nueva_fila, column=col).value = ''

            # Cargar datos básicos automáticos aunque sea guardado manual
            ws.cell(row=nueva_fila, column=64).value = detect_ip_local()

            # Cols 72-78: Mixtos (AZUL) - Vacíos por ahora
            for col in range(72, 79):
                ws.cell(row=nueva_fila, column=col).value = ''
            
            # Guardar
            wb.save(self.excel_path)
            wb.close()
            
            messagebox.showinfo(
                "Éxito",
                f"Equipo guardado exitosamente:\n\n" +
                f"Código: {next_codigo}\n" +
                f"Tipo: {datos_guardados.get('tipo_equipo', '')}\n" +
                f"Área: {datos_guardados.get('area_servicio', '')}\n\n" +
                f"Los datos de hardware se pueden agregar después con 'Recopilación Automática'."
            )
            
            # Limpiar formulario
            for widget in self.manual_widgets.values():
                try:
                    if hasattr(widget, 'winfo_exists') and widget.winfo_exists():
                        if isinstance(widget, ctk.CTkEntry):
                            widget.delete(0, 'end')
                        elif isinstance(widget, ctk.CTkComboBox):
                            widget.set('')
                        elif isinstance(widget, tk.StringVar):
                            widget.set('')
                except:
                    pass
            
            # Actualizar título del formulario con siguiente código
            next_consecutive_display = next_consecutivo + 1
            next_code_display = f"EQC-{next_consecutive_display:04d}"

            # Actualizar título en el frame del formulario
            if hasattr(self, 'form_title_label'):
                try:
                    if hasattr(self.form_title_label, 'winfo_exists') and self.form_title_label.winfo_exists():
                        self.form_title_label.configure(
                            text=f"Equipo: {next_code_display}"
                        )
                except Exception as e:
                    print(f"Error actualizando título: {e}")
            
        except Exception as e:
            messagebox.showerror(
                "Error al Guardar",
                f"Ocurrió un error al guardar los datos:\n\n{str(e)}"
            )
            import traceback
            traceback.print_exc()
    
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
                loaded_ok, load_msg = self.load_equipo_into_form_by_codigo(codigo, show_ready_message=False)
                if not loaded_ok:
                    messagebox.showerror("Error", load_msg)
                    return

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
        
        # 2. Restaurar títulos al SIGUIENTE equipo nuevo (con verificación)
        try:
            next_code = self.get_next_codigo()
        except:
            next_code = f"EQC-{self.current_row-1:04d}"
        
        # 2a. Restaurar label del frame (gris)
        if hasattr(self, 'equipo_form_frame'):
            try:
                if hasattr(self.equipo_form_frame, 'winfo_exists') and self.equipo_form_frame.winfo_exists():
                    self.equipo_form_frame.configure(
                        label_text=f"📝 EQUIPOS DE CÓMPUTO - Código: {next_code}"
                    )
            except:
                pass
        
        # 2b. Restaurar título verde
        if hasattr(self, 'form_title_label'):
            try:
                if hasattr(self.form_title_label, 'winfo_exists') and self.form_title_label.winfo_exists():
                    self.form_title_label.configure(
                        text=f"Equipo: {next_code}"
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
    
    def handle_save_equipo(self):
        """Wrapper para guardar equipo - detecta si es nuevo o actualización."""
        if hasattr(self, 'equipo_update_code') and self.equipo_update_code:
            # Modo actualización
            self.save_equipo_update()
        else:
            # Modo nuevo - guardar manual
            self.save_equipo_manual_only()
    
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
                'responsable_custodio', 'proceso', 'uso_sihos',
                'uso_office_basico', 'software_especializado', 'descripcion_software',
                'funcion_principal','horario_uso', 'estado_operativo', 'observaciones_tecnicas',
                'periodicidad_mtto', 'responsable_mtto'
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

            # Actualizar datos básicos de red/equipo en modo manual
            ws.cell(row=row, column=3, value=detect_hostname_local())
            ws.cell(row=row, column=64, value=detect_ip_local())
            
            wb.save(self.excel_path)
            wb.close()

            messagebox.showinfo("Éxito", f"✅ Equipo {codigo} actualizado correctamente")
            
            # Reseteo completo usando función unificada
            self.reset_after_update_equipos()
        
        except Exception as e:
            messagebox.showerror("Error", f"Error al actualizar equipo:\n{e}")
    
    def show_completion_message(self):
        """Mostrar mensaje de equipo completado."""
        consecutive = self.last_saved_consecutive
        code = self.last_saved_code
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

✓ Datos manuales: Guardados
✓ Datos automáticos: Guardados
✓ Datos mixtos: Guardados
✓ Excel actualizado correctamente
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
        self.show_manual_form_in_container()


    
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
            ("Observaciones", "observaciones", "entry"),
        ]
        
        for field_data in fields:
            if len(field_data) == 4:
                label, key, field_type, options = field_data
                widget = self.create_form_field_centered(scroll, label, key, field_type, options)
            else:
                label, key, field_type = field_data
                widget = self.create_form_field_centered(scroll, label, key, field_type, None)
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
                ws.cell(row=row, column=13, value=self.imp_widgets["observaciones"].get())
                
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
                ws.cell(row=next_row, column=13, value=self.imp_widgets["observaciones"].get())
                
                wb.save(self.excel_path)
                wb.close()
                
                messagebox.showinfo("Éxito", f"✅ Impresora guardada: IMP-{next_consecutive:04d}")
                
                # Detectar siguiente código y actualizar título
                next_code = self.detect_next_code("Impresoras y Escáneres", "IMP")
                self.imp_next_code = next_code
                self.imp_scroll.configure(label_text=f"🖨️ IMPRESORAS Y ESCÁNERES - Código: {next_code}")
                
                # Limpiar campos selectivamente (mantener área)
                campos_a_mantener = ['area',"codigo_asignado","modelo","serial", "area", "ubicacion", "estado"]
                
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
            ("Observaciones", "observaciones", "entry"),
        ]
        
        for field_data in fields:
            if len(field_data) == 4:
                label, key, field_type, options = field_data
                widget = self.create_form_field_centered(scroll, label, key, field_type, options)
            else:
                label, key, field_type = field_data
                widget = self.create_form_field_centered(scroll, label, key, field_type, None)
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
                ws.cell(row=row, column=10, value=self.per_widgets["observaciones"].get())
                
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
                ws.cell(row=next_row, column=10, value=self.per_widgets["observaciones"].get())
                
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
                
                self.per_widgets["observaciones"].delete(0, "end")
                self.per_widgets["observaciones"].insert(0, ws.cell(row=target_row, column=10).value or "")
                
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
            ("Observaciones", "observaciones", "entry"),
        ]
        
        for field_data in fields:
            if len(field_data) == 4:
                label, key, field_type, options = field_data
                widget = self.create_form_field_centered(scroll, label, key, field_type, options)
            else:
                label, key, field_type = field_data
                widget = self.create_form_field_centered(scroll, label, key, field_type, None)
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
                ws.cell(row=row, column=12, value=self.red_widgets["observaciones"].get())
                
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
                ws.cell(row=next_row, column=12, value=self.red_widgets["observaciones"].get())
                
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
                
                self.red_widgets["observaciones"].delete(0, "end")
                self.red_widgets["observaciones"].insert(0, ws.cell(row=target_row, column=12).value or "")
                
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
    
    def buscar_equipo_mantenimiento(self):
        """Buscar equipo y cargar datos para mantenimiento."""
        codigo = self.mtt_widgets["codigo_equipo"].get().strip().upper()
        
        if not codigo:
            messagebox.showerror("Error", "Ingresa un código de equipo")
            return
        
        try:
            wb = load_workbook(self.excel_path)
            ws = wb["Equipos de Cómputo"]
            
            found = False
            
            for row in range(2, 500):
                cell_codigo = ws.cell(row=row, column=2).value
                if cell_codigo and cell_codigo.upper() == codigo:
                    # Cargar datos
                    tipo = ws.cell(row=row, column=4).value or ""
                    marca = ws.cell(row=row, column=39).value or ""
                    modelo = ws.cell(row=row, column=40).value or ""
                    serial = ws.cell(row=row, column=41).value or ""
                    area = ws.cell(row=row, column=5).value or ""
                    periodicidad = ws.cell(row=row, column=37).value or ""
                    
                    # Actualizar campos readonly
                    for field_name, value in [
                        ("tipo_equipo_readonly", tipo),
                        ("marca_readonly", marca),
                        ("modelo_readonly", modelo),
                        ("serial_readonly", serial),
                        ("area_readonly", area),
                        ("periodicidad_readonly", periodicidad)
                    ]:
                        widget = self.mtt_widgets[field_name]
                        widget.configure(state="normal")
                        widget.delete(0, "end")
                        widget.insert(0, value)
                        widget.configure(state="readonly")
                    
                    found = True
                    break
            
            wb.close()
            
            if found:
                # Habilitar botón de guardar
                if hasattr(self, 'btn_save_mantenimiento'):
                    self.btn_save_mantenimiento.configure(state="normal")
                messagebox.showinfo("Éxito", f"✅ Equipo {codigo} cargado correctamente")
            else:
                messagebox.showerror("Error", f"No se encontró el equipo {codigo}")
        
        except Exception as e:
            messagebox.showerror("Error", f"Error al buscar equipo:\n{e}")
    
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
        
        # Campo de código con botón de búsqueda
        codigo_frame = ctk.CTkFrame(scroll, fg_color="white", corner_radius=8)
        codigo_frame.pack(fill="x", padx=40, pady=6)

        codigo_inner = ctk.CTkFrame(codigo_frame, fg_color="transparent")
        codigo_inner.pack(fill="x", padx=20, pady=10)

        codigo_inner.grid_columnconfigure(0, weight=3)
        codigo_inner.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            codigo_inner,
            text="Código Equipo *",
            font=("Segoe UI", 12, "bold"),
            anchor="w"
        ).grid(row=0, column=0, sticky="w", padx=(0, 10))

        entry_codigo = ctk.CTkEntry(
            codigo_inner,
            height=35,
            width=200,
            font=("Segoe UI", 11),
            placeholder_text="Ej: EQC-0001"
        )
        entry_codigo.grid(row=0, column=1, sticky="e", padx=(0, 10))

        btn_buscar_mtto = ctk.CTkButton(
            codigo_inner,
            text="🔍 BUSCAR Y CARGAR",
            command=self.buscar_equipo_mantenimiento,
            font=("Segoe UI", 11, "bold"),
            fg_color="#2196F3",
            hover_color="#1976D2",
            height=35,
            width=180
        )
        btn_buscar_mtto.grid(row=0, column=2, sticky="e")

        self.mtt_widgets["codigo_equipo"] = entry_codigo

        # Campos autocompletados (readonly)
        readonly_fields = [
            ("Tipo Equipo", "tipo_equipo_readonly"),
            ("Marca", "marca_readonly"),
            ("Modelo", "modelo_readonly"),
            ("Serial", "serial_readonly"),
            ("Área", "area_readonly"),
            ("Periodicidad", "periodicidad_readonly")
        ]

        for label_text, field_name in readonly_fields:
            field_frame = ctk.CTkFrame(scroll, fg_color="#E8F5E9", corner_radius=8)
            field_frame.pack(fill="x", padx=40, pady=6)
            
            inner_frame = ctk.CTkFrame(field_frame, fg_color="transparent")
            inner_frame.pack(fill="x", padx=20, pady=10)
            
            inner_frame.grid_columnconfigure(0, weight=6)
            inner_frame.grid_columnconfigure(1, weight=0, minsize=300)
            
            ctk.CTkLabel(
                inner_frame,
                text=label_text,
                font=("Segoe UI", 12, "bold"),
                anchor="w"
            ).grid(row=0, column=0, sticky="w", padx=(0, 30))
            
            widget = ctk.CTkEntry(
                inner_frame,
                height=35,
                width=300,
                font=("Segoe UI", 11),
                state="readonly",
                fg_color="#F0F0F0"
            )
            widget.grid(row=0, column=1, sticky="e")
            
            self.mtt_widgets[field_name] = widget

        # Resto de campos editables
        fields = [
            ("Tipo Mantenimiento *", "tipo", "combobox", TIPOS_MANTENIMIENTO_MTTO),
            ("Técnico Responsable *", "tecnico", "combobox", TECNICOS_RESPONSABLES),
            ("Descripción Actividades *", "descripcion", "combobox", ACTIVIDADES_MANTENIMIENTO),
            ("Repuestos/Insumos", "repuestos", "entry"),
            ("Estado Post-Mtto *", "estado_post", "combobox", ESTADO_POST_MTTO),
            ("Observaciones", "observaciones", "entry"),
        ]

        # Fecha mantenimiento - CAPTURAR y guardar en mtt_widgets
        fecha_widget = self.create_date_field_centered(scroll, "Fecha Mantenimiento *", "fecha_mtto")
        self.mtt_widgets["fecha_mtto"] = fecha_widget

        for field_data in fields:
            if len(field_data) == 4:
                label, key, field_type, options = field_data
                widget = self.create_form_field_centered(scroll, label, key, field_type, options)
            else:
                label, key, field_type = field_data
                widget = self.create_form_field_centered(scroll, label, key, field_type, None)
            self.mtt_widgets[key] = widget

        # Botón guardar (inicialmente deshabilitado)
        self.btn_save_mantenimiento = ctk.CTkButton(
            scroll,
            text="💾 GUARDAR MANTENIMIENTO",
            command=self.save_mantenimiento,
            font=("Segoe UI", 14, "bold"),
            fg_color=COLOR_VERDE_HOSPITAL,
            hover_color="#1F5039",
            height=50,
            state="disabled"  # Deshabilitado hasta buscar equipo
        )
        self.btn_save_mantenimiento.pack(pady=30)

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
            
            # Obtener fecha de mantenimiento
            fecha_mtto_str = self.get_date_value_safe(self.mtt_widgets["fecha_mtto"])
            
            # Calcular próximo mantenimiento basado en periodicidad
            periodicidad = self.mtt_widgets["periodicidad_readonly"].get().strip().upper()
            proximo_mtto = ""
            
            try:
                from datetime import datetime
                from dateutil.relativedelta import relativedelta
                
                fecha_mtto = datetime.strptime(fecha_mtto_str, "%Y-%m-%d")
                
                if "ANUAL" in periodicidad:
                    proximo_mtto = (fecha_mtto + relativedelta(years=1)).strftime("%Y-%m-%d")
                elif "SEMESTRAL" in periodicidad:
                    proximo_mtto = (fecha_mtto + relativedelta(months=6)).strftime("%Y-%m-%d")
                elif "TRIMESTRAL" in periodicidad:
                    proximo_mtto = (fecha_mtto + relativedelta(months=3)).strftime("%Y-%m-%d")
                elif "MENSUAL" in periodicidad:
                    proximo_mtto = (fecha_mtto + relativedelta(months=1)).strftime("%Y-%m-%d")
                else:
                    proximo_mtto = "No definido"
            except:
                proximo_mtto = "No calculado"
            
            # Guardar en Excel
            ws.cell(row=next_row, column=1, value=consecutive)
            ws.cell(row=next_row, column=2, value=self.mtt_widgets["codigo_equipo"].get())
            ws.cell(row=next_row, column=3, value=fecha_mtto_str)
            ws.cell(row=next_row, column=4, value=self.mtt_widgets["tipo"].get())
            ws.cell(row=next_row, column=5, value=self.mtt_widgets["tecnico"].get())
            ws.cell(row=next_row, column=6, value=self.mtt_widgets["descripcion"].get())
            ws.cell(row=next_row, column=7, value=self.mtt_widgets["repuestos"].get())
            ws.cell(row=next_row, column=8, value=self.mtt_widgets["estado_post"].get())
            ws.cell(row=next_row, column=9, value=proximo_mtto)
            ws.cell(row=next_row, column=10, value=self.mtt_widgets["observaciones"].get())
            
            wb.save(self.excel_path)
            wb.close()
            
            # Mensaje con próximo mantenimiento
            msg = f"✅ Mantenimiento registrado #{consecutive}\n\n"
            if proximo_mtto and proximo_mtto not in ["No definido", "No calculado"]:
                msg += f"📅 Próximo Mantenimiento: {proximo_mtto}"
            else:
                msg += "⚠️ No se pudo calcular próximo mantenimiento"
            
            messagebox.showinfo("Éxito", msg)
            
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
    
    def buscar_equipo_baja_mejorado(self):
        """Buscar equipo en inventarios y autocompletar datos (alias mejorado)."""
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
                tipo = ws.cell(row=target_row, column=4).value or "Computador"
                marca = ws.cell(row=target_row, column=39).value or ""
                modelo = ws.cell(row=target_row, column=40).value or ""
                serial = ws.cell(row=target_row, column=41).value or ""
                area = ws.cell(row=target_row, column=5).value or ""
                
            elif codigo.startswith("IMP-"):
                # Impresoras
                tipo = ws.cell(row=target_row, column=4).value or "Impresora"
                marca = ws.cell(row=target_row, column=5).value or ""
                modelo = ws.cell(row=target_row, column=6).value or ""
                serial = ws.cell(row=target_row, column=7).value or ""
                area = ws.cell(row=target_row, column=8).value or ""
                
            elif codigo.startswith("PER-"):
                # Periféricos
                tipo = ws.cell(row=target_row, column=4).value or "Periférico"
                marca = ws.cell(row=target_row, column=5).value or ""
                modelo = ws.cell(row=target_row, column=6).value or ""
                serial = ws.cell(row=target_row, column=7).value or ""
                area = ws.cell(row=target_row, column=8).value or ""
                
            elif codigo.startswith("RED-"):
                # Equipos de Red
                tipo = ws.cell(row=target_row, column=3).value or "Equipo de Red"
                marca = ws.cell(row=target_row, column=4).value or ""
                modelo = ws.cell(row=target_row, column=5).value or ""
                serial = ws.cell(row=target_row, column=6).value or ""
                area = ws.cell(row=target_row, column=7).value or ""
            
            wb.close()
            
            # Cargar datos en los widgets readonly
            for field_name, value in [
                ("tipo_readonly", tipo),
                ("marca_readonly", marca),
                ("modelo_readonly", modelo),
                ("serial_readonly", serial),
                ("area_readonly", area)
            ]:
                widget = self.baja_widgets[field_name]
                widget.configure(state="normal")
                widget.delete(0, "end")
                widget.insert(0, value)
                widget.configure(state="readonly")
            
            # Guardar información para actualizar después
            self.baja_origen_sheet = ws_name
            self.baja_origen_row = target_row
            
            # Habilitar botón de guardar
            if hasattr(self, 'btn_save_baja'):
                self.btn_save_baja.configure(state="normal")
            
            messagebox.showinfo("Éxito", f"✅ Datos cargados de {codigo}\n\nCompleta los campos de baja y guarda.")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al buscar equipo:\n{e}")
    
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
        
        # Campo de código con botón de búsqueda
        codigo_frame = ctk.CTkFrame(scroll, fg_color="white", corner_radius=8)
        codigo_frame.pack(fill="x", padx=40, pady=6)

        codigo_inner = ctk.CTkFrame(codigo_frame, fg_color="transparent")
        codigo_inner.pack(fill="x", padx=20, pady=10)

        codigo_inner.grid_columnconfigure(0, weight=3)
        codigo_inner.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            codigo_inner,
            text="Código Original *",
            font=("Segoe UI", 12, "bold"),
            anchor="w"
        ).grid(row=0, column=0, sticky="w", padx=(0, 10))

        entry_codigo = ctk.CTkEntry(
            codigo_inner,
            height=35,
            width=200,
            font=("Segoe UI", 11),
            placeholder_text="Ej: EQC-0001"
        )
        entry_codigo.grid(row=0, column=1, sticky="e", padx=(0, 10))

        btn_buscar_baja = ctk.CTkButton(
            codigo_inner,
            text="🔍 BUSCAR Y CARGAR",
            command=self.buscar_equipo_baja_mejorado,
            font=("Segoe UI", 11, "bold"),
            fg_color="#2196F3",
            hover_color="#1976D2",
            height=35,
            width=180
        )
        btn_buscar_baja.grid(row=0, column=2, sticky="e")

        self.baja_widgets["codigo_original"] = entry_codigo

        # Campos autocompletados (readonly)
        readonly_fields_baja = [
            ("Tipo", "tipo_readonly"),
            ("Marca", "marca_readonly"),
            ("Modelo", "modelo_readonly"),
            ("Serial", "serial_readonly"),
            ("Área", "area_readonly")
        ]

        for label_text, field_name in readonly_fields_baja:
            field_frame = ctk.CTkFrame(scroll, fg_color="#FFE8E8", corner_radius=8)
            field_frame.pack(fill="x", padx=40, pady=6)
            
            inner_frame = ctk.CTkFrame(field_frame, fg_color="transparent")
            inner_frame.pack(fill="x", padx=20, pady=10)
            
            inner_frame.grid_columnconfigure(0, weight=6)
            inner_frame.grid_columnconfigure(1, weight=0, minsize=300)
            
            ctk.CTkLabel(
                inner_frame,
                text=label_text,
                font=("Segoe UI", 12, "bold"),
                anchor="w"
            ).grid(row=0, column=0, sticky="w", padx=(0, 30))
            
            widget = ctk.CTkEntry(
                inner_frame,
                height=35,
                width=300,
                font=("Segoe UI", 11),
                state="readonly",
                fg_color="#F0F0F0"
            )
            widget.grid(row=0, column=1, sticky="e")
            
            self.baja_widgets[field_name] = widget

        # Resto de campos (fecha, motivo, etc.)
        # Fecha de baja - CAPTURAR y guardar en baja_widgets
        fecha_baja_widget = self.create_date_field_centered(scroll, "Fecha de Baja *", "fecha_baja")
        self.baja_widgets["fecha_baja"] = fecha_baja_widget

        fields = [
            ("Motivo Baja *", "motivo", "combobox", MOTIVOS_BAJA),
            ("Destino *", "destino", "combobox", DESTINOS_BAJA),
            ("Responsable Baja *", "responsable", "combobox", RESPONSABLES_BAJA),
            ("Observaciones", "observaciones", "entry"),
        ]

        for field_data in fields:
            if len(field_data) == 4:
                label, key, field_type, options = field_data
                widget = self.create_form_field_centered(scroll, label, key, field_type, options)
            else:
                label, key, field_type = field_data
                widget = self.create_form_field_centered(scroll, label, key, field_type, None)
            self.baja_widgets[key] = widget

        # Botón guardar (inicialmente deshabilitado)
        self.btn_save_baja = ctk.CTkButton(
            scroll,
            text="💾 REGISTRAR BAJA",
            command=self.save_baja,
            font=("Segoe UI", 14, "bold"),
            fg_color="#DC3545",
            hover_color="#A02828",
            height=50,
            state="disabled"  # Deshabilitado hasta buscar equipo
        )
        self.btn_save_baja.pack(pady=30)

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
            ws_baja.cell(row=next_row, column=2, value=self.baja_widgets["tipo_readonly"].get())
            ws_baja.cell(row=next_row, column=3, value=self.baja_widgets["marca_readonly"].get())
            ws_baja.cell(row=next_row, column=4, value=self.baja_widgets["modelo_readonly"].get())
            ws_baja.cell(row=next_row, column=5, value=self.baja_widgets["serial_readonly"].get())
            ws_baja.cell(row=next_row, column=6, value=self.get_date_value_safe(self.baja_widgets["fecha_baja"]))
            ws_baja.cell(row=next_row, column=7, value=self.baja_widgets["motivo"].get())
            ws_baja.cell(row=next_row, column=8, value=self.baja_widgets["destino"].get())
            ws_baja.cell(row=next_row, column=9, value=self.baja_widgets["responsable"].get())
            ws_baja.cell(row=next_row, column=10, value=self.baja_widgets["observaciones"].get())
            
            # Actualizar estado en inventario original (si fue autocompletado)
            if hasattr(self, 'baja_origen_sheet') and hasattr(self, 'baja_origen_row'):
                ws_origen = wb[self.baja_origen_sheet]
                
                # Actualizar estado operativo según el tipo
                if self.baja_origen_sheet == "Equipos de Cómputo":
                    # Columna 35 = Estado Operativo
                    ws_origen.cell(row=self.baja_origen_row, column=35, value="DADO DE BAJA")
                elif self.baja_origen_sheet == "Impresoras y Escáneres":
                    # Columna 12 = Estado Operativo
                    ws_origen.cell(row=self.baja_origen_row, column=12, value="DADO DE BAJA")
                elif self.baja_origen_sheet == "Periféricos":
                    # Columna 9 = Estado Operativo
                    ws_origen.cell(row=self.baja_origen_row, column=9, value="DADO DE BAJA")
                elif self.baja_origen_sheet == "Equipos de Red":
                    # Columna 11 = Estado Operativo
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
            
            # Limpiar campos selectivamente
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
