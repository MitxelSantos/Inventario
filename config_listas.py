# -*- coding: utf-8 -*-
"""
CONFIGURACIÓN DE LISTAS - Sistema de Inventario Tecnológico
================================================
Ing. Jose Miguel Santos Naranjo
Hospital Regional Alfonso Jaramillo Salazar

Todas las listas desplegables para los formularios
"""

# ============================================================================
# EQUIPOS DE CÓMPUTO
# ============================================================================

TIPOS_EQUIPO = ["Desktop", "Laptop", "All-in-One", "Tablet"]

# Alias para compatibilidad
TIPO_EQUIPO = TIPOS_EQUIPO


AREAS_SERVICIO = [
    "Urgencias",
    "UCI",
    "Hospitalización",
    "Quirófanos",
    "Consulta Externa",
    "Laboratorio Clínico",
    "Imágenes Diagnósticas",
    "Farmacia",
    "Facturación",
    "Admisiones",
    "Enfermería Piso 2",
    "Enfermería Piso 3",
    "Pediatría",
    "Ginecología",
    "Medicina Interna",
    "Cirugía",
    "Odontología",
    "Trabajo Social",
    "Contabilidad",
    "Recursos Humanos",
    "Sistemas",
    "Archivo",
    "Almacén",
    "Mantenimiento",
    "Seguridad",
    "Lavandería",
    "Dirección",
    "Subdirección",
]

# Alias para compatibilidad
AREAS_SERVICIOS = AREAS_SERVICIO

MACRO_PRO = ["ESTRATÉGICO", "MISIONAL", "APOYO", "EVALUACIÓN Y CONTROL"]

PROCESO_EST = ["Gerencia", "Planeación y calidad"]

PROCESO_MIS = ["Ambulatoria","Soporte diagnostico y terapeutico", "Urgencias", "Coordinación asistencial"]

PROCESO_APO = ["Contratación", "Talento humano", "Financiera", "Información y comunicaciones", "Ambiente fisico y tecnologia"]

PROCESO_EYC = ["Control interno", "Auditoria medica"]

SUBPRO_GER = ["Direccionamiento estrategico", "Asignación de recursos", "Evaluación y desempeño institucional", 
              "Rendición de cuentas", "Asesoria juridica"]

SUBPRO_PYC = ["Seguridad del paciente", "Epidemiologia", "Estadistica", "Formulación y seguimientos de planes, programas y proyectos",
              "Revision documental", "Productos no conformes"]

SUBPRO_AMB = ["Consulta externa", "Optometria", "Promoción y mantenimiento de la salud", "RIAS-MAITE", "Odontologia",
              "Atención al paciente hospitalizado", "Atención quirurgica"]

SUBPRO_SDT = ["Fisioterapia", "Laboratorio clinico", "Servicio Famarceutico", "Psicologia", "Trabajo Social", "Imagenes diagnosticas"]

USO_SIHOS = ["Local", "Web", "No usa"]

USO_SIFAX = ["Local", "Web", "No usa"]

USO_OFFICE_BASICO = ["Sí", "No"]

SOFTWARE_ESPECIALIZADO_OPCIONES = ["Sí", "No"]

SI_NO = ["Sí", "No"]

NIVEL_CRITICIDAD = ["CRÍTICO", "ALTO", "MEDIO", "BAJO"]

# Alias para compatibilidad
CRITICIDAD = NIVEL_CRITICIDAD

CLASIFICACION_CONFIDENCIALIDAD = [
    "CLASIFICADO",
    "RESERVADO",
    "CONFIDENCIAL",
    "INTERNO",
    "PÚBLICO",
]

# Alias para compatibilidad
CONFIDENCIALIDAD = CLASIFICACION_CONFIDENCIALIDAD

HORARIOS_USO = [
    "24/7",
    "Lunes a Viernes 7am-7pm",
    "Lunes a Viernes 7am-5pm",
    "Lunes a Sábado 7am-12m",
]

# Alias para compatibilidad
HORARIO_USO = HORARIOS_USO

ESTADOS_OPERATIVOS = [
    "Operativo - Óptimo",
    "Operativo - Regular",
    "Operativo - Deficiente",
    "Fuera de Servicio - Temporal",
    "En Reparación",
    "En Bodega",
    "Dado de Baja",
]

# Alias para compatibilidad
ESTADO_OPERATIVO = ESTADOS_OPERATIVOS

PERIODICIDAD_MTTO = ["Mensual", "Bimestral", "Trimestral", "Semestral", "Anual"]

TIPO_MANTENIMIENTO = ["Preventivo", "Correctivo", "Predictivo"]

# Alias para compatibilidad
TIPO_MTTO = TIPO_MANTENIMIENTO

MARCAS_EQUIPOS = [
    "HP",
    "Dell",
    "Lenovo",
    "Asus",
    "Acer",
    "Compaq",
    "Toshiba",
    "Samsung",
    "Apple",
    "MSI",
    "Otro",
]

SISTEMAS_OPERATIVOS = [
    "Windows 11 Pro",
    "Windows 10 Pro",
    "Windows 10 Home",
    "Windows 8.1",
    "Windows 7 Pro",
    "Linux Ubuntu",
    "Linux Debian",
    "macOS",
]

ARQUITECTURA_SO = ["x64", "x86"]

TIPOS_PROCESADOR = [
    "Intel Core i3",
    "Intel Core i5",
    "Intel Core i7",
    "Intel Core i9",
    "Intel Pentium",
    "Intel Celeron",
    "AMD Ryzen 3",
    "AMD Ryzen 5",
    "AMD Ryzen 7",
    "AMD Athlon",
    "Apple M1",
    "Apple M2",
]

CAPACIDADES_RAM = ["2", "4", "8", "16", "32", "64"]

TIPOS_DISCO = ["SSD", "HDD", "SSD + HDD", "NVMe"]

NAVEGADORES = ["Chrome", "Edge", "Firefox", "Opera", "Brave"]

VERSIONES_OFFICE = [
    "Office 365",
    "Office 2021",
    "Office 2019",
    "Office 2016",
    "Office 2013",
    "LibreOffice",
    "No instalado",
]

TIPOS_LICENCIA_OFFICE = ["OEM", "Retail", "Volumen", "Office 365 Suscripción"]

TIPOS_LICENCIA_WINDOWS = ["OEM", "Retail", "Volumen", "Digital"]

ESTADO_LICENCIA_WINDOWS = ["Activado", "Por Activar", "Prueba", "Sin Licencia"]

TIPOS_CONEXION = ["Cableado", "WiFi", "Ambos"]

ANTIVIRUS = [
    "Windows Defender",
    "Kaspersky",
    "ESET NOD32",
    "Avast",
    "AVG",
    "Bitdefender",
    "Norton",
    "McAfee",
    "Otro",
    "Ninguno",
]

ESTADO_ANTIVIRUS = ["Actualizado", "Desactualizado", "Desactivado", "Sin Antivirus"]

ACCESO_REMOTO = ["AnyDesk", "TeamViewer", "Chrome Remote Desktop", "Windows RDP", "No"]

TIPO_USUARIO_LOCAL = ["Administrador", "Estándar", "Invitado"]

# Alias para compatibilidad
OPCIONES_TIPO_USUARIO = TIPO_USUARIO_LOCAL
OPCIONES_CIFRADO_DISCO = SI_NO
OPCIONES_ESTADO_ANTIVIRUS = ESTADO_ANTIVIRUS

# ============================================================================
# IMPRESORAS Y ESCÁNERES
# ============================================================================

TIPOS_IMPRESORA = [
    "Impresora Multifuncional",
    "Impresora Láser",
    "Impresora de Inyección",
    "Impresora Térmica",
    "Impresora de Punto",
    "Escáner de Cama Plana",
    "Escáner Portátil",
]

MARCAS_IMPRESORA = [
    "HP",
    "Canon",
    "Epson",
    "Brother",
    "Samsung",
    "Xerox",
    "Lexmark",
    "Ricoh",
    "Kyocera",
    "Otro",
]

FUNCIONES_IMPRESORA = ["Impresión", "Digitalización", "Copia", "Fax", "Multifunción"]

ESTADOS_IMPRESORA = [
    "Operativo",
    "Fuera de Servicio",
    "En Reparación",
    "Sin Tóner/Tinta",
    "En Bodega",
    "Dado de Baja",
]

# ============================================================================
# PERIFÉRICOS
# ============================================================================

TIPOS_PERIFERICO = [
    "Mouse",
    "Teclado",
    "Monitor",
    "Webcam",
    "Diadema/Audífonos",
    "Parlantes",
    "Micrófono",
    "UPS",
    "Regulador",
    "Disco Externo",
    "USB Flash",
    "Lector Biométrico",
    "Lector de Código de Barras",
    "Televisor",
    "Proyector",
    "Otro",
]

MARCAS_PERIFERICO = [
    "Logitech",
    "HP",
    "Dell",
    "Microsoft",
    "Genius",
    "Razer",
    "Corsair",
    "Samsung",
    "LG",
    "Seagate",
    "Western Digital",
    "Kingston",
    "SanDisk",
    "APC",
    "Otro",
]

ESTADOS_PERIFERICO = [
    "Operativo",
    "Fuera de Servicio",
    "En Reparación",
    "En Bodega",
    "Dado de Baja",
]

# ============================================================================
# EQUIPOS DE RED
# ============================================================================

TIPOS_EQUIPO_RED = [
    "Switch",
    "Router",
    "Access Point",
    "Firewall",
    "Servidor Fisico",
    "Servidor Virtual",
    "Modem",
    "Patch Panel",
    "Rack",
    "Otro",
]

MARCAS_RED = [
    "Cisco",
    "TP-Link",
    "Ubiquiti",
    "Mikrotik",
    "D-Link",
    "Netgear",
    "Huawei",
    "HPE",
    "Dell",
    "Otro",
]

UBICACIONES_RED = [
    "Rack Principal",
    "Datacenter",
    "Sotano",
    "Piso 1",
    "Piso 2",
    "Piso 3",
    "Piso 4",
    "Piso 5",
    "Servidor",
    "Oficina Sistemas",
    "Otro",
]

ESTADOS_RED = [
    "Operativo",
    "Fuera de Servicio",
    "En Mantenimiento",
    "En Bodega",
    "Dado de Baja",
]

# ============================================================================
# MANTENIMIENTOS
# ============================================================================

TIPOS_MANTENIMIENTO_MTTO = [
    "Preventivo",
    "Correctivo",
    "Predictivo",
    "Actualización",
    "Instalación",
]

TECNICOS_RESPONSABLES = ["Eduar Cortez", "Heber Valero", "Tecnico Externo", "Otro"]

# Alias para compatibilidad
RESPONSABLE_MTTO = TECNICOS_RESPONSABLES

ACTIVIDADES_MANTENIMIENTO = [
    "Limpieza general",
    "Actualización SO",
    "Actualización drivers",
    "Cambio disco duro",
    "Cambio RAM",
    "Cambio fuente poder",
    "Instalación software",
    "Revisión antivirus",
    "Optimización sistema",
    "Formateo",
    "Backup",
    "Reparación hardware",
    "Otro",
]

ESTADO_POST_MTTO = [
    "Operativo",
    "Requiere seguimiento",
    "Fuera de servicio",
    "Pendiente repuesto",
    "Dado de Baja",
    "En Bodega",
]

# ============================================================================
# EQUIPOS DADOS DE BAJA
# ============================================================================

MOTIVOS_BAJA = [
    "Obsolescencia",
    "Daño irreparable",
    "Actualización tecnológica",
    "Fin de vida útil",
    "Robo/Extravío",
    "Siniestro",
    "Cambio de política",
    "Otro",
]

DESTINOS_BAJA = [
    "Reciclaje",
    "Donación",
    "Destrucción",
    "Venta",
    "Almacenamiento temporal",
    "Devolución proveedor",
    "Reutilización interna",
    "Entrega a bodega",
    "Otro",
]

RESPONSABLES_BAJA = [
    "Coordinador TI",
    "Jefe Almacén",
    "Director Administrativo",
    "Otro",
]

# ============================================================================
# VALIDACIONES
# ============================================================================


def validar_ip(ip):
    """Validar formato de dirección IP."""
    import re

    pattern = r"^(\d{1,3}\.){3}\d{1,3}$"
    if re.match(pattern, ip):
        octetos = ip.split(".")
        return all(0 <= int(octeto) <= 255 for octeto in octetos)
    return False


def validar_mac(mac):
    """Validar formato de dirección MAC."""
    import re

    pattern = r"^([0-9A-Fa-f]{2}[:-]){5}([0-9A-Fa-f]{2})$"
    return bool(re.match(pattern, mac))


def validar_fecha(fecha):
    """Validar formato de fecha YYYY-MM-DD."""
    import re

    pattern = r"^\d{4}-\d{2}-\d{2}$"
    return bool(re.match(pattern, fecha))
