# 📘 GUÍA DE USO - ESTRUCTURA DEL EXCEL

## Hospital Regional Alfonso Jaramillo Salazar
### Sistema de Inventario Tecnológico V1

---

## 📋 ESTRUCTURA GENERAL

El archivo `inventario_hospital_v1.xlsx` contiene **6 hojas** organizadas para gestionar todo el inventario tecnológico del hospital:

1. **Equipos de Cómputo** (61 columnas)
2. **Impresoras y Escáneres** (15 columnas)
3. **Periféricos** (11 columnas)
4. **Equipos de Red** (14 columnas)
5. **Mantenimientos** (10 columnas - SIN costo)
6. **Equipos Dados de Baja** (10 columnas)

---

## 💻 HOJA 1: EQUIPOS DE CÓMPUTO

### **Total: 61 Columnas**

#### **Columnas 1-3: IDENTIFICACIÓN**
| Columna | Nombre | Tipo | Descripción |
|---------|--------|------|-------------|
| A | N° Consecutivo | Numérico | Número correlativo (1, 2, 3...) |
| B | Código Inventario | Texto | Código único formato EQC-XXXX (EQC-0001, EQC-0142) |
| C | Nombre Equipo | Texto | Nombre del computador detectado automáticamente |

#### **Columnas 4-27: CAMPOS NARANJAS (Datos Manuales - 24 campos)**

**Información Administrativa:**
- D: Tipo de Equipo (Desktop, Laptop, All-in-One, etc.)
- E: Área / Servicio (Urgencias, Laboratorio, etc.)
- F: Ubicación Específica (Consultorio 1, Piso 2, etc.)
- G: Responsable / Custodio (Nombre del usuario asignado)
- H: Proceso (Asistencial, Administrativo, Apoyo)

**Uso del Equipo:**
- I: Uso SIHOS (Sí/No)
- J: Uso SIFAX (Sí/No)
- K: Uso Office Básico (Sí/No)
- L: Software Especializado (Sí/No)
- M: Descripción Software (Detalles si aplica)
- N: Función Principal (Descripción del uso principal)

**Clasificación Normativa:**
- O: Nivel Criticidad (Crítico/Alto/Medio/Bajo - según MinTIC PETI)
- P: Clasificación Confidencialidad (Reservado/Confidencial/Público - según MinSalud)
- Q: Horario Uso (24/7, Laboral, Variable)

**Estado y Mantenimiento:**
- R: Estado Operativo (Operativo, En Mantenimiento, Fuera de Servicio, DADO DE BAJA)
- S: Fecha Adquisición (YYYY-MM-DD)
- T: Valor Adquisición (COP)
- U: Fecha Vencimiento Garantía (YYYY-MM-DD)
- V: Observaciones Técnicas (Notas relevantes)
- W: Fecha Expiración Antivirus (YYYY-MM-DD)
- X: Periodicidad Mantenimiento (Mensual, Trimestral, Semestral, Anual)
- Y: Responsable Mantenimiento (Técnico asignado)
- Z: Último Mantenimiento (YYYY-MM-DD)
- AA: Tipo Último Mantenimiento (Preventivo, Correctivo, Actualización)

#### **Columnas 28-48: CAMPOS VERDES (Detección Automática - 21 campos)**

**Hardware Básico:**
- AB: Marca (Dell, HP, Lenovo, etc.)
- AC: Modelo (Modelo específico del equipo)
- AD: Serial (Número de serie del equipo)
- AE: Sistema Operativo (Windows 10, Windows 11, etc.)
- AF: Arquitectura SO (64 bits / 32 bits)

**Procesamiento y Memoria:**
- AG: Procesador (Modelo del CPU)
- AH: RAM (GB) (Memoria RAM instalada)
- AI: Almacenamiento (GB) (Disco primario)
- AJ: Tipo Disco (HDD / SSD - disco primario)

**Software Office:**
- AK: Uso Navegador Web (Sí/No)
- AL: Versión Office (Office 2016/2019/365)
- AM: Licencia Office (Retail/Volume/OEM)
- AN: Uso Teams (Sí/No)
- AO: Uso Outlook (Sí/No)

**Licencias Windows:**
- AP: Licencia Windows (Retail/OEM/Volume/Enterprise)
- AQ: Key Windows (Últimos 5 dígitos)
- AR: Estado Licencia Windows (Activado/No activado)

**Red:**
- AS: Dirección IP (192.168.X.X)
- AT: Tipo Conexión (Ethernet/Wi-Fi)

**Seguridad:**
- AU: Antivirus Instalado (Windows Defender, etc.)
- AV: Última Actualización Windows (Fecha)
- AW: Windows Update Activo (Sí/No)

#### **Columnas 49-60: CAMPOS AZULES (Mixtos con Validación - 12 campos)**

**Disco Secundario (5 campos):**
- AX: Almacenamiento Secundario (GB) (Capacidad del segundo disco o "No tiene")
- AY: Tipo Disco Secundario (HDD/SSD/"No tiene")
- AZ: Serial Disco Secundario (Número de serie o "No tiene")
- BA: Marca Disco Secundario (Fabricante o "No tiene")
- BB: Modelo Disco Secundario (Modelo específico o "No tiene")

**Infraestructura y Seguridad (7 campos):**
- BC: Switch / Puerto (Identificación del puerto de red)
- BD: VLAN Asignada (VLAN configurada)
- BE: ID AnyDesk (ID de acceso remoto)
- BF: Otro Acceso Remoto (TeamViewer, etc.)
- BG: Estado Antivirus (Actualizado, Desactualizado, Desactivado)
- BH: Cifrado de Disco (BitLocker activado/No activado)
- BI: Tipo Usuario Local (Administrador/Estándar/Restringido)

#### **Columna 61: CAMPO BLANCO (Calculado - 1 campo)**
- BJ: Antigüedad (Años) (Calculado automáticamente desde fecha de adquisición)

---

## 🖨️ HOJA 2: IMPRESORAS Y ESCÁNERES

### **Total: 15 Columnas**

| Col | Nombre | Descripción |
|-----|--------|-------------|
| A | N° Consecutivo | Número correlativo |
| B | Código Inventario | Formato IMP-XXXX (IMP-0001) |
| C | Código Asignado | Código adicional si existe |
| D | Tipo | Impresora Láser, Multifuncional, Escáner, etc. |
| E | Marca | HP, Canon, Epson, etc. |
| F | Modelo | Modelo específico |
| G | Serial | Número de serie |
| H | Área | Área donde está ubicada |
| I | Ubicación | Ubicación específica |
| J | Función | Uso principal |
| K | Dirección IP | IP asignada (si aplica) |
| L | Estado | Operativo, En Mantenimiento, DADO DE BAJA |
| M | Fecha Adquisición | YYYY-MM-DD |
| N | Valor | Costo de adquisición |
| O | Observaciones | Notas adicionales |

---

## 🖱️ HOJA 3: PERIFÉRICOS

### **Total: 11 Columnas**

| Col | Nombre | Descripción |
|-----|--------|-------------|
| A | N° Consecutivo | Número correlativo |
| B | Código Inventario | Formato PER-XXXX (PER-0001) |
| C | Código Asignado | Código adicional si existe |
| D | Tipo | Mouse, Teclado, Monitor, Webcam, etc. |
| E | Marca | Logitech, HP, Dell, etc. |
| F | Modelo | Modelo específico |
| G | Serial | Número de serie |
| H | Área | Área donde está asignado |
| I | Estado | Operativo, Dañado, DADO DE BAJA |
| J | Fecha Adquisición | YYYY-MM-DD |
| K | Observaciones | Notas adicionales |

---

## 🌐 HOJA 4: EQUIPOS DE RED

### **Total: 14 Columnas**

| Col | Nombre | Descripción |
|-----|--------|-------------|
| A | N° Consecutivo | Número correlativo |
| B | Código Inventario | Formato RED-XXXX (RED-0001) |
| C | Tipo | Switch, Router, Access Point, etc. |
| D | Marca | Cisco, TP-Link, Ubiquiti, etc. |
| E | Modelo | Modelo específico |
| F | Serial | Número de serie |
| G | Dirección IP | IP asignada |
| H | N° Puertos | Cantidad de puertos |
| I | Ubicación | Ubicación física |
| J | Área | Área que cubre |
| K | Estado | Operativo, En Mantenimiento, DADO DE BAJA |
| L | Fecha Adquisición | YYYY-MM-DD |
| M | Valor | Costo de adquisición |
| N | Observaciones | Notas adicionales |

---

## 🔧 HOJA 5: MANTENIMIENTOS

### **Total: 10 Columnas (SIN COSTO)**

| Col | Nombre | Descripción |
|-----|--------|-------------|
| A | N° Consecutivo | Número de mantenimiento |
| B | Código Equipo | Código del equipo (EQC-XXXX, IMP-XXXX, etc.) |
| C | Fecha Mantenimiento | YYYY-MM-DD |
| D | Tipo | Preventivo, Correctivo, Actualización |
| E | Técnico Responsable | Nombre del técnico |
| F | Descripción Actividades | Detalle del trabajo realizado |
| G | Repuestos/Insumos | Materiales utilizados |
| H | Estado Post-Mtto | Operativo, Requiere Seguimiento, Fuera de Servicio |
| I | Próximo Mantenimiento | YYYY-MM-DD (fecha programada) |
| J | Observaciones | Notas adicionales |

**NOTA:** La columna "Costo" fue eliminada ya que el mantenimiento es interno (técnicos del hospital + materiales disponibles).

---

## 📦 HOJA 6: EQUIPOS DADOS DE BAJA

### **Total: 10 Columnas**

| Col | Nombre | Descripción |
|-----|--------|-------------|
| A | Código Original | Código del equipo dado de baja |
| B | Tipo | Tipo de equipo |
| C | Marca | Marca |
| D | Modelo | Modelo |
| E | Serial | Número de serie |
| F | Fecha Baja | YYYY-MM-DD |
| G | Motivo | Obsolescencia, Daño irreparable, etc. |
| H | Destino Final | Reciclaje, Donación, Almacenamiento, etc. |
| I | Responsable | Quién autorizó la baja |
| J | Observaciones | Notas adicionales |

**IMPORTANTE:** Al dar de baja un equipo, su "Estado Operativo" en la hoja original se actualiza automáticamente a "DADO DE BAJA".

---

## 🔢 FORMATO DE CÓDIGOS

Todos los códigos siguen el formato de **4 DÍGITOS**:

- **Equipos de Cómputo:** EQC-0001, EQC-0002, ..., EQC-9999
- **Impresoras:** IMP-0001, IMP-0002, ..., IMP-9999
- **Periféricos:** PER-0001, PER-0002, ..., PER-9999
- **Equipos de Red:** RED-0001, RED-0002, ..., RED-9999

---

## 💡 NOTAS IMPORTANTES

### **Detección Automática de Disco Secundario**

El sistema detecta automáticamente si el equipo tiene un segundo disco duro:

1. **Si NO tiene disco secundario:** Todas las columnas se llenan con "No tiene"
2. **Si tiene disco secundario:** El sistema detecta:
   - Capacidad en GB
   - Tipo (HDD o SSD)
   - Serial del disco
   - Marca del fabricante
   - Modelo específico

3. **Validación:** Los datos se muestran en la ventana de validación mixta para confirmar o corregir

### **Mantenimiento Sin Costo**

El mantenimiento se registra sin campo de costo porque:
- Los técnicos son personal interno del hospital
- Los materiales y repuestos están disponibles en inventario
- No se generan costos adicionales por servicio

### **Actualización Automática de Estados**

Cuando se da de baja un equipo:
1. Se crea el registro en "Equipos Dados de Baja"
2. El sistema actualiza automáticamente el "Estado Operativo" a "DADO DE BAJA" en la hoja original
3. Funciona para los 4 tipos de equipos (Cómputo, Impresoras, Periféricos, Red)

### **Campos Obligatorios**

**En Equipos de Cómputo (campos naranjas):**
- Tipo de Equipo
- Área / Servicio
- Ubicación Específica
- Responsable / Custodio
- Proceso
- Uso SIHOS
- Estado Operativo

**En otros inventarios:**
- Los campos marcados con asterisco (*) en el formulario

---

## 📊 COLORES EN EL EXCEL

Los encabezados de todas las hojas usan el **color verde institucional del hospital** (#2F5233) para mantener la identidad visual.

---

## 🔄 RESPALDO Y VERSIONES

**Recomendaciones:**
1. Mantener copias de respaldo diarias
2. Usar control de versiones en el nombre del archivo
3. No modificar manualmente la estructura de columnas
4. Siempre usar el sistema para ingresar datos

---

## 📞 SOPORTE TÉCNICO

Para dudas o problemas con el sistema:
- **IT Coordinator:** Jose
- **Hospital:** Regional Alfonso Jaramillo Salazar
- **Ubicación:** Líbano, Tolima, Colombia

---

**Versión:** 1.0 - Diciembre 2025  
**Última actualización:** Incluye disco secundario y optimizaciones
