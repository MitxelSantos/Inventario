# -*- coding: utf-8 -*-
"""
TEXTOS RESUMIDOS PARA TOOLTIPS - CUESTIONARIO DE CLASIFICACIÓN
================================================================
Mapeo de textos cortos (label) a preguntas completas (tooltip)
"""

# ============================================================================
# CONFIDENCIALIDAD (9 preguntas)
# ============================================================================

CONF_LABELS = [
    "1. Información pública",
    "2. Información interna",
    "3. Datos de identidad",
    "4. Datos de contacto",
    "5. Información técnica TI",
    "6. Datos personales sensibles",
    "7. Datos financieros",
    "8. Información secreta",
    "9. Información confidencial de negocio"
]

CONF_TOOLTIPS = [
    "¿El activo contiene información pública de la entidad, como portales web, formularios o documentos destinados al público en general?",
    "¿El activo contiene información interna o de uso restringido de la compañía, como portales internos, formularios o documentos no publicados al público?",
    "¿El activo contiene datos de identidad de clientes o trabajadores, como nombres, números de identificación, fechas de nacimiento, género, lugar de nacimiento u otros identificadores similares?",
    "¿El activo contiene datos de contacto de clientes o trabajadores, como direcciones de vivienda, direcciones de trabajo, correos electrónicos, números de teléfono o direcciones postales?",
    "¿El activo contiene información técnica confidencial de TI, desarrollo o seguridad, como diagramas de arquitectura, inventarios de activos, configuraciones, reglas de seguridad o documentación técnica interna?",
    "¿El activo contiene datos personales sensibles de clientes o trabajadores, como etnia, identidad de género, identidad cultural, religión, ideología, afinidad política, antecedentes legales, estado migratorio, orientación sexual, historial médico, discapacidades, datos biométricos o genéticos?",
    "¿El activo contiene datos financieros o de nómina de clientes o trabajadores, como información de cuentas bancarias, información salarial, tarjetas de crédito u otros datos de medios de pago?",
    "¿El activo contiene información secreta como contraseñas, claves criptográficas, certificados digitales, tokens de acceso u otros secretos utilizados para autenticación o cifrado?",
    "¿El activo contiene información confidencial de negocio, como planes estratégicos, análisis internos, contratos, acuerdos con terceros, información no pública de clientes/proveedores o propiedad intelectual (por ejemplo, código fuente, algoritmos, diseños o know-how)?"
]

# ============================================================================
# INTEGRIDAD (3 preguntas)
# ============================================================================

INT_LABELS = [
    "1. Persistencia de información",
    "2. Información en tránsito",
    "3. Información en proceso"
]

INT_TOOLTIPS = [
    "Si el activo se ve comprometido, ¿podría la persistencia de la información verse afectada?",
    "Si el activo se ve comprometido, ¿podría la información en tránsito verse afectada?",
    "Si el activo se ve comprometido, ¿podría la información en proceso ser afectada?"
]

# ============================================================================
# CRITICIDAD (6 preguntas)
# ============================================================================

CRIT_LABELS = [
    "1. Afecta trabajadores",
    "2. Afecta usuarios externos",
    "3. Afecta operación principal",
    "4. Afecta procesos de apoyo",
    "5. Afecta TI/Seguridad",
    "6. Incumplimiento legal"
]

CRIT_TOOLTIPS = [
    "¿La indisponibilidad del activo afectaría de forma relevante a los trabajadores de la entidad en su operación diaria?",
    "¿La indisponibilidad del activo afectaría de forma relevante a los usuarios externos de la entidad (ciudadanos, clientes u otros)?",
    "Si el activo se encuentra indisponible, ¿se vería afectada la operación principal (misional) de la organización?",
    "¿La indisponibilidad del activo afectaría procesos de soporte clave, como finanzas, ventas/comercial, recursos humanos, logística u otros procesos administrativos críticos?",
    "¿La indisponibilidad del activo afectaría la operación de TI o de Seguridad de la Información, incluyendo el monitoreo y la gestión de otros activos de la organización?",
    "¿La indisponibilidad del activo podría generar incumplimiento de obligaciones legales, regulatorias o contractuales (incluyendo SLAs con clientes o entidades externas)?"
]
