# 🏥 Dashboard de Inventario Tecnológico v2.0

Dashboard interactivo profesional para la gestión y análisis del inventario tecnológico del **Hospital Regional Alfonso Jaramillo Salazar** - Líbano, Tolima.

---

## 🚀 Inicio Rápido

```bash
# 1. Instalar dependencias
pip install -r requirements.txt

# 2. Ejecutar (opción 1)
python app.py

# O ejecutar (opción 2)
streamlit run app.py

# 3. Abrir en navegador
http://localhost:8501
```

---

## ✨ Características Principales

### 📊 6 Módulos Principales
1. **💻 Equipos de Cómputo** - Análisis completo con 5 tabs:
   - Dashboard General (métricas, alertas, distribución)
   - Análisis de Seguridad (gauges, cifrado, antivirus)
   - Análisis Financiero (inversión, scatter plots, top equipos)
   - Ciclo de Vida (antigüedad, obsolescencia)
   - Tendencias (adquisiciones temporales)

2. **🖨️ Impresoras y Escáneres** - Distribución por tipo, función, áreas
3. **🖱️ Periféricos** - Inventario completo por tipo y estado
4. **🌐 Equipos de Red** - Análisis de puertos, ubicación, marca
5. **🔧 Mantenimientos** - Próximos 30 días, meta mensual, costos
6. **📦 Equipos Dados de Baja** - Motivos, destinos, análisis

### 🔔 Sistema de Alertas
- ✅ Antivirus vencidos/próximos a vencer (30 días)
- ✅ Licencias Windows sin activar
- ✅ Mantenimientos urgentes (código de colores)
- ✅ Meta mensual de mantenimientos preventivos

### 📈 Analíticas Avanzadas
- ✅ 15+ tipos de gráficos interactivos (Plotly)
- ✅ Gauges de seguridad
- ✅ Scatter plots financieros
- ✅ Análisis de tendencias temporales
- ✅ Distribuciones interactivas

### 📤 Exportación
- Excel (.xlsx) con formato profesional
- CSV (.csv) con UTF-8
- Timestamp automático en archivos

---

## 📁 Estructura del Proyecto

```
dashboard_hospital/
├── app.py                          # Aplicación principal (todo en 1)
├── utils.py                        # Utilidades consolidadas
├── requirements.txt                # Dependencias
├── README.md                       # Este archivo
├── .gitignore                      # Configuración Git
├── inventario_hospital_test.xlsx  # Datos de prueba
├── views/                          # Vistas modulares
│   ├── equipos_computo.py         # Vista más completa
│   ├── impresoras.py
│   ├── perifericos.py
│   ├── red.py
│   ├── mantenimientos.py
│   └── dados_baja.py
└── assets/                         # Logo opcional
    └── README.txt
```

**Total: 11 archivos Python | ~1,200 líneas de código**

---

## 🔧 Personalización

### 🎨 Cambiar Colores Institucionales
Edita `utils.py` línea 19:
```python
COLORS = {
    'primary': '#2D6A4F',    # Verde oscuro
    'secondary': '#52B788',  # Verde claro
    'success': '#40916C',    # Verde éxito
    'warning': '#FFC107',    # Amarillo
    'danger': '#DC3545',     # Rojo
    'info': '#0077B6',       # Azul
}
```

### 🖼️ Agregar Logo
1. Coloca tu logo en: `assets/logo.png`
2. Formatos: PNG, JPG o SVG
3. Tamaño recomendado: 200x200px o mayor
4. Si no existe logo, se muestra placeholder automático

### 📊 Cambiar Meta de Mantenimientos
Edita `views/mantenimientos.py` línea 96:
```python
META = 20  # Cambiar a tu meta mensual
```

### 🔍 Agregar Filtros
Cada vista tiene un selector interactivo. Ejemplo en `views/equipos_computo.py`:
```python
agrupar_por = st.selectbox("Agrupar por:", 
    ["Área / Servicio", "Proceso", "Nivel de Criticidad", ...])
```

---

## 📊 Datos Requeridos

### Hojas del Excel
El archivo debe contener estas hojas:
1. **Equipos de Cómputo**
2. **Impresoras y Escáneres**
3. **Periféricos**
4. **Equipos de Red**
5. **Mantenimientos**
6. **Equipos Dados de Baja**

### Archivos Soportados
El dashboard busca automáticamente (en orden):
- `inventario_hospital_test.xlsx`
- `inventario_hospital_v1.xlsx`
- `inventario_hospital.xlsx`

---

## 🛠️ Solución de Problemas

### Error: ModuleNotFoundError
```bash
pip install -r requirements.txt
```

### Error: Puerto 8501 ocupado
```bash
streamlit run app.py --server.port 8502
```

### Error: No encuentra Excel
```bash
# Verificar archivo
ls -la *.xlsx

# Asegúrate de tener alguno de:
# - inventario_hospital_test.xlsx
# - inventario_hospital_v1.xlsx
# - inventario_hospital.xlsx
```

### Error: Datos no se actualizan
- Presiona `C` en el navegador para limpiar cache
- O reinicia el dashboard

---

## 📚 Dependencias

### Críticas (6)
- **streamlit** 1.39.0 - Framework del dashboard
- **pandas** 2.2.3 - Procesamiento de datos
- **plotly** 5.24.1 - Visualizaciones interactivas
- **openpyxl** 3.1.5 - Lectura/escritura Excel
- **xlsxwriter** 3.2.0 - Exportación optimizada
- **python-dateutil** 2.9.0 - Manejo de fechas

---

## 🎨 Características Estéticas

### Diseño Profesional
- ✅ Fuente Google Inter
- ✅ Gradientes CSS modernos
- ✅ Sombras y efectos hover
- ✅ Tabs con animaciones
- ✅ Responsive design (móvil/tablet/desktop)
- ✅ Tarjetas con iconos emoji
- ✅ Colores institucionales personalizables

### UX/UI
- ✅ Navegación intuitiva con tabs
- ✅ Métricas con deltas visuales
- ✅ Gráficos interactivos Plotly
- ✅ Alertas con código de colores
- ✅ Expandibles para detalles
- ✅ Tooltips informativos

---

## 💡 Características Técnicas

### Arquitectura
- ✅ Modular y escalable
- ✅ Separación de concerns
- ✅ Single app.py (no necesita run_dashboard.py)
- ✅ Vistas independientes
- ✅ Utilidades consolidadas

### Rendimiento
- ✅ Cache de datos (5 minutos)
- ✅ Procesamiento optimizado
- ✅ Lazy loading de vistas
- ✅ Hot reload automático

### Calidad de Código
- ✅ Docstrings en funciones
- ✅ Type hints
- ✅ Manejo de errores robusto
- ✅ PEP 8 compliance

---

## 📈 Roadmap Futuro (Opcional)

### Corto Plazo
- [ ] Filtros por fecha en todas las vistas
- [ ] Exportación PDF de reportes
- [ ] Modo oscuro

### Mediano Plazo
- [ ] Predicción de mantenimientos con ML
- [ ] Dashboard de KPIs ejecutivo
- [ ] Integración con ERP

### Largo Plazo
- [ ] App móvil nativa
- [ ] API REST
- [ ] Notificaciones automáticas

---

## 👨‍💻 Desarrollador

**Ing. Jose Miguel Santos Naranjo**  
Área de Sistemas  
Hospital Regional Alfonso Jaramillo Salazar  
Líbano, Tolima, Colombia

© 2025 - Todos los derechos reservados

---

## 📝 Notas Importantes

- Los datos se cachean por 5 minutos para mejor rendimiento
- El dashboard recarga automáticamente al modificar código
- Para detener: presiona `Ctrl+C` en la terminal
- El archivo app.py puede ejecutarse directamente con `python app.py`
- No necesitas `run_dashboard.py` separado, todo está en `app.py`

---

## 🤝 Soporte

Para soporte técnico o consultas:
- Contacta al Área de Sistemas del Hospital
- Desarrollador: Ing. Jose Miguel Santos Naranjo

---

**Desarrollado con ❤️ para mejorar la gestión tecnológica hospitalaria**

🏥 Hospital Regional Alfonso Jaramillo Salazar • Líbano, Tolima
