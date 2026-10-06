#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Dashboard de Inventario Tecnológico v2.0
Hospital Regional Alfonso Jaramillo Salazar

Desarrollado por: Ing. Jose Miguel Santos Naranjo
© 2025

INICIO RÁPIDO:
    python app.py
    
O con Streamlit:
    streamlit run app.py
"""

import streamlit as st
import pandas as pd
import os
import sys
from pathlib import Path

# Verificar dependencias antes de importar
def check_dependencies():
    """Verificar que las dependencias estén instaladas."""
    missing = []
    for dep in ['streamlit', 'pandas', 'plotly', 'openpyxl']:
        try:
            __import__(dep)
        except ImportError:
            missing.append(dep)
    
    if missing:
        print(f"\n❌ Faltan dependencias: {', '.join(missing)}")
        print("💡 Instala con: pip install -r requirements.txt\n")
        sys.exit(1)

check_dependencies()

# Importar módulos del proyecto
from utils import configure_page, COLORS, load_all_sheets
from views.equipos_computo import show_equipos_computo
from views.impresoras import show_impresoras
from views.perifericos import show_perifericos
from views.red import show_red
from views.mantenimientos import show_mantenimientos
from views.dados_baja import show_dados_baja

# ============================================================================
# FUNCIÓN PRINCIPAL
# ============================================================================

def main():
    """Función principal del dashboard."""
    
    # Configurar página
    configure_page()
    
    # Header
    st.markdown(
        '<h1 class="main-title">🏥 DASHBOARD DE INVENTARIO TECNOLÓGICO</h1>', 
        unsafe_allow_html=True
    )
    st.markdown(
        '<p class="subtitle">Hospital Regional Alfonso Jaramillo Salazar - Líbano, Tolima</p>', 
        unsafe_allow_html=True
    )
    
    # Sidebar
    render_sidebar()
    
    # Cargar datos
    sheets = load_data()
    
    if not sheets:
        st.error("❌ No se encontró archivo de inventario")
        st.info("📂 Coloca el archivo Excel en el mismo directorio del app.py")
        st.code("""
Archivos soportados:
- inventario_hospital_test.xlsx
- inventario_hospital_v1.xlsx
- inventario_hospital.xlsx
        """)
        return
    
    # Mostrar info en sidebar
    show_data_info(sheets)
    
    st.markdown("---")
    
    # Tabs principales
    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "💻 Equipos de Cómputo",
        "🖨️ Impresoras y Escáneres",
        "🖱️ Periféricos",
        "🌐 Equipos de Red",
        "🔧 Mantenimientos",
        "📦 Dados de Baja"
    ])
    
    with tab1:
        show_equipos_computo(sheets.get("Equipos de Cómputo", pd.DataFrame()))
    
    with tab2:
        show_impresoras(sheets.get("Impresoras y Escáneres", pd.DataFrame()))
    
    with tab3:
        show_perifericos(sheets.get("Periféricos", pd.DataFrame()))
    
    with tab4:
        show_red(sheets.get("Equipos de Red", pd.DataFrame()))
    
    with tab5:
        show_mantenimientos(sheets.get("Mantenimientos", pd.DataFrame()))
    
    with tab6:
        show_dados_baja(sheets.get("Equipos Dados de Baja", pd.DataFrame()))
    
    # Footer
    render_footer()

# ============================================================================
# SIDEBAR
# ============================================================================

def render_sidebar():
    """Renderizar sidebar con logo e información."""
    with st.sidebar:
        # Logo
        logo_path = Path("assets/logo.png")
        if logo_path.exists():
            st.image(str(logo_path), use_column_width=True)
        else:
            st.markdown(f"""
            <div style="background: linear-gradient(135deg, {COLORS['primary']}, {COLORS['secondary']});
                 color: white; padding: 2.5rem 1.5rem; border-radius: 15px; text-align: center;
                 box-shadow: 0 8px 20px rgba(0,0,0,0.15);">
                <div style="font-size: 3.5rem; margin-bottom: 0.5rem;">🏥</div>
                <div style="font-size: 1.2rem; font-weight: 700; margin-bottom: 0.3rem;">Hospital AJS</div>
                <div style="font-size: 0.9rem; opacity: 0.9;">Líbano, Tolima</div>
            </div>
            """, unsafe_allow_html=True)
        
        st.markdown("---")
        
        # Info del dashboard
        st.markdown("### ℹ️ Información")
        st.info("""
        **Dashboard v2.0**
        
        Sistema de gestión y análisis del inventario tecnológico hospitalario.
        
        ✨ **Características:**
        - 📊 Analíticas avanzadas
        - 🔔 Alertas automáticas
        - 📤 Exportación de datos
        - 📱 Diseño responsive
        """)
        
        st.markdown("---")
        
        st.markdown("### 🛠️ Soporte")
        st.markdown("""
        **Área de Sistemas**  
        Hospital Regional AJS
        
        💻 Desarrollado por:  
        **Ing. Jose Miguel Santos Naranjo**
        """)

def show_data_info(sheets):
    """Mostrar información de datos cargados."""
    with st.sidebar:
        st.markdown("---")
        st.markdown("### 📂 Datos Cargados")
        
        total_registros = sum(len(df) for df in sheets.values())
        
        st.metric("Total de Registros", total_registros)
        
        for name, df in sheets.items():
            emoji = get_sheet_emoji(name)
            st.success(f"{emoji} {name}: **{len(df)}**")

def get_sheet_emoji(sheet_name):
    """Obtener emoji según el nombre de la hoja."""
    emoji_map = {
        "Equipos de Cómputo": "💻",
        "Impresoras y Escáneres": "🖨️",
        "Periféricos": "🖱️",
        "Equipos de Red": "🌐",
        "Mantenimientos": "🔧",
        "Equipos Dados de Baja": "📦"
    }
    for key, emoji in emoji_map.items():
        if key in sheet_name:
            return emoji
    return "📄"

# ============================================================================
# CARGA DE DATOS
# ============================================================================

def load_data():
    """Cargar datos del inventario."""
    files = [
        "inventario_hospital_test.xlsx",
        "inventario_hospital_v1.xlsx",
        "inventario_hospital.xlsx"
    ]
    
    for filename in files:
        if os.path.exists(filename):
            sheets = load_all_sheets(filename)
            if sheets:
                return sheets
    
    return None

# ============================================================================
# FOOTER
# ============================================================================

def render_footer():
    """Renderizar footer."""
    st.markdown(f"""
    <div style='margin-top: 4rem; padding: 2rem; border-top: 3px solid {COLORS['secondary']};
         background: linear-gradient(to right, #f8f9fa, white); border-radius: 10px;'>
        <div style="text-align: center;">
            <div style="color: {COLORS['primary']}; font-weight: 700; font-size: 1.1rem; margin-bottom: 0.8rem;">
                🏥 Dashboard de Inventario Tecnológico v2.0
            </div>
            <div style="color: {COLORS['dark']}; font-size: 0.95rem; margin-bottom: 0.5rem;">
                <strong>Desarrollado por:</strong> Ing. Jose Miguel Santos Naranjo
            </div>
            <div style="color: {COLORS['gray']}; font-size: 0.85rem;">
                Hospital Regional Alfonso Jaramillo Salazar • Líbano, Tolima
            </div>
            <div style="color: {COLORS['gray']}; font-size: 0.75rem; margin-top: 0.5rem; opacity: 0.8;">
                © 2025 - Todos los derechos reservados
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

# ============================================================================
# PUNTO DE ENTRADA
# ============================================================================

if __name__ == "__main__":
    # Si se ejecuta directamente, verificar si streamlit está corriendo
    if 'streamlit' not in sys.modules:
        print("\n" + "="*70)
        print("  🏥 DASHBOARD DE INVENTARIO TECNOLÓGICO")
        print("  Hospital Regional Alfonso Jaramillo Salazar")
        print("="*70)
        print("\n🚀 Iniciando dashboard...\n")
        
        import subprocess
        try:
            subprocess.run([sys.executable, "-m", "streamlit", "run", __file__])
        except KeyboardInterrupt:
            print("\n✅ Dashboard detenido correctamente")
        except Exception as e:
            print(f"\n❌ Error: {e}")
    else:
        main()
