"""
Módulo de Utilidades - Dashboard de Inventario Tecnológico
Hospital Regional Alfonso Jaramillo Salazar
Desarrollado por: Ing. Jose Miguel Santos Naranjo
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
from typing import Dict, Optional, List
from io import BytesIO

# ============================================================================
# COLORES Y CONFIGURACIÓN
# ============================================================================

COLORS = {
    "primary": "#2D6A4F",
    "secondary": "#52B788",
    "success": "#40916C",
    "warning": "#FFC107",
    "danger": "#DC3545",
    "info": "#0077B6",
    "light": "#F8F9FA",
    "dark": "#343A40",
    "gray": "#6C757D",
}


def configure_page():
    """Configuración inicial de la página."""
    st.set_page_config(
        page_title="Dashboard Inventario Hospital",
        page_icon="🏥",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    st.markdown(get_css(), unsafe_allow_html=True)


def get_css():
    """CSS mejorado con mejor estética."""
    return f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    
    * {{
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    }}
    
    :root {{
        --primary: {COLORS['primary']};
        --secondary: {COLORS['secondary']};
        --success: {COLORS['success']};
        --warning: {COLORS['warning']};
        --danger: {COLORS['danger']};
        --info: {COLORS['info']};
    }}
    
    .main .block-container {{
        padding: 1rem 2rem !important;
        max-width: 1400px;
    }}
    
    /* Títulos mejorados */
    .main-title {{
        background: linear-gradient(135deg, var(--primary) 0%, var(--secondary) 100%);
        color: white;
        padding: 2rem;
        border-radius: 15px;
        text-align: center;
        font-size: 2.5rem;
        font-weight: 700;
        margin-bottom: 1rem;
        box-shadow: 0 10px 30px rgba(45, 106, 79, 0.3);
        letter-spacing: -0.5px;
    }}
    
    .subtitle {{
        text-align: center;
        color: var(--dark);
        font-size: 1.1rem;
        margin-bottom: 2rem;
        font-weight: 500;
        opacity: 0.9;
    }}
    
    /* Tabs mejorados */
    .stTabs [data-baseweb="tab-list"] {{
        gap: 10px;
        background: transparent;
        padding: 0.5rem 0;
    }}
    
    .stTabs [data-baseweb="tab"] {{
        height: 60px;
        background: white;
        border-radius: 12px;
        padding: 0 28px;
        font-weight: 600;
        font-size: 0.95rem;
        border: 2px solid #e0e0e0;
        transition: all 0.3s ease;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
    }}
    
    .stTabs [data-baseweb="tab"]:hover {{
        background: #f8f9fa;
        border-color: var(--primary);
        transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(0,0,0,0.12);
    }}
    
    .stTabs [aria-selected="true"] {{
        background: linear-gradient(135deg, var(--primary) 0%, var(--secondary) 100%) !important;
        color: white !important;
        border-color: transparent !important;
        box-shadow: 0 6px 20px rgba(45, 106, 79, 0.4) !important;
    }}
    
    /* Métricas mejoradas */
    .metric-card {{
        background: white;
        padding: 1.5rem;
        border-radius: 12px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.08);
        transition: all 0.3s ease;
        border-left: 4px solid var(--primary);
    }}
    
    .metric-card:hover {{
        transform: translateY(-5px);
        box-shadow: 0 8px 25px rgba(0,0,0,0.15);
    }}
    
    /* Botones mejorados */
    .stButton > button {{
        background: linear-gradient(135deg, var(--primary) 0%, var(--secondary) 100%);
        color: white;
        border: none;
        border-radius: 10px;
        padding: 0.6rem 1.5rem;
        font-weight: 600;
        transition: all 0.3s ease;
        box-shadow: 0 4px 12px rgba(45, 106, 79, 0.3);
    }}
    
    .stButton > button:hover {{
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(45, 106, 79, 0.4);
    }}
    
    /* Expander mejorado */
    .streamlit-expanderHeader {{
        background: linear-gradient(to right, #f8f9fa, white);
        border-radius: 8px;
        font-weight: 600;
        border: 1px solid #e0e0e0;
    }}
    
    /* Dataframe mejorado */
    .dataframe {{
        border-radius: 8px;
        overflow: hidden;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
    }}
    
    /* Alertas mejoradas */
    .alert-card {{
        border-radius: 12px;
        padding: 1.5rem;
        margin: 1rem 0;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
        transition: all 0.3s ease;
    }}
    
    .alert-card:hover {{
        transform: translateY(-3px);
        box-shadow: 0 6px 20px rgba(0,0,0,0.15);
    }}
    
    /* Ocultar elementos Streamlit */
    #MainMenu, footer, header {{visibility: hidden;}}
    
    /* Responsive */
    @media (max-width: 768px) {{
        .main-title {{ font-size: 1.8rem; padding: 1.5rem; }}
        .stColumns {{ gap: 0.5rem !important; }}
        .metric-card {{ padding: 1rem; }}
    }}
    </style>
    """


# ============================================================================
# TARJETAS DE MÉTRICAS MEJORADAS
# ============================================================================


def create_metric_card(label: str, value, delta=None, gradient=None, icon="📊"):
    """Crear tarjeta de métrica mejorada con iconos y deltas."""
    if gradient is None:
        gradient = (
            f"linear-gradient(135deg, {COLORS['primary']}, {COLORS['secondary']})"
        )

    delta_html = ""
    if delta:
        delta_color = COLORS["success"] if delta > 0 else COLORS["danger"]
        delta_icon = "↑" if delta > 0 else "↓"
        delta_html = f"""
        <div style="font-size: 0.85rem; margin-top: 0.5rem; color: {delta_color}; font-weight: 600;">
            {delta_icon} {abs(delta)}%
        </div>
        """

    st.markdown(
        f"""
    <div style="background: {gradient}; padding: 1.8rem; border-radius: 15px; 
         color: white; text-align: center; box-shadow: 0 8px 20px rgba(0,0,0,0.15);
         transition: all 0.3s ease;" class="metric-card-hover">
        <div style="font-size: 2.5rem; margin-bottom: 0.3rem;">{icon}</div>
        <div style="font-size: 0.85rem; opacity: 0.95; font-weight: 600; 
             text-transform: uppercase; letter-spacing: 1px; margin-bottom: 0.8rem;">{label}</div>
        <div style="font-size: 2.8rem; font-weight: 700;">{value}</div>
        {delta_html}
    </div>
    """,
        unsafe_allow_html=True,
    )


# ============================================================================
# CARGA DE DATOS
# ============================================================================


@st.cache_data(ttl=300)
def load_all_sheets(file_path: str):
    """Cargar todas las hojas del Excel con procesamiento mejorado."""
    try:
        sheets = {}
        xls = pd.ExcelFile(file_path)
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            df.columns = df.columns.str.strip()
            sheets[sheet_name] = process_sheet(df, sheet_name)
        return sheets
    except Exception as e:
        st.error(f"❌ Error al cargar datos: {e}")
        return None


def process_sheet(df: pd.DataFrame, sheet_name: str):
    """Procesar DataFrame con conversiones de tipos."""
    # Fechas
    date_cols = [
        "Fecha de Adquisición",
        "Fecha Venc. Garantía",
        "Fecha Exp. Antivirus",
        "Último Mantenimiento",
        "Última Act. Windows",
        "Fecha Mantenimiento",
        "Próximo Mantenimiento",
        "Fecha de Baja",
    ]
    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    # Numéricos
    num_cols = [
        "RAM (GB)",
        "Almacenamiento (GB)",
        "Valor de Adquisición (COP)",
        "Costo (COP)",
        "Puertos Totales",
    ]
    for col in num_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    return df


# ============================================================================
# GRÁFICOS AVANZADOS
# ============================================================================


def create_gauge(value: float, max_value: float, title: str, suffix=""):
    """Crear medidor gauge elegante."""
    fig = go.Figure(
        go.Indicator(
            mode="gauge+number+delta",
            value=value,
            title={"text": title, "font": {"size": 20, "weight": "bold"}},
            number={"suffix": suffix, "font": {"size": 40}},
            gauge={
                "axis": {"range": [None, max_value], "tickwidth": 1},
                "bar": {"color": COLORS["primary"], "thickness": 0.75},
                "bgcolor": "white",
                "borderwidth": 2,
                "bordercolor": COLORS["gray"],
                "steps": [
                    {"range": [0, max_value * 0.33], "color": "#ffebee"},
                    {"range": [max_value * 0.33, max_value * 0.66], "color": "#fff9c4"},
                    {"range": [max_value * 0.66, max_value], "color": "#e8f5e9"},
                ],
                "threshold": {
                    "line": {"color": COLORS["danger"], "width": 4},
                    "thickness": 0.75,
                    "value": max_value * 0.9,
                },
            },
        )
    )
    fig.update_layout(height=300, margin=dict(l=20, r=20, t=60, b=20))
    return fig


def create_pie_chart(
    df: pd.DataFrame, column: str, hole=0.4, height=400, color_map=None
):
    """Crear gráfico de torta mejorado."""
    counts = df[column].value_counts()

    fig = px.pie(
        values=counts.values,
        names=counts.index,
        hole=hole,
        color_discrete_sequence=px.colors.qualitative.Set3 if not color_map else None,
        color=counts.index if color_map else None,
        color_discrete_map=color_map,
    )

    fig.update_traces(
        textposition="inside",
        textinfo="percent+label",
        hovertemplate="<b>%{label}</b><br>Cantidad: %{value}<br>%{percent}<extra></extra>",
        marker=dict(line=dict(color="white", width=2)),
    )

    fig.update_layout(
        height=height,
        showlegend=True,
        legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1.05),
        margin=dict(l=20, r=20, t=40, b=20),
        font=dict(size=12, family="Inter"),
    )

    return fig


def create_bar_chart(
    df: pd.DataFrame,
    column: str,
    orientation="v",
    height=400,
    color_scale="Greens",
    top_n=None,
    title="",
):
    """Crear gráfico de barras mejorado."""
    counts = df[column].value_counts()
    if top_n:
        counts = counts.head(top_n)

    if orientation == "h":
        fig = px.bar(
            x=counts.values,
            y=counts.index,
            orientation="h",
            color=counts.values,
            color_continuous_scale=color_scale,
            title=title,
        )
        fig.update_xaxes(title="Cantidad")
        fig.update_yaxes(title="")
    else:
        fig = px.bar(
            x=counts.index,
            y=counts.values,
            color=counts.values,
            color_continuous_scale=color_scale,
            title=title,
        )
        fig.update_xaxes(title="")
        fig.update_yaxes(title="Cantidad")

    fig.update_traces(
        hovertemplate='<b>%{x if orientation == "v" else y}</b><br>Cantidad: %{y if orientation == "v" else x}<extra></extra>',
        marker_line_color="white",
        marker_line_width=1.5,
    )

    fig.update_layout(
        height=height,
        showlegend=False,
        margin=dict(l=20, r=20, t=40, b=20),
        font=dict(size=12, family="Inter"),
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
    )

    return fig


def create_line_chart(df: pd.DataFrame, x: str, y: str, title="", height=400):
    """Crear gráfico de líneas mejorado."""
    fig = px.line(df, x=x, y=y, title=title, markers=True)

    fig.update_traces(
        line=dict(color=COLORS["primary"], width=3),
        marker=dict(
            size=10, color=COLORS["secondary"], line=dict(width=2, color="white")
        ),
    )

    fig.update_layout(
        height=height,
        margin=dict(l=20, r=20, t=40, b=20),
        font=dict(size=12, family="Inter"),
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        hovermode="x unified",
    )

    return fig


def create_sunburst(df: pd.DataFrame, path: List[str], values=None, title=""):
    """Crear gráfico sunburst jerárquico."""
    fig = px.sunburst(
        df,
        path=path,
        values=values,
        title=title,
        color_discrete_sequence=px.colors.qualitative.Pastel,
    )
    fig.update_layout(height=500, margin=dict(l=20, r=20, t=40, b=20))
    return fig


# ============================================================================
# ALERTAS MEJORADAS
# ============================================================================


def check_antivirus_alerts(df: pd.DataFrame):
    """Verificar alertas de antivirus."""
    try:
        df_copy = df.copy()
        df_copy["Fecha Exp. Antivirus"] = pd.to_datetime(
            df_copy["Fecha Exp. Antivirus"], errors="coerce"
        )
        hoy = datetime.now()

        vencidos = df_copy[df_copy["Fecha Exp. Antivirus"] < hoy]
        proximos = df_copy[
            (df_copy["Fecha Exp. Antivirus"] >= hoy)
            & (df_copy["Fecha Exp. Antivirus"] <= hoy + timedelta(days=30))
        ]

        nivel = (
            "danger"
            if len(vencidos) > 0
            else ("warning" if len(proximos) > 5 else "success")
        )

        return {
            "vencidos": vencidos,
            "proximos": proximos,
            "total": len(vencidos) + len(proximos),
            "nivel": nivel,
            "mensaje": f"{len(vencidos)} vencidos • {len(proximos)} próximos (30 días)",
        }
    except:
        return {
            "vencidos": pd.DataFrame(),
            "proximos": pd.DataFrame(),
            "total": 0,
            "nivel": "info",
            "mensaje": "No disponible",
        }


def check_license_alerts(df: pd.DataFrame):
    """Verificar alertas de licencias."""
    try:
        sin_lic = len(df[df["Estado Licencia Windows"] == "Sin Licencia"])
        por_act = len(df[df["Estado Licencia Windows"] == "Por Activar"])
        prueba = len(df[df["Estado Licencia Windows"] == "Prueba"])
        total = sin_lic + por_act + prueba

        nivel = "danger" if sin_lic > 0 else ("warning" if total > 0 else "success")
        problemas = df[
            df["Estado Licencia Windows"].isin(
                ["Sin Licencia", "Por Activar", "Prueba"]
            )
        ]

        return {
            "sin_licencia": sin_lic,
            "por_activar": por_act,
            "prueba": prueba,
            "total": total,
            "nivel": nivel,
            "problemas": problemas,
            "mensaje": f"{sin_lic} sin licencia • {por_act} por activar • {prueba} en prueba",
        }
    except:
        return {
            "total": 0,
            "nivel": "info",
            "problemas": pd.DataFrame(),
            "mensaje": "No disponible",
        }


def render_alert_card(icon: str, title: str, message: str, nivel: str, details_df=None):
    """Renderizar tarjeta de alerta mejorada."""
    color = COLORS.get(nivel, COLORS["info"])

    st.markdown(
        f"""
    <div class="alert-card" style="background: linear-gradient(135deg, {color} 0%, {color}dd 100%); 
         color: white;">
        <div style="font-size: 2.5rem; margin-bottom: 0.5rem; text-align: center;">{icon}</div>
        <div style="font-size: 1.3rem; font-weight: 700; text-align: center; margin-bottom: 0.5rem;">
            {title}
        </div>
        <div style="font-size: 1rem; text-align: center; opacity: 0.95;">
            {message}
        </div>
    </div>
    """,
        unsafe_allow_html=True,
    )

    if details_df is not None and not details_df.empty:
        with st.expander(
            f"📋 Ver detalles ({len(details_df)} registros)", expanded=False
        ):
            for idx, row in details_df.head(10).iterrows():
                codigo = row.get("Código Inventario", row.get("Código Equipo", "N/A"))
                area = row.get("Área / Servicio", "N/A")
                st.markdown(f"🔸 **{codigo}** - {area}")


# ============================================================================
# EXPORTACIÓN
# ============================================================================


def show_export_options(df: pd.DataFrame, prefix="datos"):
    """Mostrar opciones de exportación mejoradas."""
    st.markdown("---")
    st.markdown("### 📤 Exportar Datos")

    col1, col2, col3 = st.columns([1, 1, 2])
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    with col1:
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False)
        output.seek(0)

        st.download_button(
            "📥 Descargar Excel",
            output,
            f"{prefix}_{timestamp}.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with col2:
        csv = df.to_csv(index=False, encoding="utf-8-sig")
        st.download_button(
            "📥 Descargar CSV",
            csv,
            f"{prefix}_{timestamp}.csv",
            "text/csv",
            use_container_width=True,
        )

    with col3:
        st.info(
            f"📊 Total de registros: **{len(df)}** | Columnas: **{len(df.columns)}**"
        )
