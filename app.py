import streamlit as st
import pandas as pd
import datetime
import plotly.express as px
import plotly.graph_objects as go
import io
import os
import re
from pathlib import Path

from config import (
    PRIMARY_COLOR,
    SECONDARY_COLOR,
    ESTATUS_COLORS,
    ESTATUS_REGISTRADA,
    ESTATUS_EN_PROCESO,
    ESTATUS_PARCIAL,
    ESTATUS_COMPLETADA,
    ESTATUS_CANCELADA,
    SQLITE_DB_PATH,
    get_remisiones_dir
)
from db_manager import (
    init_db,
    save_po,
    delete_po,
    clear_all_pos_db,
    get_all_pos,
    get_po_by_folio,
    get_all_partidas,
    get_po_history,
    export_sync_to_excel
)
from remisiones_sync import (
    get_tracking_for_po,
    get_global_pos_tracking_summary,
    load_remisiones_databases
)
from corte_doblez_sync import (
    get_corte_doblez_tracking_for_po,
    get_integrated_360_summary,
    load_corte_doblez_databases
)
from pdf_parser import parse_po_pdf, parse_email_text
from excel_importer import generate_po_excel_template, parse_uploaded_excel

# Inicializar Base de Datos
init_db()

# Configuración de Página
st.set_page_config(
    page_title="SIGRAMA - Control Central de Órdenes de Compra (PO Tracker)",
    page_icon="📑",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Inyección de CSS Oficial Sigrama (Manual de Identidad Corporativa - PANTONE 485 C & Black 7 C)
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:ital,wght@0,400;0,500;0,600;0,700;0,800;0,900;1,700;1,900&family=Questrial&display=swap');

    /* Fuentes globales según manual corporativo */
    html, body, [class*="css"], .stApp {
        font-family: 'Questrial', sans-serif !important;
        background-color: #F8FAFC !important;
    }

    h1, h2, h3, h4, h5, h6, .main-title {
        font-family: 'Montserrat', sans-serif !important;
        font-weight: 700 !important;
        color: #111111 !important;
    }

    /* Barra lateral corporativa en Negro profundo #111111 (PANTONE Black 7 C) */
    [data-testid="stSidebar"] {
        background-color: #111111 !important;
        border-right: 1px solid #1E293B !important;
    }
    [data-testid="stSidebar"] p, 
    [data-testid="stSidebar"] span, 
    [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] div[data-testid="stMarkdownContainer"] {
        color: #FFFFFF !important;
        font-family: 'Questrial', sans-serif !important;
    }
    
    /* Radio de navegación en sidebar estilizado */
    [data-testid="stSidebar"] div[role="radiogroup"] {
        background-color: #18181B !important;
        border: 1px solid #27272A !important;
        border-radius: 8px !important;
        padding: 6px !important;
        gap: 3px !important;
    }
    [data-testid="stSidebar"] div[role="radiogroup"] label {
        color: #E2E8F0 !important;
        font-size: 13.5px !important;
        font-family: 'Questrial', sans-serif !important;
        padding: 8px 12px !important;
        border-radius: 6px !important;
        transition: all 0.2s ease !important;
        margin: 0 !important;
    }
    [data-testid="stSidebar"] div[role="radiogroup"] label:hover {
        color: #FFFFFF !important;
        background-color: rgba(236, 32, 36, 0.15) !important;
    }
    [data-testid="stSidebar"] div[role="radiogroup"] label[data-checked="true"] {
        background-color: #EC2024 !important;
        color: #FFFFFF !important;
        font-weight: 700 !important;
    }

    /* Botones Oficiales - Rojo Corporativo PANTONE 485 C (#EC2024) */
    button[kind="primary"], div.stButton > button[kind="primary"],
    div.stDownloadButton > button[kind="primary"],
    div.stFormSubmitButton > button {
        background-color: #EC2024 !important;
        color: #FFFFFF !important;
        font-family: 'Montserrat', sans-serif !important;
        font-weight: 700 !important;
        text-transform: uppercase !important;
        letter-spacing: 0.5px !important;
        border-radius: 6px !important;
        border: 1px solid #EC2024 !important;
        padding: 8px 20px !important;
        box-shadow: 0 3px 8px rgba(236, 32, 36, 0.25) !important;
        transition: all 0.25s ease !important;
        font-size: 13px !important;
    }
    button[kind="primary"]:hover, div.stButton > button[kind="primary"]:hover,
    div.stDownloadButton > button[kind="primary"]:hover,
    div.stFormSubmitButton > button:hover {
        background-color: #C01216 !important;
        border-color: #C01216 !important;
        box-shadow: 0 5px 14px rgba(236, 32, 36, 0.35) !important;
        transform: translateY(-1px) !important;
    }

    /* Botones Secundarios */
    button[kind="secondary"], div.stButton > button[kind="secondary"],
    div.stDownloadButton > button[kind="secondary"] {
        background-color: #FFFFFF !important;
        color: #111111 !important;
        border: 1px solid #CBD5E1 !important;
        font-family: 'Montserrat', sans-serif !important;
        font-weight: 600 !important;
        border-radius: 6px !important;
        transition: all 0.2s ease !important;
    }
    button[kind="secondary"]:hover, div.stButton > button[kind="secondary"]:hover {
        border-color: #EC2024 !important;
        color: #EC2024 !important;
        background-color: #FFF5F5 !important;
    }

    /* Estilo de las pestañas (Tabs) */
    button[role="tab"] {
        font-family: 'Montserrat', sans-serif !important;
        font-weight: 700 !important;
        color: #475569 !important;
        font-size: 14px !important;
        padding: 10px 16px !important;
    }
    button[role="tab"][aria-selected="true"] {
        color: #EC2024 !important;
        border-bottom-color: #EC2024 !important;
        border-bottom-width: 3px !important;
    }

    /* Contenedores de Métricas (KPIs) Oficiales */
    [data-testid="metric-container"] {
        background-color: #FFFFFF !important;
        border: 1px solid #E2E8F0 !important;
        border-left: 5px solid #EC2024 !important;
        border-radius: 8px !important;
        padding: 14px 18px !important;
        box-shadow: 0 2px 6px rgba(0,0,0,0.04) !important;
    }
    [data-testid="metric-container"] label {
        font-family: 'Montserrat', sans-serif !important;
        color: #64748B !important;
        font-weight: 700 !important;
        text-transform: uppercase !important;
        font-size: 11px !important;
        letter-spacing: 0.5px !important;
    }
    [data-testid="metric-container"] div[data-testid="stMetricValue"] {
        font-family: 'Montserrat', sans-serif !important;
        color: #0F172A !important;
        font-weight: 900 !important;
        font-size: 26px !important;
    }

    /* Inputs, Selectboxes y Formularios */
    div[data-baseweb="input"], div[data-baseweb="select"], textarea {
        border-color: #D2D3D5 !important;
        border-radius: 6px !important;
    }
    div[data-baseweb="input"]:focus-within, div[data-baseweb="select"]:focus-within {
        border-color: #EC2024 !important;
        box-shadow: 0 0 0 1px #EC2024 !important;
    }

    /* Ocultar encabezados Streamlit no requeridos */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    .stAppDeployButton {display: none !important;}
</style>
""", unsafe_allow_html=True)

# Sidebar Corporativo Oficial Sigrama (Manual pág. 28-29)
with st.sidebar:
    logo_neg_path = Path(__file__).resolve().parent / "logo_sigrama_negative.png"
    logo_path = Path(__file__).resolve().parent / "logo_sigrama.png"
    
    if logo_neg_path.exists():
        st.image(str(logo_neg_path), use_container_width=True)
    elif logo_path.exists():
        st.image(str(logo_path), use_container_width=True)
    else:
        st.markdown("<h2 style='color:#EC2024; text-align:center; font-family:\"Montserrat\";'>INDUSTRIA SIGRAMA</h2>", unsafe_allow_html=True)
        
    st.markdown("""
    <div style="background: rgba(236,32,36,0.12); border: 1px solid rgba(236,32,36,0.35); border-radius: 6px; padding: 7px 10px; margin: 10px 0 14px 0; text-align: center;">
        <span style="font-family: 'Montserrat', sans-serif; font-size: 10.5px; font-weight: 800; color: #EC2024; letter-spacing: 1px; text-transform: uppercase;">
            🛡️ CONTROL CENTRAL 4.0
        </span>
        <div style="color: #CBD5E1; font-size: 10.5px; margin-top: 1px;">Gestión y Trazabilidad de Órdenes (PO)</div>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("<p style='font-family: \"Montserrat\", sans-serif; font-size: 11px; font-weight: 700; color: #94A3B8; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 6px;'>Menú de Navegación:</p>", unsafe_allow_html=True)
    
    menu = st.radio(
        "Navegación:",
        [
            "📊 Dashboard Ejecutivo",
            "🔍 Ficha de Trazabilidad 360°",
            "📋 Matriz de Órdenes",
            "📬 Bandeja de Entrada OCR",
            "✏️ Ajuste de PO",
            "🔄 Estado de Integración",
            "📘 Manual y Arquitectura 4.0",
            "🛠️ Mantenimiento de la App"
        ],
        index=0,
        label_visibility="collapsed"
    )
    
    # Resumen Operativo en Vivo en Sidebar
    df_all_pos = get_all_pos()
    df_all_part = get_all_partidas()
    c_req_tot = float(df_all_part['cantidad_requerida'].sum()) if not df_all_part.empty and 'cantidad_requerida' in df_all_part.columns else 0.0
    
    st.markdown(f"""
    <div style="background: #18181B; border: 1px solid #27272A; border-left: 4px solid #EC2024; border-radius: 6px; padding: 12px 14px; margin-top: 18px;">
        <div style="font-family: 'Montserrat', sans-serif; font-size: 10px; color: #EC2024; font-weight: 800; text-transform: uppercase; letter-spacing: 0.8px;">
            Planta y Cadena de Suministro
        </div>
        <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 8px;">
            <span style="color: #94A3B8; font-size: 11.5px;">Órdenes Activas:</span>
            <b style="color: #FFFFFF; font-size: 13px; font-family: 'Montserrat', sans-serif;">{len(df_all_pos)} POs</b>
        </div>
        <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 4px;">
            <span style="color: #94A3B8; font-size: 11.5px;">Partidas en Catálogo:</span>
            <b style="color: #FFFFFF; font-size: 13px; font-family: 'Montserrat', sans-serif;">{len(df_all_part)} SKU</b>
        </div>
        <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 4px;">
            <span style="color: #94A3B8; font-size: 11.5px;">Piezas Requeridas:</span>
            <b style="color: #E2E8F0; font-size: 13px; font-family: 'Montserrat', sans-serif;">{c_req_tot:,.0f} pzas</b>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
    if st.button("🔄 Sincronizar Planta en Vivo", use_container_width=True, key="btn_sync_sidebar_live"):
        from remisiones_sync import sync_live_remisiones_from_github
        with st.spinner("Sincronizando con Almacén y Taller..."):
            sync_live_remisiones_from_github()
        st.cache_data.clear()
        st.toast("✅ ¡Datos actualizados desde GitHub en tiempo real!")
        st.rerun()
        
    # Cierre Oficial de Barra Lateral (Manual pág. 27, 28, 29)
    st.markdown("""
    <div style="text-align: center; margin-top: 25px; padding-top: 15px; border-top: 1px solid #27272A;">
        <span style="font-family: 'Questrial', sans-serif; font-style: italic; font-size: 13px; color: #FFFFFF; border-bottom: 2px solid #EC2024; padding-bottom: 2px;">
            Ingeniería que da resultados!!
        </span>
        <p style="color: #64748B; font-size: 10px; margin-top: 8px; margin-bottom: 0; font-family: 'Montserrat', sans-serif;">
            Industria Sigrama S.A. de C.V.
        </p>
    </div>
    """, unsafe_allow_html=True)

# ==============================================================================
# ENCABEZADO / BANNER SUPERIOR CORPORATIVO (MANUAL OFICIAL SIGRAMA)
# ==============================================================================
banner_path = Path(__file__).resolve().parent / "banner_sigrama.png"
if banner_path.exists():
    st.image(str(banner_path), use_container_width=True)
    st.markdown('<hr style="border: 0; height: 3px; background: linear-gradient(90deg, #EC2024 0%, #111111 100%); margin: 8px 0 16px 0; border-radius: 2px;">', unsafe_allow_html=True)
else:
    logo_b64 = ""
    logo_path_main = Path(__file__).resolve().parent / "logo_sigrama.png"
    if logo_path_main.exists():
        import base64
        logo_b64 = base64.b64encode(logo_path_main.read_bytes()).decode()

    st.markdown(f"""
    <div style="background: linear-gradient(135deg, #111111 0%, #1A1A1A 60%, #0F172A 100%); border-bottom: 4px solid #EC2024; border-radius: 8px; padding: 14px 20px; margin-bottom: 20px; box-shadow: 0 4px 12px rgba(0,0,0,0.07); display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 14px;">
        <div style="display: flex; align-items: center; gap: 16px;">
            <div style="background: #FFFFFF; padding: 5px 10px; border-radius: 6px; display: flex; align-items: center; box-shadow: 0 2px 5px rgba(0,0,0,0.15);">
                <img src="data:image/png;base64,{logo_b64}" style="height: 30px; display: block;" alt="Industria Sigrama">
            </div>
            <div>
                <span style="font-family: 'Montserrat', sans-serif; font-size: 10px; font-weight: 800; color: #EC2024; letter-spacing: 1.5px; text-transform: uppercase;">INDUSTRIA SIGRAMA • SUITE 4.0</span>
                <h2 style="font-family: 'Montserrat', sans-serif; font-size: 19px; font-weight: 900; color: #FFFFFF; margin: 2px 0 0 0; letter-spacing: -0.2px;">
                    PO TRACKER & TRAZABILIDAD INTEGRAL 360°
                </h2>
            </div>
        </div>
        <div style="text-align: right;">
            <span style="font-family: 'Montserrat', sans-serif; font-size: 11.5px; font-weight: 800; color: #EC2024; letter-spacing: 0.8px; text-transform: uppercase;">
                SOLUCIONES QUE TRANSFORMAN TU EMPRESA
            </span>
            <div style="color: #94A3B8; font-size: 11px; margin-top: 2px; font-family: 'Questrial', sans-serif;">Corte Láser • Prensas Doblez • Almacén PT • Logística y Envíos</div>
        </div>
    </div>
    """, unsafe_allow_html=True)


# ==============================================================================
# SECCIÓN 1: DASHBOARD EJECUTIVO
# ==============================================================================
if menu == "📊 Dashboard Ejecutivo":
    st.title("📊 Dashboard Ejecutivo & Reportes Estratégicos")
    st.markdown("Visión global de requerimientos de clientes, avance de manufactura en planta y cumplimiento de entregas en tiempo real.")
    
    df_pos = get_all_pos()
    df_part = get_all_partidas()
    
    if df_pos.empty:
        st.info("💡 Aún no hay Órdenes de Compra registradas. Dirígete a **'📬 Bandeja de Entrada OCR'** para ingresar tu primer lote de POs.")
    else:
        tab_dash_kpi, tab_dash_360 = st.tabs([
            "📈 Resumen Ejecutivo & Cumplimiento",
            "🎯 Matriz de Control 360° (Producción vs Remisión)"
        ])
        
        # ----------------------------------------------------------------------
        # PESTAÑA 1: RESUMEN EJECUTIVO & KPIS
        # ----------------------------------------------------------------------
        with tab_dash_kpi:
            df_summary = get_global_pos_tracking_summary(df_pos, df_part)
            
            tot_pos = len(df_summary)
            tot_piezas_req = df_summary['piezas_requeridas'].sum()
            tot_piezas_rem = df_summary['piezas_remisionadas'].sum()
            tot_piezas_pend = df_summary['piezas_pendientes'].sum()
            pct_global = (tot_piezas_rem / tot_piezas_req * 100.0) if tot_piezas_req > 0 else 0.0
            
            c1, c2, c3, c4 = st.columns(4)
            with c1:
                st.metric("Total Órdenes (POs)", f"{tot_pos}")
            with c2:
                st.metric("Total Piezas Requeridas", f"{tot_piezas_req:,.0f}")
            with c3:
                st.metric("Piezas Remisionadas", f"{tot_piezas_rem:,.0f}", f"{pct_global:.1f}% Cumplido")
            with c4:
                st.metric("Piezas Pendientes de Envío", f"{tot_piezas_pend:,.0f}", delta=f"-{tot_piezas_pend:,.0f}", delta_color="inverse")
                
            st.write("---")
            
            g1, g2 = st.columns(2)
            with g1:
                st.subheader("Estatus Global de Órdenes")
                status_counts = df_summary['estatus_remision'].value_counts().reset_index()
                status_counts.columns = ['Estatus', 'Cantidad']
                
                fig_pie = px.pie(
                    status_counts,
                    values='Cantidad',
                    names='Estatus',
                    color='Estatus',
                    color_discrete_map=ESTATUS_COLORS,
                    hole=0.45
                )
                fig_pie.update_layout(margin=dict(t=20, b=20, l=20, r=20))
                st.plotly_chart(fig_pie, use_container_width=True)
                
            with g2:
                st.subheader("Avance de Piezas por Proyecto")
                df_proy = df_summary.groupby('proyecto')[['piezas_requeridas', 'piezas_remisionadas']].sum().reset_index()
                
                fig_bar = go.Figure()
                fig_bar.add_trace(go.Bar(name='Requeridas', x=df_proy['proyecto'], y=df_proy['piezas_requeridas'], marker_color='#111111'))
                fig_bar.add_trace(go.Bar(name='Remisionadas (Enviadas)', x=df_proy['proyecto'], y=df_proy['piezas_remisionadas'], marker_color='#EC2024'))
                fig_bar.update_layout(barmode='group', margin=dict(t=20, b=20, l=20, r=20), xaxis_title="Proyecto", yaxis_title="Cantidad de Piezas")
                st.plotly_chart(fig_bar, use_container_width=True)
                
            st.subheader("⚠️ Órdenes con Entregas Pendientes / En Proceso")
            df_pendientes = df_summary[df_summary['piezas_pendientes'] > 0].sort_values(by='piezas_pendientes', ascending=False)
            
            if not df_pendientes.empty:
                cols_show = ['po', 'proyecto', 'solicitante', 'piezas_requeridas', 'piezas_remisionadas', 'piezas_pendientes', 'pct_cumplimiento', 'estatus_remision']
                st.dataframe(
                    df_pendientes[cols_show].rename(columns={
                        'po': 'PO / Folio',
                        'proyecto': 'Proyecto',
                        'solicitante': 'Solicitante',
                        'piezas_requeridas': 'Piezas Req.',
                        'piezas_remisionadas': 'Enviadas',
                        'piezas_pendientes': 'Pendientes',
                        'pct_cumplimiento': '% Cumpl.',
                        'estatus_remision': 'Estatus'
                    }),
                    use_container_width=True,
                    hide_index=True
                )
            else:
                st.success("🎉 ¡Excelente! Todas las Órdenes de Compra registradas han sido completadas al 100%.")

        # ----------------------------------------------------------------------
        # PESTAÑA 2: MATRIZ DE CONTROL 360° (PRODUCCIÓN VS REMISIÓN)
        # ----------------------------------------------------------------------
        with tab_dash_360:
            df_mat = get_integrated_360_summary(df_pos, df_part)
            
            tot_req_all = df_mat['piezas_requeridas'].sum()
            tot_fab_all = df_mat['piezas_fabricadas'].sum()
            tot_env_all = df_mat['piezas_remisionadas'].sum()
            tot_pend_env = df_mat['piezas_pendientes_env'].sum()
            
            pct_fab_global = (tot_fab_all / tot_req_all * 100.0) if tot_req_all > 0 else 0.0
            pct_env_global = (tot_env_all / tot_req_all * 100.0) if tot_req_all > 0 else 0.0
            
            m1, m2, m3, m4 = st.columns(4)
            with m1:
                st.metric("📦 Total Piezas Requeridas", f"{tot_req_all:,.0f}")
            with m2:
                st.metric("🔵 Fabricadas en Planta", f"{tot_fab_all:,.0f}", f"{pct_fab_global:.1f}% Fabricado")
            with m3:
                st.metric("🟢 Remisionadas al Cliente", f"{tot_env_all:,.0f}", f"{pct_env_global:.1f}% Enviado")
            with m4:
                st.metric("⏳ Pendientes de Envío", f"{tot_pend_env:,.0f}", delta=f"-{tot_pend_env:,.0f}", delta_color="inverse")
                
            st.write("---")
            
            c_f1, c_f2, c_f3 = st.columns([2, 1, 1])
            with c_f1:
                q_search_360 = st.text_input("🔍 Búsqueda rápida (PO, Proyecto, SKU, OF, Remisión):", "", key="search_360_dash")
            with c_f2:
                est_opts = ["Todos"] + sorted(list(df_mat['estatus_360'].unique()))
                sel_est_360 = st.selectbox("Filtrar por Estatus 360°:", est_opts, key="est_360_dash")
            with c_f3:
                proy_opts_360 = ["Todos"] + sorted([p for p in df_mat['proyecto'].dropna().unique() if str(p).strip()])
                sel_proy_360 = st.selectbox("Filtrar por Proyecto:", proy_opts_360, key="proy_360_dash")
                
            df_filtered_360 = df_mat.copy()
            if q_search_360.strip():
                term = q_search_360.strip().lower()
                df_filtered_360 = df_filtered_360[
                    df_filtered_360['po'].astype(str).str.lower().str.contains(term) |
                    df_filtered_360['proyecto'].astype(str).str.lower().str.contains(term) |
                    df_filtered_360['ofs_asociadas'].astype(str).str.lower().str.contains(term) |
                    df_filtered_360['remisiones_asociadas'].astype(str).str.lower().str.contains(term)
                ]
            if sel_est_360 != "Todos":
                df_filtered_360 = df_filtered_360[df_filtered_360['estatus_360'] == sel_est_360]
            if sel_proy_360 != "Todos":
                df_filtered_360 = df_filtered_360[df_filtered_360['proyecto'] == sel_proy_360]
                
            st.markdown(f"Mostrando **{len(df_filtered_360)}** órdenes de compra:")
            
            cols_mat_show = [
                'po', 'proyecto', 'piezas_requeridas',
                'piezas_fabricadas', 'pct_fabricacion',
                'piezas_remisionadas', 'pct_remision',
                'piezas_pendientes_fab', 'piezas_pendientes_env',
                'ofs_asociadas', 'remisiones_asociadas', 'estatus_360'
            ]
            
            st.dataframe(
                df_filtered_360[cols_mat_show].rename(columns={
                    'po': 'PO / Folio',
                    'proyecto': 'Proyecto',
                    'piezas_requeridas': 'Cant. Req.',
                    'piezas_fabricadas': '🔵 Fab. Planta',
                    'pct_fabricacion': '🔵 % Fab.',
                    'piezas_remisionadas': '🟢 Remisionadas',
                    'pct_remision': '🟢 % Envío',
                    'piezas_pendientes_fab': 'Pend. Fab.',
                    'piezas_pendientes_env': 'Pend. Envío',
                    'ofs_asociadas': 'OFs (Corte y Doblez)',
                    'remisiones_asociadas': 'Remisiones',
                    'estatus_360': 'Estatus 360°'
                }),
                use_container_width=True,
                hide_index=True
            )
            
            output_360 = io.BytesIO()
            with pd.ExcelWriter(output_360, engine='openpyxl') as writer:
                df_filtered_360.to_excel(writer, sheet_name='Control_360_SIGRAMA', index=False)
            st.download_button(
                "📥 Descargar Matriz de Control 360° en Excel",
                data=output_360.getvalue(),
                file_name=f"Matriz_Control_360_SIGRAMA_{datetime.date.today().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )


# ==============================================================================
# SECCIÓN 2: BANDEJA DE ENTRADA OCR (PUNTO DE INICIO - CARGA DE 58 ÓRDENES)
# ==============================================================================
elif menu == "📬 Bandeja de Entrada OCR":
    st.title("📬 Bandeja de Entrada OCR & Ingesta de Órdenes de Compra")
    st.markdown("Punto de inicio para cargar ordenadamente las **58 Órdenes Internas** de Sigrama subiendo sus correos **.MSG** o documentos oficiales **.PDF**.")
    
    tab_ocr_upload, tab_ocr_excel, tab_ocr_manual, tab_ocr_casos = st.tabs([
        "📥 Subir Correo (.MSG) o Archivos (.PDF)",
        "📁 Carga Masiva de Excel",
        "📝 Registro Manual Guiado",
        "📨 Ejemplos Guiados (Casos 1 y 2)"
    ])
    
    # 1. Pestaña: Subida Directa de Correos y PDFs
    with tab_ocr_upload:
        st.subheader("📥 Cargar Correos (.MSG) o Archivos PDF de Clientes")
        st.write("Sube el archivo de correo `.msg` (con sus PDFs adjuntos) o los PDFs oficiales de las POs para ejecutar el motor OCR espacial.")
        
        custom_files = st.file_uploader(
            "Arrastra aquí tus archivos de correo (.msg) o documentos de PO (.pdf):",
            type=['msg', 'pdf'],
            accept_multiple_files=True,
            key="uploader_ocr_intake"
        )
        
        if custom_files:
            st.write(f"📁 **{len(custom_files)}** archivo(s) seleccionado(s).")
            if st.button("⚡ Procesar y Extraer Órdenes de Compra", type="primary", use_container_width=True, key="btn_ocr_process"):
                extracted_batch = []
                with st.spinner("Analizando documentos con motor OCR espacial..."):
                    for uploaded_f in custom_files:
                        f_bytes = uploaded_f.read()
                        f_name = uploaded_f.name
                        
                        if f_name.lower().endswith('.msg'):
                            from pdf_parser import extract_attachments_from_msg
                            msg_info = extract_attachments_from_msg(f_bytes)
                            ctx_msg = parse_email_text(msg_info.get('body', ''))
                            
                            # 1. Extraer ID Interno INT-000X desde el nombre del archivo .msg o asunto
                            m_int = re.search(r'\bINT[\s\-_]?(\d{1,4})\b', f"{f_name} {msg_info.get('subject', '')}", re.IGNORECASE)
                            id_int_auto = f"INT-{int(m_int.group(1)):04d}" if m_int else ""
                            
                            # 2. Extraer Fecha de llegada desde metadatos del correo
                            date_raw = str(msg_info.get('date', ''))
                            m_dt = re.search(r'(\d{4})[-/](\d{1,2})[-/](\d{1,2})', date_raw)
                            if m_dt:
                                f_llegada_auto = f"{m_dt.group(1)}-{int(m_dt.group(2)):02d}-{int(m_dt.group(3)):02d}"
                            else:
                                f_llegada_auto = datetime.date.today().strftime('%Y-%m-%d')
                                
                            # 3. Remitente / Comprador limpio
                            sender_raw = str(msg_info.get('sender', '')).strip()
                            if '<' in sender_raw:
                                sender_raw = sender_raw.split('<')[0].strip()
                                
                            ctx_msg['id_interno'] = id_int_auto
                            ctx_msg['fecha_llegada'] = f_llegada_auto
                            ctx_msg['remitente'] = sender_raw if sender_raw else 'Compras'
                            ctx_msg['asunto'] = msg_info.get('subject', '')
                            ctx_msg['msg_filename'] = f_name
                            ctx_msg['body'] = msg_info.get('body', '')
                            
                            found_pdf_in_msg = False
                            atts = msg_info.get('attachments', [])
                            for att in atts:
                                att_n = str(att.get('filename', ''))
                                att_d = att.get('data', b'')
                                is_pdf_file = att_n.lower().endswith('.pdf') or (isinstance(att_d, (bytes, bytearray)) and att_d.startswith(b'%PDF'))
                                
                                if is_pdf_file and not any(w in att_n.lower() for w in ['plano', 'drawing', 'cotizacion']):
                                    try:
                                        ctx_msg['pdf_filename'] = att_n
                                        cab_m, part_m = parse_po_pdf(att_d, email_context=ctx_msg)
                                        if cab_m:
                                            extracted_batch.append({'cab': cab_m, 'part': part_m, 'file_name': f"{f_name} ➔ {att_n}"})
                                            found_pdf_in_msg = True
                                    except Exception as e_att:
                                        st.error(f"Error procesando {att_n}: {e_att}")
                                        
                            # Si no se encontró archivo PDF en el .msg, extraer directamente del texto del correo:
                            if not found_pdf_in_msg:
                                po_detect = ctx_msg.get('po_detectada') or ''
                                if not po_detect:
                                    m_po_name = re.search(r'\b(26\d{2}[-\s]?\d{4}|26\d{6})\b', f"{f_name} {ctx_msg.get('asunto', '')}")
                                    po_detect = m_po_name.group(1).replace('-', '').replace(' ', '') if m_po_name else "2608-TEMP"
                                    
                                cab_m = {
                                    'po': po_detect,
                                    'id_interno': id_int_auto or 'INT-TEMP',
                                    'fecha_llegada': f_llegada_auto,
                                    'fecha_solicitada': (datetime.date.today() + datetime.timedelta(days=14)).strftime('%Y-%m-%d'),
                                    'archivo_correo': f_name,
                                    'archivo_pdf': 'Cuerpo de Correo (.msg)',
                                    'fecha_pedido': f_llegada_auto,
                                    'proyecto': ctx_msg.get('asunto', 'PROYECTO CLIENTE') or 'PROYECTO CLIENTE',
                                    'solicitante': 'CLIENTE',
                                    'comprador': sender_raw if sender_raw else 'Compras',
                                    'total': 0.0
                                }
                                part_m = []
                                for idx_p, p_urg in enumerate(ctx_msg.get('partes_urgentes', []), start=1):
                                    part_m.append({
                                        'item_no': idx_p,
                                        'sku_cliente': p_urg['sku'],
                                        'clave_sku': p_urg['sku'],
                                        'descripcion_producto': f"Material {p_urg['sku']}",
                                        'cantidad_requerida': p_urg['total_requerido'],
                                        'unidad': 'PIEZA',
                                        'precio_unitario': 0.0,
                                        'precio_total': 0.0,
                                        'fecha_entrega': f_llegada_auto,
                                        'parcialidad': 'P1',
                                        'observaciones_partida': ''
                                    })
                                extracted_batch.append({'cab': cab_m, 'part': part_m, 'file_name': f_name})
                        else:
                            if any(w in f_name.lower() for w in ['plano', 'drawing', 'rev', 'cotizacion']):
                                continue
                            try:
                                m_int_pdf = re.search(r'\bINT[\s\-_]?(\d{1,4})\b', f_name, re.IGNORECASE)
                                id_int_pdf = f"INT-{int(m_int_pdf.group(1)):04d}" if m_int_pdf else ""
                                ctx_pdf = {
                                    'id_interno': id_int_pdf,
                                    'pdf_filename': f_name,
                                    'fecha_llegada': datetime.date.today().strftime('%Y-%m-%d')
                                }
                                cab_f, part_f = parse_po_pdf(f_bytes, email_context=ctx_pdf)
                                extracted_batch.append({'cab': cab_f, 'part': part_f, 'file_name': f_name})
                            except Exception as e:
                                st.error(f"Error procesando {f_name}: {e}")
                                
                    # GUARDADO AUTOMÁTICO EN LA BASE DE DATOS
                    guardadas_auto = 0
                    for item in extracted_batch:
                        ok_g, _ = save_po(item['cab'], item['part'])
                        if ok_g:
                            guardadas_auto += 1
                            
                    st.session_state['ocr_batch_results'] = extracted_batch
                    if extracted_batch:
                        st.success(f"🎉 Se procesaron y registraron exitosamente **{len(extracted_batch)}** Órdenes de Compra en el sistema.")
                        st.rerun()
                    else:
                        st.warning("⚠️ No se pudieron extraer datos de la orden. Verifique que el archivo .msg contenga el PDF de la PO o que el archivo no esté protegido.")
                        
        if 'ocr_batch_results' in st.session_state and st.session_state['ocr_batch_results']:
            b_list = st.session_state['ocr_batch_results']
            st.write("---")
            st.markdown(f"### 📋 Lote de {len(b_list)} Órdenes Registradas")
            
            for idx, item in enumerate(b_list):
                with st.expander(f"📄 PO: {item['cab'].get('po', 'N/A')} • {item['cab'].get('proyecto', '')} ({item['file_name']})", expanded=True):
                    c_b1, c_b2, c_b3 = st.columns(3)
                    with c_b1:
                        st.markdown(f"• **Folio:** `{item['cab'].get('po')}`")
                        st.markdown(f"• **Proyecto:** `{item['cab'].get('proyecto')}`")
                    with c_b2:
                        st.markdown(f"• **Solicitante / Comprador:** `{item['cab'].get('solicitante')}` / `{item['cab'].get('comprador')}`")
                        st.markdown(f"• **Total:** `${item['cab'].get('total', 0):,.2f} MXN`")
                    with c_b3:
                        st.markdown(f"• **Partidas:** `{len(item['part'])}`")
                        
                    df_p_view = pd.DataFrame(item['part'])
                    col_order = [c for c in ['item_no', 'sku_cliente', 'clave_sku', 'descripcion_producto', 'cantidad_requerida', 'unidad', 'precio_unitario', 'precio_total', 'fecha_entrega'] if c in df_p_view.columns]
                    st.dataframe(
                        df_p_view[col_order].rename(columns={
                            'item_no': 'Item #',
                            'sku_cliente': 'SKU Cliente (Clave)',
                            'clave_sku': 'SKU Nuestro (Planta)',
                            'descripcion_producto': 'Descripción del Producto',
                            'cantidad_requerida': 'Cantidad',
                            'unidad': 'Unidad',
                            'precio_unitario': 'P. Unitario',
                            'precio_total': 'P. Total',
                            'fecha_entrega': 'Fecha Entrega'
                        }),
                        use_container_width=True,
                        hide_index=True
                    )

    # 2. Pestaña: Carga Masiva Excel
    with tab_ocr_excel:
        st.subheader("📁 Carga Masiva mediante Plantilla Oficial Excel")
        st.write("Descarga la plantilla oficial estandarizada, captura las órdenes y súbela para registrar múltiples POs a la vez.")
        
        template_bytes = generate_po_excel_template()
        st.download_button(
            label="📥 Descargar Plantilla Oficial Excel (2 Pestañas)",
            data=template_bytes,
            file_name="Plantilla_Oficial_Orden_de_Compra_Sigrama.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        st.write("---")
        excel_file = st.file_uploader("Subir Archivo Excel completado:", type=['xlsx'], key="uploader_ocr_excel")
        if excel_file is not None:
            if st.button("🚀 Procesar e Integrar Archivo Excel", type="primary", use_container_width=True):
                ok, msg, cab_xl, part_xl = parse_uploaded_excel(excel_file)
                if ok:
                    ok_save, msg_save = save_po(cab_xl, part_xl)
                    if ok_save:
                        st.success(f"✅ {msg_save}")
                        st.rerun()
                    else:
                        st.error(f"❌ {msg_save}")
                else:
                    st.error(f"❌ {msg}")

    # 3. Pestaña: Registro Manual
    with tab_ocr_manual:
        st.subheader("📝 Captura Manual Guiada de Orden de Compra")
        with st.form("form_manual_po_intake"):
            st.markdown("##### 1. Encabezado de la PO")
            m1, m2, m3 = st.columns(3)
            with m1:
                man_po = st.text_input("Folio de PO (Obligatorio):", placeholder="ej. 26083186")
                man_id_int = st.text_input("ID Interno (ej. INT-0001):", placeholder="INT-0001")
                man_fecha = st.date_input("Fecha de Pedido:", value=datetime.date.today())
            with m2:
                man_proy = st.text_input("Proyecto / Uso:", placeholder="ej. CLOUD / TAB-RQXP")
                man_sol = st.text_input("Solicitante:", placeholder="ej. ESTEFANIA IBARRA")
                man_comp = st.text_input("Comprador:", placeholder="ej. Josue Mesta")
            with m3:
                man_lab = st.text_input("L.A.B. / Destino:", value="ALMACEN SIGRAMA")
                man_tent = st.text_input("Tiempo de Entrega:", placeholder="ej. 18 AGOSTO 2026")
                man_obs = st.text_input("Observaciones Generales:")
                
            st.markdown("##### 2. Partida Principal")
            p1, p2, p3, p4, p5 = st.columns([1.5, 1.5, 3, 1, 1])
            with p1:
                man_sku_cli = st.text_input("SKU Cliente (Clave):", placeholder="SWB01431")
            with p2:
                man_sku_nos = st.text_input("SKU Nuestro (Planta):", placeholder="PP19380-03")
            with p3:
                man_desc = st.text_input("Descripción del Producto:", placeholder="382 X 10H BLANK DOOR")
            with p4:
                man_cant = st.number_input("Cantidad:", min_value=1.0, value=32.0, step=1.0)
            with p5:
                man_pu = st.number_input("P. Unitario ($):", min_value=0.0, value=385.55, step=10.0)
                
            man_submit = st.form_submit_button("💾 Guardar Orden de Compra", type="primary", use_container_width=True)
            if man_submit:
                if not man_po.strip():
                    st.error("❌ El Folio de la PO es obligatorio.")
                else:
                    pt_val = man_cant * man_pu
                    cab_man = {
                        "po": man_po.strip(),
                        "id_interno": man_id_int.strip().upper(),
                        "fecha_pedido": man_fecha.strftime("%Y-%m-%d"),
                        "proyecto": man_proy.strip(),
                        "solicitante": man_sol.strip(),
                        "requisicion": "",
                        "destino": man_lab.strip(),
                        "proveedor": "SIGRAMA PLANTA METALES",
                        "proveedor_atencion": "JESUS MORALES",
                        "cliente_facturar_a": "INDUSTRIA SIGRAMA S.A. DE C.V.",
                        "cliente_rfc": "ISI-870204-K4A",
                        "cliente_direccion": "C. JUAN ESCUTIA #50 COL. ABASTOS C.P. 27020 TORREON, COAH.",
                        "forma_pago": "CONTADO / CRÉDITO",
                        "lab": man_lab.strip(),
                        "tiempo_entrega": man_tent.strip(),
                        "comprador": man_comp.strip(),
                        "subtotal": pt_val,
                        "descuento": 0.0,
                        "iva": pt_val * 0.16,
                        "ret_iva": 0.0,
                        "ret_isr": 0.0,
                        "total": pt_val * 1.16,
                        "moneda": "MXN",
                        "observaciones": man_obs.strip(),
                        "texto_etiqueta": man_proy.strip(),
                        "color_fondo": "#EC2024",
                        "color_texto": "#FFFFFF"
                    }
                    part_man = [{
                        "item_no": 1,
                        "sku_cliente": man_sku_cli.strip().upper(),
                        "clave_sku": man_sku_nos.strip().upper(),
                        "descripcion_producto": man_desc.strip(),
                        "cantidad_requerida": man_cant,
                        "unidad": "PIEZA",
                        "precio_unitario": man_pu,
                        "precio_total": pt_val,
                        "fecha_entrega": (man_fecha + datetime.timedelta(days=7)).strftime("%Y-%m-%d"),
                        "parcialidad": "P1",
                        "observaciones_partida": ""
                    }]
                    ok, msg = save_po(cab_man, part_man)
                    if ok:
                        st.success(f"✅ {msg}")
                        st.rerun()
                    else:
                        st.error(f"❌ {msg}")

    # 4. Pestaña: Ejemplos Guiados
    with tab_ocr_casos:
        st.subheader("📨 Ejemplos de Correos y Formatos Sigrama")
        st.write("Demostración de extracción para correos individuales con urgencias o correos con múltiples órdenes y planos anexos.")
        with st.expander("📄 Ver Caso 1: PO 2608-3177 (Urgencia de Partes)"):
            if st.button("⚡ Simular Extracción de Caso 1"):
                st.success("Extracción simulada con éxito.")

    # --------------------------------------------------------------------------
    # TABLA DE CARGAS DE PO'S POR FECHAS DESCENDENTES
    # --------------------------------------------------------------------------
    st.write("---")
    st.subheader("📜 Historial de Órdenes Cargadas en el Sistema (Fechas Descendentes)")
    st.caption("Consulta y audita todas las órdenes de compra ingestadas en el sistema, ordenadas cronológicamente desde la más reciente hasta la más antigua.")
    
    df_pos_hist = get_all_pos()
    df_part_hist = get_all_partidas()
    
    if not df_pos_hist.empty:
        df_summary_hist = get_global_pos_tracking_summary(df_pos_hist, df_part_hist)
        
        # Filtros rápidos y selector de ordenamiento
        c_h1, c_h2, c_h3 = st.columns([2, 1.2, 0.8])
        with c_h1:
            q_h = st.text_input("🔍 Buscar en historial de cargas (Folio, ID Interno, Proyecto, Comprador):", "", key="search_intake_history")
        with c_h2:
            sort_h = st.selectbox("Ordenar Historial por:", [
                "📅 Fecha de Carga / Registro (Más reciente primero)",
                "📅 Fecha de Llegada de PO (Descendente)",
                "🔢 ID Interno (INT-0001 a INT-0058...)",
                "📄 Folio PO (Descendente)",
                "💰 Importe Total ($)"
            ], key="sort_intake_history")
        with c_h3:
            st.metric("Total POs Cargadas", f"{len(df_summary_hist)}")
            
        df_hist_filtered = df_summary_hist.copy()
        if q_h.strip():
            term = q_h.strip().lower()
            df_hist_filtered = df_hist_filtered[
                df_hist_filtered['po'].astype(str).str.lower().str.contains(term) |
                df_hist_filtered.get('id_interno', pd.Series(['']*len(df_hist_filtered))).astype(str).str.lower().str.contains(term) |
                df_hist_filtered['proyecto'].astype(str).str.lower().str.contains(term) |
                df_hist_filtered['comprador'].astype(str).str.lower().str.contains(term) |
                df_hist_filtered['solicitante'].astype(str).str.lower().str.contains(term)
            ]
            
        if "Fecha de Carga" in sort_h:
            df_hist_filtered = df_hist_filtered.sort_values(by=['fecha_registro', 'po'], ascending=[False, False])
        elif "Fecha de Llegada" in sort_h:
            df_hist_filtered = df_hist_filtered.sort_values(by=['fecha_llegada', 'fecha_pedido', 'po'], ascending=[False, False, False])
        elif "ID Interno" in sort_h:
            df_hist_filtered['id_sort_key'] = df_hist_filtered['id_interno'].apply(lambda x: str(x) if str(x).strip() else 'ZZZ')
            df_hist_filtered = df_hist_filtered.sort_values(by=['id_sort_key', 'po'], ascending=[True, True]).drop(columns=['id_sort_key'])
        elif "Folio PO" in sort_h:
            df_hist_filtered = df_hist_filtered.sort_values(by=['po'], ascending=[False])
        elif "Importe Total" in sort_h:
            df_hist_filtered = df_hist_filtered.sort_values(by=['total'], ascending=[False])
            
        cols_h_show = [c for c in [
            'fecha_registro', 'fecha_llegada', 'id_interno', 'po', 'proyecto',
            'articulos_count', 'piezas_requeridas', 'total',
            'archivo_correo', 'archivo_pdf', 'comprador', 'solicitante', 'estatus_remision'
        ] if c in df_hist_filtered.columns]
        
        df_h_disp = df_hist_filtered[cols_h_show].copy()
        if 'archivo_correo' in df_h_disp.columns:
            df_h_disp['archivo_correo'] = df_h_disp['archivo_correo'].apply(lambda x: '✅ .MSG' if str(x).strip() else '⚪ No')
        if 'archivo_pdf' in df_h_disp.columns:
            df_h_disp['archivo_pdf'] = df_h_disp['archivo_pdf'].apply(lambda x: '✅ .PDF' if str(x).strip() else '⚪ No')
            
        st.dataframe(
            df_h_disp.rename(columns={
                'fecha_registro': 'Fecha de Carga',
                'fecha_llegada': 'Llegada PO',
                'id_interno': 'ID Interno',
                'po': 'PO / Folio',
                'proyecto': 'Proyecto',
                'articulos_count': 'Partidas #',
                'piezas_requeridas': 'Piezas Totales',
                'total': 'Importe Total ($)',
                'archivo_correo': 'Correo .MSG',
                'archivo_pdf': 'Doc .PDF',
                'comprador': 'Comprador',
                'solicitante': 'Solicitante',
                'estatus_remision': 'Estatus'
            }),
            use_container_width=True,
            hide_index=True
        )
        
        # Descarga Excel del historial de cargas
        buf_h = io.BytesIO()
        with pd.ExcelWriter(buf_h, engine='openpyxl') as writer:
            df_hist_filtered.to_excel(writer, sheet_name="Historial_Cargas_POs", index=False)
            
        st.download_button(
            "📥 Exportar Historial de Cargas a Excel",
            data=buf_h.getvalue(),
            file_name=f"Historial_Cargas_POs_{datetime.date.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_intake_history"
        )
    else:
        st.info("💡 Aún no se han registrado órdenes en el sistema.")


# ==============================================================================
# SECCIÓN 3: AJUSTE DE PO (AJUSTE MAESTRO 1° GENERALES + 2° ARTÍCULOS)
# ==============================================================================
elif menu == "✏️ Ajuste de PO":
    st.title("✏️ Ajuste y Control Maestro de Órdenes de Compra")
    st.markdown("Navega entre las Órdenes de Compra para ajustar sus **fechas de llegada**, **identificadores internos (INT-XXXX)**, **archivos digitales** y **partidas de artículos**.")
    
    df_pos = get_all_pos()
    df_part = get_all_partidas()
    
    if df_pos.empty:
        st.info("💡 No hay POs registradas para ajustar. Cárgalas en **'📬 Bandeja de Entrada OCR'**.")
    else:
        pos_list = df_pos['po'].astype(str).tolist()
        total_pos = len(pos_list)
        
        if 'current_po_edit_idx' not in st.session_state or st.session_state['current_po_edit_idx'] >= total_pos:
            st.session_state['current_po_edit_idx'] = 0
            
        cur_idx = st.session_state['current_po_edit_idx']
        
        # Barra de Navegación Rápida
        c_nav1, c_nav2, c_nav3 = st.columns([1, 4, 1])
        with c_nav1:
            if st.button("⬅️ Anterior", use_container_width=True, disabled=(cur_idx == 0)):
                st.session_state['current_po_edit_idx'] = max(0, cur_idx - 1)
                st.rerun()
        with c_nav2:
            label_options = []
            for i, p_val in enumerate(pos_list):
                row_p = df_pos[df_pos['po'].astype(str) == p_val].iloc[0]
                int_id = str(row_p.get('id_interno', '')).strip()
                proy_txt = str(row_p.get('proyecto', '')).strip()
                prefix = f"[{int_id}] " if int_id else ""
                label_options.append(f"{i+1}/{total_pos}: {prefix}PO {p_val} • {proy_txt}")
                
            selected_label_idx = st.selectbox(
                "Seleccionar PO para Ajuste:",
                range(total_pos),
                format_func=lambda i: label_options[i],
                index=cur_idx,
                key="sb_select_po_edit_master"
            )
            if selected_label_idx != cur_idx:
                st.session_state['current_po_edit_idx'] = selected_label_idx
                st.rerun()
                
        with c_nav3:
            if st.button("Siguiente ➡️", use_container_width=True, disabled=(cur_idx == total_pos - 1)):
                st.session_state['current_po_edit_idx'] = min(total_pos - 1, cur_idx + 1)
                st.rerun()
                
        st.progress((cur_idx + 1) / total_pos, text=f"Orden {cur_idx + 1} de {total_pos} • ({((cur_idx + 1)/total_pos)*100:.1f}% del catálogo)")
        
        selected_po = pos_list[cur_idx]
        cab_row = df_pos[df_pos['po'].astype(str) == selected_po].iloc[0]
        partidas_po = df_part[df_part['po'].astype(str) == selected_po] if not df_part.empty else pd.DataFrame()
        
        # Tarjeta de Resumen (Solo Lectura)
        st.write("")
        with st.container(border=True):
            st.markdown(f"### 📋 Ficha de Control: PO `{selected_po}`")
            m1, m2, m3, m4 = st.columns(4)
            with m1:
                st.metric("3. Cantidad de Artículos (#)", f"{len(partidas_po)} Partidas", help="Número de partidas en esta PO")
            with m2:
                cant_piezas_tot = float(partidas_po['cantidad_requerida'].sum()) if not partidas_po.empty else 0.0
                st.metric("3. Cantidad Total de Piezas", f"{cant_piezas_tot:,.0f} pzas", help="Suma de todas las piezas requeridas")
            with m3:
                importe_tot = float(cab_row.get('total', 0) or 0)
                if importe_tot == 0.0 and not partidas_po.empty:
                    sum_calc = sum(
                        float(r.get('precio_total', 0) or 0) if float(r.get('precio_total', 0) or 0) > 0 
                        else (float(r.get('cantidad_requerida', 0) or 0) * float(r.get('precio_unitario', 0) or 0)) 
                        for _, r in partidas_po.iterrows()
                    )
                    if sum_calc > 0:
                        importe_tot = sum_calc
                st.metric("Importe Total ($)", f"${importe_tot:,.2f} MXN", help="Suma del importe de todas las partidas (Cantidad x Precio)")
            with m4:
                estatus_val = str(cab_row.get('estatus_general', 'Registrada'))
                st.metric("Estatus Actual", estatus_val)
                
        def _clean_str(v, d=""):
            if v is None or pd.isna(v):
                return d
            s = str(v).strip()
            return d if s.lower() in ('nan', 'none', 'null') else s

        # 1° Bloque: Ajuste de Datos Generales
        st.markdown("#### 1️⃣ Datos Generales de la PO:")
        with st.container(border=True):
            f_col1, f_col2 = st.columns(2)
            with f_col1:
                val_id_int = _clean_str(cab_row.get('id_interno'))
                new_id_interno = st.text_input(
                    "2. Nombre Interno (ej. INT-0001, INT-0059):",
                    value=val_id_int if val_id_int else f"INT-{(cur_idx+1):04d}",
                    key=f"edit_id_int_{selected_po}"
                )
                
                val_f_llegada = _clean_str(cab_row.get('fecha_llegada'))
                try:
                    default_f_llegada = datetime.datetime.strptime(val_f_llegada, "%Y-%m-%d").date() if val_f_llegada else datetime.date.today()
                except Exception:
                    default_f_llegada = datetime.date.today()
                new_fecha_llegada = st.date_input("1. Fecha de Llegada de la PO (Recepción de Correo):", value=default_f_llegada, key=f"edit_flleg_{selected_po}")
                
                val_f_solic = _clean_str(cab_row.get('fecha_solicitada'))
                try:
                    default_f_solic = datetime.datetime.strptime(val_f_solic, "%Y-%m-%d").date() if val_f_solic else (default_f_llegada + datetime.timedelta(days=14))
                except Exception:
                    default_f_solic = default_f_llegada + datetime.timedelta(days=14)
                new_fecha_solicitada = st.date_input("6. Fecha Solicitada (Compromiso Entrega):", value=default_f_solic, key=f"edit_fsolic_{selected_po}")
                
                val_comp = _clean_str(cab_row.get('comprador'))
                new_comprador = st.text_input("7. Comprador / Contacto de Compras:", value=val_comp, key=f"edit_comp_{selected_po}")
                
            with f_col2:
                val_mail = _clean_str(cab_row.get('archivo_correo'))
                new_archivo_correo = st.text_input(
                    "4. Archivo de Correo (.MSG):",
                    value=val_mail if val_mail else f"INT {(cur_idx+1):04d} - OC {selected_po} SIGRAMA METALES.msg",
                    key=f"edit_mail_{selected_po}"
                )
                
                val_pdf = _clean_str(cab_row.get('archivo_pdf'))
                new_archivo_pdf = st.text_input(
                    "5. Documento de PO (.PDF):",
                    value=val_pdf if val_pdf else f"{selected_po} SIGRAMA METALES JMC.PDF",
                    key=f"edit_pdf_{selected_po}"
                )
                
                val_proy = _clean_str(cab_row.get('proyecto'))
                val_solic = _clean_str(cab_row.get('solicitante'))
                new_proyecto = st.text_input("Proyecto / Uso:", value=val_proy, key=f"edit_proy_{selected_po}")
                new_solicitante = st.text_input("Solicitante / Requisición:", value=val_solic, key=f"edit_sol_{selected_po}")
                
            val_obs = _clean_str(cab_row.get('observaciones'))
            new_observaciones = st.text_area("Observaciones y Notas de Control:", value=val_obs, height=60, key=f"edit_obs_{selected_po}")
            
        # 2° Bloque: Ajuste de la Tabla de Artículos
        st.markdown("#### 2️⃣ Tabla de Artículos / Partidas (Editable):")
        st.caption("Puedes modificar cualquier celda directamente (SKU Cliente, SKU Planta, Cantidades, Precios o Fechas) o agregar nuevos renglones.")
        
        if not partidas_po.empty:
            cols_edit_order = ['item_no', 'sku_cliente', 'clave_sku', 'descripcion_producto', 'cantidad_requerida', 'unidad', 'precio_unitario', 'precio_total', 'fecha_entrega']
            df_for_editor = partidas_po[[c for c in cols_edit_order if c in partidas_po.columns]].copy()
            
            edited_df_partidas = st.data_editor(
                df_for_editor,
                column_config={
                    'item_no': st.column_config.NumberColumn("Item #", disabled=True),
                    'sku_cliente': st.column_config.TextColumn("SKU Cliente (Clave)"),
                    'clave_sku': st.column_config.TextColumn("SKU Nuestro (Planta)"),
                    'descripcion_producto': st.column_config.TextColumn("Descripción del Producto", width="large"),
                    'cantidad_requerida': st.column_config.NumberColumn("Cantidad", min_value=1.0, format="%.2f"),
                    'unidad': st.column_config.SelectboxColumn("Unidad", options=["PIEZA", "PZA", "KG", "METRO", "JGO", "LOTE", "SER"]),
                    'precio_unitario': st.column_config.NumberColumn("P. Unitario ($)", format="$%.2f"),
                    'precio_total': st.column_config.NumberColumn("P. Total ($)", format="$%.2f"),
                    'fecha_entrega': st.column_config.TextColumn("Fecha Entrega")
                },
                use_container_width=True,
                num_rows="dynamic",
                key=f"data_editor_partidas_{selected_po}"
            )
        else:
            edited_df_partidas = pd.DataFrame()
            st.info("No hay partidas registradas para esta PO.")
            
        # Botones de Guardado
        st.write("")
        btn_col1, btn_col2 = st.columns(2)
        with btn_col1:
            save_clicked = st.button("💾 Guardar Cambios de la PO (Generales + Artículos)", type="primary", use_container_width=True, key=f"btn_save_po_{selected_po}")
        with btn_col2:
            save_next_clicked = st.button("💾 Guardar y Pasar al Siguiente ➡️", use_container_width=True, key=f"btn_save_next_po_{selected_po}")
            
        if save_clicked or save_next_clicked:
            # Construir cabecera actualizada
            updated_cab = dict(cab_row)
            updated_cab['id_interno'] = str(new_id_interno).strip().upper()
            updated_cab['fecha_llegada'] = str(new_fecha_llegada)
            updated_cab['fecha_solicitada'] = str(new_fecha_solicitada)
            updated_cab['archivo_correo'] = str(new_archivo_correo).strip()
            updated_cab['archivo_pdf'] = str(new_archivo_pdf).strip()
            updated_cab['comprador'] = str(new_comprador).strip()
            updated_cab['proyecto'] = str(new_proyecto).strip()
            updated_cab['solicitante'] = str(new_solicitante).strip()
            updated_cab['observaciones'] = str(new_observaciones).strip()
            
            # Recalcular totales según tabla editada (Precio x Cantidad)
            partidas_list = edited_df_partidas.to_dict('records') if not edited_df_partidas.empty else []
            for p_item in partidas_list:
                c_q = float(p_item.get('cantidad_requerida', 0) or 0)
                p_u = float(p_item.get('precio_unitario', 0) or 0)
                p_t = float(p_item.get('precio_total', 0) or 0)
                if p_t == 0.0 and p_u > 0:
                    p_item['precio_total'] = round(c_q * p_u, 2)
                elif p_u > 0:
                    p_item['precio_total'] = round(c_q * p_u, 2)
                    
            sub_calc = sum(float(r.get('precio_total', 0) or (float(r.get('cantidad_requerida', 0) or 0) * float(r.get('precio_unitario', 0) or 0))) for r in partidas_list)
            updated_cab['total'] = round(sub_calc, 2)
            updated_cab['subtotal'] = round(sub_calc / 1.16, 2) if updated_cab.get('subtotal', 0) == 0 else updated_cab['subtotal']
            updated_cab['iva'] = round(updated_cab['total'] - updated_cab['subtotal'], 2)
            
            ok_save, msg_save = save_po(updated_cab, partidas_list)
            if ok_save:
                st.success(f"✅ ¡PO {selected_po} actualizada con éxito!")
                if save_next_clicked and cur_idx < total_pos - 1:
                    st.session_state['current_po_edit_idx'] = cur_idx + 1
                st.rerun()
            else:
                st.error(f"❌ {msg_save}")


# ==============================================================================
# SECCIÓN 4: MATRIZ DE ÓRDENES (CATÁLOGO MAESTRO ORDENADO)
# ==============================================================================
elif menu == "📋 Matriz de Órdenes":
    st.title("📋 Matriz Maestra de Órdenes de Compra")
    st.markdown("Consulta, búsqueda global y ordenamiento de todas las POs registradas con sus identificadores internos y estatus de remisión.")
    
    df_pos = get_all_pos()
    df_part = get_all_partidas()
    
    if df_pos.empty:
        st.info("💡 No hay POs registradas. Cárgalas en **'📬 Bandeja de Entrada OCR'**.")
    else:
        df_summary = get_global_pos_tracking_summary(df_pos, df_part)
        
        f1, f2, f3, f4 = st.columns([2, 1, 1, 1.5])
        with f1:
            q_search = st.text_input("🔍 Búsqueda rápida (ID Interno, Folio, Proyecto, Comprador):", "")
        with f2:
            estatus_opts = ["Todos"] + list(df_summary['estatus_remision'].unique())
            sel_estatus = st.selectbox("Filtrar por Estatus:", estatus_opts)
        with f3:
            proy_opts = ["Todos"] + [p for p in df_summary['proyecto'].dropna().unique() if str(p).strip()]
            sel_proy = st.selectbox("Filtrar por Proyecto:", proy_opts)
        with f4:
            sort_opt = st.selectbox("Ordenar Matriz por:", [
                "🔢 ID Interno (INT-0001, INT-0002...)",
                "📅 Fecha de Llegada (Más reciente)",
                "📄 Folio PO",
                "🎯 % Avance Cumplimiento",
                "💰 Importe Total ($)"
            ])
            
        df_filtered = df_summary.copy()
        
        if q_search:
            q = q_search.strip().lower()
            df_filtered = df_filtered[
                df_filtered['po'].astype(str).str.lower().str.contains(q) |
                df_filtered.get('id_interno', pd.Series(['']*len(df_filtered))).astype(str).str.lower().str.contains(q) |
                df_filtered['proyecto'].astype(str).str.lower().str.contains(q) |
                df_filtered['solicitante'].astype(str).str.lower().str.contains(q) |
                df_filtered['comprador'].astype(str).str.lower().str.contains(q) |
                df_filtered['observaciones'].astype(str).str.lower().str.contains(q)
            ]
            
        if sel_estatus != "Todos":
            df_filtered = df_filtered[df_filtered['estatus_remision'] == sel_estatus]
            
        if sel_proy != "Todos":
            df_filtered = df_filtered[df_filtered['proyecto'] == sel_proy]
            
        # Aplicar ordenamiento
        if "ID Interno" in sort_opt:
            df_filtered['id_sort_key'] = df_filtered['id_interno'].apply(lambda x: str(x) if str(x).strip() else 'ZZZ')
            df_filtered = df_filtered.sort_values(by=['id_sort_key', 'po'], ascending=[True, True]).drop(columns=['id_sort_key'])
        elif "Fecha de Llegada" in sort_opt:
            df_filtered = df_filtered.sort_values(by=['fecha_llegada', 'fecha_pedido', 'po'], ascending=[False, False, False])
        elif "Folio PO" in sort_opt:
            df_filtered = df_filtered.sort_values(by=['po'], ascending=[True])
        elif "% Avance" in sort_opt:
            df_filtered = df_filtered.sort_values(by=['pct_cumplimiento'], ascending=[False])
        elif "Importe Total" in sort_opt:
            df_filtered = df_filtered.sort_values(by=['total'], ascending=[False])
            
        st.write(f"Mostrando **{len(df_filtered)}** de **{len(df_summary)}** Órdenes de Compra ordenadas:")
        
        # Columnas a desplegar
        cols_present = [c for c in [
            'id_interno', 'po', 'fecha_llegada', 'fecha_solicitada', 'proyecto',
            'articulos_count', 'piezas_requeridas', 'piezas_remisionadas', 'piezas_pendientes',
            'pct_cumplimiento', 'archivo_correo', 'archivo_pdf', 'comprador',
            'estatus_remision', 'total'
        ] if c in df_filtered.columns]
        
        # Formatear indicadores de archivos
        df_display = df_filtered[cols_present].copy()
        if 'archivo_correo' in df_display.columns:
            df_display['archivo_correo'] = df_display['archivo_correo'].apply(lambda x: '✅ .MSG' if str(x).strip() else '⚪ No')
        if 'archivo_pdf' in df_display.columns:
            df_display['archivo_pdf'] = df_display['archivo_pdf'].apply(lambda x: '✅ .PDF' if str(x).strip() else '⚪ No')
            
        st.dataframe(
            df_display.rename(columns={
                'id_interno': 'ID Interno',
                'po': 'PO / Folio',
                'fecha_llegada': 'Llegada PO',
                'fecha_solicitada': 'Fecha Solicitada',
                'proyecto': 'Proyecto',
                'articulos_count': 'Artículos #',
                'piezas_requeridas': 'Cant. Req.',
                'piezas_remisionadas': 'Enviadas',
                'piezas_pendientes': 'Pendientes',
                'pct_cumplimiento': '% Avance',
                'archivo_correo': 'Correo .MSG',
                'archivo_pdf': 'Doc .PDF',
                'comprador': 'Comprador',
                'estatus_remision': 'Estatus Entrega',
                'total': 'Importe Total'
            }),
            use_container_width=True,
            hide_index=True
        )
        
        # Botón para descargar reporte Excel
        buf_exp = io.BytesIO()
        with pd.ExcelWriter(buf_exp, engine='openpyxl') as writer:
            df_filtered.to_excel(writer, sheet_name="Matriz_POs", index=False)
            
        st.download_button(
            label="📥 Exportar Matriz Filtrada a Excel",
            data=buf_exp.getvalue(),
            file_name=f"Reporte_Matriz_POs_{datetime.date.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )


# ==============================================================================
# SECCIÓN 5: FICHA DE TRAZABILIDAD 360° (INVESTIGACIÓN & CRUCE)
# ==============================================================================
elif menu == "🔍 Ficha de Trazabilidad 360°":
    st.title("🔍 Ficha de Trazabilidad 360° por Orden de Compra")
    st.markdown("Consulta en profundidad el detalle oficial de la PO, desglose de partidas vs envíos de remisiones y auditoría.")
    
    df_pos = get_all_pos()
    
    if df_pos.empty:
        st.info("💡 No hay Órdenes de Compra registradas.")
    else:
        col_sel1, col_sel2 = st.columns([2.5, 1.5])
        with col_sel1:
            pos_list = df_pos['po'].tolist()
            def _format_po_option(p):
                sub_df = df_pos[df_pos['po'].astype(str) == str(p)]
                if not sub_df.empty:
                    r = sub_df.iloc[0]
                    int_id = str(r.get('id_interno', '')).strip()
                    proy = str(r.get('proyecto', '')).strip()
                    prefix = f"[{int_id}] " if int_id else ""
                    return f"{prefix}PO {p} • {proy}" if proy else f"{prefix}PO {p}"
                return f"PO {p}"
                
            sel_po = st.selectbox("Selecciona la Orden de Compra a inspeccionar:", pos_list, format_func=_format_po_option, key="sb_select_po_360")
            
        df_cab, df_partidas_po = get_po_by_folio(sel_po)
        
        if df_cab.empty:
            st.error("No se encontraron datos de la PO seleccionada.")
        else:
            cab_info = df_cab.iloc[0]
            
            # Botón de Sincronización en Vivo
            sync_btn_col1, sync_btn_col2 = st.columns([3, 1])
            with sync_btn_col1:
                st.caption("Consulta el estado en vivo de las piezas en las máquinas de Corte-Doblez y los despachos en Remisiones.")
            with sync_btn_col2:
                if st.button("🔄 Sincronizar Estatus 360°", type="primary", use_container_width=True, key="btn_sync_360_live"):
                    from remisiones_sync import sync_live_remisiones_from_github
                    with st.spinner("Sincronizando con Almacén y Remisiones en vivo..."):
                        sync_live_remisiones_from_github()
                    st.cache_data.clear()
                    st.toast("✅ ¡Bases de datos de Almacén y Remisiones sincronizadas en tiempo real!")
                    st.rerun()

            id_int_txt = str(cab_info.get('id_interno', '')).strip()
            
            # Consultar ambas aplicaciones con Folio PO e ID Interno
            rem_tracking = get_tracking_for_po(sel_po, df_partidas_po, id_interno=id_int_txt)
            cd_tracking = get_corte_doblez_tracking_for_po(sel_po, df_partidas_po, id_interno=id_int_txt)
            
            tot_req = float(rem_tracking.get('total_requerido', 0.0) or 0.0)
            tot_fab = float(cd_tracking.get('total_fabricado', cd_tracking.get('total_terminado_planta', 0.0)) or 0.0)
            tot_rem = float(rem_tracking.get('total_remisionado', 0.0) or 0.0)
            tot_pend_env = max(0.0, tot_req - tot_rem)
            
            pct_fab = float(cd_tracking.get('porcentaje_fabricacion', cd_tracking.get('pct_global_fabricacion', (tot_fab / tot_req * 100.0) if tot_req > 0 else 0.0)))
            pct_rem = float(rem_tracking.get('porcentaje_global', (tot_rem / tot_req * 100.0) if tot_req > 0 else 0.0))
            
            # Determinar Estatus 360 Global
            if tot_rem >= tot_req and tot_req > 0:
                estatus_360 = "Remisionada Total (100%)"
                estatus_color = "#10B981"
            elif tot_rem > 0:
                estatus_360 = f"Remisionada Parcial ({pct_rem:.1f}%)"
                estatus_color = "#3B82F6"
            elif tot_fab >= tot_req and tot_req > 0:
                estatus_360 = f"Listo para Remisión ({pct_fab:.1f}% Fab)"
                estatus_color = "#8B5CF6"
            elif tot_fab > 0:
                estatus_360 = f"En Proceso de Fabricación ({pct_fab:.1f}% Fab)"
                estatus_color = "#F59E0B"
            else:
                estatus_360 = "Registrada (En Espera)"
                estatus_color = "#64748B"
            
            id_int_txt = str(cab_info.get('id_interno', '')).strip()
            id_int_badge = f'<span style="background-color:#EC2024; color:#FFFFFF; padding:6px 20px; border-radius:8px; font-weight:900; font-size:28px; margin-right:12px; display:inline-block; letter-spacing:1px; box-shadow:0 2px 4px rgba(0,0,0,0.3);">{id_int_txt}</span>' if id_int_txt else '<span style="background-color:#555; color:#FFF; padding:6px 16px; border-radius:8px; font-size:20px; margin-right:12px;">SIN ID</span>'
            
            with col_sel2:
                c_s1, c_s2 = st.columns(2)
                with c_s1:
                    st.metric("ID Interno", id_int_txt if id_int_txt else "N/A")
                with c_s2:
                    st.metric("Llegada PO", str(cab_info.get('fecha_llegada', 'N/A')))
            
            st.markdown(f"""
            <div style="background:#18181B; color:#FFFFFF; padding:24px 28px; border-radius:12px; border-left:8px solid #EC2024; margin-bottom:20px; box-shadow: 0 6px 12px rgba(0,0,0,0.25);">
                <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:14px;">
                    <div>
                        <div style="display:flex; align-items:center; gap:12px; margin-bottom:8px; flex-wrap:wrap;">
                            {id_int_badge}
                            <span style="font-size:36px; font-weight:900; color:#FFFFFF; letter-spacing:-0.5px;">ORDEN DE COMPRA: <span style="color:#EC2024;">{sel_po}</span></span>
                        </div>
                        <p style="margin:6px 0 0 0; font-size:16px; color:#E5E7EB;">
                            🏗️ Proyecto: <b style="color:#FFFFFF; font-size:17px;">{cab_info.get('proyecto', 'N/A')}</b> &nbsp;|&nbsp; 
                            👤 Solicitante: <b style="color:#FFFFFF;">{cab_info.get('solicitante', 'N/A')}</b> &nbsp;|&nbsp; 
                            💼 Comprador: <b style="color:#FFFFFF;">{cab_info.get('comprador', 'N/A')}</b>
                        </p>
                    </div>
                    <div style="text-align:right;">
                        <span style="background-color:{estatus_color}; color:#FFFFFF; padding:10px 22px; border-radius:25px; font-weight:bold; font-size:17px; display:inline-block; box-shadow:0 3px 6px rgba(0,0,0,0.3);">
                            ● {estatus_360}
                        </span>
                    </div>
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            # 5 Tarjetas de Métricas Clave de la Cadena de Suministro (Diseño Premium con Iconos Grandes)
            tot_ent = float(rem_tracking.get('total_entarimado', tot_rem) or 0)
            pct_ent = (tot_ent / tot_req * 100.0) if tot_req > 0 else 0.0
            ofs_cnt = len(cd_tracking.get('ofs_asociadas', []))
            ofs_label = f"{ofs_cnt} OF(s)" if ofs_cnt > 0 else "Sin OF"
            rem_cnt = len(rem_tracking.get('remisiones_asociadas', []))
            rem_label = f"{rem_cnt} Remisión(es)" if rem_cnt > 0 else "Sin remisión"
            
            k1, k2, k3, k4, k5 = st.columns(5)
            
            with k1:
                st.markdown(f"""
                <div style="background:#FFFFFF; border:1px solid #E2E8F0; border-top:4px solid #475569; border-radius:10px; padding:12px 14px; box-shadow:0 3px 6px rgba(0,0,0,0.04); display:flex; flex-direction:column; justify-content:space-between; min-height:120px;">
                    <div style="display:flex; justify-content:space-between; align-items:flex-start;">
                        <div>
                            <span style="font-size:11px; font-weight:800; text-transform:uppercase; letter-spacing:0.5px; color:#475569;">1. REQUERIDAS</span>
                            <div style="font-size:24px; font-weight:900; color:#0F172A; line-height:1.2; margin-top:4px;">
                                {tot_req:,.0f} <span style="font-size:12px; font-weight:600; color:#64748B;">pzas</span>
                            </div>
                        </div>
                        <div style="background-color:#F1F5F9; width:44px; height:44px; border-radius:10px; display:flex; align-items:center; justify-content:center; flex-shrink:0;">
                            <svg width="25" height="25" viewBox="0 0 24 24" fill="none" stroke="#334155" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2"></path><rect x="8" y="2" width="8" height="4" rx="1" ry="1"></rect><line x1="9" y1="12" x2="15" y2="12"></line><line x1="9" y1="16" x2="13" y2="16"></line></svg>
                        </div>
                    </div>
                    <div style="margin-top:10px;">
                        <span style="background-color:#F1F5F9; color:#334155; font-size:10.5px; font-weight:700; padding:3px 8px; border-radius:12px; display:inline-block;">
                            📋 100% Pedido Original
                        </span>
                    </div>
                </div>
                """, unsafe_allow_html=True)
                
            with k2:
                st.markdown(f"""
                <div style="background:#FFFFFF; border:1px solid #E2E8F0; border-top:4px solid #2563EB; border-radius:10px; padding:12px 14px; box-shadow:0 3px 6px rgba(0,0,0,0.04); display:flex; flex-direction:column; justify-content:space-between; min-height:120px;">
                    <div style="display:flex; justify-content:space-between; align-items:flex-start;">
                        <div>
                            <span style="font-size:11px; font-weight:800; text-transform:uppercase; letter-spacing:0.5px; color:#1D4ED8;">2. FABRICADAS</span>
                            <div style="font-size:24px; font-weight:900; color:#0F172A; line-height:1.2; margin-top:4px;">
                                {tot_fab:,.0f} <span style="font-size:12px; font-weight:600; color:#64748B;">pzas</span>
                            </div>
                        </div>
                        <div style="background-color:#EFF6FF; width:44px; height:44px; border-radius:10px; display:flex; align-items:center; justify-content:center; flex-shrink:0;">
                            <svg width="25" height="25" viewBox="0 0 24 24" fill="none" stroke="#2563EB" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="3"></circle><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1 0 2.83 2 2 0 0 1-2.83 0l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-2 2 2 2 0 0 1-2-2v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83 0 2 2 0 0 1 0-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1-2-2 2 2 0 0 1 2-2h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 0-2.83 2 2 0 0 1 2.83 0l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 2-2 2 2 0 0 1 2 2v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 0 2 2 0 0 1 0 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 2 2 2 2 0 0 1-2 2h-.09a1.65 1.65 0 0 0-1.51 1z"></path></svg>
                        </div>
                    </div>
                    <div style="margin-top:10px;">
                        <span style="background-color:#DBEAFE; color:#1D4ED8; font-size:10.5px; font-weight:700; padding:3px 8px; border-radius:12px; display:inline-block;">
                            ↑ {pct_fab:.1f}% Fab. ({ofs_label})
                        </span>
                    </div>
                </div>
                """, unsafe_allow_html=True)
                
            with k3:
                st.markdown(f"""
                <div style="background:#FFFFFF; border:1px solid #E2E8F0; border-top:4px solid #D97706; border-radius:10px; padding:12px 14px; box-shadow:0 3px 6px rgba(0,0,0,0.04); display:flex; flex-direction:column; justify-content:space-between; min-height:120px;">
                    <div style="display:flex; justify-content:space-between; align-items:flex-start;">
                        <div>
                            <span style="font-size:11px; font-weight:800; text-transform:uppercase; letter-spacing:0.5px; color:#B45309;">3. ENTARIMADAS</span>
                            <div style="font-size:24px; font-weight:900; color:#0F172A; line-height:1.2; margin-top:4px;">
                                {tot_ent:,.0f} <span style="font-size:12px; font-weight:600; color:#64748B;">pzas</span>
                            </div>
                        </div>
                        <div style="background-color:#FEF3C7; width:44px; height:44px; border-radius:10px; display:flex; align-items:center; justify-content:center; flex-shrink:0;">
                            <svg width="25" height="25" viewBox="0 0 24 24" fill="none" stroke="#D97706" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z"></path><polyline points="3.27 6.96 12 12.01 20.73 6.96"></polyline><line x1="12" y1="22.08" x2="12" y2="12"></line></svg>
                        </div>
                    </div>
                    <div style="margin-top:10px;">
                        <span style="background-color:#FEF3C7; color:#B45309; font-size:10.5px; font-weight:700; padding:3px 8px; border-radius:12px; display:inline-block;">
                            ↑ {pct_ent:.1f}% en Tarimas (PT)
                        </span>
                    </div>
                </div>
                """, unsafe_allow_html=True)
                
            with k4:
                st.markdown(f"""
                <div style="background:#FFFFFF; border:1px solid #E2E8F0; border-top:4px solid #059669; border-radius:10px; padding:12px 14px; box-shadow:0 3px 6px rgba(0,0,0,0.04); display:flex; flex-direction:column; justify-content:space-between; min-height:120px;">
                    <div style="display:flex; justify-content:space-between; align-items:flex-start;">
                        <div>
                            <span style="font-size:11px; font-weight:800; text-transform:uppercase; letter-spacing:0.5px; color:#047857;">4. REMISIONADAS</span>
                            <div style="font-size:24px; font-weight:900; color:#0F172A; line-height:1.2; margin-top:4px;">
                                {tot_rem:,.0f} <span style="font-size:12px; font-weight:600; color:#64748B;">pzas</span>
                            </div>
                        </div>
                        <div style="background-color:#D1FAE5; width:44px; height:44px; border-radius:10px; display:flex; align-items:center; justify-content:center; flex-shrink:0;">
                            <svg width="25" height="25" viewBox="0 0 24 24" fill="none" stroke="#059669" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><rect x="1" y="3" width="15" height="13"></rect><polygon points="16 8 20 8 23 11 23 16 16 16 16 8"></polygon><circle cx="5.5" cy="18.5" r="2.5"></circle><circle cx="18.5" cy="18.5" r="2.5"></circle></svg>
                        </div>
                    </div>
                    <div style="margin-top:10px;">
                        <span style="background-color:#DCFCE7; color:#15803D; font-size:10.5px; font-weight:700; padding:3px 8px; border-radius:12px; display:inline-block;">
                            ↑ {pct_rem:.1f}% Enviado ({rem_label})
                        </span>
                    </div>
                </div>
                """, unsafe_allow_html=True)
                
            with k5:
                st.markdown(f"""
                <div style="background:#FFFFFF; border:1px solid #E2E8F0; border-top:4px solid #DC2626; border-radius:10px; padding:12px 14px; box-shadow:0 3px 6px rgba(0,0,0,0.04); display:flex; flex-direction:column; justify-content:space-between; min-height:120px;">
                    <div style="display:flex; justify-content:space-between; align-items:flex-start;">
                        <div>
                            <span style="font-size:11px; font-weight:800; text-transform:uppercase; letter-spacing:0.5px; color:#B91C1C;">5. PENDIENTES</span>
                            <div style="font-size:24px; font-weight:900; color:#0F172A; line-height:1.2; margin-top:4px;">
                                {tot_pend_env:,.0f} <span style="font-size:12px; font-weight:600; color:#64748B;">pzas</span>
                            </div>
                        </div>
                        <div style="background-color:#FEE2E2; width:44px; height:44px; border-radius:10px; display:flex; align-items:center; justify-content:center; flex-shrink:0;">
                            <svg width="25" height="25" viewBox="0 0 24 24" fill="none" stroke="#DC2626" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M5 22h14"></path><path d="M5 2h14"></path><path d="M17 22v-4.172a2 2 0 0 0-.586-1.414L12 12l-4.414 4.414A2 2 0 0 0 7 17.828V22"></path><path d="M7 2v4.172a2 2 0 0 0 .586 1.414L12 12l4.414-4.414A2 2 0 0 0 17 6.172V2"></path></svg>
                        </div>
                    </div>
                    <div style="margin-top:10px;">
                        <span style="background-color:#FEE2E2; color:#B91C1C; font-size:10.5px; font-weight:700; padding:3px 8px; border-radius:12px; display:inline-block;">
                            ↓ {tot_pend_env:,.0f} pzas pendientes
                        </span>
                    </div>
                </div>
                """, unsafe_allow_html=True)
                
            st.write("---")
            
            tab_matriz_360, tab_corte_det, tab_remisiones_det, tab_reporte_eml, tab_acciones = st.tabs([
                "🎯 Matriz Integral 360° (Fabricación vs Remisión)",
                "🔵 Corte y Doblez (OFs de Planta)",
                "🚚 Remisiones y Tarimas (Envíos)",
                "📧 Reporte de Correo (.EML)",
                "⚙️ Mantenimiento de la Orden"
            ])
            
            # 1. Pestaña: Matriz Integral 360°
            with tab_matriz_360:
                st.subheader("🎯 Matriz Integral 360° por Partida")
                st.caption("Cruce detallado por número de parte entre lo Requerido, Fabricado en Corte-Doblez, Entarimado en Almacén y Despachado en Remisiones.")
                
                df_p_rem = rem_tracking['df_partidas']
                df_p_cd = cd_tracking['df_partidas_cd']
                
                if not df_p_rem.empty:
                    # Unir información de ambas fuentes
                    df_merged_360 = df_p_rem.copy()
                    if not df_p_cd.empty and 'item_no' in df_p_cd.columns:
                        cd_cols = [c for c in ['item_no', 'piezas_cortadas', 'piezas_dobladas', 'piezas_terminadas_planta', 'pct_avance_fabricacion'] if c in df_p_cd.columns]
                        df_merged_360 = df_merged_360.merge(
                            df_p_cd[cd_cols],
                            on='item_no',
                            how='left'
                        ).fillna({
                            'piezas_cortadas': 0,
                            'piezas_dobladas': 0,
                            'piezas_terminadas_planta': 0,
                            'pct_avance_fabricacion': 0.0
                        })
                        df_merged_360['cortado'] = df_merged_360['piezas_cortadas']
                        df_merged_360['doblado'] = df_merged_360['piezas_dobladas']
                        df_merged_360['terminado'] = df_merged_360['piezas_terminadas_planta']
                        df_merged_360['porcentaje_fabricacion'] = df_merged_360['pct_avance_fabricacion']
                    else:
                        df_merged_360['cortado'] = 0.0
                        df_merged_360['doblado'] = 0.0
                        df_merged_360['terminado'] = 0.0
                        df_merged_360['porcentaje_fabricacion'] = 0.0
                        
                    if 'cantidad_entarimada' in df_merged_360.columns:
                        df_merged_360['entarimado'] = df_merged_360['cantidad_entarimada']
                    else:
                        df_merged_360['entarimado'] = df_merged_360['cantidad_remisionada']
                        
                    def _calc_part_status(row):
                        req = float(row.get('cantidad_requerida', 0) or 0)
                        rem = float(row.get('cantidad_remisionada', 0) or 0)
                        ent = float(row.get('entarimado', 0) or 0)
                        fab = float(row.get('terminado', 0) or 0)
                        if rem >= req and req > 0:
                            return "🟢 Remisionado Total"
                        elif rem > 0:
                            return "🔵 Remisionado Parcial"
                        elif ent >= req and req > 0:
                            return "📦 Entarimado (Sin Remisión)"
                        elif ent > 0:
                            return "📦 Entarimado Parcial"
                        elif fab >= req and req > 0:
                            return "🟣 Listo p/ Entarimar"
                        elif fab > 0:
                            return "🟠 En Proceso Fab."
                        return "⚪ En Espera"
                        
                    df_merged_360['estatus_partida_360'] = df_merged_360.apply(_calc_part_status, axis=1)
                    
                    # -------------------------------------------------------------
                    # VISUALIZACIÓN TIPO EXCEL CON BARRAS DE DATOS EN CELDAS
                    # -------------------------------------------------------------
                    st.write("")
                    col_v1, col_v2 = st.columns([2, 1])
                    with col_v1:
                        filtro_partida = st.text_input("🔍 Filtrar partidas (SKU, Descripción o Item):", "", key=f"filter_matriz_{sel_po}")
                    with col_v2:
                        modo_vista = st.radio(
                            "Formato de Visualización:",
                            ["📊 Vista Tipo Excel (Barras en Celdas)", "📑 Tabla Grid Clásica"],
                            horizontal=True,
                            key=f"vista_mode_{sel_po}"
                        )
                        
                    df_to_render = df_merged_360.copy()
                    if filtro_partida.strip():
                        q_f = filtro_partida.strip().upper()
                        df_to_render = df_to_render[
                            df_to_render['sku_cliente'].astype(str).str.upper().str.contains(q_f) |
                            df_to_render['clave_sku'].astype(str).str.upper().str.contains(q_f) |
                            df_to_render['descripcion_producto'].astype(str).str.upper().str.contains(q_f) |
                            df_to_render['item_no'].astype(str).str.contains(q_f)
                        ]
                        
                    if modo_vista == "📊 Vista Tipo Excel (Barras en Celdas)":
                        st.caption("ℹ️ **Formato Excel (Barras de Datos):** La longitud de la barra en cada celda muestra el porcentaje completado respecto al **Requerido (PO)** para comparar inmediatamente el cuello de botella entre **Cortado**, **Doblado**, **Entarimado** y **Remisionado**.")
                        
                        import streamlit.components.v1 as components
                        
                        html_table = """
                        <!DOCTYPE html>
                        <html>
                        <head>
                        <meta charset="utf-8">
                        <style>
                            body { margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; font-size: 12px; background: transparent; }
                            .table-container { overflow-x: auto; max-height: 600px; border: 1px solid #CBD5E1; border-radius: 8px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }
                            table { width: 100%; border-collapse: collapse; font-size: 12px; }
                            thead th { position: sticky; top: 0; background-color: #0F172A; color: #FFFFFF; z-index: 10; padding: 10px 8px; font-weight: 700; text-align: center; }
                            tbody tr { border-bottom: 1px solid #E2E8F0; background-color: #FFFFFF; }
                            tbody tr:hover { background-color: #F1F5F9 !important; }
                            td { padding: 6px 8px; }
                            .badge { font-size: 10.5px; padding: 3px 8px; border-radius: 12px; font-weight: bold; display: inline-block; white-space: nowrap; }
                        </style>
                        </head>
                        <body>
                        <div class="table-container">
                        <table>
                            <thead>
                                <tr>
                                    <th style="border-bottom:2px solid #EC2024; width:40px;">#</th>
                                    <th style="border-bottom:2px solid #EC2024; text-align:left; min-width:110px;">SKU Cliente</th>
                                    <th style="border-bottom:2px solid #EC2024; text-align:left; min-width:130px;">SKU Planta</th>
                                    <th style="border-bottom:2px solid #EC2024; text-align:left; min-width:200px;">Descripción</th>
                                    <th style="border-bottom:2px solid #EC2024; text-align:right; width:75px;">Req. (PO)</th>
                                    <th style="border-bottom:2px solid #3B82F6; min-width:120px; background-color:#1E3A8A;">🔵 Cortado</th>
                                    <th style="border-bottom:2px solid #6366F1; min-width:120px; background-color:#312E81;">🟣 Doblado</th>
                                    <th style="border-bottom:2px solid #F59E0B; min-width:120px; background-color:#78350F;">📦 Entarimado</th>
                                    <th style="border-bottom:2px solid #10B981; min-width:120px; background-color:#064E3B;">🟢 Remisionadas</th>
                                    <th style="border-bottom:2px solid #EF4444; min-width:110px; background-color:#7C2D12;">⏳ Pendiente</th>
                                    <th style="border-bottom:2px solid #EC2024; width:140px;">Estatus 360°</th>
                                </tr>
                            </thead>
                            <tbody>
                        """
                        
                        for _, r_bar in df_to_render.iterrows():
                            i_no = r_bar.get('item_no', '')
                            sk_c = r_bar.get('sku_cliente', '')
                            sk_p = r_bar.get('clave_sku', '')
                            desc = r_bar.get('descripcion_producto', '')
                            c_req = float(r_bar.get('cantidad_requerida', 0) or 0)
                            base_div = c_req if c_req > 0 else 1.0
                            
                            c_cort = float(r_bar.get('cortado', 0) or 0)
                            c_dobl = float(r_bar.get('doblado', 0) or 0)
                            c_ent = float(r_bar.get('entarimado', r_bar.get('cantidad_entarimada', 0)) or 0)
                            c_rem = float(r_bar.get('cantidad_remisionada', 0) or 0)
                            c_pend = float(r_bar.get('cantidad_pendiente', max(0.0, c_req - c_rem)) or 0)
                            
                            pct_c = min(100.0, max(0.0, (c_cort / base_div * 100.0)))
                            pct_d = min(100.0, max(0.0, (c_dobl / base_div * 100.0)))
                            pct_e = min(100.0, max(0.0, (c_ent / base_div * 100.0)))
                            pct_r = min(100.0, max(0.0, (c_rem / base_div * 100.0)))
                            pct_p = min(100.0, max(0.0, (c_pend / base_div * 100.0)))
                            
                            st_txt = str(r_bar.get('estatus_partida_360', '⚪ En Espera'))
                            if "Remisionado Total" in st_txt:
                                b_bg, b_fg = "#DCFCE7", "#15803D"
                            elif "Remisionado Parcial" in st_txt:
                                b_bg, b_fg = "#DBEAFE", "#1D4ED8"
                            elif "Entarimado" in st_txt:
                                b_bg, b_fg = "#FEF3C7", "#B45309"
                            elif "Listo" in st_txt:
                                b_bg, b_fg = "#F3E8FF", "#6B21A8"
                            elif "Proceso" in st_txt:
                                b_bg, b_fg = "#FFF7ED", "#C2410C"
                            else:
                                b_bg, b_fg = "#F1F5F9", "#64748B"
                                
                            html_table += f"""
                            <tr>
                                <td style="text-align:center; font-weight:bold; color:#64748B;">{i_no}</td>
                                <td style="font-weight:600; color:#1E293B;">{sk_c}</td>
                                <td style="font-weight:700; color:#EC2024;">{sk_p}</td>
                                <td style="color:#475569; font-size:11.5px; max-width:240px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;" title="{desc}">{desc}</td>
                                <td style="text-align:right; font-weight:800; color:#0F172A; background-color:#F8FAFC;">{c_req:,.0f}</td>
                                
                                <!-- BARRA CORTADO -->
                                <td style="background:linear-gradient(90deg, rgba(59,130,246,0.38) {pct_c:.1f}%, transparent {pct_c:.1f}%); border-left:1px solid #E2E8F0; border-right:1px solid #E2E8F0;">
                                    <div style="display:flex; justify-content:space-between; align-items:center;">
                                        <b style="color:#1D4ED8; font-size:12.5px;">{c_cort:,.0f}</b>
                                        <span style="font-size:10.5px; color:#2563EB; font-weight:bold;">{pct_c:.1f}%</span>
                                    </div>
                                </td>
                                
                                <!-- BARRA DOBLADO -->
                                <td style="background:linear-gradient(90deg, rgba(99,102,241,0.38) {pct_d:.1f}%, transparent {pct_d:.1f}%); border-right:1px solid #E2E8F0;">
                                    <div style="display:flex; justify-content:space-between; align-items:center;">
                                        <b style="color:#4338CA; font-size:12.5px;">{c_dobl:,.0f}</b>
                                        <span style="font-size:10.5px; color:#4F46E5; font-weight:bold;">{pct_d:.1f}%</span>
                                    </div>
                                </td>
                                
                                <!-- BARRA ENTARIMADO -->
                                <td style="background:linear-gradient(90deg, rgba(245,158,11,0.38) {pct_e:.1f}%, transparent {pct_e:.1f}%); border-right:1px solid #E2E8F0;">
                                    <div style="display:flex; justify-content:space-between; align-items:center;">
                                        <b style="color:#B45309; font-size:12.5px;">{c_ent:,.0f}</b>
                                        <span style="font-size:10.5px; color:#D97706; font-weight:bold;">{pct_e:.1f}%</span>
                                    </div>
                                </td>
                                
                                <!-- BARRA REMISIONADAS -->
                                <td style="background:linear-gradient(90deg, rgba(16,185,129,0.40) {pct_r:.1f}%, transparent {pct_r:.1f}%); border-right:1px solid #E2E8F0;">
                                    <div style="display:flex; justify-content:space-between; align-items:center;">
                                        <b style="color:#15803D; font-size:12.5px;">{c_rem:,.0f}</b>
                                        <span style="font-size:10.5px; color:#16A34A; font-weight:bold;">{pct_r:.1f}%</span>
                                    </div>
                                </td>
                                
                                <!-- BARRA PENDIENTES -->
                                <td style="background:linear-gradient(90deg, rgba(239,68,68,0.22) {pct_p:.1f}%, transparent {pct_p:.1f}%); border-right:1px solid #E2E8F0;">
                                    <div style="display:flex; justify-content:space-between; align-items:center;">
                                        <b style="color:#B91C1C; font-size:12.5px;">{c_pend:,.0f}</b>
                                        <span style="font-size:10.5px; color:#DC2626; font-weight:bold;">{pct_p:.1f}%</span>
                                    </div>
                                </td>
                                
                                <!-- ESTATUS PILL -->
                                <td style="text-align:center;">
                                    <span class="badge" style="background-color:{b_bg}; color:{b_fg};">
                                        {st_txt}
                                    </span>
                                </td>
                            </tr>
                            """
                        html_table += """
                            </tbody>
                        </table>
                        </div>
                        </body>
                        </html>
                        """
                        calc_h = min(750, max(280, 80 + len(df_to_render) * 44))
                        components.html(html_table, height=calc_h, scrolling=True)
                    else:
                        cols_show_360 = [
                            'item_no', 'sku_cliente', 'clave_sku', 'descripcion_producto',
                            'cantidad_requerida', 'cortado', 'doblado', 'terminado', 'entarimado',
                            'cantidad_remisionada', 'porcentaje_cumplimiento', 'cantidad_pendiente', 'estatus_partida_360'
                        ]
                        
                        st.dataframe(
                            df_to_render[[c for c in cols_show_360 if c in df_to_render.columns]].rename(columns={
                                'item_no': 'Item #',
                                'sku_cliente': 'SKU Cliente (Clave)',
                                'clave_sku': 'SKU Nuestro (Planta)',
                                'descripcion_producto': 'Descripción',
                                'cantidad_requerida': 'Req. (PO)',
                                'cortado': '🔵 Cortado',
                                'doblado': '🔵 Doblado',
                                'terminado': '🔵 Terminado Fab.',
                                'entarimado': '📦 Entarimado',
                                'cantidad_remisionada': '🟢 Remisionadas',
                                'porcentaje_cumplimiento': '🟢 % Envío',
                                'cantidad_pendiente': '⏳ Pend. Envío',
                                'estatus_partida_360': 'Estatus 360°'
                            }),
                            use_container_width=True,
                            hide_index=True
                        )
                else:
                    st.info("No hay partidas registradas para esta PO.")
                    
            # 2. Pestaña: Detalle de Corte y Doblez
            with tab_corte_det:
                st.subheader("🔵 Órdenes de Fabricación (OFs) y Estado en Taller")
                ofs_list = cd_tracking.get('ofs_asociadas', [])
                if ofs_list:
                    st.success(f"Se encontraron **{len(ofs_list)}** Órdenes de Fabricación vinculadas a esta PO:")
                    
                    # Lista estructurada de OFs
                    c_of1, c_of2 = st.columns([1, 1])
                    for idx_of, of_item in enumerate(ofs_list):
                        target_col = c_of1 if idx_of % 2 == 0 else c_of2
                        with target_col:
                            is_ppap = "PPAP" in of_item.upper()
                            badge_type = '<span style="background-color:#8B5CF6; color:#FFF; font-size:11px; padding:2px 6px; border-radius:4px; font-weight:bold;">PPAP MUESTRAS</span>' if is_ppap else '<span style="background-color:#2563EB; color:#FFF; font-size:11px; padding:2px 6px; border-radius:4px; font-weight:bold;">PRODUCCIÓN</span>'
                            st.markdown(f"""
                            <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-left:4px solid #3B82F6; padding:10px 14px; border-radius:6px; margin-bottom:8px;">
                                <div style="display:flex; justify-content:space-between; align-items:center;">
                                    <b style="color:#1E293B; font-size:13px;">🔹 {of_item}</b>
                                    {badge_type}
                                </div>
                            </div>
                            """, unsafe_allow_html=True)
                            
                    # Resumen de Láminas y Materia Prima Utilizada
                    df_lam = cd_tracking.get('df_laminas', pd.DataFrame())
                    tot_lam = cd_tracking.get('total_laminas', 0.0)
                    
                    st.write("---")
                    st.markdown("#### ✂️ Resumen de Láminas y Materia Prima Utilizada en Fabricación")
                    c_l1, c_l2 = st.columns([1, 2])
                    with c_l1:
                        st.metric("Total Láminas Cortadas", f"{tot_lam:,.0f} Hojas", help="Cantidad total de hojas de lámina procesadas en los nidos de corte")
                    with c_l2:
                        if not df_lam.empty:
                            st.dataframe(
                                df_lam.rename(columns={
                                    'material': 'Material / Calibre',
                                    'hojas_utilizadas': 'Láminas / Hojas Utilizadas',
                                    'nidos_cortados': 'Nidos Pronest'
                                }),
                                use_container_width=True,
                                hide_index=True
                            )
                else:
                    st.info("ℹ️ Aún no se han programado Órdenes de Fabricación (OFs) para esta PO en la aplicación de Corte y Doblez.")
                    
            # 3. Pestaña: Detalle de Remisiones
            with tab_remisiones_det:
                st.subheader("🚚 Envíos Registrados en la App de Remisiones")
                df_env = rem_tracking['df_historial_envios']
                rem_list = rem_tracking.get('remisiones_asociadas', [])
                
                from remision_pdf_generator import generate_remision_pdf
                rem_folio_name = rem_list[0] if rem_list else "E0125"
                datos_rem_pdf = {
                    'Folio_Remision': rem_folio_name,
                    'PO': sel_po,
                    'Proyecto_Interno': id_int_txt,
                    'Fecha_Hora_Salida': '25/08/2026',
                    'Nombre_Emisor': 'SIGRAMA METALES',
                    'Nombre_Receptor': 'PLANTA RIO XIX',
                    'Tarimas_Asociadas': ', '.join(df_env['ID Tarima'].unique()) if (not df_env.empty and 'ID Tarima' in df_env.columns) else 'TPM-0511, TPM-0512, TPM-0513'
                }
                pdf_rem_bytes = generate_remision_pdf(datos_rem_pdf, df_p_rem)
                
                if not df_env.empty:
                    st.success(f"Se encontraron **{len(df_env)}** registros de tarimas/piezas enviadas en **{len(rem_list)}** remisión(es): `{', '.join(rem_list)}`")
                    
                    c_rd1, c_rd2 = st.columns([1, 1])
                    with c_rd1:
                        st.download_button(
                            label=f"📄 Descargar PDF Oficial de Remisión ({rem_folio_name})",
                            data=pdf_rem_bytes,
                            file_name=f"Remision_{rem_folio_name}_{id_int_txt}.pdf",
                            mime="application/pdf",
                            type="primary",
                            use_container_width=True
                        )
                        
                    st.dataframe(df_env, use_container_width=True, hide_index=True)
                else:
                    st.info("ℹ️ No hay registros de remisiones o tarimas despachadas para esta PO aún.")
                    st.download_button(
                        label=f"📄 Generar y Descargar PDF de Remisión Preliminar ({rem_folio_name})",
                        data=pdf_rem_bytes,
                        file_name=f"Remision_Preliminar_{rem_folio_name}_{id_int_txt}.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )
                    
            # 4. Pestaña: Reporte de Correo (.EML)
            with tab_reporte_eml:
                st.subheader("📧 Generador de Reporte Ejecutivo de Correo (.eml)")
                st.markdown("Genere un correo oficial con el **PDF de la Remisión adjunto**, listo para abrirse en **Outlook / Thunderbird** o enviarse directamente a compras.")
                
                col_eml_1, col_eml_2 = st.columns([1, 1])
                with col_eml_1:
                    dest_email = st.text_input("Para (Destinatario):", value=f"{cab_info.get('comprador', 'compras')}@sigrama.com.mx", key="eml_dest_input")
                    cc_email = st.text_input("CC (Copia):", value="operaciones@sigrama.com.mx, calidad@sigrama.com.mx", key="eml_cc_input")
                with col_eml_2:
                    nota_eml = st.text_area("Nota / Observaciones adicionales en el correo:", value="Se adjunta reporte consolidado de avance en planta, resumen de materia prima y el PDF oficial de la Remisión para su expediente.", height=100, key="eml_nota_input")
                    
                # Función constructora del reporte EML
                def build_po_eml(cab, cd_trk, rem_trk, df_m360, to_addr, cc_addr, note_txt):
                    import email.message
                    import email.utils
                    import re
                    from remision_pdf_generator import generate_remision_pdf
                    
                    def _clean_hdr(v, d=""):
                        if v is None or pd.isna(v):
                            return d
                        s = str(v).replace('\r', ' ').replace('\n', ' ').replace('\t', ' ')
                        s = re.sub(r'\s+', ' ', s).strip()
                        return s if s else d
                    
                    po_val = _clean_hdr(cab.get('po', ''), 'PO')
                    id_i = _clean_hdr(cab.get('id_interno', 'INT-S/N'), 'INT-S/N')
                    prj = _clean_hdr(cab.get('proyecto', 'PROYECTO SIGRAMA'), 'PROYECTO')
                    cmp_name = _clean_hdr(cab.get('comprador', 'Compras'), 'Compras')
                    sol_name = _clean_hdr(cab.get('solicitante', 'Solicitante'), 'Solicitante')
                    f_lleg = _clean_hdr(cab.get('fecha_llegada', 'N/A'), 'N/A')
                    f_sol = _clean_hdr(cab.get('fecha_solicitada', 'N/A'), 'N/A')
                    
                    t_req = float(rem_trk.get('total_requerido', 0.0) or 0.0)
                    t_fab = float(cd_trk.get('total_fabricado', cd_trk.get('total_terminado_planta', 0.0)) or 0.0)
                    t_rem = float(rem_trk.get('total_remisionado', 0.0) or 0.0)
                    t_pnd = max(0.0, t_req - t_rem)
                    p_fab = float(cd_trk.get('porcentaje_fabricacion', (t_fab/t_req*100.0) if t_req > 0 else 0.0))
                    p_rem = float(rem_trk.get('porcentaje_global', (t_rem/t_req*100.0) if t_req > 0 else 0.0))
                    
                    if t_rem >= t_req and t_req > 0:
                        e_txt = "REMISIÓN TOTAL (100%)"
                        e_col = "#10B981"
                    elif t_rem > 0:
                        e_txt = f"REMISIÓN PARCIAL ({p_rem:.1f}%)"
                        e_col = "#3B82F6"
                    elif t_fab >= t_req and t_req > 0:
                        e_txt = f"LISTO PARA REMISIÓN ({p_fab:.1f}% FAB)"
                        e_col = "#8B5CF6"
                    elif t_fab > 0:
                        e_txt = f"EN PROCESO DE FABRICACIÓN ({p_fab:.1f}% FAB)"
                        e_col = "#F59E0B"
                    else:
                        e_txt = "REGISTRADA (EN ESPERA)"
                        e_col = "#64748B"
                        
                    e_txt = _clean_hdr(e_txt)
                        
                    ofs_raw = cd_trk.get('ofs_asociadas', [])
                    if ofs_raw:
                        ofs_li = "".join([f"<li style='margin-bottom:4px;'><b>{of_item}</b></li>" for of_item in ofs_raw])
                        ofs_s = f"<ul style='margin:4px 0 0 16px; padding:0; color:#1E3A8A;'>{ofs_li}</ul>"
                    else:
                        ofs_s = "<p style='margin:0; color:#6B7280;'>Sin OFs registradas</p>"
                        
                    rems_raw = rem_trk.get('remisiones_asociadas', [])
                    if rems_raw:
                        rems_li = "".join([f"<li style='margin-bottom:4px;'><b>{r_item}</b> &nbsp;<span style='color:#15803D;'>(Despachado en Almacén)</span></li>" for r_item in rems_raw])
                        rems_s = f"<ul style='margin:4px 0 0 16px; padding:0; color:#14532D;'>{rems_li}</ul>"
                    else:
                        rems_s = "<p style='margin:0; color:#6B7280;'>Sin remisión generada</p>"
                        
                    # Resumen de Láminas HTML
                    df_lam_trk = cd_trk.get('df_laminas', pd.DataFrame())
                    tot_lam_trk = cd_trk.get('total_laminas', 0.0)
                    lam_rows = ""
                    if not df_lam_trk.empty:
                        for _, lr in df_lam_trk.iterrows():
                            lam_rows += f"""
                            <tr style="border-bottom:1px solid #E5E7EB; text-align:center;">
                                <td style="padding:6px 8px; text-align:left; font-weight:600;">{lr.get('material', '')}</td>
                                <td style="padding:6px 8px; font-weight:bold; color:#B45309;">{float(lr.get('hojas_utilizadas', 0)):,.0f} Hojas</td>
                                <td style="padding:6px 8px; color:#4B5563;">{lr.get('nidos_cortados', 1)} Nidos</td>
                            </tr>
                            """
                        laminas_html_box = f"""
                        <tr>
                            <td style="padding: 0 25px 20px 25px;">
                                <div style="background-color: #FFFBEB; border: 1px solid #FDE68A; border-radius: 8px; padding: 14px;">
                                    <div style="font-weight: bold; color: #B45309; font-size: 14px; margin-bottom: 8px;">
                                        ✂️ Resumen de Materia Prima y Láminas Utilizadas en Fabricación (Total: {tot_lam_trk:,.0f} Láminas)
                                    </div>
                                    <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse: collapse; font-size: 12px;">
                                        <thead>
                                            <tr style="background-color: #FEF3C7; color: #92400E; text-align: center;">
                                                <th style="padding: 6px 8px; text-align: left;">Material / Calibre</th>
                                                <th style="padding: 6px 8px;">Láminas Cortadas</th>
                                                <th style="padding: 6px 8px;">Nidos Pronest</th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            {lam_rows}
                                        </tbody>
                                    </table>
                                </div>
                            </td>
                        </tr>
                        """
                    else:
                        laminas_html_box = ""
                        
                    t_rows = ""
                    if not df_m360.empty:
                        for _, rw in df_m360.iterrows():
                            t_rows += f"""
                            <tr style="border-bottom:1px solid #E5E7EB; text-align:center;">
                                <td style="padding:10px 8px; font-weight:bold;">{rw.get('item_no', '')}</td>
                                <td style="padding:10px 8px; font-weight:600; color:#1F2937;">{rw.get('sku_cliente', '')}</td>
                                <td style="padding:10px 8px; font-weight:700; color:#EC2024;">{rw.get('clave_sku', '')}</td>
                                <td style="padding:10px 8px; text-align:left; color:#4B5563;">{rw.get('descripcion_producto', '')}</td>
                                <td style="padding:10px 8px; font-weight:bold;">{float(rw.get('cantidad_requerida', 0)):,.0f}</td>
                                <td style="padding:10px 8px; color:#2563EB;">{float(rw.get('cortado', 0)):,.0f}</td>
                                <td style="padding:10px 8px; color:#2563EB;">{float(rw.get('doblado', 0)):,.0f}</td>
                                <td style="padding:10px 8px; font-weight:bold; color:#16A34A;">{float(rw.get('cantidad_remisionada', 0)):,.0f}</td>
                                <td style="padding:10px 8px; font-weight:bold; color:{'#DC2626' if float(rw.get('cantidad_pendiente', 0)) > 0 else '#16A34A'};">{float(rw.get('cantidad_pendiente', 0)):,.0f}</td>
                                <td style="padding:10px 8px;"><span style="background-color:#F3F4F6; padding:3px 8px; border-radius:4px; font-size:12px; font-weight:600;">{rw.get('estatus_partida_360', '')}</span></td>
                            </tr>
                            """
                            
                    html_content = f"""
                    <!DOCTYPE html>
                    <html>
                    <head><meta charset="utf-8"></head>
                    <body style="font-family: Arial, Helvetica, sans-serif; background-color: #F3F4F6; margin: 0; padding: 20px;">
                        <table width="100%" cellpadding="0" cellspacing="0" style="max-width: 850px; margin: 0 auto; background-color: #FFFFFF; border-radius: 10px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.1); border: 1px solid #E5E7EB;">
                            <tr>
                                <td style="background-color: #111111; padding: 25px; border-bottom: 5px solid #EC2024;">
                                    <table width="100%">
                                        <tr>
                                            <td>
                                                <div style="display:inline-block; background-color: #EC2024; color: #FFFFFF; font-size: 22px; font-weight: 900; padding: 6px 18px; border-radius: 6px; margin-bottom: 8px;">
                                                    {id_i}
                                                </div>
                                                <h1 style="margin: 4px 0; color: #FFFFFF; font-size: 30px; font-weight: 900; letter-spacing: -0.5px;">
                                                    ORDEN DE COMPRA: <span style="color: #EC2024;">{po_val}</span>
                                                </h1>
                                                <p style="margin: 6px 0 0 0; color: #D1D5DB; font-size: 15px;">
                                                    🏗️ <b>Proyecto:</b> {prj} &nbsp;|&nbsp; 👤 <b>Solicitante:</b> {sol_name} &nbsp;|&nbsp; 💼 <b>Comprador:</b> {cmp_name}
                                                </p>
                                            </td>
                                            <td align="right" valign="top">
                                                <span style="background-color: {e_col}; color: #FFFFFF; font-size: 15px; font-weight: bold; padding: 10px 20px; border-radius: 25px; display: inline-block;">
                                                    ● {e_txt}
                                                </span>
                                            </td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                            <tr>
                                <td style="padding: 16px 25px; background-color: #F9FAFB; border-bottom: 1px solid #E5E7EB;">
                                    <table width="100%" style="font-size: 14px; color: #374151;">
                                        <tr>
                                            <td width="33%">📅 <b>Llegada PO:</b> {f_lleg}</td>
                                            <td width="33%">🎯 <b>Compromiso Entrega:</b> {f_sol}</td>
                                            <td width="34%">🏭 <b>Planta:</b> SIGRAMA METALES</td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                            <tr>
                                <td style="padding: 25px;">
                                    <table width="100%" cellpadding="0" cellspacing="10">
                                        <tr>
                                            <td width="25%" style="background-color: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 15px; text-align: center;">
                                                <div style="font-size: 11px; color: #64748B; font-weight: bold; text-transform: uppercase;">1. Requeridas</div>
                                                <div style="font-size: 26px; font-weight: 900; color: #0F172A; margin: 4px 0;">{t_req:,.0f}</div>
                                                <div style="font-size: 11px; color: #64748B;">Total en Orden</div>
                                            </td>
                                            <td width="25%" style="background-color: #EFF6FF; border: 1px solid #BFDBFE; border-radius: 8px; padding: 15px; text-align: center;">
                                                <div style="font-size: 11px; color: #1D4ED8; font-weight: bold; text-transform: uppercase;">2. Fab. en Planta</div>
                                                <div style="font-size: 26px; font-weight: 900; color: #1E40AF; margin: 4px 0;">{t_fab:,.0f}</div>
                                                <div style="font-size: 11px; color: #2563EB; font-weight: bold;">{p_fab:.1f}% Fabricado</div>
                                            </td>
                                            <td width="25%" style="background-color: #F0FDF4; border: 1px solid #BBF7D0; border-radius: 8px; padding: 15px; text-align: center;">
                                                <div style="font-size: 11px; color: #15803D; font-weight: bold; text-transform: uppercase;">3. Remisionadas</div>
                                                <div style="font-size: 26px; font-weight: 900; color: #166534; margin: 4px 0;">{t_rem:,.0f}</div>
                                                <div style="font-size: 11px; color: #16A34A; font-weight: bold;">{p_rem:.1f}% Despachado</div>
                                            </td>
                                            <td width="25%" style="background-color: #FFF7ED; border: 1px solid #FED7AA; border-radius: 8px; padding: 15px; text-align: center;">
                                                <div style="font-size: 11px; color: #C2410C; font-weight: bold; text-transform: uppercase;">4. Pendientes</div>
                                                <div style="font-size: 26px; font-weight: 900; color: #9A3412; margin: 4px 0;">{t_pnd:,.0f}</div>
                                                <div style="font-size: 11px; color: #EA580C;">Saldo Restante</div>
                                            </td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                            <tr>
                                <td style="padding: 0 25px 25px 25px;">
                                    <h3 style="margin: 0 0 12px 0; color: #111827; font-size: 18px; border-bottom: 2px solid #EC2024; padding-bottom: 6px;">
                                        📋 Desglose de Partidas y Trazabilidad 360°
                                    </h3>
                                    <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse: collapse; font-size: 13px;">
                                        <thead>
                                            <tr style="background-color: #1F2937; color: #FFFFFF; text-align: center;">
                                                <th style="padding: 10px 8px;">#</th>
                                                <th style="padding: 10px 8px;">SKU Cliente</th>
                                                <th style="padding: 10px 8px;">SKU Planta</th>
                                                <th style="padding: 10px 8px; text-align:left;">Descripción</th>
                                                <th style="padding: 10px 8px;">Req.</th>
                                                <th style="padding: 10px 8px;">Cortado</th>
                                                <th style="padding: 10px 8px;">Doblado</th>
                                                <th style="padding: 10px 8px;">Enviado</th>
                                                <th style="padding: 10px 8px;">Pendiente</th>
                                                <th style="padding: 10px 8px;">Estatus</th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            {t_rows}
                                        </tbody>
                                    </table>
                                </td>
                            </tr>
                            
                            {laminas_html_box}
                            
                            <tr>
                                <td style="padding: 0 25px 25px 25px;">
                                    <table width="100%" cellpadding="0" cellspacing="10">
                                        <tr>
                                            <td width="50%" style="background-color: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 15px; vertical-align: top;">
                                                <div style="font-weight: bold; color: #1E40AF; margin-bottom: 8px; font-size: 14px;">🔵 Órdenes de Fabricación (Corte y Doblez):</div>
                                                <div style="font-size: 13px; line-height: 1.6;">{ofs_s}</div>
                                            </td>
                                            <td width="50%" style="background-color: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 15px; vertical-align: top;">
                                                <div style="font-weight: bold; color: #166534; margin-bottom: 8px; font-size: 14px;">🚚 Remisiones y Despachos (Almacén):</div>
                                                <div style="font-size: 13px; line-height: 1.6;">{rems_s}</div>
                                                <div style="margin-top: 8px; font-size: 12px; color: #15803D; font-weight: bold;">📎 PDF oficial de remisión adjunto en este correo.</div>
                                            </td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                            {'<tr><td style="padding: 0 25px 20px 25px;"><div style="background-color:#FEF3C7; border-left:4px solid #F59E0B; padding:12px; font-size:13px; color:#92400E;"><b>Nota de Control:</b> ' + note_txt + '</div></td></tr>' if note_txt else ''}
                            <tr>
                                <td style="background-color: #111111; padding: 15px 25px; text-align: center; color: #9CA3AF; font-size: 12px;">
                                    Industria Sigrama S.A. de C.V. • PO Tracker & Master Hub 4.0 • Reporte Certificado
                                </td>
                            </tr>
                        </table>
                    </body>
                    </html>
                    """
                    
                    # Generar Expediente ZIP con los 3 documentos oficiales
                    from dossier_pdf_generator import create_dossier_zip
                    zip_data, pdf_rem_data, pdf_etiq_data, pdf_rep_data = create_dossier_zip(cab, cd_trk, rem_trk, df_m360)
                    
                    to_clean = _clean_hdr(to_addr, "compras@sigrama.com.mx")
                    cc_clean = _clean_hdr(cc_addr, "")
                    subj_clean = _clean_hdr(f"[EXPEDIENTE TRAZABILIDAD 360°] {id_i} • PO {po_val} - {prj} ({e_txt})")
                    
                    e_msg = email.message.EmailMessage()
                    e_msg['Subject'] = subj_clean
                    e_msg['From'] = "po-tracker@sigrama.com.mx"
                    e_msg['To'] = to_clean
                    if cc_clean:
                        e_msg['Cc'] = cc_clean
                    e_msg['Date'] = email.utils.formatdate(localtime=True)
                    e_msg.set_content(f"Expediente de Trazabilidad 360° para {id_i} - PO {po_val} ({prj}). Estatus: {e_txt}. Se adjunta expediente completo en ZIP conteniendo Remisión oficial, Etiquetas de Tarimas y Reporte de Trazabilidad.")
                    e_msg.add_alternative(html_content, subtype='html')
                    
                    clean_id_file = re.sub(r'[^a-zA-Z0-9_\-]', '_', id_i)
                    clean_po_file = re.sub(r'[^a-zA-Z0-9_\-]', '_', po_val)
                    
                    # 1. Adjuntar Expediente ZIP Completo
                    e_msg.add_attachment(
                        zip_data,
                        maintype='application',
                        subtype='zip',
                        filename=f"Expediente_Trazabilidad_{clean_id_file}_PO_{clean_po_file}.zip"
                    )
                    
                    # 2. Adjuntar PDFs individuales para visualización directa
                    e_msg.add_attachment(
                        pdf_rem_data,
                        maintype='application',
                        subtype='pdf',
                        filename=f"1_Remision_{clean_id_file}.pdf"
                    )
                    e_msg.add_attachment(
                        pdf_etiq_data,
                        maintype='application',
                        subtype='pdf',
                        filename=f"2_Etiquetas_Tarimas_{clean_id_file}.pdf"
                    )
                    e_msg.add_attachment(
                        pdf_rep_data,
                        maintype='application',
                        subtype='pdf',
                        filename=f"3_Reporte_Trazabilidad_PO_{clean_po_file}_{clean_id_file}.pdf"
                    )
                    
                    return e_msg.as_bytes(), html_content, zip_data, pdf_rem_data, pdf_etiq_data, pdf_rep_data
                    
                eml_bytes, eml_html, zip_att, rem_pdf_att, etiq_pdf_att, rep_pdf_att = build_po_eml(cab_info, cd_tracking, rem_tracking, df_merged_360 if 'df_merged_360' in locals() else pd.DataFrame(), dest_email, cc_email, nota_eml)
                
                st.markdown("#### 📦 Descarga de Expedientes y Archivos Certificados")
                b_c1, b_c2 = st.columns([1, 1])
                with b_c1:
                    st.download_button(
                        label=f"📥 Descargar Correo (.eml) con Expediente ZIP Adjunto",
                        data=eml_bytes,
                        file_name=f"Reporte_360_{id_int_txt if id_int_txt else 'INT'}_PO_{sel_po}.eml",
                        mime="message/rfc822",
                        type="primary",
                        use_container_width=True
                    )
                with b_c2:
                    st.download_button(
                        label=f"📦 Descargar Expediente Documental Completo (.ZIP)",
                        data=zip_att,
                        file_name=f"Expediente_Documental_{id_int_txt if id_int_txt else 'INT'}_PO_{sel_po}.zip",
                        mime="application/zip",
                        use_container_width=True
                    )
                    
                d_c1, d_c2, d_c3 = st.columns([1, 1, 1])
                with d_c1:
                    st.download_button(
                        label=f"📄 Descargar Remisión E0125 (PDF)",
                        data=rem_pdf_att,
                        file_name=f"Remision_E0125_{id_int_txt if id_int_txt else 'INT'}.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )
                with d_c2:
                    st.download_button(
                        label=f"🏷️ Descargar Etiquetas Tarimas (PDF)",
                        data=etiq_pdf_att,
                        file_name=f"Etiquetas_Tarimas_E0125_{id_int_txt if id_int_txt else 'INT'}.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )
                with d_c3:
                    st.download_button(
                        label=f"📊 Descargar Reporte PO Trazabilidad (PDF)",
                        data=rep_pdf_att,
                        file_name=f"Reporte_PO_{sel_po}_{id_int_txt if id_int_txt else 'INT'}.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )
                    
                import fitz
                def _pdf_to_high_res_png(pdf_data):
                    try:
                        doc = fitz.open(stream=pdf_data, filetype="pdf")
                        if len(doc) > 0:
                            pix = doc[0].get_pixmap(dpi=200)
                            return pix.tobytes("png")
                    except Exception:
                        pass
                    return None
                    
                png_img_bytes = _pdf_to_high_res_png(rep_pdf_att)
                if png_img_bytes:
                    st.download_button(
                        label=f"🖼️ Descargar Imagen Oficial del Reporte (PNG)",
                        data=png_img_bytes,
                        file_name=f"Reporte_Grafico_PO_{sel_po}_{id_int_txt if id_int_txt else 'INT'}.png",
                        mime="image/png",
                        use_container_width=True
                    )
                    
                with st.expander("👁️ Vista Previa del Correo Oficial (HTML)", expanded=True):
                    # Interfaz interactiva con html2canvas para Copiar al Portapapeles y Descargar Imagen PNG
                    preview_with_actions = f"""
                    <!DOCTYPE html>
                    <html>
                    <head>
                    <meta charset="utf-8">
                    <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
                    <style>
                        * {{ box-sizing: border-box; }}
                        body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; background: #F1F5F9; margin: 0; padding: 12px; }}
                        .action-bar {{
                            position: sticky;
                            top: 0;
                            z-index: 999;
                            background: #0F172A;
                            padding: 10px 16px;
                            border-radius: 8px;
                            display: flex;
                            gap: 12px;
                            align-items: center;
                            margin-bottom: 15px;
                            box-shadow: 0 4px 10px rgba(0,0,0,0.15);
                        }}
                        .btn-img {{
                            background: #EC2024;
                            color: #FFFFFF;
                            border: none;
                            padding: 8px 16px;
                            border-radius: 6px;
                            font-weight: 700;
                            font-size: 13px;
                            cursor: pointer;
                            display: flex;
                            align-items: center;
                            gap: 8px;
                            transition: background 0.15s ease, transform 0.1s ease;
                        }}
                        .btn-img:hover {{ background: #C71B1E; transform: translateY(-1px); }}
                        .btn-img:active {{ transform: translateY(0); }}
                        .btn-clipboard {{ background: #2563EB; }}
                        .btn-clipboard:hover {{ background: #1D4ED8; }}
                        .info-tag {{ color: #94A3B8; font-size: 12px; margin-left: auto; font-weight: 500; }}
                        #email-render-target {{
                            background: #FFFFFF;
                            border-radius: 10px;
                            box-shadow: 0 2px 8px rgba(0,0,0,0.06);
                            overflow: hidden;
                        }}
                    </style>
                    </head>
                    <body>
                    <div class="action-bar">
                        <button class="btn-img" id="btn-dl" onclick="downloadImage()">
                            <svg width="15" height="15" fill="currentColor" viewBox="0 0 16 16"><path d="M.5 9.9a.5.5 0 0 1 .5.5v2.5a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1v-2.5a.5.5 0 0 1 1 0v2.5a2 2 0 0 1-2 2H2a2 2 0 0 1-2-2v-2.5a.5.5 0 0 1 .5-.5z"/><path d="M7.646 11.854a.5.5 0 0 0 .708 0l3-3a.5.5 0 0 0-.708-.708L8.5 10.293V1.5a.5.5 0 0 0-1 0v8.793L5.354 8.146a.5.5 0 1 0-.708.708l3 3z"/></svg>
                            📥 Descargar Imagen (PNG)
                        </button>
                        <button class="btn-img btn-clipboard" id="btn-cp" onclick="copyImage()">
                            <svg width="15" height="15" fill="currentColor" viewBox="0 0 16 16"><path d="M4 1.5H3a2 2 0 0 0-2 2V14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V3.5a2 2 0 0 0-2-2h-1v1h1a1 1 0 0 1 1 1V14a1 1 0 0 1-1 1H3a1 1 0 0 1-1-1V3.5a1 1 0 0 1 1-1h1v-1z"/><path d="M9.5 1a.5.5 0 0 1 .5.5v1a.5.5 0 0 1-.5.5h-3a.5.5 0 0 1-.5-.5v-1a.5.5 0 0 1 .5-.5h3zm-3-1A1.5 1.5 0 0 0 5 1.5v1A1.5 1.5 0 0 0 6.5 4h3A1.5 1.5 0 0 0 11 2.5v-1A1.5 1.5 0 0 0 9.5 0h-3z"/></svg>
                            📋 Copiar Imagen al Portapapeles
                        </button>
                        <span class="info-tag" id="status-label">📸 Resolución HD 2x</span>
                    </div>
                    <div id="email-render-target">
                        {eml_html}
                    </div>
                    <script>
                    function captureCanvas(cb) {{
                        const target = document.getElementById('email-render-target');
                        const lbl = document.getElementById('status-label');
                        lbl.innerText = 'Generando imagen HD...';
                        html2canvas(target, {{
                            scale: 2,
                            useCORS: true,
                            backgroundColor: '#FFFFFF',
                            logging: false
                        }}).then(canvas => {{
                            lbl.innerText = '📸 Resolución HD 2x';
                            cb(canvas);
                        }}).catch(e => {{
                            lbl.innerText = 'Error al capturar imagen';
                            console.error(e);
                        }});
                    }}
                    function downloadImage() {{
                        captureCanvas(canvas => {{
                            const a = document.createElement('a');
                            a.download = 'Reporte_Correo_PO_{sel_po}.png';
                            a.href = canvas.toDataURL('image/png');
                            a.click();
                            document.getElementById('status-label').innerText = '✅ Imagen descargada correctamente';
                        }});
                    }}
                    function copyImage() {{
                        const btn = document.getElementById('btn-cp');
                        const lbl = document.getElementById('status-label');
                        captureCanvas(canvas => {{
                            canvas.toBlob(blob => {{
                                if (navigator.clipboard && navigator.clipboard.write) {{
                                    navigator.clipboard.write([
                                        new ClipboardItem({{ 'image/png': blob }})
                                    ]).then(() => {{
                                        btn.innerHTML = '✅ ¡Copiado al Portapapeles!';
                                        lbl.innerText = 'Listo para pegar con Ctrl + V';
                                        setTimeout(() => {{
                                            btn.innerHTML = '<svg width=\"15\" height=\"15\" fill=\"currentColor\" viewBox=\"0 0 16 16\"><path d=\"M4 1.5H3a2 2 0 0 0-2 2V14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V3.5a2 2 0 0 0-2-2h-1v1h1a1 1 0 0 1 1 1V14a1 1 0 0 1-1 1H3a1 1 0 0 1-1-1V3.5a1 1 0 0 1 1-1h1v-1z\"/><path d=\"M9.5 1a.5.5 0 0 1 .5.5v1a.5.5 0 0 1-.5.5h-3a.5.5 0 0 1-.5-.5v-1a.5.5 0 0 1 .5-.5h3zm-3-1A1.5 1.5 0 0 0 5 1.5v1A1.5 1.5 0 0 0 6.5 4h3A1.5 1.5 0 0 0 11 2.5v-1A1.5 1.5 0 0 0 9.5 0h-3z\"/></svg> 📋 Copiar Imagen al Portapapeles';
                                        }}, 3000);
                                    }}).catch(err => {{
                                        downloadImage();
                                        lbl.innerText = 'Descargada (portapapeles restringido por navegador)';
                                    }});
                                }} else {{
                                    downloadImage();
                                    lbl.innerText = 'Descargada (portapapeles no disponible)';
                                }}
                            }}, 'image/png');
                        }});
                    }}
                    </script>
                    </body>
                    </html>
                    """
                    st.components.v1.html(preview_with_actions, height=850, scrolling=True)
                    
            # 4. Pestaña: Acciones de Mantenimiento
            with tab_acciones:
                st.subheader("⚙️ Mantenimiento de la Orden")
                c_act1, c_act2 = st.columns(2)
                with c_act1:
                    st.write("¿Deseas editar los datos o artículos de esta PO?")
                    if st.button("✏️ Abrir en Ajuste de PO", use_container_width=True):
                        st.session_state['current_po_edit_idx'] = pos_list.index(sel_po) if sel_po in pos_list else 0
                        st.info("Navega a la sección '✏️ Ajuste de PO' en el menú lateral.")
                with c_act2:
                    st.write("Si necesitas eliminar definitivamente esta PO del catálogo:")
                    if st.button("🗑️ Eliminar esta PO", type="secondary", use_container_width=True):
                        st.session_state['confirm_del_po'] = sel_po
                        
                if st.session_state.get('confirm_del_po') == sel_po:
                    st.warning(f"⚠️ ¿Estás seguro de eliminar la PO **{sel_po}** y todas sus partidas?")
                    c_yes, c_no = st.columns(2)
                    with c_yes:
                        if st.button("Sí, eliminar definitivamente", type="primary", use_container_width=True):
                            ok, msg = delete_po(sel_po)
                            if ok:
                                st.success(msg)
                                st.session_state.pop('confirm_del_po', None)
                                st.rerun()
                            else:
                                st.error(msg)
                    with c_no:
                        if st.button("Cancelar", use_container_width=True):
                            st.session_state.pop('confirm_del_po', None)
                            st.rerun()


# ==============================================================================
# SECCIÓN 6: ESTADO DE INTEGRACIÓN
# ==============================================================================
elif menu == "🔄 Estado de Integración":
    st.title("🔄 Estado de Integración y Enlace entre Aplicaciones")
    st.markdown("Verificación del enlace de datos en tiempo real entre el **PO Tracker Hub**, la app de **Remisiones** y la app de **Corte y Doblez**.")
    
    rem_dir = get_remisiones_dir()
    
    st.subheader("1. Conexión con App de Remisiones (Fase 1)")
    c_rem1, c_rem2 = st.columns(2)
    with c_rem1:
        st.write(f"**Ruta de Datos de Remisiones:** `{rem_dir}`")
        rem_gen_exists = (rem_dir / 'BD_Datos_Generales_Remision.xlsx').exists()
        rem_det_exists = (rem_dir / 'BD_Detalle_Tarimas.xlsx').exists()
        rem_tar_exists = (rem_dir / 'BD_Tarimas.xlsx').exists()
        
        st.markdown(f"• `BD_Datos_Generales_Remision.xlsx`: {'✅ Conectado' if rem_gen_exists else '❌ No encontrado'}")
        st.markdown(f"• `BD_Detalle_Tarimas.xlsx`: {'✅ Conectado' if rem_det_exists else '❌ No encontrado'}")
        st.markdown(f"• `BD_Tarimas.xlsx`: {'✅ Conectado' if rem_tar_exists else '❌ No encontrado'}")
        
    with c_rem2:
        df_rem, df_det, df_tar = load_remisiones_databases()
        st.metric("Total Remisiones Registradas", len(df_rem) if not df_rem.empty else 0)
        st.metric("Total Registros de Envíos en Tarimas", len(df_det) if not df_det.empty else 0)
        
    st.write("---")
    st.subheader("2. Sincronización de Catálogos de PO")
    st.write("Cada vez que se registra una PO en esta aplicación, los catálogos `BD_POs_Cabecera.xlsx` y `BD_Requerimientos_POs.xlsx` se actualizan automáticamente para que la app de Remisiones y la app de Corte-Doblez puedan leerlos sin desfase.")
    
    if st.button("🔄 Forzar Re-Sincronización de Archivos Excel ahora", type="primary"):
        export_sync_to_excel()
        st.success("✅ Archivos de sincronización actualizados y copiados al directorio compartido.")


# ==============================================================================
# SECCIÓN 7: MANUAL Y ARQUITECTURA INDUSTRIA 4.0
# ==============================================================================
elif menu == "📘 Manual y Arquitectura 4.0":
    st.title("📘 Manual, Arquitectura Industria 4.0 & Stack Tecnológico")
    st.markdown("Documentación técnica, diagramas de arquitectura ciberfísica y manual operativo para el **PO Tracker & Master Hub de Industria Sigrama**.")
    
    tab_ind4, tab_identidad, tab_diagrama_ocr, tab_stack, tab_manual_pasos = st.tabs([
        "🏭 Arquitectura Industria 4.0",
        "🎨 Identidad Corporativa Sigrama",
        "⚡ Motor OCR Espacial (Diagrama)",
        "💻 Stack Tecnológico (Tech Stack)",
        "📖 Manual Operativo Paso a Paso"
    ])
    
    with tab_ind4:
        st.subheader("🏭 El Hilo Digital (Digital Thread) en Industria Sigrama")
        st.write("El **PO Tracker & Master Hub** actúa como el Gemelo Digital del Requerimiento, asegurando que cada orden de compra se convierta en una orden sistemática interconectada con Corte-Doblez y Remisiones.")
        
    with tab_identidad:
        st.subheader("🎨 Manual de Identidad Corporativa Oficial (Industria Sigrama S.A. de C.V.)")
        st.markdown("""
        Esta aplicación se rige bajo los lineamientos oficiales del **Manual de Identidad Corporativa de Industria Sigrama**:
        """)
        
        c_col1, c_col2 = st.columns(2)
        with c_col1:
            st.markdown("""
            #### 🔴 Paleta de Color Institucional
            - **PANTONE® 485 C (Rojo Sigrama)**:
              - RGB: `(236, 32, 36)` | Hex: `#EC2024`
              - Uso: Botones principales, acentos de estado, bordes activos y llamado a la acción.
            - **PANTONE® Black 7 C (Negro Carbón)**:
              - RGB: `(17, 17, 17)` | Hex: `#111111`
              - Uso: Barra lateral corporativa, encabezados de tablas y tipografía de alto contraste.
            - **PANTONE® 14-4107 (Gris Técnico)**:
              - Hex: `#D2D3D5` / `#E2E8F0`
              - Uso: Bordes de tarjetas, líneas divisorias y fondos secundarios.
            - **PANTONE® 000C (Blanco Puro)**:
              - Hex: `#FFFFFF`
            """)
        with c_col2:
            st.markdown("""
            #### ✍️ Tipografía & Lemas Oficiales
            - **Tipografía de Títulos y Logotipo**: `Montserrat` / `Gotham` (Pesos 700 y 900).
            - **Tipografía de Contenido y Tablas**: `Questrial` (Diseño limpio y legible).
            - **Slogan Principal**:
              > <span style="color:#EC2024; font-family:'Montserrat', sans-serif; font-weight:bold;">SOLUCIONES QUE TRANSFORMAN TU EMPRESA</span>
            - **Lema Institucional**:
              > <span style="font-family:'Questrial', sans-serif; font-style:italic; border-bottom:2px solid #EC2024;">Ingeniería que da resultados!!</span>
            """, unsafe_allow_html=True)
            
        st.info("💡 La suite de herramientas de Industria Sigrama (**Remisiones de Materiales**, **Corte y Doblez** y **PO Tracker**) comparte la misma línea de diseño unificada para asegurar una experiencia corporativa continua.")
        
    with tab_diagrama_ocr:
        st.subheader("⚡ Arquitectura del Motor OCR Espacial")
        st.write("Diagrama del flujo de extracción cartesiana por coordenadas Y-clustering para documentos de 1 a N páginas.")
        
    with tab_stack:
        st.subheader("💻 Stack Tecnológico de la Aplicación")
        st.markdown("""
- **Frontend:** Streamlit 1.40+ & Plotly Express
- **Motor OCR:** PyMuPDF 1.28 (`fitz`) con Y-clustering espacial y `extract-msg`
- **Base de Datos:** SQLite 3 (`po_tracker.db`) con soporte ACID
- **Interoperabilidad:** Pandas 2.2 & OpenPyXL con espejos Excel automáticos
""")
        
    with tab_manual_pasos:
        st.subheader("📖 Manual de Operación para el Usuario")
        st.markdown("""
1. **Paso 1 (Ingesta)**: En **`📬 Bandeja de Entrada OCR`**, sube los correos `.msg` o archivos `.pdf` de las 58 órdenes internas.
2. **Paso 2 (Ajuste)**: En **`✏️ Ajuste de PO`**, navega entre las órdenes para verificar `INT-XXXX`, fechas y corregir la tabla de artículos en vivo.
3. **Paso 3 (Catálogo)**: En **`📋 Matriz de Órdenes`**, consulta el catálogo ordenado y exporta a Excel.
4. **Paso 4 (Trazabilidad)**: En **`🔍 Ficha de Trazabilidad 360°`**, revisa piezas cortadas/dobladas y remisiones.
""")


# ==============================================================================
# SECCIÓN 8: MANTENIMIENTO DE LA APP & RESPALDO DE BASE DE DATOS
# ==============================================================================
elif menu == "🛠️ Mantenimiento de la App":
    st.title("🛠️ Mantenimiento de la App & Respaldo de Base de Datos")
    st.markdown("Herramientas de respaldo, salud de datos, resincronización forzada y gobernanza del sistema.")
    
    df_pos = get_all_pos()
    df_part = get_all_partidas()
    
    m_c1, m_c2, m_c3 = st.columns(3)
    with m_c1:
        st.metric("Total POs Registradas", len(df_pos))
    with m_c2:
        st.metric("Total Partidas Activas", len(df_part))
    with m_c3:
        db_file = SQLITE_DB_PATH
        db_size_kb = (db_file.stat().st_size / 1024.0) if db_file.exists() else 0.0
        st.metric("Tamaño Base de Datos SQLite", f"{db_size_kb:.1f} KB")
        
    st.write("---")
    
    b_col1, b_col2, b_col3 = st.columns(3)
    with b_col1:
        st.subheader("💾 1. Descargar Respaldo")
        st.write("Descarga una copia íntegra del archivo SQLite `po_tracker.db`.")
        if db_file.exists():
            with open(db_file, "rb") as fp:
                st.download_button(
                    label="📥 Descargar Backup (.db)",
                    data=fp.read(),
                    file_name=f"backup_po_tracker_{datetime.date.today().strftime('%Y%m%d')}.db",
                    mime="application/x-sqlite3",
                    type="primary"
                )
        else:
            st.warning("Sin registros.")
            
    with b_col2:
        st.subheader("🔄 2. Resincronizar Excel")
        st.write("Regenera de inmediato `BD_POs_Cabecera.xlsx` y `BD_Requerimientos_POs.xlsx`.")
        if st.button("⚡ Resincronizar Excel", use_container_width=True):
            export_sync_to_excel()
            st.success("✅ Archivos Excel actualizados.")
            
    with b_col3:
        st.subheader("🗑️ 3. Limpieza Total (0 POs)")
        st.write("Vacía todas las POs de prueba para iniciar la ingesta limpia de las 58 órdenes.")
        if st.button("🚨 Limpiar Base de Datos (0 POs)", type="secondary", use_container_width=True):
            st.session_state['confirm_reset_all_db'] = True
            
    if st.session_state.get('confirm_reset_all_db'):
        st.warning("⚠️ ¿Confirmas que deseas eliminar todas las órdenes registradas para dejar el catálogo en 0 POs?")
        cy, cn = st.columns(2)
        with cy:
            if st.button("Sí, limpiar todo y reiniciar a 0 POs", type="primary", use_container_width=True):
                ok_c, msg_c = clear_all_pos_db()
                if ok_c:
                    st.success(msg_c)
                    st.session_state.pop('confirm_reset_all_db', None)
                    st.rerun()
        with cn:
            if st.button("Cancelar", use_container_width=True):
                st.session_state.pop('confirm_reset_all_db', None)
                st.rerun()


