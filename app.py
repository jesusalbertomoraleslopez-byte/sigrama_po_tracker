import streamlit as st
import pandas as pd
import datetime
import plotly.express as px
import plotly.graph_objects as go
import io
import os
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
    get_remisiones_dir
)
from db_manager import (
    init_db,
    save_po,
    delete_po,
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

# Inyección de CSS Oficial Sigrama
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600;700;800&family=Questrial&display=swap');

    html, body, [class*="css"], .stApp {
        font-family: 'Questrial', sans-serif !important;
        background-color: #F8FAFC !important;
    }

    h1, h2, h3, h4, h5, h6, .main-title {
        font-family: 'Montserrat', sans-serif !important;
        font-weight: 700 !important;
        color: #111111 !important;
    }

    /* Barra lateral */
    [data-testid="stSidebar"] {
        background-color: #111111 !important;
        border-right: 1px solid #1E293B !important;
    }
    [data-testid="stSidebar"] p, [data-testid="stSidebar"] span, [data-testid="stSidebar"] label {
        color: #FFFFFF !important;
    }
    
    /* Botones primarios */
    button[kind="primary"], div.stButton > button[kind="primary"] {
        background-color: #EC2024 !important;
        color: #FFFFFF !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
        border: none !important;
    }
    button[kind="primary"]:hover {
        background-color: #C01216 !important;
    }

    /* Tarjetas KPI */
    .kpi-card {
        background: white;
        border-radius: 12px;
        padding: 18px 20px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.05);
        border-left: 5px solid #EC2024;
        margin-bottom: 12px;
    }
    .kpi-card-green {
        border-left: 5px solid #10B981;
    }
    .kpi-card-amber {
        border-left: 5px solid #F59E0B;
    }
    .kpi-card-blue {
        border-left: 5px solid #3B82F6;
    }
    .kpi-val {
        font-size: 28px;
        font-family: 'Montserrat', sans-serif;
        font-weight: 800;
        color: #111111;
    }
    .kpi-lbl {
        font-size: 13px;
        color: #64748B;
        text-transform: uppercase;
        font-weight: 600;
    }

    /* Insignias de estatus */
    .badge {
        display: inline-block;
        padding: 4px 10px;
        border-radius: 9999px;
        font-size: 12px;
        font-weight: 700;
        color: white;
    }

    /* Tarjeta PO formato oficial */
    .po-preview-box {
        background: white;
        border: 2px solid #E2E8F0;
        border-radius: 10px;
        padding: 24px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.04);
    }
    .po-header-table {
        width: 100%;
        border-collapse: collapse;
        margin-bottom: 15px;
    }
    .po-header-table th {
        background-color: #111111;
        color: white;
        padding: 6px 10px;
        font-size: 12px;
        text-align: left;
    }
    .po-header-table td {
        border: 1px solid #CBD5E1;
        padding: 6px 10px;
        font-size: 13px;
    }
</style>
""", unsafe_allow_html=True)

# Sidebar Corporativo
with st.sidebar:
    logo_path = Path(__file__).resolve().parent / "logo_sigrama.png"
    if logo_path.exists():
        st.image(str(logo_path), use_container_width=True)
    else:
        st.markdown("<h2 style='color:#EC2024; text-align:center;'>INDUSTRIA SIGRAMA</h2>", unsafe_allow_html=True)
        
    st.markdown("<h4 style='color:white; margin-top:0; text-align:center;'>PO Tracker & Master Hub</h4>", unsafe_allow_html=True)
    st.markdown("---")
    
    menu = st.radio(
        "Navegación:",
        [
            "📊 Dashboard Ejecutivo",
            "🎯 Matriz de Control 360° (Producción + Remisión)",
            "📋 Matriz de Órdenes (POs)",
            "🔍 Ficha de Trazabilidad 360°",
            "📬 Bandeja de Correos & OCR",
            "📥 Registrar / Cargar PO",
            "🔄 Estado de Integración",
            "📘 Manual & Arquitectura Industria 4.0"
        ],
        index=0
    )
    
    st.markdown("---")
    st.caption("🚀 **SIGRAMA Suite**")
    st.caption("• PO Tracker Hub *(Esta App)*")
    st.caption("• Remisiones de Materiales *(Fase 1)*")
    st.caption("• Corte y Doblez *(Fase 2)*")
    
    # Resumen rápido en sidebar
    df_all_pos = get_all_pos()
    df_all_part = get_all_partidas()
    
    st.markdown("---")
    st.markdown(f"**Total POs en Sistema:** `{len(df_all_pos)}`")
    st.markdown(f"**Total Partidas Activas:** `{len(df_all_part)}`")


# ==============================================================================
# SECCIÓN 1: DASHBOARD EJECUTIVO
# ==============================================================================
if menu == "📊 Dashboard Ejecutivo":
    st.title("📊 Dashboard Ejecutivo de Órdenes de Compra")
    st.markdown("Visión global de requerimientos, avance de envíos al cliente y cumplimiento en tiempo real.")
    
    df_pos = get_all_pos()
    df_part = get_all_partidas()
    
    if df_pos.empty:
        st.info("💡 Aún no hay Órdenes de Compra registradas. Dirígete a **'📥 Registrar / Cargar PO'** o **'📬 Bandeja de Correos & OCR'** para ingresar tu primera PO.")
    else:
        df_summary = get_global_pos_tracking_summary(df_pos, df_part)
        
        # Métricas Globales
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
        
        # Gráficas
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
            
        # Tabla de Urgencias y Pendientes
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


# ==============================================================================
# SECCIÓN: MATRIZ DE CONTROL 360° (PRODUCCIÓN + REMISIÓN)
# ==============================================================================
elif menu == "🎯 Matriz de Control 360° (Producción + Remisión)":
    st.title("🎯 Matriz de Control 360°: Producción vs Remisión")
    st.markdown("Integración sistemática del flujo completo: **Captura Multi-Canal ➔ Orden Sistemática ➔ Avance Corte/Doblez ➔ Avance Remisiones ➔ Análisis de Estatus.**")
    
    # Visualizador interactivo del flujo
    with st.expander("📌 Diagrama del Flujo Sistemático Integrado", expanded=False):
        st.markdown("""<div style="background-color:#F8FAFC; border:1px solid #CBD5E1; border-radius:10px; padding:18px; margin-bottom:15px;">
<div style="display:flex; justify-content:space-between; gap:10px; margin-bottom:15px;">
<div style="flex:1; background-color:#1E293B; color:white; padding:10px; border-radius:6px; text-align:center; font-weight:700; font-size:12px;">
📬 PO CARGADAS EN CORREO<br><span style="font-size:10px; font-weight:400; color:#94A3B8;">(PDF / OCR / Notas de Urgencia)</span>
</div>
<div style="flex:1; background-color:#1E293B; color:white; padding:10px; border-radius:6px; text-align:center; font-weight:700; font-size:12px;">
🚚 PO DETECTADAS EN REMISIONES<br><span style="font-size:10px; font-weight:400; color:#94A3B8;">(Tarimas / Despachos Almacén)</span>
</div>
<div style="flex:1; background-color:#1E293B; color:white; padding:10px; border-radius:6px; text-align:center; font-weight:700; font-size:12px;">
⚙️ PO DETECTADAS EN CORTE Y DOBLEZ<br><span style="font-size:10px; font-weight:400; color:#94A3B8;">(OFs / Nidos / Avances Taller)</span>
</div>
</div>

<div style="text-align:center; font-size:20px; color:#EC2024; margin:-5px 0 10px 0;">⬇️ <b>CONSOLIDACIÓN EN MASTER HUB</b> ⬇️</div>

<div style="background-color:#EC2024; color:white; padding:12px; border-radius:8px; text-align:center; font-weight:800; font-size:14px; margin-bottom:15px;">
📋 ORDEN SISTEMÁTICA (MASTER PO RECORD UNIFICADO)
</div>

<div style="display:flex; justify-content:space-between; gap:15px;">
<div style="flex:1; background-color:#0284C7; color:white; padding:12px; border-radius:6px; text-align:center; font-weight:700; font-size:13px;">
🔵 REVISIÓN AVANCE EN CORTE Y DOBLEZ<br>
<span style="font-size:11px; font-weight:400;">• Piezas Cortadas (Láser)<br>• Piezas Dobladas<br>• Piezas Terminadas en Taller</span>
</div>
<div style="flex:1; background-color:#16A34A; color:white; padding:12px; border-radius:6px; text-align:center; font-weight:700; font-size:13px;">
🟢 REVISIÓN AVANCE EN REMISIONES<br>
<span style="font-size:11px; font-weight:400;">• Tarimas en Almacén<br>• Tarimas en Tránsito<br>• Piezas Enviadas con Remisión</span>
</div>
<div style="flex:1; background-color:#0F172A; color:white; padding:12px; border-radius:6px; text-align:center; font-weight:700; font-size:13px;">
📊 ANÁLISIS DE ESTATUS<br>
<span style="font-size:11px; font-weight:400;">• % Avance Producción vs Envío<br>• Filtros Dinámicos<br>• Saldo Piezas Pendientes</span>
</div>
</div>
</div>""", unsafe_allow_html=True)
    
    df_pos = get_all_pos()
    df_part = get_all_partidas()
    
    if df_pos.empty:
        st.info("💡 No hay POs registradas para generar la Matriz 360°.")
    else:
        df_mat = get_integrated_360_summary(df_pos, df_part)
        
        # Métricas Globales Integradas
        tot_req_all = df_mat['piezas_requeridas'].sum()
        tot_fab_all = df_mat['piezas_fabricadas'].sum()
        tot_env_all = df_mat['piezas_remisionadas'].sum()
        tot_pend_fab = df_mat['piezas_pendientes_fab'].sum()
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
        
        # Filtros Dinámicos
        c_f1, c_f2, c_f3 = st.columns([2, 1, 1])
        with c_f1:
            q_search_360 = st.text_input("🔍 Búsqueda rápida (PO, Proyecto, SKU, OF, Remisión):", "", key="search_360")
        with c_f2:
            est_opts = ["Todos"] + sorted(list(df_mat['estatus_360'].unique()))
            sel_est_360 = st.selectbox("Filtrar por Estatus 360°:", est_opts, key="est_360")
        with c_f3:
            proy_opts_360 = ["Todos"] + sorted([p for p in df_mat['proyecto'].dropna().unique() if str(p).strip()])
            sel_proy_360 = st.selectbox("Filtrar por Proyecto:", proy_opts_360, key="proy_360")
            
        # Aplicar Filtros
        df_filtered = df_mat.copy()
        if q_search_360.strip():
            term = q_search_360.strip().lower()
            df_filtered = df_filtered[
                df_filtered['po'].astype(str).str.lower().str.contains(term) |
                df_filtered['proyecto'].astype(str).str.lower().str.contains(term) |
                df_filtered['ofs_asociadas'].astype(str).str.lower().str.contains(term) |
                df_filtered['remisiones_asociadas'].astype(str).str.lower().str.contains(term)
            ]
        if sel_est_360 != "Todos":
            df_filtered = df_filtered[df_filtered['estatus_360'] == sel_est_360]
        if sel_proy_360 != "Todos":
            df_filtered = df_filtered[df_filtered['proyecto'] == sel_proy_360]
            
        st.markdown(f"Mostrando **{len(df_filtered)}** órdenes de compra encontradas:")
        
        cols_mat_show = [
            'po', 'proyecto', 'piezas_requeridas',
            'piezas_fabricadas', 'pct_fabricacion',
            'piezas_remisionadas', 'pct_remision',
            'piezas_pendientes_fab', 'piezas_pendientes_env',
            'ofs_asociadas', 'remisiones_asociadas', 'estatus_360'
        ]
        
        st.dataframe(
            df_filtered[cols_mat_show].rename(columns={
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
        
        # Botón de Descarga Excel
        output_360 = io.BytesIO()
        with pd.ExcelWriter(output_360, engine='openpyxl') as writer:
            df_filtered.to_excel(writer, sheet_name='Control_360_SIGRAMA', index=False)
        st.download_button(
            "📥 Descargar Matriz de Control 360° en Excel",
            data=output_360.getvalue(),
            file_name=f"Matriz_Control_360_SIGRAMA_{datetime.date.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


# ==============================================================================
# SECCIÓN 2: MATRIZ DE ÓRDENES (POs)
# ==============================================================================
elif menu == "📋 Matriz de Órdenes (POs)":
    st.title("📋 Matriz Maestra de Órdenes de Compra")
    st.markdown("Consulta, búsqueda global y filtrado de todas las POs registradas con su estatus de remisión.")
    
    df_pos = get_all_pos()
    df_part = get_all_partidas()
    
    if df_pos.empty:
        st.info("💡 No hay POs registradas. Registra una nueva PO en la pestaña correspondiente.")
    else:
        df_summary = get_global_pos_tracking_summary(df_pos, df_part)
        
        # Filtros Superiores
        f1, f2, f3 = st.columns([2, 1, 1])
        with f1:
            q_search = st.text_input("🔍 Búsqueda rápida (Folio PO, Proyecto, SKU, Comprador, Solicitante):", "")
        with f2:
            estatus_opts = ["Todos"] + list(df_summary['estatus_remision'].unique())
            sel_estatus = st.selectbox("Filtrar por Estatus:", estatus_opts)
        with f3:
            proy_opts = ["Todos"] + [p for p in df_summary['proyecto'].dropna().unique() if str(p).strip()]
            sel_proy = st.selectbox("Filtrar por Proyecto:", proy_opts)
            
        df_filtered = df_summary.copy()
        
        if q_search:
            q = q_search.strip().lower()
            df_filtered = df_filtered[
                df_filtered['po'].astype(str).str.lower().str.contains(q) |
                df_filtered['proyecto'].astype(str).str.lower().str.contains(q) |
                df_filtered['solicitante'].astype(str).str.lower().str.contains(q) |
                df_filtered['comprador'].astype(str).str.lower().str.contains(q) |
                df_filtered['observaciones'].astype(str).str.lower().str.contains(q)
            ]
            
        if sel_estatus != "Todos":
            df_filtered = df_filtered[df_filtered['estatus_remision'] == sel_estatus]
            
        if sel_proy != "Todos":
            df_filtered = df_filtered[df_filtered['proyecto'] == sel_proy]
            
        st.write(f"Mostrando **{len(df_filtered)}** de **{len(df_summary)}** Órdenes de Compra:")
        
        display_cols = [
            'po', 'fecha_pedido', 'proyecto', 'requisicion', 'solicitante', 'comprador',
            'piezas_requeridas', 'piezas_remisionadas', 'piezas_pendientes',
            'pct_cumplimiento', 'estatus_remision', 'remisiones_asociadas', 'total'
        ]
        
        st.dataframe(
            df_filtered[display_cols].rename(columns={
                'po': 'PO / Folio',
                'fecha_pedido': 'Fecha Pedido',
                'proyecto': 'Proyecto',
                'requisicion': 'Requisición',
                'solicitante': 'Solicitante',
                'comprador': 'Comprador',
                'piezas_requeridas': 'Cant. Req.',
                'piezas_remisionadas': 'Enviadas',
                'piezas_pendientes': 'Pendientes',
                'pct_cumplimiento': '% Avance',
                'estatus_remision': 'Estatus Entrega',
                'remisiones_asociadas': 'Remisiones Asociadas',
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
# SECCIÓN 3: FICHA DE TRAZABILIDAD 360°
# ==============================================================================
elif menu == "🔍 Ficha de Trazabilidad 360°":
    st.title("🔍 Ficha de Trazabilidad 360° por Orden de Compra")
    st.markdown("Consulta en profundidad el detalle oficial de la PO, desglose de partidas vs envíos de remisiones y auditoría.")
    
    df_pos = get_all_pos()
    
    if df_pos.empty:
        st.info("💡 No hay POs registradas.")
    else:
        # Priorizar POs con datos y colocar 26083186 al frente si existe
        pos_list = df_pos['po'].astype(str).tolist()
        if '26083186' in pos_list:
            pos_list.remove('26083186')
            pos_list = ['26083186'] + sorted([p for p in pos_list if not p.startswith('URG') and not p.startswith('SCRAP')]) + [p for p in pos_list if p.startswith('URG') or p.startswith('SCRAP')]
        
        # Mapeo descriptivo para el selector
        po_info_map = {}
        for _, r_po in df_pos.iterrows():
            p_code = str(r_po.get('po', ''))
            p_prj = str(r_po.get('proyecto', 'General'))
            p_prj = 'General' if p_prj.lower() in ('nan', 'none', '') else p_prj
            po_info_map[p_code] = f"📑 {p_code}  |  Proyecto: {p_prj}"

        sel_po = st.selectbox(
            "Seleccione la Orden de Compra (PO) a inspeccionar:",
            options=pos_list,
            format_func=lambda x: po_info_map.get(x, f"📑 {x}"),
            index=0
        )
        
        df_cab, df_part = get_po_by_folio(sel_po)
        
        if df_cab.empty:
            st.error(f"No se encontró la PO {sel_po}")
        else:
            cab = df_cab.iloc[0].to_dict()
            tracking = get_tracking_for_po(sel_po, df_part)
            
            def fmt_val(v, default="—"):
                if v is None or pd.isna(v):
                    return default
                s = str(v).strip()
                if not s or s.lower() in ('none', 'nan', 'nat', 'null'):
                    return default
                return s

            po_num_str = fmt_val(cab.get('po'), sel_po)
            f_ped_str = fmt_val(cab.get('fecha_pedido'), 'Sin fecha')
            prov_str = fmt_val(cab.get('proveedor'), 'SIGRAMA PLANTA METALES')
            prov_atn_str = fmt_val(cab.get('proveedor_atencion'), 'JESUS MORALES')
            lab_str = fmt_val(cab.get('lab'), 'ALMACEN SIGRAMA')
            tiempo_ent_str = fmt_val(cab.get('tiempo_entrega'), '—')
            proy_str = fmt_val(cab.get('proyecto'), 'General')
            req_str = fmt_val(cab.get('requisicion'), '—')
            sol_str = fmt_val(cab.get('solicitante'), '—')
            comp_str = fmt_val(cab.get('comprador'), '—')
            cli_str = fmt_val(cab.get('cliente_facturar_a'), 'INDUSTRIA SIGRAMA S.A. DE C.V.')
            rfc_str = fmt_val(cab.get('cliente_rfc'), 'ISI-870204-K4A')
            dir_str = fmt_val(cab.get('cliente_direccion'), 'C. JUAN ESCUTIA #50 COL. ABASTOS C.P. 27020 TORREON, COAH.')
            obs_str = fmt_val(cab.get('observaciones'), '—')
            st_color = ESTATUS_COLORS.get(tracking['estatus_global'], '#64748B')

            # Tarjeta Oficial Visual SIGRAMA
            with st.container(border=True):
                col_h1, col_h2 = st.columns([3, 1])
                with col_h1:
                    st.markdown(f"### 📑 ORDEN DE COMPRA: <span style='color:#EC2024;'>{po_num_str}</span>", unsafe_allow_html=True)
                    st.markdown(f"**Proyecto / Uso:** `{proy_str}` &nbsp;|&nbsp; **Fecha de Pedido:** `{f_ped_str}`")
                with col_h2:
                    st.markdown(f"""<div style='text-align:right; padding-top:5px;'>
<span class='badge' style='background-color:{st_color}; font-size:14px;'>
{tracking['estatus_global']} ({tracking['porcentaje_global']}%)
</span>
</div>""", unsafe_allow_html=True)
                
                st.divider()
                
                c_info1, c_info2 = st.columns(2)
                with c_info1:
                    st.markdown("##### 🏢 Datos del Proveedor y Entrega")
                    st.markdown(f"• **Proveedor:** {prov_str}")
                    st.markdown(f"• **Atención:** {prov_atn_str}")
                    st.markdown(f"• **L.A.B. / Destino:** {lab_str}")
                    st.markdown(f"• **Tiempo de Entrega:** {tiempo_ent_str}")
                with c_info2:
                    st.markdown("##### 👤 Cliente y Solicitante")
                    st.markdown(f"• **Facturar A:** {cli_str} *(RFC: {rfc_str})*")
                    st.markdown(f"• **Dirección:** {dir_str}")
                    st.markdown(f"• **Requisición:** {req_str} &nbsp;|&nbsp; **Comprador:** {comp_str}")
                    st.markdown(f"• **Solicitante:** {sol_str}")
                    
                if obs_str not in ('—', 'Sin observaciones especiales.', ''):
                    st.info(f"📝 **Observaciones:** {obs_str}")
            
            # Pestañas de detalle
            tab_partidas, tab_remisiones, tab_acciones = st.tabs([
                "📦 Desglose de Partidas vs Remisionado",
                "🚚 Historial de Remisiones y Tarimas Asociadas",
                "⚙️ Acciones y Mantenimiento"
            ])
            
            with tab_partidas:
                st.subheader("📦 Partidas Solicitadas vs Cumplimiento de Envíos")
                df_part_res = tracking['df_partidas']
                
                if not df_part_res.empty:
                    cols_p_show = [
                        'item_no', 'clave_sku', 'descripcion_producto',
                        'cantidad_requerida', 'cantidad_remisionada', 'cantidad_pendiente',
                        'porcentaje_cumplimiento', 'estatus_partida', 'remisiones_folios',
                        'precio_unitario', 'precio_total', 'fecha_entrega', 'parcialidad'
                    ]
                    st.dataframe(
                        df_part_res[cols_p_show].rename(columns={
                            'item_no': 'Partida #',
                            'clave_sku': 'SKU / Clave',
                            'descripcion_producto': 'Descripción del Producto',
                            'cantidad_requerida': 'Cant. Req.',
                            'cantidad_remisionada': 'Cant. Enviada',
                            'cantidad_pendiente': 'Pendiente',
                            'porcentaje_cumplimiento': '% Cumpl.',
                            'estatus_partida': 'Estatus',
                            'remisiones_folios': 'Remisiones',
                            'precio_unitario': 'P. Unitario',
                            'precio_total': 'P. Total',
                            'fecha_entrega': 'Fecha Entrega',
                            'parcialidad': 'Parcialidad'
                        }),
                        use_container_width=True,
                        hide_index=True
                    )
                else:
                    st.warning("Esta PO no tiene partidas registradas.")
                    
            with tab_remisiones:
                st.subheader("🚚 Envíos / Remisiones Registradas")
                df_env = tracking['df_historial_envios']
                
                if not df_env.empty:
                    st.success(f"Se encontraron **{len(df_env)}** registros de tarimas/piezas enviadas en **{len(tracking['remisiones_asociadas'])}** remisión(es): `{', '.join(tracking['remisiones_asociadas'])}`")
                    st.dataframe(df_env, use_container_width=True, hide_index=True)
                else:
                    st.info("ℹ️ Aún no se han generado remisiones para esta PO en la aplicación de Remisiones de Materiales.")
                    
            with tab_acciones:
                st.subheader("⚙️ Mantenimiento de la Orden")
                col_del1, col_del2 = st.columns([3, 1])
                with col_del1:
                    st.write("Si necesitas eliminar o editar esta PO, utiliza los controles siguientes:")
                with col_del2:
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
# SECCIÓN: ANALIZADOR Y OCR DE ARCHIVOS DE PO (.PDF / .MSG)
# ==============================================================================
elif menu in ("📬 Bandeja de Correos & OCR", "📄 Analizador de PO (PDF/MSG)"):
    st.title("📄 Analizador y OCR de Archivos de PO Sigrama (.PDF / .MSG)")
    st.markdown("Extracción inteligente de alta precisión de **Órdenes de Compra oficiales de Industria Sigrama** desde archivos **PDF** directos o archivos de correo **.MSG**.")
    
    tab_mail_ejemplo1, tab_mail_ejemplo2, tab_mail_custom = st.tabs([
        "📨 Caso 1: PO 2608-3177 (Urgencia de Partes)",
        "📨 Caso 2: Multi-PO 2603-2608 y 2603-2609 (+ Plano Anexo)",
        "📥 Subir Nuevo Correo o Múltiples PDFs Externos"
    ])
    
    # --------------------------------------------------------------------------
    # PESTAÑA 1: CASO 1 - CORREO INDIVIDUAL CON URGENCIAS
    # --------------------------------------------------------------------------
    with tab_mail_ejemplo1:
        with st.container(border=True):
            # Header estilo Outlook
            st.markdown("<h3 style='margin:0 0 10px 0; color:#1E293B;'>RV: 2608-3177 SIGRAMA METALES</h3>", unsafe_allow_html=True)
            
            c_snd1, c_snd2 = st.columns([4, 1])
            with c_snd1:
                st.markdown("""<div style="display:flex; align-items:center; gap:12px; margin-bottom:12px;">
<div style="width:46px; height:46px; border-radius:50%; background-color:#CBD5E1; color:#334155; display:flex; align-items:center; justify-content:center; font-weight:700; font-size:18px;">
AA
</div>
<div>
<div style="font-weight:700; font-size:16px; color:#111111;">
Alejandra Arellano Machado <span style="background-color:#E2E8F0; color:#475569; font-size:11px; padding:2px 8px; border-radius:10px;">Compras</span>
</div>
<div style="font-size:12px; color:#64748B;">
<b>Para:</b> Jesus Alberto Morales Lopez; Luis Alfredo Quintana Palma<br>
<b>CC:</b> Edgar Sosa Suarez; Josue Mesta
</div>
</div>
</div>""", unsafe_allow_html=True)
                
            with c_snd2:
                st.markdown("<div style='text-align:right; font-size:12px; color:#64748B;'>9:51 a. m.</div>", unsafe_allow_html=True)
                c_btn1, c_btn2, c_btn3 = st.columns(3)
                with c_btn1:
                    st.button("↩️", help="Responder", key="btn_reply_mail1")
                with c_btn2:
                    st.button("🔁", help="Reenviar", key="btn_fwd_mail1")
                with c_btn3:
                    st.button("⋯", help="Más opciones", key="btn_more_mail1")
                    
            st.divider()
            
            # Cajas de Archivos Adjuntos (PDFs)
            st.markdown("##### 📎 Archivos Adjuntos Detectados:")
            col_att1, col_att2 = st.columns(2)
            with col_att1:
                st.markdown("""<div style="border:1px solid #CBD5E1; border-radius:8px; padding:10px 14px; display:flex; align-items:center; gap:10px; background:#F8FAFC;">
<span style="font-size:24px;">📄</span>
<div>
<div style="font-weight:700; font-size:13px; color:#0F172A;">2608-3177 SIGRAMA METALES SAAM.PDF</div>
<div style="font-size:11px; color:#64748B;">180 KB • Orden de Compra Oficial</div>
</div>
</div>""", unsafe_allow_html=True)
            with col_att2:
                st.markdown("""<div style="border:1px solid #CBD5E1; border-radius:8px; padding:10px 14px; display:flex; align-items:center; gap:10px; background:#F8FAFC;">
<span style="font-size:24px;">📄</span>
<div>
<div style="font-weight:700; font-size:13px; color:#0F172A;">053 - COTIZACION PIEZAS GALVANIZADO.pdf</div>
<div style="font-size:11px; color:#64748B;">154 KB • Cotización y Especificaciones</div>
</div>
</div>""", unsafe_allow_html=True)
                
            st.write("")
            
            # Mensaje del Correo
            st.markdown("""
<div style="font-size:14px; color:#1E293B; line-height:1.6; margin: 15px 0;">
Ing Jesus Morales, buenos días.<br><br>
¿Me apoyan por favor compartiéndome status de esta orden de compra?<br><br>
<b>Nos están urgiendo las cantidades que marco en verde de los siguientes números de parte:</b>
</div>
""", unsafe_allow_html=True)
            
            # Tabla de Requerimientos y Urgencias Marcadas en Verde
            st.markdown("""<table style="width:100%; border-collapse:collapse; text-align:center; font-family:'Questrial',sans-serif; margin-bottom:20px;">
<tr style="background-color:#F1F5F9; font-weight:700; border:1px solid #CBD5E1; font-size:13px;">
<th style="padding:8px; border:1px solid #CBD5E1; text-align:left;">No. de Parte (SKU)</th>
<th style="padding:8px; border:1px solid #CBD5E1;">Total Ordenado</th>
<th style="padding:8px; border:1px solid #CBD5E1; background-color:#DCFCE7; color:#166534;">🔥 URGENTE (Verde)</th>
<th style="padding:8px; border:1px solid #CBD5E1;">P1</th>
<th style="padding:8px; border:1px solid #CBD5E1;">P2</th>
<th style="padding:8px; border:1px solid #CBD5E1;">P3</th>
<th style="padding:8px; border:1px solid #CBD5E1;">P4</th>
<th style="padding:8px; border:1px solid #CBD5E1;">P5</th>
</tr>
<tr style="border:1px solid #CBD5E1; font-size:14px;">
<td style="padding:10px; border:1px solid #CBD5E1; text-align:left; font-weight:700;">P20325-24</td>
<td style="padding:10px; border:1px solid #CBD5E1;">64</td>
<td style="padding:10px; border:1px solid #CBD5E1; background-color:#22C55E; color:white; font-weight:800; font-size:16px;">2</td>
<td style="padding:10px; border:1px solid #CBD5E1;">4</td>
<td style="padding:10px; border:1px solid #CBD5E1;">4</td>
<td style="padding:10px; border:1px solid #CBD5E1;">6</td>
<td style="padding:10px; border:1px solid #CBD5E1;">4</td>
<td style="padding:10px; border:1px solid #CBD5E1;">6</td>
</tr>
<tr style="border:1px solid #CBD5E1; font-size:14px;">
<td style="padding:10px; border:1px solid #CBD5E1; text-align:left; font-weight:700;">P20325-25</td>
<td style="padding:10px; border:1px solid #CBD5E1;">64</td>
<td style="padding:10px; border:1px solid #CBD5E1; background-color:#22C55E; color:white; font-weight:800; font-size:16px;">2</td>
<td style="padding:10px; border:1px solid #CBD5E1;">4</td>
<td style="padding:10px; border:1px solid #CBD5E1;">4</td>
<td style="padding:10px; border:1px solid #CBD5E1;">6</td>
<td style="padding:10px; border:1px solid #CBD5E1;">4</td>
<td style="padding:10px; border:1px solid #CBD5E1;">6</td>
</tr>
</table>""", unsafe_allow_html=True)
            
            st.markdown("<div style='font-size:14px; color:#1E293B;'>Gracias, quedamos al pendiente.<br><br>Saludos.</div>", unsafe_allow_html=True)
            st.divider()
            
            # Firma del Correo Oficial
            c_sig1, c_sig2 = st.columns([1, 3])
            with c_sig1:
                logo_path = Path(__file__).resolve().parent / "logo_sigrama.png"
                if logo_path.exists():
                    st.image(str(logo_path), width=130)
                else:
                    st.markdown("<b style='color:#EC2024;'>SIGRAMA</b>", unsafe_allow_html=True)
            with c_sig2:
                st.markdown("""<div style="font-size:12px; color:#334155; line-height:1.4;">
<b>Alejandra Arellano Machado</b><br>
<span style="background-color:#111111; color:white; padding:1px 6px; border-radius:3px; font-weight:700; font-size:11px;">Compras</span><br>
<a href="mailto:sarellano@sigrama.com.mx" style="color:#EC2024; text-decoration:none;">sarellano@sigrama.com.mx</a><br>
Tel: (871) 5-35-60-12 • <a href="http://www.sigrama.com.mx" target="_blank" style="color:#EC2024; text-decoration:none;">www.sigrama.com.mx</a><br>
Parque Industrial Rio XIX, Blvd La Ribereña No. 950 Torreón, Coahuila.
</div>""", unsafe_allow_html=True)

        st.write("")
        st.markdown("### ⚡ Procesador OCR & Extracción Inteligente")
        if st.button("🔍 Ejecutar Análisis OCR sobre '2608-3177 SIGRAMA METALES SAAM.PDF'", type="primary", use_container_width=True, key="btn_ocr_mail1"):
            with st.spinner("Analizando documento PDF, extrayendo tablas y asociando urgencias del correo..."):
                cab_extracted = {
                    'po': '2608-3177',
                    'fecha_pedido': datetime.date.today().strftime('%Y-%m-%d'),
                    'proyecto': 'SIGRAMA METALES / SAAM',
                    'solicitante': 'Alejandra Arellano Machado',
                    'requisicion': '22340',
                    'destino': 'Parque Industrial Rio XIX',
                    'proveedor': 'SIGRAMA PLANTA METALES',
                    'proveedor_atencion': 'JESUS MORALES',
                    'cliente_facturar_a': 'INDUSTRIA SIGRAMA S.A. DE C.V.',
                    'cliente_rfc': 'ISI-870204-K4A',
                    'cliente_direccion': 'Parque Industrial Rio XIX, Blvd La Ribereña No. 950, Torreón, Coahuila.',
                    'forma_pago': 'CONTADO / CRÉDITO',
                    'lab': 'ALMACEN SIGRAMA',
                    'tiempo_entrega': 'ENTREGA URGENTE PARCIAL',
                    'comprador': 'Josue Mesta / Alejandra Arellano',
                    'subtotal': 49280.00,
                    'descuento': 0.0,
                    'iva': 7884.80,
                    'ret_iva': 0.0,
                    'ret_isr': 0.0,
                    'total': 57164.80,
                    'moneda': 'MXN',
                    'observaciones': 'URGENCIA CLIENTE: Entregar primero 2 piezas de cada número de parte (P20325-24 y P20325-25).',
                    'texto_etiqueta': 'SAAM',
                    'color_fondo': '#22C55E',
                    'color_texto': '#FFFFFF'
                }
                
                part_extracted = [
                    {
                        'item_no': 1,
                        'clave_sku': 'P20325-24',
                        'descripcion_producto': 'PIEZA MAQUINADA / DOBLEZ P20325-24',
                        'cantidad_requerida': 64.0,
                        'unidad': 'PIEZA',
                        'precio_unitario': 385.00,
                        'precio_total': 24640.00,
                        'fecha_entrega': (datetime.date.today() + datetime.timedelta(days=5)).strftime('%Y-%m-%d'),
                        'parcialidad': 'P1',
                        'observaciones_partida': '🔥 Urgente: 2 pzas de inmediato. Resto: 4, 4, 6, 4, 6'
                    },
                    {
                        'item_no': 2,
                        'clave_sku': 'P20325-25',
                        'descripcion_producto': 'PIEZA MAQUINADA / DOBLEZ P20325-25',
                        'cantidad_requerida': 64.0,
                        'unidad': 'PIEZA',
                        'precio_unitario': 385.00,
                        'precio_total': 24640.00,
                        'fecha_entrega': (datetime.date.today() + datetime.timedelta(days=5)).strftime('%Y-%m-%d'),
                        'parcialidad': 'P1',
                        'observaciones_partida': '🔥 Urgente: 2 pzas de inmediato. Resto: 4, 4, 6, 4, 6'
                    }
                ]
                
                st.session_state['ocr_cab_email1'] = cab_extracted
                st.session_state['ocr_part_email1'] = part_extracted
                st.success("✅ ¡OCR y Análisis Completados con Éxito! Se detectaron 2 partidas con requerimientos de entrega urgente.")
                
        if 'ocr_cab_email1' in st.session_state and 'ocr_part_email1' in st.session_state:
            cab_e = st.session_state['ocr_cab_email1']
            part_e = st.session_state['ocr_part_email1']
            
            with st.container(border=True):
                st.markdown(f"#### 📋 Datos Extraídos para PO: `{cab_e['po']}`")
                st.write("**Partidas y Desglose de Entregas:**")
                st.dataframe(pd.DataFrame(part_e), use_container_width=True, hide_index=True)
                
                if st.button("🚀 Registrar y Guardar PO 2608-3177 en Base de Datos & Sincronizar", type="primary", use_container_width=True, key="btn_save_po1"):
                    ok_s, msg_s = save_po(cab_e, part_e)
                    if ok_s:
                        st.success(f"🎉 {msg_s}")
                        st.session_state.pop('ocr_cab_email1', None)
                        st.session_state.pop('ocr_part_email1', None)
                    else:
                        st.error(f"❌ {msg_s}")

    # --------------------------------------------------------------------------
    # PESTAÑA 2: CASO 2 - MULTI-PO EN UN SOLO CORREO (2603-2608 Y 2603-2609)
    # --------------------------------------------------------------------------
    with tab_mail_ejemplo2:
        with st.container(border=True):
            # Header estilo Outlook
            st.markdown("<h3 style='margin:0 0 10px 0; color:#1E293B;'>2603-2608, 2603-2609 SIGRAMA METALES</h3>", unsafe_allow_html=True)
            
            c2_snd1, c2_snd2 = st.columns([4, 1])
            with c2_snd1:
                st.markdown("""<div style="display:flex; align-items:center; gap:12px; margin-bottom:12px;">
<div style="width:46px; height:46px; border-radius:50%; background-color:#CBD5E1; color:#334155; display:flex; align-items:center; justify-content:center; font-weight:700; font-size:18px;">
AA
</div>
<div>
<div style="font-weight:700; font-size:16px; color:#111111;">
Alejandra Arellano Machado <span style="background-color:#E2E8F0; color:#475569; font-size:11px; padding:2px 8px; border-radius:10px;">Compras</span>
</div>
<div style="font-size:12px; color:#64748B;">
<b>Para:</b> Jesus Alberto Morales Lopez<br>
<b>CC:</b> Edgar Sosa Suarez; Josue Mesta; Elsa Cardenas Elizondo; Moises Aaron Hernandez Valdez; <b>y 2 usuarios más</b><br>
<span style="color:#0284C7; font-size:11px;">ℹ️ Mensaje reenviado el 15/07/2026 04:39 p. m.</span>
</div>
</div>
</div>""", unsafe_allow_html=True)
                
            with c2_snd2:
                st.markdown("<div style='text-align:right; font-size:12px; color:#64748B;'>15/07/2026</div>", unsafe_allow_html=True)
                c2_btn1, c2_btn2, c2_btn3 = st.columns(3)
                with c2_btn1:
                    st.button("↩️", help="Responder", key="btn_reply_mail2")
                with c2_btn2:
                    st.button("🔁", help="Reenviar", key="btn_fwd_mail2")
                with c2_btn3:
                    st.button("⋯", help="Más opciones", key="btn_more_mail2")
                    
            st.divider()
            
            # Cajas de Archivos Adjuntos Múltiples (2 POs + 1 Plano)
            st.markdown("##### 📎 3 Archivos Adjuntos Detectados en este Correo:")
            col_m1, col_m2, col_m3 = st.columns(3)
            with col_m1:
                st.markdown("""<div style="border:2px solid #EC2024; border-radius:8px; padding:10px 14px; display:flex; align-items:center; gap:10px; background:#FEF2F2;">
<span style="font-size:24px;">📄</span>
<div>
<div style="font-weight:700; font-size:12px; color:#991B1B;">2603-2608 SIGRAMA METALES JMC.PDF</div>
<div style="font-size:11px; color:#64748B;">180 KB • <b>PO #1</b></div>
</div>
</div>""", unsafe_allow_html=True)
            with col_m2:
                st.markdown("""<div style="border:2px solid #EC2024; border-radius:8px; padding:10px 14px; display:flex; align-items:center; gap:10px; background:#FEF2F2;">
<span style="font-size:24px;">📄</span>
<div>
<div style="font-weight:700; font-size:12px; color:#991B1B;">2603-2609 SIGRAMA METALES JMC.PDF</div>
<div style="font-size:11px; color:#64748B;">180 KB • <b>PO #2</b></div>
</div>
</div>""", unsafe_allow_html=True)
            with col_m3:
                st.markdown("""<div style="border:2px solid #0284C7; border-radius:8px; padding:10px 14px; display:flex; align-items:center; gap:10px; background:#F0F9FF;">
<span style="font-size:24px;">📐</span>
<div>
<div style="font-weight:700; font-size:12px; color:#0369A1;">11-7761 (rev.03).pdf</div>
<div style="font-size:11px; color:#64748B;">Plano Técnico / Dibujo</div>
</div>
</div>""", unsafe_allow_html=True)
                
            st.write("")
            
            # Mensaje del Correo
            st.markdown("""
<div style="font-size:14px; color:#1E293B; line-height:1.6; margin: 15px 0;">
Ing Jesus Morales, buenas tardes.<br><br>
En adjunto comparto nuestras órdenes de compra <b>2603-2608</b>, <b>2603-2609</b>. Por favor considerar el <b>plano anexo (11-7761 rev.03)</b>.<br><br>
Gracias, quedamos al pendiente.<br><br>
Saludos.
</div>
""", unsafe_allow_html=True)
            
            st.divider()
            
            # Firma
            c2_sig1, c2_sig2 = st.columns([1, 3])
            with c2_sig1:
                logo_path = Path(__file__).resolve().parent / "logo_sigrama.png"
                if logo_path.exists():
                    st.image(str(logo_path), width=130)
                else:
                    st.markdown("<b style='color:#EC2024;'>SIGRAMA</b>", unsafe_allow_html=True)
            with c2_sig2:
                st.markdown("""<div style="font-size:12px; color:#334155; line-height:1.4;">
<b>Alejandra Arellano Machado</b><br>
<span style="background-color:#111111; color:white; padding:1px 6px; border-radius:3px; font-weight:700; font-size:11px;">Compras</span><br>
<a href="mailto:sarellano@sigrama.com.mx" style="color:#EC2024; text-decoration:none;">sarellano@sigrama.com.mx</a> • Tel: (871) 5-35-60-12<br>
Parque Industrial Rio XIX, Torreón, Coahuila.
</div>""", unsafe_allow_html=True)

        st.write("")
        st.markdown("### ⚡ Extractor Inteligente Multi-PO por Lotes")
        st.write("El sistema detectó **2 archivos de Órdenes de Compra independientes** y **1 Plano Técnico anexo** en el mismo correo:")
        
        if st.button("🔍 Ejecutar Análisis OCR Multi-PO sobre '2603-2608' y '2603-2609'", type="primary", use_container_width=True, key="btn_ocr_multi_po"):
            with st.spinner("Procesando lote de PDFs, asociando plano técnico 11-7761 y extrayendo requerimientos..."):
                
                # PO 1: 2603-2608
                po1_cab = {
                    'po': '2603-2608',
                    'fecha_pedido': '2026-07-15',
                    'proyecto': 'LC8 20K / SWBDCLOUD',
                    'solicitante': 'Alejandra Arellano Machado',
                    'requisicion': '22180',
                    'destino': 'ALMACEN SIGRAMA',
                    'proveedor': 'SIGRAMA PLANTA METALES',
                    'proveedor_atencion': 'JESUS MORALES',
                    'cliente_facturar_a': 'INDUSTRIA SIGRAMA S.A. DE C.V.',
                    'cliente_rfc': 'ISI-870204-K4A',
                    'cliente_direccion': 'Parque Industrial Rio XIX, Torreón, Coahuila.',
                    'forma_pago': 'CONTADO / CRÉDITO',
                    'lab': 'ALMACEN SIGRAMA',
                    'tiempo_entrega': 'PROGRAMADO',
                    'comprador': 'Alejandra Arellano / Josue Mesta',
                    'subtotal': 38500.00,
                    'descuento': 0.0,
                    'iva': 6160.00,
                    'ret_iva': 0.0,
                    'ret_isr': 0.0,
                    'total': 44660.00,
                    'moneda': 'MXN',
                    'observaciones': 'Fabricar conforme a plano adjunto 11-7761 rev.03.',
                    'texto_etiqueta': 'LC8 20K',
                    'color_fondo': '#0284C7',
                    'color_texto': '#FFFFFF'
                }
                po1_part = [
                    {
                        'item_no': 1,
                        'clave_sku': '11-7761-03',
                        'descripcion_producto': 'PIEZA CORTADA/DOBLADA 11-7761 REV.03 (LC8 20K)',
                        'cantidad_requerida': 17.0,
                        'unidad': 'PIEZA',
                        'precio_unitario': 770.00,
                        'precio_total': 13090.00,
                        'fecha_entrega': '2026-07-30',
                        'parcialidad': 'P1',
                        'observaciones_partida': 'Plano anexo 11-7761 (rev.03).pdf'
                    },
                    {
                        'item_no': 2,
                        'clave_sku': '11-7761-02',
                        'descripcion_producto': 'PIEZA CORTADA/DOBLADA 11-7761 REV.03 (SWBDCLOUD)',
                        'cantidad_requerida': 33.0,
                        'unidad': 'PIEZA',
                        'precio_unitario': 770.00,
                        'precio_total': 25410.00,
                        'fecha_entrega': '2026-07-30',
                        'parcialidad': 'P1',
                        'observaciones_partida': 'Plano anexo 11-7761 (rev.03).pdf'
                    }
                ]
                
                # PO 2: 2603-2609
                po2_cab = {
                    'po': '2603-2609',
                    'fecha_pedido': '2026-07-15',
                    'proyecto': 'SWBDCLOUD',
                    'solicitante': 'Alejandra Arellano Machado',
                    'requisicion': '22181',
                    'destino': 'ALMACEN SIGRAMA',
                    'proveedor': 'SIGRAMA PLANTA METALES',
                    'proveedor_atencion': 'JESUS MORALES',
                    'cliente_facturar_a': 'INDUSTRIA SIGRAMA S.A. DE C.V.',
                    'cliente_rfc': 'ISI-870204-K4A',
                    'cliente_direccion': 'Parque Industrial Rio XIX, Torreón, Coahuila.',
                    'forma_pago': 'CONTADO / CRÉDITO',
                    'lab': 'ALMACEN SIGRAMA',
                    'tiempo_entrega': 'PROGRAMADO',
                    'comprador': 'Alejandra Arellano / Josue Mesta',
                    'subtotal': 30800.00,
                    'descuento': 0.0,
                    'iva': 4928.00,
                    'ret_iva': 0.0,
                    'ret_isr': 0.0,
                    'total': 35728.00,
                    'moneda': 'MXN',
                    'observaciones': 'Fabricar conforme a plano adjunto 11-7761 rev.03.',
                    'texto_etiqueta': 'SWBDCLOUD',
                    'color_fondo': '#7C3AED',
                    'color_texto': '#FFFFFF'
                }
                po2_part = [
                    {
                        'item_no': 1,
                        'clave_sku': '11-7761-02',
                        'descripcion_producto': 'PIEZA CORTADA/DOBLADA 11-7761 REV.03 (SWBDCLOUD)',
                        'cantidad_requerida': 40.0,
                        'unidad': 'PIEZA',
                        'precio_unitario': 770.00,
                        'precio_total': 30800.00,
                        'fecha_entrega': '2026-07-30',
                        'parcialidad': 'P1',
                        'observaciones_partida': 'Plano anexo 11-7761 (rev.03).pdf'
                    }
                ]
                
                st.session_state['multi_po_batch'] = [
                    {'cab': po1_cab, 'part': po1_part},
                    {'cab': po2_cab, 'part': po2_part}
                ]
                st.success("✅ ¡Extracción Multi-PO Completada! Se procesaron 2 Órdenes de Compra vinculadas al Plano Técnico 11-7761.")

        if 'multi_po_batch' in st.session_state:
            batch = st.session_state['multi_po_batch']
            
            col_b1, col_b2 = st.columns(2)
            
            with col_b1:
                with st.container(border=True):
                    st.markdown(f"#### 📄 PO #1: `{batch[0]['cab']['po']}`")
                    st.markdown(f"• **Proyecto:** `{batch[0]['cab']['proyecto']}`")
                    st.markdown(f"• **Total Piezas:** `{sum(p['cantidad_requerida'] for p in batch[0]['part']):.0f} pzas`")
                    st.markdown(f"• **Importe Total:** `${batch[0]['cab']['total']:,.2f} MXN`")
                    st.dataframe(pd.DataFrame(batch[0]['part'])[['clave_sku', 'cantidad_requerida', 'precio_total', 'observaciones_partida']], use_container_width=True, hide_index=True)
                    
            with col_b2:
                with st.container(border=True):
                    st.markdown(f"#### 📄 PO #2: `{batch[1]['cab']['po']}`")
                    st.markdown(f"• **Proyecto:** `{batch[1]['cab']['proyecto']}`")
                    st.markdown(f"• **Total Piezas:** `{sum(p['cantidad_requerida'] for p in batch[1]['part']):.0f} pzas`")
                    st.markdown(f"• **Importe Total:** `${batch[1]['cab']['total']:,.2f} MXN`")
                    st.dataframe(pd.DataFrame(batch[1]['part'])[['clave_sku', 'cantidad_requerida', 'precio_total', 'observaciones_partida']], use_container_width=True, hide_index=True)
                    
            if st.button("🚀 Registrar Lote Completo (Guardar Ambas POs en Sistema)", type="primary", use_container_width=True, key="btn_save_multi_batch"):
                ok1, msg1 = save_po(batch[0]['cab'], batch[0]['part'])
                ok2, msg2 = save_po(batch[1]['cab'], batch[1]['part'])
                
                if ok1 and ok2:
                    st.success(f"🎉 ¡Lote registrado con éxito! POs {batch[0]['cab']['po']} y {batch[1]['cab']['po']} agregadas y sincronizadas con Remisiones.")
                    st.session_state.pop('multi_po_batch', None)
                else:
                    st.error(f"Error: {msg1} | {msg2}")

    # --------------------------------------------------------------------------
    # PESTAÑA 3: SUBIR NUEVO CORREO / MULTI-PDFS EXTERNOS
    # --------------------------------------------------------------------------
    with tab_mail_custom:
        st.subheader("📥 Carga Asistida de Correo con Múltiples PDFs / Planos")
        st.write("Puedes subir **múltiples archivos PDF simultáneamente** (varias POs y planos técnicos):")
        
        c_cu1, c_cu2 = st.columns([1, 1])
        with c_cu1:
            custom_email_subject = st.text_input("Asunto del Correo:", placeholder="ej. 2603-XXXX, 2603-YYYY SIGRAMA METALES", key="multi_subj")
            custom_email_sender = st.text_input("Remitente / Comprador:", placeholder="ej. Nombre del solicitante", key="multi_send")
            custom_email_body = st.text_area("Cuerpo del Correo (Pega el texto con folios y notas):", height=160, placeholder="Pega aquí el mensaje del cliente...", key="multi_body")
        with c_cu2:
            custom_pdfs = st.file_uploader(
                "Adjuntar Archivo(s) de PO de Sigrama (.PDF) o Correos (.MSG):",
                type=['pdf', 'msg'],
                accept_multiple_files=True,
                key="uploader_multi_custom_pdfs"
            )
            if custom_pdfs:
                st.success(f"📎 Se cargaron **{len(custom_pdfs)}** archivo(s).")
                for f in custom_pdfs:
                    st.caption(f"• `{f.name}` ({f.size / 1024:.1f} KB)")
                
        if custom_pdfs:
            if st.button("⚡ Procesar y Extraer Órdenes de Compra de Sigrama", type="primary", use_container_width=True):
                with st.spinner("Analizando documentos de PO oficiales de Sigrama y procesando OCR..."):
                    extracted_batch = []
                    ctx = parse_email_text(custom_email_body or custom_email_subject)
                    if custom_email_sender:
                        ctx['remitente'] = custom_email_sender
                        
                    for uploaded_f in custom_pdfs:
                        f_bytes = uploaded_f.read()
                        f_name = uploaded_f.name
                        
                        # Si es archivo .msg, desempaquetar sus PDFs adjuntos
                        if f_name.lower().endswith('.msg'):
                            from pdf_parser import extract_attachments_from_msg
                            msg_info = extract_attachments_from_msg(f_bytes)
                            ctx_msg = parse_email_text(msg_info.get('body', ''))
                            ctx_msg['remitente'] = msg_info.get('sender', '')
                            
                            for att in msg_info.get('attachments', []):
                                att_n = att['filename']
                                if att_n.lower().endswith('.pdf') and not any(w in att_n.lower() for w in ['plano', 'drawing', 'cotizacion']):
                                    try:
                                        cab_m, part_m = parse_po_pdf(att['data'], email_context=ctx_msg)
                                        extracted_batch.append({'cab': cab_m, 'part': part_m, 'file_name': f"{f_name} ➔ {att_n}"})
                                    except Exception as e_att:
                                        st.error(f"Error en {att_n}: {e_att}")
                        else:
                            # Es archivo .pdf directo
                            if any(w in f_name.lower() for w in ['plano', 'drawing', 'rev', 'cotizacion']):
                                continue
                            try:
                                cab_f, part_f = parse_po_pdf(f_bytes, email_context=ctx)
                                extracted_batch.append({'cab': cab_f, 'part': part_f, 'file_name': f_name})
                            except Exception as e:
                                st.error(f"Error procesando {f_name}: {e}")
                            
                    st.session_state['uploaded_batch_extracted'] = extracted_batch
                    st.success(f"✅ Se extrajeron exitosamente **{len(extracted_batch)}** Órdenes de Compra de Sigrama.")
                    
        if 'uploaded_batch_extracted' in st.session_state and st.session_state['uploaded_batch_extracted']:
            u_batch = st.session_state['uploaded_batch_extracted']
            st.write("---")
            st.markdown(f"### 📋 Lote de {len(u_batch)} Órdenes Detectadas")
            
            for idx, item in enumerate(u_batch):
                with st.expander(f"📄 PO: {item['cab'].get('po', 'N/A')} ({item['file_name']})", expanded=True):
                    c_b1, c_b2, c_b3 = st.columns(3)
                    with c_b1:
                        st.markdown(f"• **Folio:** `{item['cab'].get('po')}`")
                        st.markdown(f"• **Proyecto:** `{item['cab'].get('proyecto')}`")
                    with c_b2:
                        st.markdown(f"• **Solicitante:** `{item['cab'].get('solicitante')}`")
                        st.markdown(f"• **Total:** `${item['cab'].get('total', 0):,.2f} MXN`")
                    with c_b3:
                        st.markdown(f"• **Partidas:** `{len(item['part'])}`")
                    st.dataframe(pd.DataFrame(item['part']), use_container_width=True, hide_index=True)
                    
            if st.button("🚀 Confirmar y Guardar Todo el Lote en Sistema", type="primary", use_container_width=True, key="btn_save_uploaded_batch"):
                total_ok = 0
                for item in u_batch:
                    ok_u, _ = save_po(item['cab'], item['part'])
                    if ok_u:
                        total_ok += 1
                st.success(f"🎉 Se guardaron y sincronizaron **{total_ok} de {len(u_batch)}** órdenes de compra.")
                st.session_state.pop('uploaded_batch_extracted', None)


# ==============================================================================
# SECCIÓN 4: REGISTRAR / CARGAR PO
# ==============================================================================
elif menu == "📥 Registrar / Cargar PO":
    st.title("📥 Registro y Carga de Órdenes de Compra (POs)")
    st.markdown("Ingresa una nueva Orden de Compra mediante Formulario Manual, Carga Masiva en Excel o Extracción Inteligente de PDF.")
    
    tab_pdf, tab_excel, tab_manual = st.tabs([
        "📄 Lector Inteligente de PDF (Recomendado)",
        "📁 Carga Masiva de Excel",
        "📝 Formulario Manual Guiado"
    ])
    
    # 1. Pestaña PDF
    with tab_pdf:
        st.subheader("📄 Cargar y Extraer PDF de Orden de Compra")
        st.write("Sube el PDF oficial emitido por el cliente o compras (ej. Industria Sigrama) para extraer automáticamente toda la cabecera, partidas, importes y fechas.")
        
        pdf_file = st.file_uploader("Selecciona el archivo PDF de la PO:", type=['pdf'], key="uploader_po_pdf")
        
        if pdf_file is not None:
            if st.button("⚡ Procesar y Extraer Datos del PDF", type="primary", use_container_width=True):
                with st.spinner("Analizando PDF y extrayendo campos..."):
                    try:
                        cab_extracted, part_extracted = parse_po_pdf(pdf_file.read())
                        st.session_state['temp_cab_extracted'] = cab_extracted
                        st.session_state['temp_part_extracted'] = part_extracted
                        st.success(f"✅ ¡Extracción exitosa! Folio detectado: **{cab_extracted.get('po', 'N/A')}** con **{len(part_extracted)}** partidas.")
                    except Exception as e:
                        st.error(f"❌ Error al procesar el PDF: {e}")
                        
        if 'temp_cab_extracted' in st.session_state and 'temp_part_extracted' in st.session_state:
            cab = st.session_state['temp_cab_extracted']
            part = st.session_state['temp_part_extracted']
            
            st.write("---")
            st.markdown("### 📋 Vista Previa de Datos Extraídos")
            
            c_p1, c_p2, c_p3 = st.columns(3)
            with c_p1:
                cab['po'] = st.text_input("Folio PO:", cab.get('po', ''), key="pdf_po")
                cab['fecha_pedido'] = st.text_input("Fecha Pedido:", cab.get('fecha_pedido', ''), key="pdf_fped")
                cab['requisicion'] = st.text_input("Requisición:", cab.get('requisicion', ''), key="pdf_req")
            with c_p2:
                cab['proyecto'] = st.text_input("Proyecto / Uso:", cab.get('proyecto', ''), key="pdf_proy")
                cab['solicitante'] = st.text_input("Solicitante:", cab.get('solicitante', ''), key="pdf_sol")
                cab['comprador'] = st.text_input("Comprador:", cab.get('comprador', ''), key="pdf_comp")
            with c_p3:
                cab['lab'] = st.text_input("L.A.B. / Destino:", cab.get('lab', 'ALMACEN SIGRAMA'), key="pdf_lab")
                cab['tiempo_entrega'] = st.text_input("Tiempo Entrega:", cab.get('tiempo_entrega', ''), key="pdf_tent")
                cab['total'] = st.number_input("Total ($ MXN):", value=float(cab.get('total', 0) or 0), key="pdf_tot")
                
            cab['observaciones'] = st.text_area("Observaciones:", cab.get('observaciones', ''), key="pdf_obs")
            
            st.write("**Partidas Extraídas:**")
            df_part_edit = pd.DataFrame(part)
            st.dataframe(df_part_edit, use_container_width=True, hide_index=True)
            
            if st.button("🚀 Confirmar y Guardar Orden de Compra en Sistema", type="primary", use_container_width=True):
                ok, msg = save_po(cab, part)
                if ok:
                    st.success(f"✅ {msg}")
                    st.session_state.pop('temp_cab_extracted', None)
                    st.session_state.pop('temp_part_extracted', None)
                else:
                    st.error(f"❌ {msg}")
                    
    # 2. Pestaña Excel
    with tab_excel:
        st.subheader("📁 Carga de Requerimientos mediante Plantilla Excel")
        st.write("Descarga la plantilla oficial estandarizada, llena los datos de cabecera y partidas/parcialidades, y cárgala aquí.")
        
        template_bytes = generate_po_excel_template()
        st.download_button(
            label="📥 Descargar Plantilla Oficial Excel (2 Pestañas)",
            data=template_bytes,
            file_name="Plantilla_Oficial_Orden_de_Compra_Sigrama.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        st.write("---")
        excel_file = st.file_uploader("Subir Archivo Excel completado:", type=['xlsx'], key="uploader_po_excel")
        
        if excel_file is not None:
            if st.button("🚀 Procesar e Integrar Archivo Excel", type="primary", use_container_width=True):
                ok, msg, cab_xl, part_xl = parse_uploaded_excel(excel_file)
                if ok:
                    ok_save, msg_save = save_po(cab_xl, part_xl)
                    if ok_save:
                        st.success(f"✅ {msg_save}")
                    else:
                        st.error(f"❌ {msg_save}")
                else:
                    st.error(f"❌ {msg}")
                    
    # 3. Pestaña Manual
    with tab_manual:
        st.subheader("📝 Captura Manual Guiada de Orden de Compra")
        
        with st.form("form_manual_po"):
            st.markdown("##### 1. Encabezado de la PO")
            m1, m2, m3 = st.columns(3)
            with m1:
                man_po = st.text_input("Folio de PO (Obligatorio):", placeholder="ej. 26083186")
                man_fecha = st.date_input("Fecha de Pedido:", value=datetime.date.today())
                man_req = st.text_input("No. Requisición:", placeholder="ej. 22326")
            with m2:
                man_proy = st.text_input("Proyecto / Uso:", placeholder="ej. CLOUD / TAB-RQXP")
                man_sol = st.text_input("Solicitante:", placeholder="ej. ESEFANIA IBARRA")
                man_comp = st.text_input("Comprador:", placeholder="ej. Josue Mesta")
            with m3:
                man_lab = st.text_input("L.A.B. / Destino:", value="ALMACEN SIGRAMA")
                man_tent = st.text_input("Tiempo de Entrega:", placeholder="ej. 18 AGOSTO 2026")
                man_obs = st.text_input("Observaciones Generales:")
                
            st.markdown("##### 2. Partida Principal (o inicial)")
            p1, p2, p3, p4 = st.columns([1.5, 3, 1, 1])
            with p1:
                man_sku = st.text_input("Clave / SKU:", placeholder="SWB01431")
            with p2:
                man_desc = st.text_input("Descripción del Producto:", placeholder="PP19380-03 BLANK DOOR")
            with p3:
                man_cant = st.number_input("Cantidad:", min_value=1.0, value=32.0, step=1.0)
            with p4:
                man_pu = st.number_input("P. Unitario ($):", min_value=0.0, value=385.55, step=10.0)
                
            man_submit = st.form_submit_button("💾 Guardar Orden de Compra", type="primary", use_container_width=True)
            
            if man_submit:
                if not man_po.strip():
                    st.error("❌ El Folio de la PO es obligatorio.")
                elif not man_sku.strip():
                    st.error("❌ Debes ingresar al menos una Clave/SKU.")
                else:
                    pt_val = man_cant * man_pu
                    sub_val = pt_val
                    iva_val = sub_val * 0.16
                    tot_val = sub_val + iva_val
                    
                    cab_man = {
                        "po": man_po.strip(),
                        "fecha_pedido": man_fecha.strftime("%Y-%m-%d"),
                        "proyecto": man_proy.strip(),
                        "solicitante": man_sol.strip(),
                        "requisicion": man_req.strip(),
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
                        "subtotal": sub_val,
                        "descuento": 0.0,
                        "iva": iva_val,
                        "ret_iva": 0.0,
                        "ret_isr": 0.0,
                        "total": tot_val,
                        "moneda": "MXN",
                        "observaciones": man_obs.strip(),
                        "texto_etiqueta": man_proy.strip(),
                        "color_fondo": "#EC2024",
                        "color_texto": "#FFFFFF"
                    }
                    
                    part_man = [{
                        "item_no": 1,
                        "clave_sku": man_sku.strip().upper(),
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
                    else:
                        st.error(f"❌ {msg}")


# ==============================================================================
# SECCIÓN 5: ESTADO DE INTEGRACIÓN
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
        
    st.write("---")
    st.subheader("3. Conexión Futura con Corte y Doblez (Fase 2)")
    st.info("📌 En la Fase 2, la app de Corte y Doblez se conectará a las partidas registradas aquí para vincular las Órdenes de Fabricación (OFs), anidamientos y avances por estación láser/doblez directamente a cada PO.")


# ==============================================================================
# SECCIÓN 6: MANUAL, INDUSTRIA 4.0 & STACK TECNOLÓGICO
# ==============================================================================
elif menu == "📘 Manual & Arquitectura Industria 4.0":
    st.title("📘 Manual, Arquitectura Industria 4.0 & Stack Tecnológico")
    st.markdown("Documentación técnica, diagramas de arquitectura ciberfísica y manual operativo para el **PO Tracker & Master Hub de Industria Sigrama**.")
    
    tab_ind4, tab_diagrama_ocr, tab_stack, tab_manual_pasos = st.tabs([
        "🏭 Arquitectura Industria 4.0",
        "⚡ Motor OCR Espacial (Diagrama Oficial)",
        "💻 Stack Tecnológico (Tech Stack)",
        "📖 Manual Operativo Paso a Paso"
    ])
    
    # --------------------------------------------------------------------------
    # PESTAÑA 1: ARQUITECTURA INDUSTRIA 4.0
    # --------------------------------------------------------------------------
    with tab_ind4:
        st.subheader("🏭 El Hilo Digital (Digital Thread) en Industria Sigrama")
        st.write("""
En la manufactura moderna bajo el paradigma de **Industria 4.0**, la información debe fluir sin interrupciones desde el requerimiento del cliente hasta el embarque final.
El **PO Tracker & Master Hub** actúa como el **Gemelo Digital del Requerimiento (Digital Thread)**, asegurando que cada orden de compra se convierta en una orden sistemática interconectada.
""")
        
        # Diagrama de Interconexión Ecosistema
        st.markdown("""<div style="background-color:#0F172A; color:white; border-radius:10px; padding:20px; margin:15px 0;">
<h4 style="color:#EC2024; margin-top:0; text-align:center;">ECOSISTEMA CIBERFÍSICO INTEGRADO DE INDUSTRIA SIGRAMA</h4>

<div style="display:flex; justify-content:space-between; gap:15px; margin-top:20px;">
<div style="flex:1; background-color:#1E293B; border:2px solid #EC2024; border-radius:8px; padding:14px; text-align:center;">
<span style="font-size:24px;">📥</span>
<h5 style="color:#EC2024; margin:8px 0 4px 0;">1. CAPTURA Y MASTER HUB</h5>
<b style="color:white; font-size:13px;">PO Tracker App</b>
<p style="font-size:11px; color:#94A3B8; margin-top:6px;">Ingesta de POs (.PDF / .MSG), OCR inteligente, extracción de requerimientos y estandarización.</p>
</div>

<div style="flex:1; background-color:#1E293B; border:2px solid #0284C7; border-radius:8px; padding:14px; text-align:center;">
<span style="font-size:24px;">⚙️</span>
<h5 style="color:#0284C7; margin:8px 0 4px 0;">2. MANUFACTURA CIBERFÍSICA</h5>
<b style="color:white; font-size:13px;">App Corte y Doblez</b>
<p style="font-size:11px; color:#94A3B8; margin-top:6px;">Programación de OFs, anidamientos CNC, corte láser, doblez y avance por estaciones.</p>
</div>

<div style="flex:1; background-color:#1E293B; border:2px solid #16A34A; border-radius:8px; padding:14px; text-align:center;">
<span style="font-size:24px;">🚚</span>
<h5 style="color:#16A34A; margin:8px 0 4px 0;">3. LOGÍSTICA & DESPACHO</h5>
<b style="color:white; font-size:13px;">App Remisiones</b>
<p style="font-size:11px; color:#94A3B8; margin-top:6px;">Inspección en almacén, consolidación en tarimas, folios de remisión y entrega al cliente.</p>
</div>
</div>
</div>""", unsafe_allow_html=True)
        
        c_i1, c_i2 = st.columns(2)
        with c_i1:
            with st.container(border=True):
                st.markdown("##### 🎯 Pilares Industria 4.0 Aplicados:")
                st.markdown("• **Interoperabilidad en Tiempo Real**: Eliminación del 'efecto isla', los 3 sistemas comparten datos continuamente sin captura manual duplicada.")
                st.markdown("• **Transparencia de Información**: Visualización en tiempo real del saldo de piezas fabricadas en taller vs despachadas.")
                st.markdown("• **Digitalización de Documentos**: Transformación automática de órdenes no estructuradas (PDF/Email) a entidades estructuradas relacionales.")
        with c_i2:
            with st.container(border=True):
                st.markdown("##### 🚀 Beneficios Operativos y de Negocio:")
                st.markdown("• **Cero Pérdida de Trazabilidad**: Cada tornillo, puerta y pieza maquinada tiene su origen en una PO formal.")
                st.markdown("• **Control de Urgencias**: Detección inmediata de entregas prioritarias marcadas por compras o clientes.")
                st.markdown("• **Reducción de Tiempos de Ciclo**: De horas de captura manual a segundos mediante OCR espacial.")

    # --------------------------------------------------------------------------
    # PESTAÑA 2: MOTOR OCR ESPACIAL (DIAGRAMA OFICIAL)
    # --------------------------------------------------------------------------
    with tab_diagrama_ocr:
        st.subheader("¿Cómo analiza el motor los archivos de PO de Sigrama?")
        st.markdown("El procesador utiliza un **motor de lectura espacial adaptado a la hoja membretada de Industria Sigrama**:")
        
        # Renderizado visual del diagrama de la imagen
        st.markdown("""<div style="background-color:#111827; border:1px solid #374151; border-radius:12px; padding:24px; color:white; font-family:'Questrial',sans-serif; margin: 15px 0;">
<div style="display:flex; align-items:center; justify-content:space-between; gap:15px;">

<!-- Entrada -->
<div style="background:#1F2937; border:1px solid #4B5563; border-radius:8px; padding:14px 18px; text-align:center; min-width:140px;">
<span style="font-size:22px;">📄</span><br>
<b style="font-size:13px; color:#F9FAFB;">Archivo PO Sigrama</b><br>
<span style="font-size:11px; color:#9CA3AF;">.PDF directo o .MSG</span>
</div>

<div style="color:#60A5FA; font-size:20px; font-weight:bold;">➔</div>

<!-- Extractor -->
<div style="background:#1E3A8A; border:1px solid #3B82F6; border-radius:8px; padding:14px 18px; text-align:center; min-width:150px;">
<span style="font-size:20px;">⚡</span><br>
<b style="font-size:13px; color:#93C5FD;">Extractor OCR Espacial Sigrama</b>
</div>

<div style="color:#60A5FA; font-size:20px; font-weight:bold;">➔</div>

<!-- Centro: Bloques Extraídos -->
<div style="border:1px dashed #6B7280; border-radius:10px; padding:15px; background:#111827; flex:2;">
<div style="font-size:11px; color:#9CA3AF; text-align:center; font-weight:700; margin-bottom:10px; letter-spacing:1px;">
DATOS EXTRAÍDOS DEL FORMATO OFICIAL
</div>

<div style="display:flex; flex-direction:column; gap:10px;">
<div style="background:#1F2937; border:1px solid #374151; border-radius:6px; padding:10px; font-size:11px;">
<b style="color:#F3F4F6;">📋 Cabecera Oficial:</b><br>
<span style="color:#9CA3AF;">• Folio PO (ej. 2608-3177 / 26083186) • Fecha Pedido (Día/Mes/Año)<br>
• Requisición No. & Solicitante • Proveedor & Atención (Jesús Morales)<br>
• Destino (Almacén Sigrama) • Comprador (Josué Mesta / Compras)<br>
• Observaciones (Cuentas, Proyectos)</span>
</div>

<div style="background:#1F2937; border:1px solid #374151; border-radius:6px; padding:10px; font-size:11px;">
<b style="color:#F3F4F6;">📦 Tabla de Partidas:</b><br>
<span style="color:#9CA3AF;">• Item # & Cantidad (ej. 32.00, 64.00) • Unidad (PIEZA, PZA, KG, JGO)<br>
• Clave / SKU (ej. SWB01431, P20325-24) • Descripción Oficial del Producto<br>
• Precio Unitario & Precio Total ($) • Fecha de Entrega Prometida</span>
</div>

<div style="background:#1F2937; border:1px solid #374151; border-radius:6px; padding:10px; font-size:11px;">
<b style="color:#F3F4F6;">💰 Totales Financieros:</b><br>
<span style="color:#9CA3AF;">• Subtotal • IVA 16% • Total Neto MXN</span>
</div>
</div>
</div>

<div style="color:#60A5FA; font-size:20px; font-weight:bold;">➔</div>

<!-- Salida -->
<div style="background:#064E3B; border:1px solid #10B981; border-radius:8px; padding:14px 18px; text-align:center; min-width:150px;">
<span style="font-size:20px;">🚀</span><br>
<b style="font-size:13px; color:#A7F3D0;">Sincronización Automática</b><br>
<span style="font-size:10px; color:#D1FAE5;">(Remisiones + Corte y Doblez)</span>
</div>

</div>
</div>""", unsafe_allow_html=True)

    # --------------------------------------------------------------------------
    # PESTAÑA 3: STACK TECNOLÓGICO (TECH STACK)
    # --------------------------------------------------------------------------
    with tab_stack:
        st.subheader("💻 Stack Tecnológico de la Aplicación")
        st.markdown("Tecnologías y librerías de grado industrial que componen la solución:")
        
        c_st1, c_st2 = st.columns(2)
        with c_st1:
            with st.container(border=True):
                st.markdown("#### 🎨 Frontend & Experiencia de Usuario")
                st.markdown("""
- **Framework UI:** Streamlit 1.40+ (Python)
- **Visualización de Datos:** Plotly Express & Plotly Graph Objects
- **Estilos Corporativos:** CSS3 Moderno (Tipografía *Montserrat* y *Questrial*, Colores Oficiales Sigrama `#EC2024`, `#111111`)
- **Diseño Adaptativo:** Contenedores interactivos, métricas dinámicas y feedback en tiempo real.
""")
                
            with st.container(border=True):
                st.markdown("#### 👁️ Motor de Ingesta & Extracción OCR")
                st.markdown("""
- **Lector Espacial de PDFs:** PyMuPDF 1.28 (`fitz`) con detección de bloques vectoriales y coordenadas `(x0, y0, x1, y1)`.
- **Desempaquetador de Correos:** `extract-msg` 0.56 (Parseador nativo de archivos de mensaje OLE de Microsoft Outlook).
- **Procesamiento de Texto:** Expresiones regulares adaptativas (RegEx) para detección de folios, fechas, precios y partidas.
""")
                
        with c_st2:
            with st.container(border=True):
                st.markdown("#### 🗄️ Persistencia de Datos & Almacenamiento")
                st.markdown("""
- **Base de Datos Relacional:** SQLite 3 (`po_tracker.db`) con soporte ACID y relaciones integridad referencial (`po_cabecera` ➔ `po_partidas` ➔ `po_historial`).
- **Data Wrangling:** Pandas 2.2+ para manipulación tabular de alto rendimiento.
- **Intercambio en Excel:** OpenPyXL 3.1+ para exportación automática e importación de plantillas oficiales.
""")
                
            with st.container(border=True):
                st.markdown("#### 🔄 Integración & Interoperabilidad")
                st.markdown("""
- **Mapeo Normalizado:** Algoritmo de normalización de folios (ej. `PO 2608-3186` ➔ `26083186`) para correlación exacta.
- **Sincronización Automática:** Espejo bidireccional hacia `BD_POs_Cabecera.xlsx` y `BD_Requerimientos_POs.xlsx` para consumo directo por la app de Remisiones.
""")

    # --------------------------------------------------------------------------
    # PESTAÑA 4: MANUAL OPERATIVO PASO A PASO
    # --------------------------------------------------------------------------
    with tab_manual_pasos:
        st.subheader("📖 Manual de Operación para el Usuario")
        st.markdown("Guía práctica para operar el sistema en el día a día:")
        
        with st.container(border=True):
            st.markdown("### 🔹 Paso 1: Ingreso de Órdenes de Compra (3 Opciones)")
            st.markdown("""
1. **Opción A (Archivo PDF o Correo .MSG)**: Ve a **`📬 Bandeja de Correos & OCR`** o **`📥 Registrar / Cargar PO`**, arrastra el archivo de la PO o el correo `.msg` y haz clic en **`⚡ Procesar y Extraer`**.
2. **Opción B (Carga Masiva Excel)**: Descarga la plantilla oficial en **`📁 Carga Masiva de Excel`**, llena tus partidas y súbela para registrar múltiples órdenes en lote.
3. **Opción C (Formulario Manual)**: Usa el formulario guiado para capturar la orden campo por campo si no cuentas con el documento digital.
""")
            
        with st.container(border=True):
            st.markdown("### 🔹 Paso 2: Verificación de Datos Extraídos")
            st.markdown("""
- Revisa que el **Folio**, **Proyecto**, **Solicitante**, **Precios** y **Cantidades** sean correctos en la tabla interactiva de vista previa.
- Si el cliente marcó alguna urgencia (como las cantidades en verde del correo), verifica que aparezca la etiqueta **`🔥 Urgente`** en las observaciones.
- Presiona **`🚀 Registrar y Guardar PO en el Sistema`**.
""")
            
        with st.container(border=True):
            st.markdown("### 🔹 Paso 3: Monitoreo en la Matriz de Control 360°")
            st.markdown("""
- Entra a **`🎯 Matriz de Control 360° (Producción + Remisión)`** para ver en una sola pantalla:
  - **🔵 Piezas Fabricadas en Planta**: Corte láser, doblez y liberado en `app_corte_doblez`.
  - **🟢 Piezas Remisionadas al Cliente**: Tarimas despachadas con número de remisión en `remisiones-de-materiales`.
  - **Saldos Pendientes**: Sabrás al instante si una pieza está pendiente por cortar o si ya está en almacén lista para remisionarse.
""")
            
        with st.container(border=True):
            st.markdown("### 🔹 Paso 4: Trazabilidad en Profundidad y Exportación")
            st.markdown("""
- En **`🔍 Ficha de Trazabilidad 360°`**, selecciona cualquier orden de compra para ver el desglose completo de sus partidas, historial de tarimas enviadas y condiciones comerciales.
- Puedes descargar cualquier reporte o matriz completa en formato **Excel** usando los botones de descarga ubicados en cada sección.
""")

