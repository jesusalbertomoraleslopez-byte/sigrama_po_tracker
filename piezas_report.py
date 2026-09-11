import streamlit as st
import pandas as pd
import io
import datetime
from pathlib import Path

from config import (
    SQLITE_DB_PATH,
    is_historical_completed,
    normalize_po,
    PRIMARY_COLOR,
    SECONDARY_COLOR
)
from db_manager import get_all_pos, get_all_partidas
from remisiones_sync import (
    load_remisiones_databases,
    normalize_sku,
    clean_pronest_piece_name
)
from corte_doblez_sync import load_corte_doblez_databases

@st.cache_data(ttl=120)
def get_cached_pieces_by_project_summary():
    """Genera el reporte consolidado de piezas enriquecido con avances de taller y remisiones de forma vectorizada."""
    import sqlite3
    
    conn = sqlite3.connect(SQLITE_DB_PATH)
    df_all = pd.read_sql_query('''
        SELECT 
            c.id_interno,
            c.po,
            c.proyecto,
            c.estatus_general,
            p.item_no,
            p.clave_sku,
            p.sku_cliente,
            p.descripcion_producto,
            p.cantidad_requerida,
            p.unidad,
            p.precio_unitario,
            p.precio_total,
            p.fecha_entrega
        FROM po_partidas p
        INNER JOIN po_cabecera c ON p.po = c.po
        WHERE c.estatus_general NOT IN ('Cancelada', 'Cancelado')
        ORDER BY c.id_interno, p.item_no
    ''', conn)
    conn.close()

    if df_all.empty:
        return pd.DataFrame()

    # Cargar bases de datos externas de remisiones y taller
    df_rem, df_det, df_tar = load_remisiones_databases()
    df_ord, df_pie, df_ava, df_tar_cd, df_nid = load_corte_doblez_databases()

    # Pre-procesar Detalle_Tarimas para búsquedas vectorizadas O(1)
    rem_totals = {}
    if not df_det.empty:
        df_det_c = df_det.copy()
        df_det_c['norm_po'] = df_det_c['PO'].apply(normalize_po)
        df_det_c['norm_sku'] = df_det_c['SKU'].apply(lambda s: normalize_sku(clean_pronest_piece_name(s)))
        rem_totals = df_det_c.groupby(['norm_po', 'norm_sku'])['Cantidad'].sum().to_dict()

    # Pre-procesar avances y piezas en taller (Corte y Doblez)
    ava_lib, ava_dob, ava_cor = {}, {}, {}
    ofs_by_po = {}
    if not df_ava.empty and 'no_pieza' in df_ava.columns and 'of_number' in df_ava.columns:
        df_ava_c = df_ava.copy()
        df_ava_c['_norm_pieza'] = df_ava_c['no_pieza'].apply(lambda x: normalize_sku(clean_pronest_piece_name(x)))
        
        # Mapear OF a PO normalizada
        of_to_po = {}
        if not df_ord.empty and 'of_number' in df_ord.columns:
            for _, o in df_ord.iterrows():
                of_k = str(o.get('of_number', '')).strip()
                po_k = normalize_po(str(o.get('po', '')).strip())
                if of_k and po_k:
                    of_to_po[of_k] = po_k
                    if po_k not in ofs_by_po:
                        ofs_by_po[po_k] = set()
                    ofs_by_po[po_k].add(of_k)

        df_ava_c['_norm_po'] = df_ava_c['of_number'].astype(str).str.strip().map(of_to_po)
        df_ava_c['_area_lc'] = df_ava_c['area'].astype(str).str.lower()
        
        ava_lib = df_ava_c[df_ava_c['_area_lc'].isin(['liberado', 'empaque'])].groupby(['_norm_po', '_norm_pieza'])['cantidad'].sum().to_dict()
        ava_dob = df_ava_c[df_ava_c['_area_lc'] == 'doblez'].groupby(['_norm_po', '_norm_pieza'])['cantidad'].sum().to_dict()
        ava_cor = df_ava_c[df_ava_c['_area_lc'] == 'corte'].groupby(['_norm_po', '_norm_pieza'])['cantidad'].sum().to_dict()

    rows_enriched = []
    for _, r in df_all.iterrows():
        id_int = str(r['id_interno']).strip()
        po = str(r['po']).strip()
        proy = str(r['proyecto']).strip() or 'PROYECTO SIGRAMA'
        sku = str(r['clave_sku']).strip()
        sku_c = str(r.get('sku_cliente', '') or '').strip()
        desc = str(r['descripcion_producto']).strip() or f"Material {sku}"
        c_req = float(r['cantidad_requerida'] or 0)

        if is_historical_completed(id_interno=id_int, po=po):
            c_fab = c_req
            c_rem = c_req
            c_pen = 0.0
            ofs = "Entrega Histórica Validada"
            rem_fol = "Entrega Histórica (100% Remisionada)"
        else:
            norm_p = normalize_po(po)
            n_sku = normalize_sku(clean_pronest_piece_name(sku))
            n_skuc = normalize_sku(clean_pronest_piece_name(sku_c)) if sku_c else ""

            # Remisiones
            c_rem = max(rem_totals.get((norm_p, n_sku), 0.0), rem_totals.get((norm_p, n_skuc), 0.0) if n_skuc else 0.0)

            # Avances en taller
            c_lib = max(ava_lib.get((norm_p, n_sku), 0.0), ava_lib.get((norm_p, n_skuc), 0.0) if n_skuc else 0.0)
            c_dob = max(ava_dob.get((norm_p, n_sku), 0.0), ava_dob.get((norm_p, n_skuc), 0.0) if n_skuc else 0.0)
            c_cor = max(ava_cor.get((norm_p, n_sku), 0.0), ava_cor.get((norm_p, n_skuc), 0.0) if n_skuc else 0.0)

            c_fab = max(c_lib, c_dob, c_cor, c_rem)
            c_pen = max(0.0, c_req - c_rem)

            of_list = sorted(list(ofs_by_po.get(norm_p, [])))
            ofs = f"OF {', '.join(of_list)}" if of_list else ("Fabricado (En Almacén)" if c_fab > 0 else "Por Programar OF")
            rem_fol = f"Remisionado ({c_rem:.0f} pzas)" if c_rem > 0 else "Pendiente de Salida"

        c_fab = min(c_req, c_fab)
        c_rem = min(c_req, c_rem)

        pct_f = (c_fab / c_req * 100.0) if c_req > 0 else 0.0
        pct_r = (c_rem / c_req * 100.0) if c_req > 0 else 0.0

        if c_rem >= c_req and c_req > 0:
            st_p = "🟢 Remisionada Total (100%)"
            st_cat = "Remisionada"
        elif c_rem > 0:
            st_p = f"🔵 Parcial Enviada ({pct_r:.0f}%)"
            st_cat = "Parcial"
        elif c_fab >= c_req and c_req > 0:
            st_p = "🟣 Fabricada 100% (En Planta)"
            st_cat = "Fabricada"
        elif c_fab > 0:
            st_p = f"🟠 En Fabricación ({pct_f:.0f}%)"
            st_cat = "En Proceso"
        else:
            st_p = "⚪ Por Fabricar"
            st_cat = "Por Fabricar"

        rows_enriched.append({
            'id_interno': id_int,
            'po': po,
            'proyecto': proy,
            'item_no': r['item_no'],
            'clave_sku': sku,
            'sku_cliente': sku_c if sku_c else sku,
            'descripcion_producto': desc,
            'cantidad_requerida': c_req,
            'piezas_fabricadas': c_fab,
            'piezas_remisionadas': c_rem,
            'piezas_pendientes': c_pen,
            'pct_fabricacion': round(pct_f, 1),
            'pct_remision': round(pct_r, 1),
            'estatus_pieza': st_p,
            'categoria_estatus': st_cat,
            'ofs_asociadas': ofs,
            'remisiones_folios': rem_fol,
            'fecha_entrega': r.get('fecha_entrega', '—')
        })

    return pd.DataFrame(rows_enriched)


def generate_pieces_excel(df_export, project_name="Todos_los_Proyectos", mode="Consolidado"):
    """Genera un archivo Excel formateado con estilo corporativo Sigrama listo para descarga."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_export.to_excel(writer, index=False, sheet_name="Piezas")
        ws = writer.sheets["Piezas"]
        
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        
        header_fill = PatternFill(start_color="111111", end_color="111111", fill_type="solid")
        header_font = Font(name="Montserrat", size=11, bold=True, color="FFFFFF")
        border_thin = Border(
            left=Side(style='thin', color='E2E8F0'),
            right=Side(style='thin', color='E2E8F0'),
            top=Side(style='thin', color='E2E8F0'),
            bottom=Side(style='thin', color='E2E8F0')
        )
        
        for col_num, col_name in enumerate(df_export.columns, 1):
            cell = ws.cell(row=1, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws.column_dimensions[cell.column_letter].width = max(len(str(col_name)) + 6, 15)

        for row in ws.iter_rows(min_row=2, max_row=len(df_export)+1, min_col=1, max_col=len(df_export.columns)):
            for cell in row:
                cell.border = border_thin
                cell.alignment = Alignment(vertical="center")

    output.seek(0)
    return output


def render_catalogo_piezas_por_proyecto():
    """Renderiza el módulo completo de exploración y extracción de piezas filtradas por Proyecto."""
    st.markdown("""
    <div style="background: #FFFFFF; border: 1px solid #CBD5E1; border-left: 6px solid #EC2024; border-radius: 10px; padding: 16px 22px; margin-bottom: 16px; box-shadow: 0 2px 8px rgba(0,0,0,0.05);">
        <h2 style="color: #0F172A !important; font-family: 'Montserrat', sans-serif; font-size: 21px; font-weight: 800; margin: 0 0 4px 0;">
            📦 Catálogo y Lista de Piezas por Proyecto
        </h2>
        <p style="color: #475569 !important; font-size: 13px; margin: 0; font-family: 'Questrial', sans-serif;">
            Extrae y audita la lista completa de piezas (Nombre, Clave SKU, Cantidades requeridas, fabricadas y entregadas) filtradas por familia de Proyecto (ej. <b>LC8, RENO 4, CLOUD, META</b>).
        </p>
    </div>
    """, unsafe_allow_html=True)

    df_piezas = get_cached_pieces_by_project_summary()
    if df_piezas.empty:
        st.info("💡 No hay piezas registradas en el catálogo.")
        return

    # Proyectos disponibles y conteos
    proyectos_list = sorted(list(df_piezas['proyecto'].unique()))
    
    # 1. Filtros Principales
    f_c1, f_c2, f_c3 = st.columns([2.2, 1.4, 1.4])
    
    with f_c1:
        # Selector de Proyecto con conteo de piezas
        proy_options = ["🌟 Todos los Proyectos"] + [
            f"{p} ({len(df_piezas[df_piezas['proyecto'] == p]):,} pzas)" for p in proyectos_list
        ]
        sel_proy_raw = st.selectbox(
            "🏗️ Seleccionar Proyecto a Consultar:",
            proy_options,
            index=0,
            key="sb_proy_piezas_filter"
        )
        
    with f_c2:
        estatus_piezas_opts = [
            "🌟 Todos los Estatus",
            "🟢 Solo Remisionadas (100%)",
            "🟣 Solo Fabricadas (100%)",
            "🟠 Con Avance (> 0%)",
            "⏳ Con Piezas Pendientes (> 0)"
        ]
        sel_est_p = st.selectbox("🎯 Filtrar por Estado:", estatus_piezas_opts, key="sb_est_piezas_filter")

    with f_c3:
        txt_search = st.text_input("🔍 Buscar SKU o Nombre de Pieza:", placeholder="Ej. DOOR, PP10315, 11-A-6014...", key="txt_search_piezas")

    # Píldoras de Acceso Rápido para los Proyectos más Usados
    top_projects = ["🌟 Todos", "LC8", "RENO 4", "RENO 5", "RENO 6", "CLOUD", "META", "SOUTH VALLEY", "TORONTO"]
    pill_proy = st.pills(
        "Acceso rápido por familia:",
        options=top_projects,
        default="🌟 Todos",
        key="pills_proy_rapido"
    )

    # Determinar proyecto activo
    proy_activo = None
    if pill_proy and pill_proy != "🌟 Todos":
        proy_activo = pill_proy
    elif sel_proy_raw and not sel_proy_raw.startswith("🌟 Todos"):
        proy_activo = sel_proy_raw.split(" (")[0].strip()

    # Aplicar Filtros
    df_filtered = df_piezas.copy()
    if proy_activo:
        df_filtered = df_filtered[df_filtered['proyecto'] == proy_activo]

    if sel_est_p and not sel_est_p.startswith("🌟 Todos"):
        s_low = sel_est_p.lower()
        if "remisionada" in s_low:
            df_filtered = df_filtered[df_filtered['piezas_remisionadas'] >= df_filtered['cantidad_requerida']]
        elif "fabricada" in s_low:
            df_filtered = df_filtered[df_filtered['piezas_fabricadas'] >= df_filtered['cantidad_requerida']]
        elif "avance" in s_low:
            df_filtered = df_filtered[df_filtered['piezas_fabricadas'] > 0]
        elif "pendiente" in s_low:
            df_filtered = df_filtered[df_filtered['piezas_pendientes'] > 0]

    if txt_search:
        q = txt_search.strip().lower()
        df_filtered = df_filtered[
            df_filtered['clave_sku'].astype(str).str.lower().str.contains(q, na=False) |
            df_filtered['sku_cliente'].astype(str).str.lower().str.contains(q, na=False) |
            df_filtered['descripcion_producto'].astype(str).str.lower().str.contains(q, na=False) |
            df_filtered['po'].astype(str).str.lower().str.contains(q, na=False) |
            df_filtered['id_interno'].astype(str).str.lower().str.contains(q, na=False)
        ]

    # Tarjetas KPI del Filtro Activo
    k_skus = df_filtered['clave_sku'].nunique()
    k_req = float(df_filtered['cantidad_requerida'].sum())
    k_fab = float(df_filtered['piezas_fabricadas'].sum())
    k_rem = float(df_filtered['piezas_remisionadas'].sum())
    k_pen = float(df_filtered['piezas_pendientes'].sum())
    pct_cumpl = (k_rem / k_req * 100.0) if k_req > 0 else 0.0

    st.markdown("<div style='height: 6px;'></div>", unsafe_allow_html=True)
    m1, m2, m3, m4, m5, m6 = st.columns(6)
    with m1:
        st.metric("📦 SKUs Únicos", f"{k_skus:,}")
    with m2:
        st.metric("🎯 Total Requeridas", f"{k_req:,.0f} pz")
    with m3:
        st.metric("⚙️ Fabricadas", f"{k_fab:,.0f} pz")
    with m4:
        st.metric("🚚 Remisionadas", f"{k_rem:,.0f} pz")
    with m5:
        st.metric("⏳ Pendientes", f"{k_pen:,.0f} pz")
    with m6:
        st.metric("📊 % Cumplimiento", f"{pct_cumpl:.1f}%")

    st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)

    # Selector de Modo de Agrupación
    c_mode1, c_mode2 = st.columns([3.2, 1.8])
    with c_mode1:
        vista_modo = st.radio(
            "Modo de Visualización:",
            [
                "📊 Resumen Consolidado por SKU (Pieza Única Acumulada)",
                "📋 Desglose Detallado por Partida (Orden por Orden)"
            ],
            horizontal=True,
            key="radio_modo_vista_piezas"
        )

    # Preparar DataFrame según el modo
    if "Consolidado" in vista_modo:
        # Agrupar por SKU y Nombre
        df_display = df_filtered.groupby(['clave_sku', 'descripcion_producto']).agg(
            proyecto=('proyecto', lambda x: ', '.join(sorted(list(set(x))))),
            ordenes_asociadas=('po', lambda x: ', '.join(sorted(list(set(str(v) for v in x))))),
            total_ordenes=('po', 'nunique'),
            total_requerido=('cantidad_requerida', 'sum'),
            total_fabricado=('piezas_fabricadas', 'sum'),
            total_remisionado=('piezas_remisionadas', 'sum'),
            total_pendiente=('piezas_pendientes', 'sum')
        ).reset_index()

        df_display['pct_fabricacion'] = df_display.apply(
            lambda r: round((r['total_fabricado'] / r['total_requerido'] * 100.0) if r['total_requerido'] > 0 else 0.0, 1), axis=1
        )
        df_display['pct_remision'] = df_display.apply(
            lambda r: round((r['total_remisionado'] / r['total_requerido'] * 100.0) if r['total_requerido'] > 0 else 0.0, 1), axis=1
        )

        def _st_cons(r):
            if r['total_remisionado'] >= r['total_requerido'] and r['total_requerido'] > 0:
                return "🟢 100% Remisionado"
            elif r['total_remisionado'] > 0:
                return f"🔵 Parcial Rem. ({r['pct_remision']}%)"
            elif r['total_fabricado'] >= r['total_requerido'] and r['total_requerido'] > 0:
                return "🟣 100% Fabricado"
            elif r['total_fabricado'] > 0:
                return f"🟠 En Taller ({r['pct_fabricacion']}%)"
            return "⚪ Por Fabricar"

        df_display['estatus'] = df_display.apply(_st_cons, axis=1)
        df_display = df_display.sort_values('total_requerido', ascending=False)

        # Renombrar columnas para visualización impecable
        df_table = df_display.rename(columns={
            'clave_sku': 'SKU (Planta)',
            'descripcion_producto': 'Nombre / Descripción de la Pieza',
            'proyecto': 'Proyecto',
            'total_ordenes': 'POs',
            'ordenes_asociadas': 'Órdenes de Compra (POs)',
            'total_requerido': 'Requeridas',
            'total_fabricado': 'Fabricadas',
            'total_remisionado': 'Remisionadas',
            'total_pendiente': 'Pendientes',
            'pct_fabricacion': '% Fab.',
            'pct_remision': '% Rem.',
            'estatus': 'Estatus'
        })
        
        column_order = [
            'SKU (Planta)', 'Nombre / Descripción de la Pieza', 'Proyecto',
            'Requeridas', 'Fabricadas', 'Remisionadas', 'Pendientes',
            '% Fab.', '% Rem.', 'Estatus', 'POs', 'Órdenes de Compra (POs)'
        ]
        df_table = df_table[column_order]

    else:
        # Modo Detallado
        df_table = df_filtered[[
            'id_interno', 'po', 'proyecto', 'item_no',
            'clave_sku', 'descripcion_producto',
            'cantidad_requerida', 'piezas_fabricadas', 'piezas_remisionadas', 'piezas_pendientes',
            'pct_fabricacion', 'pct_remision', 'estatus_pieza', 'ofs_asociadas', 'remisiones_folios', 'fecha_entrega'
        ]].copy()
        
        df_table = df_table.rename(columns={
            'id_interno': 'ID Interno',
            'po': 'Folio PO',
            'proyecto': 'Proyecto',
            'item_no': 'Item',
            'clave_sku': 'SKU',
            'descripcion_producto': 'Nombre / Descripción de la Pieza',
            'cantidad_requerida': 'Requeridas',
            'piezas_fabricadas': 'Fabricadas',
            'piezas_remisionadas': 'Remisionadas',
            'piezas_pendientes': 'Pendientes',
            'pct_fabricacion': '% Fab.',
            'pct_remision': '% Rem.',
            'estatus_pieza': 'Estatus',
            'ofs_asociadas': 'Avance en Planta',
            'remisiones_folios': 'Folios de Remisión',
            'fecha_entrega': 'Fecha Entrega'
        })

    # Botones de Exportación en la esquina superior derecha
    with c_mode2:
        exp_col1, exp_col2 = st.columns(2)
        tag_file = (proy_activo or "Todos_los_Proyectos").replace(" ", "_")
        
        with exp_col1:
            excel_data = generate_pieces_excel(df_table, project_name=tag_file)
            st.download_button(
                label="📥 Bajar Excel",
                data=excel_data,
                file_name=f"Reporte_Piezas_{tag_file}_{datetime.date.today().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )
        with exp_col2:
            csv_data = df_table.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📄 Bajar CSV",
                data=csv_data,
                file_name=f"Reporte_Piezas_{tag_file}_{datetime.date.today().strftime('%Y%m%d')}.csv",
                mime="text/csv",
                use_container_width=True
            )

    # Renderizar la Tabla Interactiva con Streamlit Dataframe
    st.markdown(f"""
    <div style="font-family: 'Montserrat', sans-serif; font-size: 13px; font-weight: 700; color: #1E293B; margin-bottom: 6px;">
        Mostrando <b style="color: #EC2024;">{len(df_table):,}</b> registros para: <u>{proy_activo or 'Todos los Proyectos'}</u>
    </div>
    """, unsafe_allow_html=True)

    column_config = {
        "SKU (Planta)": st.column_config.TextColumn("SKU (Planta)", width="medium"),
        "SKU": st.column_config.TextColumn("SKU", width="medium"),
        "Nombre / Descripción de la Pieza": st.column_config.TextColumn("Nombre / Descripción de la Pieza", width="large"),
        "Requeridas": st.column_config.NumberColumn("Requeridas", format="%d pz"),
        "Fabricadas": st.column_config.NumberColumn("Fabricadas", format="%d pz"),
        "Remisionadas": st.column_config.NumberColumn("Remisionadas", format="%d pz"),
        "Pendientes": st.column_config.NumberColumn("Pendientes", format="%d pz"),
        "% Fab.": st.column_config.ProgressColumn("% Fab.", format="%.1f%%", min_value=0.0, max_value=100.0),
        "% Rem.": st.column_config.ProgressColumn("% Rem.", format="%.1f%%", min_value=0.0, max_value=100.0),
        "Estatus": st.column_config.TextColumn("Estatus", width="medium"),
        "Órdenes de Compra (POs)": st.column_config.TextColumn("Órdenes de Compra (POs)", width="large"),
        "Avance en Planta": st.column_config.TextColumn("Avance en Planta", width="medium"),
        "Folios de Remisión": st.column_config.TextColumn("Folios de Remisión", width="medium")
    }

    st.dataframe(
        df_table,
        use_container_width=True,
        hide_index=True,
        height=540,
        column_config=column_config
    )
