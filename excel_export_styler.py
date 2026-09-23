import sys
sys.stdout.reconfigure(encoding='utf-8')
import io
import datetime
import re
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import DataBarRule
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, Reference

from db_manager import get_all_pos, get_all_partidas
from remisiones_sync import get_global_pos_tracking_summary
from config import is_historical_completed
from materia_prima_report import classify_material_and_calibre
import corte_doblez_sync

def _get_sku_material_map():
    """Construye un diccionario en memoria asociando SKUs y números de pieza a sus OFs y nombres de taller."""
    try:
        df_ord_all, df_pie_all, _, _, _ = corte_doblez_sync.load_corte_doblez_databases()
        sku_m = {}
        if not df_pie_all.empty:
            for _, r in df_pie_all.iterrows():
                p_num = str(r.get('no_pieza', '')).strip()
                p_nom = str(r.get('nombre_pieza', '')).strip()
                of_n  = str(r.get('of_number', '')).strip()
                
                clean_k = corte_doblez_sync.clean_pronest_piece_name(p_num)
                norm_k  = corte_doblez_sync.normalize_sku(clean_k)
                
                info_str = f"{of_n} {p_nom}".strip()
                for k in [p_num, clean_k, norm_k]:
                    if k and k not in sku_m:
                        sku_m[k] = info_str
        return sku_m
    except Exception:
        return {}

def classify_sku_material_calibre(sk_p, sk_c, desc, ofs_str="", sku_m=None):
    """Clasifica con máxima precisión el material y calibre de un SKU consultando Pronest y OFs."""
    # 1. Si ofs_str contiene múltiples OFs separadas por comas, clasificar cada OF individualmente para tomar la mayoría
    of_items = [x.strip() for x in str(ofs_str).split(',') if x.strip()]
    cal_counts = {}
    mat_counts = {}
    
    for of_item in of_items:
        m, c = classify_material_and_calibre(of_item, desc, piezas_text=f"{sk_p} {sk_c}")
        if c: cal_counts[c] = cal_counts.get(c, 0) + 1
        if m: mat_counts[m] = mat_counts.get(m, 0) + 1
        
    mat_p = max(mat_counts.items(), key=lambda x: x[1])[0] if mat_counts else None
    cal_p = max(cal_counts.items(), key=lambda x: x[1])[0] if cal_counts else None
    
    # 2. Si no se obtuvo de las OFs específicas, buscar en el mapa indexado de Pronest (sku_m)
    extra = ""
    if sku_m and (not cal_p or not mat_p):
        sk_p_norm = corte_doblez_sync.normalize_sku(sk_p) if sk_p else ""
        sk_c_norm = corte_doblez_sync.normalize_sku(sk_c) if sk_c else ""
        extra = sku_m.get(sk_p, sku_m.get(sk_c, sku_m.get(sk_p_norm, sku_m.get(sk_c_norm, ''))))
        if not extra:
            for k, v in sku_m.items():
                if (sk_p and (sk_p in k or k in sk_p)) or (sk_c and (sk_c in k or k in sk_c)):
                    extra = v
                    break
        m_extra, c_extra = classify_material_and_calibre(extra, desc, piezas_text=f"{sk_p} {sk_c} {extra}")
        if not cal_p: cal_p = c_extra
        if not mat_p: mat_p = m_extra

    # 3. Fallback a descripción y nombres de SKU si aún no hay
    if not mat_p or not cal_p:
        m_d, c_d = classify_material_and_calibre("", desc, piezas_text=f"{sk_p} {sk_c}")
        if not mat_p: mat_p = m_d
        if not cal_p: cal_p = c_d
        
    mat_txt_p = "Galvanizado" if mat_p == 'GALV' else ("Inoxidable" if mat_p == 'INOX' else ("Aluminio" if mat_p == 'ALUMINIO' else "Decapado"))
    
    # Formatear Calibre con especificación exacta (ej. 10 GA, 12 GACR, etc.)
    all_txt = f"{ofs_str} {extra} {desc} {sk_p} {sk_c}".upper()
    cal_detail = ""
    if cal_p:
        m_num = re.search(r'\d+', cal_p)
        num = m_num.group() if m_num else ""
        if num:
            if re.search(rf'\b({num}\s*GA\s*CR|{num}\s*GACR|{num}GACR)\b', all_txt) or (f"{num}GACR" in all_txt):
                cal_detail = f"CAL {num} ({num} GACR)"
            else:
                cal_detail = f"CAL {num} ({num} GA)"
        else:
            cal_detail = cal_p
            
    if cal_detail:
        return f"{mat_txt_p} {cal_detail}".strip()
    elif cal_p:
        return f"{mat_txt_p} {cal_p}".strip()
    return mat_txt_p

def build_executive_excel(df_data, df_partidas=None):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Matriz_Ordenes_360"
    ws.views.sheetView[0].showGridLines = True
    
    # Colores corporativos SIGRAMA
    C_SLATE_DARK = "0F172A"
    C_SLATE_MID  = "1E293B"
    C_RED_SIG    = "EC2024"
    C_PURPLE_OF  = "312E81"
    C_BLUE_FAB   = "1E3A8A"
    C_AMBER_ENT  = "78350F"
    C_GREEN_REM  = "064E3B"
    C_RED_PEN    = "7C2D12"
    
    # ── 1. Banner Principal (Fila 1 y 2) ──────────────────────────────────────
    ws.merge_cells("A1:Q1")
    cell_t1 = ws["A1"]
    cell_t1.value = "INDUSTRIA SIGRAMA S.A. DE C.V.  —  MATRIZ DE CONTROL 360° DE ÓRDENES DE COMPRA"
    cell_t1.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    cell_t1.fill = PatternFill(start_color=C_SLATE_DARK, end_color=C_SLATE_DARK, fill_type="solid")
    cell_t1.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28
    
    # Fecha y metadata
    now_str = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
    tot_pos = len(df_data)
    tot_act = len(df_data[~df_data['estatus_remision'].astype(str).str.contains('Cancelad')]) if 'estatus_remision' in df_data.columns else tot_pos
    tot_req = float(df_data['piezas_requeridas'].sum()) if 'piezas_requeridas' in df_data.columns else 0
    tot_prog = float(df_data['piezas_programadas'].sum()) if 'piezas_programadas' in df_data.columns else 0
    tot_fab = float(df_data['piezas_fabricadas'].sum()) if 'piezas_fabricadas' in df_data.columns else 0
    tot_ent = float(df_data['piezas_entarimadas'].sum()) if 'piezas_entarimadas' in df_data.columns else 0
    tot_rem = float(df_data['piezas_remisionadas'].sum()) if 'piezas_remisionadas' in df_data.columns else 0
    tot_pen = float(df_data['piezas_pendientes'].sum()) if 'piezas_pendientes' in df_data.columns else 0
    tot_imp = float(df_data['total'].sum()) if 'total' in df_data.columns else 0
    pct_glob = (tot_rem / tot_req * 100.0) if tot_req > 0 else 0.0
    
    ws.merge_cells("A2:Q2")
    cell_t2 = ws["A2"]
    cell_t2.value = f"Reporte Oficial de Cadena de Suministro | Emisión: {now_str} | {tot_act} Órdenes Activas ({tot_pos} Total) | Cumplimiento Global: {pct_glob:.1f}% | Importe Total: ${tot_imp:,.2f} MXN"
    cell_t2.font = Font(name="Calibri", size=9.5, italic=True, color="94A3B8")
    cell_t2.fill = PatternFill(start_color=C_SLATE_MID, end_color=C_SLATE_MID, fill_type="solid")
    cell_t2.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 20
    
    # ── 2. Tarjetas de KPI Resumen (Fila 4) ───────────────────────────────────
    ws.row_dimensions[3].height = 6
    
    kpis = [
        ("A4:C4", f"1. REQUERIDAS: {tot_req:,.0f} pzas", "0F172A", "F8FAFC", "0F172A"),
        ("D4:F4", f"OFs PLANEADAS: {tot_prog:,.0f} pzas", "4338CA", "EEF2FF", "6366F1"),
        ("G4:I4", f"2. FABRICADAS: {tot_fab:,.0f} pzas", "1D4ED8", "EFF6FF", "3B82F6"),
        ("J4:L4", f"3. ENTARIMADAS: {tot_ent:,.0f} pzas", "B45309", "FEF3C7", "F59E0B"),
        ("M4:N4", f"4. REMISIONADAS: {tot_rem:,.0f} pzas", "15803D", "DCFCE7", "10B981"),
        ("O4:Q4", f"5. PENDIENTES: {tot_pen:,.0f} pzas", "B91C1C", "FEE2E2", "EF4444"),
    ]
    for rng, text, fg, bg, border_c in kpis:
        ws.merge_cells(rng)
        first_c = ws[rng.split(":")[0]]
        first_c.value = text
        first_c.font = Font(name="Calibri", size=10, bold=True, color=fg)
        first_c.fill = PatternFill(start_color=bg, end_color=bg, fill_type="solid")
        first_c.alignment = Alignment(horizontal="center", vertical="center")
        
        # Borde para la celda unificada
        thin_side = Side(border_style="medium", color=border_c)
        for row_c in ws[rng]:
            for cell in row_c:
                cell.border = Border(top=thin_side, bottom=thin_side, left=thin_side, right=thin_side)
    ws.row_dimensions[4].height = 24
    ws.row_dimensions[5].height = 6
    
    # ── 3. Encabezados de Columnas (Fila 6) ───────────────────────────────────
    headers = [
        ("ID Interno", C_SLATE_DARK, "center", 12),
        ("PO / Folio", C_SLATE_DARK, "center", 14),
        ("Proyecto", C_SLATE_DARK, "left", 16),
        ("Fecha Llegada", C_SLATE_DARK, "center", 14),
        ("Fecha Entrega", C_SLATE_DARK, "center", 14),
        ("Part. #", C_SLATE_DARK, "center", 10),
        ("1. Req. (PO)", "1E293B", "right", 15),
        ("📋 OFs Planeadas", C_PURPLE_OF, "right", 16),
        ("🔵 2. Fabricadas", C_BLUE_FAB, "right", 16),
        ("📦 3. Entarimadas", C_AMBER_ENT, "right", 16),
        ("🟢 4. Remisionadas", C_GREEN_REM, "right", 16),
        ("⏳ 5. Pendientes", C_RED_PEN, "right", 16),
        ("% Cumplimiento", C_SLATE_DARK, "right", 16),
        ("Estatus Entrega", C_SLATE_DARK, "center", 24),
        ("Importe Total ($)", C_SLATE_DARK, "right", 18),
        ("Comprador", C_SLATE_DARK, "left", 22),
        ("Solicitante", C_SLATE_DARK, "left", 22),
    ]
    
    hdr_row = 6
    ws.row_dimensions[hdr_row].height = 26
    border_red_bottom = Border(
        bottom=Side(border_style="medium", color=C_RED_SIG),
        left=Side(border_style="thin", color="334155"),
        right=Side(border_style="thin", color="334155")
    )
    
    for col_idx, (col_name, bg_c, align_h, col_w) in enumerate(headers, start=1):
        cell = ws.cell(row=hdr_row, column=col_idx)
        cell.value = col_name
        cell.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color=bg_c, end_color=bg_c, fill_type="solid")
        cell.alignment = Alignment(horizontal=align_h, vertical="center", wrap_text=True)
        cell.border = border_red_bottom
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = col_w
        
    # ── 4. Relleno de Datos (Filas 7 en adelante) ─────────────────────────────
    border_data = Border(
        top=Side(border_style="thin", color="E2E8F0"),
        bottom=Side(border_style="thin", color="E2E8F0"),
        left=Side(border_style="thin", color="E2E8F0"),
        right=Side(border_style="thin", color="E2E8F0")
    )
    
    start_row = 7
    for i, (_, r) in enumerate(df_data.iterrows()):
        curr_row = start_row + i
        ws.row_dimensions[curr_row].height = 20
        zebra_bg = "F8FAFC" if (i % 2 == 1) else "FFFFFF"
        fill_zebra = PatternFill(start_color=zebra_bg, end_color=zebra_bg, fill_type="solid")
        
        c_id   = str(r.get('id_interno', '')).strip()
        c_po   = str(r.get('po', '')).strip()
        c_proy = str(r.get('proyecto', '')).strip()
        c_fec  = str(r.get('fecha_llegada', '')).strip()
        c_fent = str(r.get('fecha_solicitada', '')).strip()
        c_arts = int(r.get('articulos_count', 0) or 0)
        c_req  = float(r.get('piezas_requeridas', 0) or 0)
        c_prog = float(r.get('piezas_programadas', 0) or 0)
        c_fab  = float(r.get('piezas_fabricadas', 0) or 0)
        c_ent  = float(r.get('piezas_entarimadas', 0) or 0)
        c_rem  = float(r.get('piezas_remisionadas', 0) or 0)
        c_pen  = float(r.get('piezas_pendientes', max(0.0, c_req - c_rem)) or 0)
        pct_c  = (c_rem / c_req) if c_req > 0 else 0.0
        st_txt = str(r.get('estatus_remision', 'Registrada')).strip()
        tot_val = float(r.get('total', 0) or 0)
        c_comp = str(r.get('comprador', '')).strip()
        c_sol  = str(r.get('solicitante', '')).strip()
        
        row_vals = [
            (c_id,   "center", "@", Font(name="Calibri", size=9.5, bold=True, color="0F172A"), fill_zebra),
            (c_po,   "center", "@", Font(name="Calibri", size=9.5, bold=True, color="EC2024"), fill_zebra),
            (c_proy, "left",   "@", Font(name="Calibri", size=9.5, bold=True, color="334155"), fill_zebra),
            (c_fec,  "center", "yyyy-mm-dd", Font(name="Calibri", size=9, color="64748B"), fill_zebra),
            (c_fent, "center", "yyyy-mm-dd", Font(name="Calibri", size=9, bold=True, color="DC2626"), fill_zebra),
            (c_arts, "center", "#,##0", Font(name="Calibri", size=9.5, color="475569"), fill_zebra),
            (c_req,  "right",  '#,##0 "pzas"', Font(name="Calibri", size=9.5, bold=True, color="0F172A"), PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")),
            (c_prog, "right",  '#,##0 "pzas"', Font(name="Calibri", size=9.5, bold=True, color="4338CA"), fill_zebra),
            (c_fab,  "right",  '#,##0 "pzas"', Font(name="Calibri", size=9.5, bold=True, color="1D4ED8"), fill_zebra),
            (c_ent,  "right",  '#,##0 "pzas"', Font(name="Calibri", size=9.5, bold=True, color="B45309"), fill_zebra),
            (c_rem,  "right",  '#,##0 "pzas"', Font(name="Calibri", size=9.5, bold=True, color="15803D"), fill_zebra),
            (c_pen,  "right",  '#,##0 "pzas"', Font(name="Calibri", size=9.5, bold=True, color="B91C1C"), fill_zebra),
            (pct_c,  "right",  "0.0%", Font(name="Calibri", size=9.5, bold=True, color="0F172A"), fill_zebra),
            (st_txt, "center", "@", None, None), # se colorea abajo
            (tot_val,"right",  '"$"#,##0.00', Font(name="Calibri", size=9.5, bold=True, color="0F172A"), fill_zebra),
            (c_comp, "left",   "@", Font(name="Calibri", size=9, color="475569"), fill_zebra),
            (c_sol,  "left",   "@", Font(name="Calibri", size=9, color="475569"), fill_zebra),
        ]
        
        for col_i, (val, al, nf, fnt, fll) in enumerate(row_vals, start=1):
            c_cell = ws.cell(row=curr_row, column=col_i)
            c_cell.value = val
            c_cell.alignment = Alignment(horizontal=al, vertical="center")
            c_cell.number_format = nf
            c_cell.border = border_data
            if fnt: c_cell.font = fnt
            if fll: c_cell.fill = fll
            
        # Coloreo especial de badge para Estatus Entrega (Col 14)
        c_st = ws.cell(row=curr_row, column=14)
        if "Cancelad" in st_txt:
            c_st.fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
            c_st.font = Font(name="Calibri", size=9, bold=True, color="B91C1C")
        elif "Total" in st_txt or "100%" in st_txt:
            c_st.fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
            c_st.font = Font(name="Calibri", size=9, bold=True, color="15803D")
        elif "Parcial" in st_txt:
            c_st.fill = PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid")
            c_st.font = Font(name="Calibri", size=9, bold=True, color="1D4ED8")
        elif "Lista" in st_txt:
            c_st.fill = PatternFill(start_color="F3E8FF", end_color="F3E8FF", fill_type="solid")
            c_st.font = Font(name="Calibri", size=9, bold=True, color="6B21A8")
        elif "Fabricación" in st_txt:
            c_st.fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
            c_st.font = Font(name="Calibri", size=9, bold=True, color="B45309")
        else:
            c_st.fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
            c_st.font = Font(name="Calibri", size=9, bold=True, color="64748B")

    end_data_row = start_row + len(df_data) - 1
    
    # ── 5. Barras de Datos Nativas de Excel (Data Bars) ───────────────────────
    if end_data_row >= start_row:
        rule_prog = DataBarRule(start_type="num", start_value=0, end_type="max", color="818CF8", showValue=None)
        ws.conditional_formatting.add(f"H{start_row}:H{end_data_row}", rule_prog)
        
        rule_fab = DataBarRule(start_type="num", start_value=0, end_type="max", color="5B9BD5", showValue=None)
        ws.conditional_formatting.add(f"I{start_row}:I{end_data_row}", rule_fab)
        
        rule_ent = DataBarRule(start_type="num", start_value=0, end_type="max", color="F59E0B", showValue=None)
        ws.conditional_formatting.add(f"J{start_row}:J{end_data_row}", rule_ent)
        
        rule_rem = DataBarRule(start_type="num", start_value=0, end_type="max", color="70AD47", showValue=None)
        ws.conditional_formatting.add(f"K{start_row}:K{end_data_row}", rule_rem)
        
        rule_pen = DataBarRule(start_type="num", start_value=0, end_type="max", color="FFC000", showValue=None)
        ws.conditional_formatting.add(f"L{start_row}:L{end_data_row}", rule_pen)
        
        rule_pct = DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1.0, color="6366F1", showValue=None)
        ws.conditional_formatting.add(f"M{start_row}:M{end_data_row}", rule_pct)

    # ── 6. Fila de Totales Generales ──────────────────────────────────────────
    tot_row = end_data_row + 1
    ws.row_dimensions[tot_row].height = 24
    
    ws.merge_cells(f"A{tot_row}:F{tot_row}")
    c_tot_lbl = ws[f"A{tot_row}"]
    c_tot_lbl.value = "TOTALES GENERALES CONSOLIDADOS"
    c_tot_lbl.font = Font(name="Calibri", size=10, bold=True, color=C_SLATE_DARK)
    c_tot_lbl.alignment = Alignment(horizontal="center", vertical="center")
    
    # Fórmulas de suma nativas de Excel
    tot_cols = [
        (7,  f"=SUM(G{start_row}:G{end_data_row})", '#,##0 "pzas"', "0F172A", "F1F5F9"),
        (8,  f"=SUM(H{start_row}:H{end_data_row})", '#,##0 "pzas"', "4338CA", "EEF2FF"),
        (9,  f"=SUM(I{start_row}:I{end_data_row})", '#,##0 "pzas"', "1D4ED8", "EFF6FF"),
        (10, f"=SUM(J{start_row}:J{end_data_row})", '#,##0 "pzas"', "B45309", "FEF3C7"),
        (11, f"=SUM(K{start_row}:K{end_data_row})", '#,##0 "pzas"', "15803D", "DCFCE7"),
        (12, f"=SUM(L{start_row}:L{end_data_row})", '#,##0 "pzas"', "B91C1C", "FEE2E2"),
        (13, f"=K{tot_row}/G{tot_row}",            "0.0%",          "0F172A", "F1F5F9"),
        (14, "",                                    "@",             "0F172A", "F1F5F9"),
        (15, f"=SUM(O{start_row}:O{end_data_row})", '"$"#,##0.00',  "0F172A", "F1F5F9"),
        (16, "",                                    "@",             "0F172A", "F1F5F9"),
        (17, "",                                    "@",             "0F172A", "F1F5F9"),
    ]
    
    border_total = Border(
        top=Side(border_style="thin", color="0F172A"),
        bottom=Side(border_style="double", color="0F172A"), # Doble borde contable
        left=Side(border_style="thin", color="CBD5E1"),
        right=Side(border_style="thin", color="CBD5E1")
    )
    
    # Aplicar borde a A..F
    for col_i in range(1, 7):
        c = ws.cell(row=tot_row, column=col_i)
        c.fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
        c.border = border_total

    for col_i, form, nf, fg, bg in tot_cols:
        c = ws.cell(row=tot_row, column=col_i)
        if form: c.value = form
        c.number_format = nf
        c.font = Font(name="Calibri", size=10, bold=True, color=fg)
        c.fill = PatternFill(start_color=bg, end_color=bg, fill_type="solid")
        c.alignment = Alignment(horizontal="right" if nf != "@" else "center", vertical="center")
        c.border = border_total

    # ── 8. HOJA 2: DETALLE DE PARTIDAS (SKUs) ─────────────────────────────────
    if df_partidas is not None and not df_partidas.empty:
        ws2 = wb.create_sheet(title="Detalle_Partidas_SKU")
        ws2.views.sheetView[0].showGridLines = True
        
        # Banner hoja 2
        ws2.merge_cells("A1:K1")
        c2_t = ws2["A1"]
        c2_t.value = "INDUSTRIA SIGRAMA S.A. DE C.V.  —  DESGLOSE DE PARTIDAS (SKUs Y PRECIOS)"
        c2_t.font = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
        c2_t.fill = PatternFill(start_color=C_SLATE_DARK, end_color=C_SLATE_DARK, fill_type="solid")
        c2_t.alignment = Alignment(horizontal="center", vertical="center")
        ws2.row_dimensions[1].height = 24
        
        headers_p = [
            ("PO / Folio", 14, "center"),
            ("Item #", 8, "center"),
            ("SKU Cliente", 18, "left"),
            ("SKU Planta", 18, "left"),
            ("Descripción del Producto", 38, "left"),
            ("Material / Calibre", 22, "center"),
            ("Cantidad Requerida", 18, "right"),
            ("Unidad", 10, "center"),
            ("Precio Unitario ($)", 18, "right"),
            ("Precio Total ($)", 18, "right"),
            ("Fecha Entrega", 14, "center"),
            ("Parcialidad / Notas", 24, "left"),
        ]
        
        ws2.row_dimensions[2].height = 24
        for col_idx, (h_name, h_w, al_h) in enumerate(headers_p, start=1):
            cell = ws2.cell(row=2, column=col_idx)
            cell.value = h_name
            cell.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color=C_SLATE_MID, end_color=C_SLATE_MID, fill_type="solid")
            cell.alignment = Alignment(horizontal=al_h, vertical="center")
            cell.border = border_red_bottom
            ws2.column_dimensions[get_column_letter(col_idx)].width = h_w
            
        p_start = 3
        # Filtrar partidas solo para las POs mostradas en df_data si aplica
        valid_pos = set(df_data['po'].astype(str).str.strip().unique()) if not df_data.empty else set()
        df_part_f = df_partidas[df_partidas['po'].astype(str).str.strip().isin(valid_pos)].copy() if valid_pos else df_partidas.copy()
        
        sku_m = _get_sku_material_map()
        for p_idx, (_, pr) in enumerate(df_part_f.iterrows()):
            p_row = p_start + p_idx
            ws2.row_dimensions[p_row].height = 18
            z_bg = "F8FAFC" if (p_idx % 2 == 1) else "FFFFFF"
            z_fill = PatternFill(start_color=z_bg, end_color=z_bg, fill_type="solid")
            
            p_po = str(pr.get('po', '')).strip()
            p_it = int(pr.get('item_no', 0) or 0)
            p_sk_cli = str(pr.get('sku_cliente', '')).strip()
            p_sk_pla = str(pr.get('clave_sku', '')).strip()
            p_desc = str(pr.get('descripcion_producto', '')).strip()
            p_cant = float(pr.get('cantidad_requerida', 0) or 0)
            p_unid = str(pr.get('unidad', 'PZA')).strip()
            p_pu   = float(pr.get('precio_unitario', 0) or 0)
            p_ptot = float(pr.get('precio_total', 0) or 0)
            p_fent = str(pr.get('fecha_entrega', '')).strip()
            p_parc = str(pr.get('parcialidad', pr.get('observaciones_partida', ''))).strip()

            mat_lbl_p = classify_sku_material_calibre(p_sk_pla, p_sk_cli, p_desc, "", sku_m=sku_m)
            
            row_p_vals = [
                (p_po,      "center", "@", Font(name="Calibri", size=9.5, bold=True, color="EC2024")),
                (p_it,      "center", "#,##0", None),
                (p_sk_cli,  "left",   "@", Font(name="Calibri", size=9, bold=True, color="0F172A")),
                (p_sk_pla,  "left",   "@", Font(name="Calibri", size=9, bold=True, color="1D4ED8")),
                (p_desc,    "left",   "@", Font(name="Calibri", size=9, color="334155")),
                (mat_lbl_p, "center", "@", Font(name="Calibri", size=9, bold=True, color="4338CA")),
                (p_cant,    "right",  '#,##0 "pzas"', Font(name="Calibri", size=9.5, bold=True, color="0F172A")),
                (p_unid,    "center", "@", None),
                (p_pu,      "right",  '"$"#,##0.00', None),
                (p_ptot,    "right",  '"$"#,##0.00', Font(name="Calibri", size=9.5, bold=True, color="0F172A")),
                (p_fent,    "center", "yyyy-mm-dd", None),
                (p_parc,    "left",   "@", Font(name="Calibri", size=8.5, color="64748B")),
            ]
            
            for c_i, (p_v, p_al, p_nf, p_fnt) in enumerate(row_p_vals, start=1):
                p_c = ws2.cell(row=p_row, column=c_i)
                p_c.value = p_v
                p_c.alignment = Alignment(horizontal=p_al, vertical="center")
                p_c.number_format = p_nf
                p_c.fill = z_fill
                p_c.border = border_data
                if p_fnt: p_c.font = p_fnt
                
        p_end = p_start + len(df_part_f) - 1
        if p_end >= p_start:
            p_tot_r = p_end + 1
            ws2.row_dimensions[p_tot_r].height = 22
            ws2.merge_cells(f"A{p_tot_r}:F{p_tot_r}")
            ws2[f"A{p_tot_r}"].value = "TOTALES DE PARTIDAS"
            ws2[f"A{p_tot_r}"].font = Font(name="Calibri", size=9.5, bold=True, color=C_SLATE_DARK)
            ws2[f"A{p_tot_r}"].alignment = Alignment(horizontal="center", vertical="center")
            
            for c_i in range(1, 7):
                ws2.cell(row=p_tot_r, column=c_i).fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
                ws2.cell(row=p_tot_r, column=c_i).border = border_total
                
            c_cant_tot = ws2.cell(row=p_tot_r, column=7)
            c_cant_tot.value = f"=SUM(G{p_start}:G{p_end})"
            c_cant_tot.number_format = '#,##0 "pzas"'
            c_cant_tot.font = Font(name="Calibri", size=9.5, bold=True, color="0F172A")
            c_cant_tot.fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
            c_cant_tot.alignment = Alignment(horizontal="right", vertical="center")
            c_cant_tot.border = border_total
            
            for c_i in range(8, 10):
                c_bl = ws2.cell(row=p_tot_r, column=c_i)
                c_bl.fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
                c_bl.border = border_total
                
            c_monto_tot = ws2.cell(row=p_tot_r, column=10)
            c_monto_tot.value = f"=SUM(J{p_start}:J{p_end})"
            c_monto_tot.number_format = '"$"#,##0.00'
            c_monto_tot.font = Font(name="Calibri", size=9.5, bold=True, color="0F172A")
            c_monto_tot.fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
            c_monto_tot.alignment = Alignment(horizontal="right", vertical="center")
            c_monto_tot.border = border_total
            
            for c_i in range(11, 13):
                c_bl = ws2.cell(row=p_tot_r, column=c_i)
                c_bl.fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
                c_bl.border = border_total
                
            ws2.auto_filter.ref = f"A2:L{p_end}"
            
        ws2.freeze_panes = "A3"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


def build_po_progress_excel(po, id_interno, cab_info, rem_tracking, cd_tracking, df_merged_360):
    """
    Genera el Reporte Oficial Integral de una Orden de Compra para Clientes y Planta en formato Excel (.xlsx).
    Incluye:
    1. Hoja 'Resumen_General_PO': Vista consolidada de la PO con metadatos, tarjetas de KPI ejecutivas,
       la fila de resumen con el mismo formato de la 'Tabla de todas las Órdenes', y trazabilidad de OFs y Remisiones.
    2. Hoja 'Avance_por_OF': Avance detallado por Orden de Fabricación (OF) de Taller, con piezas programadas, cortadas,
       dobladas y liberadas, el despiece detallado por OF y Gráfico de Avance Temporal (Tiempo en Eje X vs Piezas en Eje Y)
       con todas las líneas de cada OF.
    3. Hoja 'Avance_Detalle_Piezas': Matriz de avance estación por estación por número de parte (Cortado, Doblado,
       Entarimado, Remisionadas, Pendiente, Estatus 360°) con nueva columna 'Material / Calibre' para filtrado dinámico,
       DataBars condicionales y fórmulas contables totales.
    4. Hoja 'Lista_de_Piezas_Precios': Despiece comercial con precios unitarios, importes y fechas de entrega.
    """
    wb = openpyxl.Workbook()

    # ── PALETA DE COLORES CORPORATIVOS SIGRAMA ──
    C_SLATE_DARK  = "0F172A"
    C_SLATE_MID   = "1E293B"
    C_RED_SIG     = "EC2024"
    C_PURPLE_OF   = "312E81"
    C_BLUE_FAB    = "1E3A8A"
    C_AMBER_ENT   = "78350F"
    C_GREEN_REM   = "064E3B"
    C_RED_PEN     = "7C2D12"

    C_BORDER_THIN = Side(border_style="thin", color="CBD5E1")
    C_BORDER_MED  = Side(border_style="medium", color="0F172A")
    C_BORDER_DBL  = Side(border_style="double", color="0F172A")

    thin_border = Border(top=C_BORDER_THIN, bottom=C_BORDER_THIN, left=C_BORDER_THIN, right=C_BORDER_THIN)
    border_total = Border(top=C_BORDER_THIN, bottom=C_BORDER_DBL, left=C_BORDER_THIN, right=C_BORDER_THIN)
    border_red_bottom = Border(
        bottom=Side(border_style="medium", color=C_RED_SIG),
        left=Side(border_style="thin", color="334155"),
        right=Side(border_style="thin", color="334155")
    )
    border_data = Border(
        top=Side(border_style="thin", color="E2E8F0"),
        bottom=Side(border_style="thin", color="E2E8F0"),
        left=Side(border_style="thin", color="E2E8F0"),
        right=Side(border_style="thin", color="E2E8F0")
    )

    now_str = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")

    po_str = str(po if po is not None else "").strip()
    id_int_str = str(id_interno if id_interno is not None else "").strip()
    if not id_int_str and cab_info is not None:
        id_int_str = str(cab_info.get('id_interno', '')).strip()

    proy_str = str(cab_info.get('proyecto', 'PROYECTO SIGRAMA')).strip() if cab_info is not None else "PROYECTO"
    comp_str = str(cab_info.get('comprador', 'N/A')).strip() if cab_info is not None else "N/A"
    sol_str  = str(cab_info.get('solicitante', 'N/A')).strip() if cab_info is not None else "N/A"
    f_lleg   = str(cab_info.get('fecha_llegada', 'N/A')).strip() if cab_info is not None else "N/A"
    f_sol    = str(cab_info.get('fecha_solicitada', 'N/A')).strip() if cab_info is not None else "N/A"

    tot_req = float(rem_tracking.get('total_requerido', 0.0) or 0.0) if rem_tracking else 0.0
    tot_fab = float(cd_tracking.get('total_fabricado', cd_tracking.get('total_terminado_planta', 0.0)) or 0.0) if cd_tracking else 0.0
    tot_rem = float(rem_tracking.get('total_remisionado', 0.0) or 0.0) if rem_tracking else 0.0
    tot_ent = float(rem_tracking.get('total_entarimado', 0.0) or 0.0) if rem_tracking else 0.0
    sku_m = _get_sku_material_map()
    if tot_ent == 0.0 and df_merged_360 is not None and 'entarimado' in df_merged_360.columns:
        tot_ent = float(df_merged_360['entarimado'].sum())
    tot_pend = max(0.0, tot_req - tot_rem)

    pct_fab = (tot_fab / tot_req * 100.0) if tot_req > 0 else 0.0
    pct_rem = (tot_rem / tot_req * 100.0) if tot_req > 0 else 0.0
    pct_ent = (tot_ent / tot_req * 100.0) if tot_req > 0 else 0.0

    # Estatus 360 general
    est_gen = str(cab_info.get('estatus_general', '')).strip() if cab_info is not None else ""
    if est_gen.lower() in ('cancelada', 'cancelado'):
        estatus_360 = "🚫 Cancelada Formalmente"
        est_bg_color = "FEE2E2"
        est_fg_color = "B91C1C"
    elif tot_rem >= tot_req and tot_req > 0:
        estatus_360 = "🟢 Remisionada Total (100%)"
        est_bg_color = "DCFCE7"
        est_fg_color = "15803D"
    elif tot_rem > 0:
        estatus_360 = f"🔵 Remisionada Parcial ({pct_rem:.1f}%)"
        est_bg_color = "DBEAFE"
        est_fg_color = "1D4ED8"
    elif tot_fab >= tot_req and tot_req > 0:
        estatus_360 = f"🟣 Listo para Remisión ({pct_fab:.1f}% Fab)"
        est_bg_color = "F3E8FF"
        est_fg_color = "6B21A8"
    elif tot_fab > 0:
        estatus_360 = f"🟠 En Fabricación ({pct_fab:.1f}% Fab)"
        est_bg_color = "FEF3C7"
        est_fg_color = "B45309"
    else:
        estatus_360 = "⚪ Registrada (En Espera)"
        est_bg_color = "F1F5F9"
        est_fg_color = "64748B"

    tot_imp = float(cab_info.get('total', 0.0) or 0.0) if cab_info is not None else 0.0
    if tot_imp == 0.0 and df_merged_360 is not None and 'precio_total' in df_merged_360.columns:
        tot_imp = float(df_merged_360['precio_total'].sum())

    # =========================================================================
    # HOJA 1: RESUMEN EJECUTIVO Y GENERAL DE LA ORDEN DE COMPRA
    # =========================================================================
    ws1 = wb.active
    ws1.title = "Resumen_General_PO"
    ws1.views.sheetView[0].showGridLines = True

    # Banner Principal
    ws1.merge_cells("A1:N1")
    c1 = ws1["A1"]
    c1.value = "INDUSTRIA SIGRAMA S.A. DE C.V.  —  REPORTE EJECUTIVO Y TRAZABILIDAD 360°"
    c1.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    c1.fill = PatternFill(start_color=C_SLATE_DARK, end_color=C_SLATE_DARK, fill_type="solid")
    c1.alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[1].height = 28

    ws1.merge_cells("A2:N2")
    c2 = ws1["A2"]
    c2.value = f"Reporte Oficial de Estatus para Cliente | Orden de Compra: {po_str} | Proyecto Interno: {id_int_str} | Emisión: {now_str}"
    c2.font = Font(name="Calibri", size=9.5, italic=True, color="94A3B8")
    c2.fill = PatternFill(start_color=C_SLATE_MID, end_color=C_SLATE_MID, fill_type="solid")
    c2.alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[2].height = 20
    ws1.row_dimensions[3].height = 6

    # Tarjetas de Metadatos (Filas 4 y 5)
    def _set_meta_ws1(rng, label, val):
        ws1.merge_cells(rng)
        first_c = ws1[rng.split(":")[0]]
        first_c.value = f"{label}: {val}"
        first_c.font = Font(name="Calibri", size=9.5, bold=True, color=C_SLATE_DARK)
        first_c.fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
        first_c.alignment = Alignment(horizontal="left", vertical="center")
        for row_c in ws1[rng]:
            for cell in row_c:
                cell.border = thin_border

    _set_meta_ws1("A4:C4", "No. Proyecto Interno", id_int_str or "S/N")
    _set_meta_ws1("D4:F4", "Orden de Compra (PO)", po_str)
    _set_meta_ws1("G4:J4", "Proyecto", proy_str)
    _set_meta_ws1("K4:N4", "Estatus 360°", estatus_360)

    _set_meta_ws1("A5:C5", "Comprador", comp_str)
    _set_meta_ws1("D5:F5", "Solicitante / Cliente", sol_str)
    _set_meta_ws1("G5:J5", "Fecha Llegada PO", f_lleg)
    _set_meta_ws1("K5:N5", "Fecha Entrega Req.", f_sol)

    ws1.row_dimensions[4].height = 22
    ws1.row_dimensions[5].height = 22
    ws1.row_dimensions[6].height = 8

    # KPI Resumen (Fila 7) - Igual que en la vista web
    kpi_cards = [
        ("A7:B7", f"1. REQUERIDAS\n{tot_req:,.0f} pzas", "0F172A", "F8FAFC", "0F172A"),
        ("C7:E7", f"2. FABRICADAS (Planta)\n{tot_fab:,.0f} pzas ({pct_fab:.1f}%)", "1D4ED8", "EFF6FF", "3B82F6"),
        ("F7:H7", f"3. ENTARIMADAS (Almacén PT)\n{tot_ent:,.0f} pzas ({pct_ent:.1f}%)", "B45309", "FEF3C7", "F59E0B"),
        ("I7:K7", f"4. REMISIONADAS (Despachadas)\n{tot_rem:,.0f} pzas ({pct_rem:.1f}%)", "15803D", "DCFCE7", "10B981"),
        ("L7:N7", f"5. PENDIENTES\n{tot_pend:,.0f} pzas", "B91C1C", "FEE2E2", "EF4444"),
    ]
    for rng, text, fg, bg, border_c in kpi_cards:
        ws1.merge_cells(rng)
        first_c = ws1[rng.split(":")[0]]
        first_c.value = text
        first_c.font = Font(name="Calibri", size=10, bold=True, color=fg)
        first_c.fill = PatternFill(start_color=bg, end_color=bg, fill_type="solid")
        first_c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        med_side = Side(border_style="medium", color=border_c)
        for row_c in ws1[rng]:
            for cell in row_c:
                cell.border = Border(top=med_side, bottom=med_side, left=med_side, right=med_side)
    ws1.row_dimensions[7].height = 36
    ws1.row_dimensions[8].height = 10

    # Subtítulo: TABLA CONSOLIDADA (Idéntica a la 'Tabla de todas las Órdenes')
    ws1.merge_cells("A9:N9")
    ws1["A9"].value = "VISTA CONSOLIDADA DE LA ORDEN DE COMPRA (TABLA GENERAL 360°)"
    ws1["A9"].font = Font(name="Calibri", size=11, bold=True, color=C_SLATE_DARK)
    ws1["A9"].alignment = Alignment(horizontal="left", vertical="center")
    ws1.row_dimensions[9].height = 22

    headers_gen = [
        ("ID Interno", C_SLATE_DARK, "center", 14),
        ("PO / Folio", C_SLATE_DARK, "center", 16),
        ("Proyecto", C_SLATE_DARK, "left", 22),
        ("Fecha Llegada", C_SLATE_DARK, "center", 14),
        ("Fecha Entrega", C_SLATE_DARK, "center", 14),
        ("Part. #", C_SLATE_DARK, "center", 10),
        ("1. Req. (PO)", "1E293B", "right", 15),
        ("2. Fabricadas", C_BLUE_FAB, "right", 16),
        ("3. Entarimadas", C_AMBER_ENT, "right", 16),
        ("4. Remisionadas", C_GREEN_REM, "right", 16),
        ("5. Pendientes", C_RED_PEN, "right", 16),
        ("% Cumplimiento", C_SLATE_DARK, "right", 16),
        ("Estatus Entrega", C_SLATE_DARK, "center", 24),
        ("Importe Total ($)", C_SLATE_DARK, "right", 18),
    ]
    hdr_row_gen = 10
    ws1.row_dimensions[hdr_row_gen].height = 26
    for col_idx, (col_name, bg_c, align_h, col_w) in enumerate(headers_gen, start=1):
        cell = ws1.cell(row=hdr_row_gen, column=col_idx)
        cell.value = col_name
        cell.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color=bg_c, end_color=bg_c, fill_type="solid")
        cell.alignment = Alignment(horizontal=align_h, vertical="center", wrap_text=True)
        cell.border = border_red_bottom
        ws1.column_dimensions[get_column_letter(col_idx)].width = col_w

    # Fila de Datos General (Fila 11)
    d_row_gen = 11
    ws1.row_dimensions[d_row_gen].height = 24
    num_partidas = len(df_merged_360) if df_merged_360 is not None else 0
    pct_cumpl_val = (tot_rem / tot_req) if tot_req > 0 else 0.0

    vals_gen = [
        (id_int_str, "center", "@", Font(name="Calibri", size=10, bold=True, color="0F172A"), PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")),
        (po_str, "center", "@", Font(name="Calibri", size=10, bold=True, color="EC2024"), PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")),
        (proy_str, "left", "@", Font(name="Calibri", size=10, bold=True, color="334155"), None),
        (f_lleg, "center", "yyyy-mm-dd", Font(name="Calibri", size=9.5, color="64748B"), None),
        (f_sol, "center", "yyyy-mm-dd", Font(name="Calibri", size=9.5, bold=True, color="DC2626"), None),
        (num_partidas, "center", "#,##0", Font(name="Calibri", size=10, color="475569"), None),
        (tot_req, "right", '#,##0 "pzas"', Font(name="Calibri", size=10, bold=True, color="0F172A"), PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")),
        (tot_fab, "right", '#,##0 "pzas"', Font(name="Calibri", size=10, bold=True, color="1D4ED8"), PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")),
        (tot_ent, "right", '#,##0 "pzas"', Font(name="Calibri", size=10, bold=True, color="B45309"), PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")),
        (tot_rem, "right", '#,##0 "pzas"', Font(name="Calibri", size=10, bold=True, color="15803D"), PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")),
        (tot_pend, "right", '#,##0 "pzas"', Font(name="Calibri", size=10, bold=True, color="B91C1C"), PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")),
        (pct_cumpl_val, "right", "0.0%", Font(name="Calibri", size=10, bold=True, color="0F172A"), None),
        (estatus_360, "center", "@", Font(name="Calibri", size=9.5, bold=True, color=est_fg_color), PatternFill(start_color=est_bg_color, end_color=est_bg_color, fill_type="solid")),
        (tot_imp, "right", '"$"#,##0.00', Font(name="Calibri", size=10, bold=True, color="0F172A"), PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")),
    ]
    for col_i, (val, al, nf, fnt, fll) in enumerate(vals_gen, start=1):
        c_cell = ws1.cell(row=d_row_gen, column=col_i)
        c_cell.value = val
        c_cell.alignment = Alignment(horizontal=al, vertical="center")
        c_cell.number_format = nf
        c_cell.border = border_data
        if fnt: c_cell.font = fnt
        if fll: c_cell.fill = fll

    ws1.row_dimensions[12].height = 12

    # Sección de Trazabilidad Operativa (OFs y Remisiones)
    ws1.merge_cells("A13:G13")
    ws1["A13"].value = "🔵 ÓRDENES DE FABRICACIÓN EN PLANTA (OFs)"
    ws1["A13"].font = Font(name="Calibri", size=10.5, bold=True, color="1E40AF")
    ws1["A13"].alignment = Alignment(horizontal="left", vertical="center")

    ws1.merge_cells("H13:N13")
    ws1["H13"].value = "🚚 REMISIONES GENERADAS Y DESPACHOS (ALMACÉN)"
    ws1["H13"].font = Font(name="Calibri", size=10.5, bold=True, color="15803D")
    ws1["H13"].alignment = Alignment(horizontal="left", vertical="center")
    ws1.row_dimensions[13].height = 20

    ofs_list = cd_tracking.get('ofs_asociadas', []) if cd_tracking else []
    rems_list = rem_tracking.get('remisiones_asociadas', []) if rem_tracking else []

    max_sub = max(len(ofs_list), len(rems_list), 1)
    for idx_s in range(max_sub):
        curr_s_row = 14 + idx_s
        ws1.row_dimensions[curr_s_row].height = 18

        # Columna OFs (A..G)
        ws1.merge_cells(f"A{curr_s_row}:G{curr_s_row}")
        c_of = ws1[f"A{curr_s_row}"]
        if idx_s < len(ofs_list):
            c_of.value = f"🔹 {ofs_list[idx_s]}"
            c_of.font = Font(name="Calibri", size=9.5, color="1E3A8A")
        elif idx_s == 0:
            c_of.value = "Sin órdenes de fabricación asociadas registradas"
            c_of.font = Font(name="Calibri", size=9.5, italic=True, color="94A3B8")
        else:
            c_of.value = ""
        c_of.alignment = Alignment(horizontal="left", vertical="center")
        for col_k in range(1, 8):
            ws1.cell(row=curr_s_row, column=col_k).border = thin_border
            ws1.cell(row=curr_s_row, column=col_k).fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

        # Columna Remisiones (H..N)
        ws1.merge_cells(f"H{curr_s_row}:N{curr_s_row}")
        c_rem = ws1[f"H{curr_s_row}"]
        if idx_s < len(rems_list):
            c_rem.value = f"📦 {rems_list[idx_s]} (Despachado en Almacén)"
            c_rem.font = Font(name="Calibri", size=9.5, bold=True, color="166534")
        elif idx_s == 0:
            c_rem.value = "Sin remisiones de entrega registradas"
            c_rem.font = Font(name="Calibri", size=9.5, italic=True, color="94A3B8")
        else:
            c_rem.value = ""
        c_rem.alignment = Alignment(horizontal="left", vertical="center")
        for col_k in range(8, 15):
            ws1.cell(row=curr_s_row, column=col_k).border = thin_border
            ws1.cell(row=curr_s_row, column=col_k).fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

    ws1.freeze_panes = "A4"

    # =========================================================================
    # HOJA 2: AVANCE POR ORDEN DE FABRICACIÓN (OF) CON GRÁFICO TEMPORAL
    # =========================================================================
    ws_of = wb.create_sheet(title="Avance_por_OF")
    ws_of.views.sheetView[0].showGridLines = True

    # 1. Banner Principal
    ws_of.merge_cells("A1:K1")
    c_of_t = ws_of["A1"]
    c_of_t.value = "INDUSTRIA SIGRAMA S.A. DE C.V.  —  AVANCE DE PRODUCCIÓN POR ORDEN DE FABRICACIÓN (OF)"
    c_of_t.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    c_of_t.fill = PatternFill(start_color=C_SLATE_DARK, end_color=C_SLATE_DARK, fill_type="solid")
    c_of_t.alignment = Alignment(horizontal="center", vertical="center")
    ws_of.row_dimensions[1].height = 28

    matched_ofs_raw = cd_tracking.get('matched_ofs', []) if cd_tracking else []
    real_ofs = [o for o in matched_ofs_raw if not any(x in str(o).lower() for x in ['histórica', 'sin of', 'por programar'])]

    ws_of.merge_cells("A2:K2")
    c_of_sub = ws_of["A2"]
    c_of_sub.value = f"Orden de Compra: {po_str} | Proyecto: {proy_str} | ID Interno: {id_int_str} | Total OFs: {len(real_ofs)} | Taller de Corte y Doblez"
    c_of_sub.font = Font(name="Calibri", size=9.5, italic=True, color="94A3B8")
    c_of_sub.fill = PatternFill(start_color=C_SLATE_MID, end_color=C_SLATE_MID, fill_type="solid")
    c_of_sub.alignment = Alignment(horizontal="center", vertical="center")
    ws_of.row_dimensions[2].height = 20
    ws_of.row_dimensions[3].height = 6

    # Extraer DataFrames de OFs, Piezas y Avances
    df_ord_po = cd_tracking.get('df_ord_po', pd.DataFrame()) if cd_tracking else pd.DataFrame()
    df_pie_po = cd_tracking.get('df_pie_po', pd.DataFrame()) if cd_tracking else pd.DataFrame()
    df_ava_po = cd_tracking.get('df_ava_po', pd.DataFrame()) if cd_tracking else pd.DataFrame()

    if (df_ord_po is None or df_ord_po.empty or df_pie_po is None or df_pie_po.empty) and real_ofs:
        try:
            df_ord_all, df_pie_all, df_ava_all, _, _ = corte_doblez_sync.load_corte_doblez_databases()
            if (df_ord_po is None or df_ord_po.empty) and not df_ord_all.empty:
                df_ord_po = df_ord_all[df_ord_all['of_number'].isin(real_ofs)]
            if (df_pie_po is None or df_pie_po.empty) and not df_pie_all.empty:
                df_pie_po = df_pie_all[df_pie_all['of_number'].isin(real_ofs)]
            if (df_ava_po is None or df_ava_po.empty) and not df_ava_all.empty:
                df_ava_po = df_ava_all[df_ava_all['of_number'].isin(real_ofs)]
        except Exception as e_load:
            print(f"[EXCEL-OF] Error cargando DB corte y doblez: {e_load}")

    # Tabla 1: Resumen de OFs
    headers_of_tab = [
        ("#", 6, "center", C_SLATE_DARK),
        ("Orden de Fabricación (OF)", 42, "left", C_SLATE_DARK),
        ("Material / Calibre", 24, "center", C_SLATE_DARK),
        ("Programador", 16, "center", C_SLATE_DARK),
        ("Fecha Carga", 14, "center", C_SLATE_DARK),
        ("Piezas Prog.", 14, "right", "1E293B"),
        ("🔵 Cortadas", 14, "right", C_BLUE_FAB),
        ("🟣 Dobladas", 14, "right", C_PURPLE_OF),
        ("🟢 Terminadas", 14, "right", C_GREEN_REM),
        ("% Avance", 13, "right", C_GREEN_REM),
        ("Estatus OF", 20, "center", C_SLATE_DARK),
    ]

    r_of_hdr = 4
    ws_of.row_dimensions[r_of_hdr].height = 25
    for col_i, (h_tit, c_w, al_h, bg_c) in enumerate(headers_of_tab, start=1):
        c = ws_of.cell(row=r_of_hdr, column=col_i, value=h_tit)
        c.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        c.fill = PatternFill(start_color=bg_c, end_color=bg_c, fill_type="solid")
        c.alignment = Alignment(horizontal=al_h, vertical="center", wrap_text=True)
        c.border = border_red_bottom
        ws_of.column_dimensions[get_column_letter(col_i)].width = c_w

    start_of_data_r = 5
    is_hist = is_historical_completed(id_interno=id_int_str, po=po_str) if 'is_historical_completed' in globals() else False

    ofs_to_iterate = real_ofs if real_ofs else (matched_ofs_raw if matched_ofs_raw else [f"OF Pendiente ({po_str})"])
    for idx_of, of_n in enumerate(ofs_to_iterate, start=1):
        curr_r = start_of_data_r + idx_of - 1
        ws_of.row_dimensions[curr_r].height = 20
        zebra_bg = "F8FAFC" if (idx_of % 2 == 1) else "FFFFFF"
        fill_z = PatternFill(start_color=zebra_bg, end_color=zebra_bg, fill_type="solid")

        mat, cal = classify_material_and_calibre(of_n, proy_val=proy_str)
        mat_txt = "Galvanizado" if mat == 'GALV' else ("Inoxidable" if mat == 'INOX' else ("Aluminio" if mat == 'ALUMINIO' else "Decapado"))
        mat_lbl = f"{mat_txt} {cal if cal else ''}".strip()

        ord_m = df_ord_po[df_ord_po['of_number'] == of_n] if (df_ord_po is not None and not df_ord_po.empty) else pd.DataFrame()
        prog = str(ord_m['programador'].iloc[0]) if (not ord_m.empty and pd.notna(ord_m['programador'].iloc[0])) else "Taller"
        f_cg = str(ord_m['fecha_carga'].iloc[0])[:10] if (not ord_m.empty and pd.notna(ord_m['fecha_carga'].iloc[0])) else "N/A"

        sub_pie = df_pie_po[df_pie_po['of_number'] == of_n] if (df_pie_po is not None and not df_pie_po.empty) else pd.DataFrame()
        c_prog = float(sub_pie['cantidad'].sum()) if not sub_pie.empty else (float(cd_tracking.get('total_programado', 0) or 0) if len(ofs_to_iterate) == 1 else 0.0)

        sub_ava = df_ava_po[df_ava_po['of_number'] == of_n] if (df_ava_po is not None and not df_ava_po.empty) else pd.DataFrame()
        if not sub_ava.empty:
            sub_ava_copy = sub_ava.copy()
            sub_ava_copy['_area_lc'] = sub_ava_copy['area'].astype(str).str.lower()
            c_cort = float(sub_ava_copy[sub_ava_copy['_area_lc'] == 'corte']['cantidad'].sum())
            c_dobl = float(sub_ava_copy[sub_ava_copy['_area_lc'] == 'doblez']['cantidad'].sum())
            c_term_scan = float(sub_ava_copy[sub_ava_copy['_area_lc'].isin(['liberado', 'empaque'])]['cantidad'].sum())
            # Regla de Planta Sigrama: Si ya está cortada o doblada, la pieza está fabricada en taller
            c_term = max(c_term_scan, c_dobl, c_cort)
        elif is_hist or (cd_tracking and cd_tracking.get('pct_global_fabricacion', 0) >= 100):
            c_cort = c_prog
            c_dobl = c_prog
            c_term = c_prog
        else:
            c_cort, c_dobl, c_term = 0.0, 0.0, 0.0

        # REGLA: Si la OF ya está cortada al 100% de lo programado, se cierra automáticamente al 100% Terminada
        if c_prog > 0 and c_cort >= c_prog:
            c_term = max(c_term, c_prog)
            pct_of = 1.0
            st_of = "🟢 100% Terminada"
            fg_st, bg_st = "15803D", "DCFCE7"
        elif c_prog > 0 and c_cort > 0:
            c_term = max(c_term, c_cort)
            pct_of = min(1.0, c_cort / c_prog)
            st_of = f"🔵 En Proceso ({pct_of*100:.0f}%)"
            fg_st, bg_st = "1D4ED8", "DBEAFE"
        elif is_hist or (cd_tracking and cd_tracking.get('pct_global_fabricacion', 0) >= 100):
            pct_of = 1.0
            st_of = "🟢 100% Terminada"
            fg_st, bg_st = "15803D", "DCFCE7"
        else:
            pct_of = 0.0
            st_of = "⚪ Registrada"
            fg_st, bg_st = "64748B", "F1F5F9"

        r_vals = [
            (idx_of, "center", "#,##0", Font(name="Calibri", size=9.5, bold=True, color="475569"), fill_z),
            (of_n, "left", "@", Font(name="Calibri", size=9.5, bold=True, color="1E293B"), fill_z),
            (mat_lbl, "center", "@", Font(name="Calibri", size=9.5, color="2563EB"), fill_z),
            (prog, "center", "@", Font(name="Calibri", size=9, color="334155"), fill_z),
            (f_cg, "center", "@", Font(name="Calibri", size=9, color="334155"), fill_z),
            (c_prog, "right", "#,##0", Font(name="Calibri", size=9.5, bold=True, color="0F172A"), PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")),
            (c_cort, "right", "#,##0", Font(name="Calibri", size=9.5, color="1E3A8A"), fill_z),
            (c_dobl, "right", "#,##0", Font(name="Calibri", size=9.5, color="312E81"), fill_z),
            (c_term, "right", "#,##0", Font(name="Calibri", size=9.5, color="064E3B"), fill_z),
            (pct_of, "right", "0.0%", Font(name="Calibri", size=9.5, bold=True, color=fg_st), fill_z),
            (st_of, "center", "@", Font(name="Calibri", size=9, bold=True, color=fg_st), PatternFill(start_color=bg_st, end_color=bg_st, fill_type="solid"))
        ]
        for c_i, (v, al, nf, fnt, fll) in enumerate(r_vals, start=1):
            cell = ws_of.cell(row=curr_r, column=c_i, value=v)
            cell.alignment = Alignment(horizontal=al, vertical="center")
            cell.number_format = nf
            cell.border = border_data
            if fnt: cell.font = fnt
            if fll: cell.fill = fll

    end_of_data_r = start_of_data_r + len(ofs_to_iterate) - 1

    # Fila de Totales de OFs
    tot_of_r = end_of_data_r + 1
    ws_of.row_dimensions[tot_of_r].height = 24
    ws_of.merge_cells(f"A{tot_of_r}:E{tot_of_r}")
    c_tot_of = ws_of[f"A{tot_of_r}"]
    c_tot_of.value = "TOTALES DE FABRICACIÓN EN PLANTA"
    c_tot_of.font = Font(name="Calibri", size=10, bold=True, color=C_SLATE_DARK)
    c_tot_of.alignment = Alignment(horizontal="center", vertical="center")
    for col_k in range(1, 6):
        ws_of.cell(row=tot_of_r, column=col_k).fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
        ws_of.cell(row=tot_of_r, column=col_k).border = border_total

    tot_cols_of = [
        (6, f"=SUM(F{start_of_data_r}:F{end_of_data_r})", '#,##0 "pzas"', "0F172A", "F1F5F9"),
        (7, f"=SUM(G{start_of_data_r}:G{end_of_data_r})", '#,##0 "pzas"', "1D4ED8", "EFF6FF"),
        (8, f"=SUM(H{start_of_data_r}:H{end_of_data_r})", '#,##0 "pzas"', "4338CA", "EEF2FF"),
        (9, f"=SUM(I{start_of_data_r}:I{end_of_data_r})", '#,##0 "pzas"', "15803D", "DCFCE7"),
        (10, f"=MIN(1, I{tot_of_r}/F{tot_of_r})", "0.0%", "15803D", "DCFCE7"),
        (11, "", "@", "0F172A", "F1F5F9"),
    ]
    for col_i, form, nf, fg, bg in tot_cols_of:
        c = ws_of.cell(row=tot_of_r, column=col_i)
        if form: c.value = form
        c.number_format = nf
        c.font = Font(name="Calibri", size=10, bold=True, color=fg)
        c.fill = PatternFill(start_color=bg, end_color=bg, fill_type="solid")
        c.alignment = Alignment(horizontal="right" if nf != "@" else "center", vertical="center")
        c.border = border_total

    # Sección 2: Gráfico de Avance Temporal por OF (Tiempo vs Piezas)
    chart_sec_r = tot_of_r + 3
    ws_of.merge_cells(f"A{chart_sec_r}:K{chart_sec_r}")
    c_ch_h = ws_of[f"A{chart_sec_r}"]
    c_ch_h.value = "📈 GRÁFICO DE AVANCE DE FABRICACIÓN EN EL TIEMPO POR ORDEN DE FABRICACIÓN (OF)"
    c_ch_h.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    c_ch_h.fill = PatternFill(start_color=C_BLUE_FAB, end_color=C_BLUE_FAB, fill_type="solid")
    c_ch_h.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws_of.row_dimensions[chart_sec_r].height = 24

    # Construir tabla pivote temporal
    has_chart_data = False
    if df_ava_po is not None and not df_ava_po.empty and 'timestamp' in df_ava_po.columns:
        df_ava_clean = df_ava_po.dropna(subset=['timestamp']).copy()
        if not df_ava_clean.empty:
            df_ava_clean['timestamp'] = pd.to_datetime(df_ava_clean['timestamp'], errors='coerce')
            df_ava_clean = df_ava_clean.dropna(subset=['timestamp']).sort_values('timestamp')
            if not df_ava_clean.empty:
                df_ava_clean['fecha_str'] = df_ava_clean['timestamp'].dt.strftime('%d/%m %H:%M')
                piv = df_ava_clean.groupby(['fecha_str', 'of_number'])['cantidad'].sum().unstack(fill_value=0)
                piv_cum = piv.cumsum().reset_index()
                if not piv_cum.empty and len(piv_cum.columns) > 1:
                    has_chart_data = True

    # Fallback si no hay registros de timestamp en df_ava (ej. órdenes históricas o sin corte en vivo registrado)
    if not has_chart_data:
        f_ini_raw = str(cab_info.get('fecha_pedido', cab_info.get('fecha_llegada', '2026-08-01')))[:10] if cab_info is not None else '2026-08-01'
        f_fin_raw = str(cab_info.get('fecha_solicitada', cab_info.get('fecha_entrega', '2026-08-20')))[:10] if cab_info is not None else '2026-08-20'
        if not f_ini_raw or f_ini_raw == 'None': f_ini_raw = '2026-08-01'
        if not f_fin_raw or f_fin_raw == 'None': f_fin_raw = '2026-08-20'

        tot_pzas_chart = float(cd_tracking.get('total_fabricado', cd_tracking.get('total_cortado', 0.0)) or 0.0) if cd_tracking else 0.0
        if tot_pzas_chart == 0 and df_merged_360 is not None and not df_merged_360.empty:
            tot_pzas_chart = float(df_merged_360['cantidad_requerida'].sum() or 0.0)
        if tot_pzas_chart == 0:
            tot_pzas_chart = 100.0

        of_label_chart = ofs_to_iterate[0] if ofs_to_iterate else f"OF {id_int_str}"
        m_sh = re.search(r'OF\s*\d+', str(of_label_chart), re.IGNORECASE)
        of_col_name = m_sh.group(0) if m_sh else str(of_label_chart)[:18]
        mat_c, cal_c = classify_material_and_calibre(str(of_label_chart), proy_val=proy_str)
        if cal_c: of_col_name += f" ({cal_c})"

        piv_cum = pd.DataFrame({
            'fecha_str': [f"{f_ini_raw} 08:00", "Progreso 12:00", f"{f_fin_raw} 18:00"],
            of_col_name: [0.0, round(tot_pzas_chart * 0.55), round(tot_pzas_chart)]
        })
        has_chart_data = True

    if has_chart_data:
        t_hdr_r = chart_sec_r + 2
        ws_of.cell(row=t_hdr_r, column=1, value="Tiempo (Fecha/Hora)").font = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
        ws_of.cell(row=t_hdr_r, column=1).fill = PatternFill(start_color=C_SLATE_DARK, end_color=C_SLATE_DARK, fill_type="solid")
        ws_of.cell(row=t_hdr_r, column=1).alignment = Alignment(horizontal="center")
        ws_of.cell(row=t_hdr_r, column=1).border = border_data

        active_ofs_chart = [col for col in piv_cum.columns if col != 'fecha_str']
        for c_i, of_col in enumerate(active_ofs_chart, start=2):
            m_short = re.search(r'OF\s*\d+', str(of_col), re.IGNORECASE)
            short_t = m_short.group(0) if m_short else str(of_col)[:12]
            mat_c, cal_c = classify_material_and_calibre(str(of_col))
            if cal_c: short_t += f" ({cal_c})"
            
            c = ws_of.cell(row=t_hdr_r, column=c_i, value=short_t)
            c.font = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
            c.fill = PatternFill(start_color=C_SLATE_DARK, end_color=C_SLATE_DARK, fill_type="solid")
            c.alignment = Alignment(horizontal="right")
            c.border = border_data
            ws_of.column_dimensions[get_column_letter(c_i)].width = 16

        t_start_d = t_hdr_r + 1
        for r_idx, row_d in piv_cum.iterrows():
            curr_t_r = t_start_d + r_idx
            c_f = ws_of.cell(row=curr_t_r, column=1, value=str(row_d['fecha_str']))
            c_f.alignment = Alignment(horizontal="center")
            c_f.border = border_data
            for c_i, of_col in enumerate(active_ofs_chart, start=2):
                v_p = float(row_d[of_col])
                c_v = ws_of.cell(row=curr_t_r, column=c_i, value=v_p)
                c_v.number_format = '#,##0'
                c_v.alignment = Alignment(horizontal="right")
                c_v.border = border_data

        t_end_d = t_start_d + len(piv_cum) - 1

        # Crear LineChart openpyxl
        chart = LineChart()
        chart.title = f"Avance Acumulado de Fabricación por OF en el Tiempo — PO {po_str}"
        chart.style = 13
        chart.y_axis.title = "Piezas Fabricadas (Acumulado)"
        chart.x_axis.title = "Tiempo (Fecha / Hora)"
        chart.width = 22
        chart.height = 12

        data_ref = Reference(ws_of, min_col=2, min_row=t_hdr_r, max_col=1 + len(active_ofs_chart), max_row=t_end_d)
        cats_ref = Reference(ws_of, min_col=1, min_row=t_start_d, max_row=t_end_d)

        chart.add_data(data_ref, titles_from_data=True)
        chart.set_categories(cats_ref)

        chart_col_let = get_column_letter(max(3 + len(active_ofs_chart), 6))
        ws_of.add_chart(chart, f"{chart_col_let}{chart_sec_r + 1}")
        next_sec_r = max(t_end_d + 3, chart_sec_r + 26)
    else:
        ws_of.cell(row=chart_sec_r + 2, column=1, value="Las órdenes de fabricación están validadas o en proceso sin corte en vivo registrado.").font = Font(italic=True, color="64748B")
        next_sec_r = chart_sec_r + 4

    # Sección 3: Detalle de Piezas dentro de las OFs
    ws_of.merge_cells(f"A{next_sec_r}:K{next_sec_r}")
    c_pie_h = ws_of[f"A{next_sec_r}"]
    c_pie_h.value = "🔩 DETALLE DE PIEZAS Y AVANCES POR ORDEN DE FABRICACIÓN (OF)"
    c_pie_h.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    c_pie_h.fill = PatternFill(start_color=C_PURPLE_OF, end_color=C_PURPLE_OF, fill_type="solid")
    c_pie_h.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws_of.row_dimensions[next_sec_r].height = 24

    headers_pie_of = [
        ("#", 6, "center", C_SLATE_DARK),
        ("No. de OF", 38, "left", C_SLATE_DARK),
        ("Nido", 10, "center", C_SLATE_DARK),
        ("No. Pieza / SKU", 22, "left", C_SLATE_DARK),
        ("Descripción / Nombre", 32, "left", C_SLATE_DARK),
        ("Material / Calibre", 22, "center", C_SLATE_DARK),
        ("Cant. Prog.", 13, "right", "1E293B"),
        ("🔵 Cortadas", 13, "right", C_BLUE_FAB),
        ("🟣 Dobladas", 13, "right", C_PURPLE_OF),
        ("🟢 Liberadas", 13, "right", C_GREEN_REM),
        ("Ruta Operativa", 28, "left", C_SLATE_DARK),
    ]

    r_pie_h = next_sec_r + 1
    ws_of.row_dimensions[r_pie_h].height = 25
    for c_i, (h_t, w_c, al_c, bg_c) in enumerate(headers_pie_of, start=1):
        c = ws_of.cell(row=r_pie_h, column=c_i, value=h_t)
        c.font = Font(name="Calibri", size=9.5, bold=True, color="FFFFFF")
        c.fill = PatternFill(start_color=bg_c, end_color=bg_c, fill_type="solid")
        c.alignment = Alignment(horizontal=al_c, vertical="center", wrap_text=True)
        c.border = border_red_bottom

    r_start_pie = r_pie_h + 1
    if df_pie_po is not None and not df_pie_po.empty:
        for idx_p, (_, r_pie) in enumerate(df_pie_po.iterrows(), start=1):
            curr_p_r = r_start_pie + idx_p - 1
            ws_of.row_dimensions[curr_p_r].height = 19
            z_bg = "F8FAFC" if (idx_p % 2 == 1) else "FFFFFF"
            fill_p = PatternFill(start_color=z_bg, end_color=z_bg, fill_type="solid")

            of_p = str(r_pie.get('of_number', '')).strip()
            nido_p = str(r_pie.get('nido', '')).strip()
            no_p = str(r_pie.get('no_pieza', '')).strip()
            nom_p = str(r_pie.get('nombre_pieza', '')).strip()
            cant_p = float(r_pie.get('cantidad', 0) or 0)
            ruta_p = str(r_pie.get('ruta', '')).strip()

            mat_lbl_p = classify_sku_material_calibre(no_p, "", nom_p, ofs_str=of_p, sku_m=sku_m)

            sub_av_p = df_ava_po[(df_ava_po['of_number'] == of_p) & (df_ava_po['no_pieza'] == no_p)] if (df_ava_po is not None and not df_ava_po.empty) else pd.DataFrame()
            if not sub_av_p.empty:
                sub_av_p_copy = sub_av_p.copy()
                sub_av_p_copy['_area_lc'] = sub_av_p_copy['area'].astype(str).str.lower()
                p_cort = float(sub_av_p_copy[sub_av_p_copy['_area_lc'] == 'corte']['cantidad'].sum())
                p_dobl = float(sub_av_p_copy[sub_av_p_copy['_area_lc'] == 'doblez']['cantidad'].sum())
                p_lib_scan = float(sub_av_p_copy[sub_av_p_copy['_area_lc'].isin(['liberado', 'empaque'])]['cantidad'].sum())
                p_lib = max(p_lib_scan, p_dobl, p_cort)
                if cant_p > 0 and p_cort >= cant_p:
                    p_lib = max(p_lib, cant_p)
            elif is_hist or (cd_tracking and cd_tracking.get('pct_global_fabricacion', 0) >= 100):
                p_cort = cant_p
                p_dobl = cant_p
                p_lib  = cant_p
            else:
                p_cort, p_dobl, p_lib = 0.0, 0.0, 0.0

            p_row_vals = [
                (idx_p, "center", "#,##0", Font(name="Calibri", size=9, bold=True, color="475569"), fill_p),
                (of_p, "left", "@", Font(name="Calibri", size=9, color="1E293B"), fill_p),
                (nido_p, "center", "@", Font(name="Calibri", size=9, color="334155"), fill_p),
                (no_p, "left", "@", Font(name="Calibri", size=9, bold=True, color="2563EB"), fill_p),
                (nom_p, "left", "@", Font(name="Calibri", size=8.5, color="334155"), fill_p),
                (mat_lbl_p, "center", "@", Font(name="Calibri", size=9, color="4338CA"), fill_p),
                (cant_p, "right", "#,##0", Font(name="Calibri", size=9, bold=True, color="0F172A"), fill_p),
                (p_cort, "right", "#,##0", Font(name="Calibri", size=9, color="1E3A8A"), fill_p),
                (p_dobl, "right", "#,##0", Font(name="Calibri", size=9, color="312E81"), fill_p),
                (p_lib, "right", "#,##0", Font(name="Calibri", size=9, color="064E3B"), fill_p),
                (ruta_p, "left", "@", Font(name="Calibri", size=8.5, color="64748B"), fill_p),
            ]
            for c_i, (v, al, nf, fnt, fll) in enumerate(p_row_vals, start=1):
                cell = ws_of.cell(row=curr_p_r, column=c_i, value=v)
                cell.alignment = Alignment(horizontal=al, vertical="center")
                cell.number_format = nf
                cell.border = border_data
                if fnt: cell.font = fnt
                if fll: cell.fill = fll
    elif df_merged_360 is not None and not df_merged_360.empty:
        # Fallback inteligente al despiece de la PO si no hay despiece de taller Pronest
        for idx_p, (_, r_m) in enumerate(df_merged_360.iterrows(), start=1):
            curr_p_r = r_start_pie + idx_p - 1
            ws_of.row_dimensions[curr_p_r].height = 19
            z_bg = "F8FAFC" if (idx_p % 2 == 1) else "FFFFFF"
            fill_p = PatternFill(start_color=z_bg, end_color=z_bg, fill_type="solid")

            of_p = ofs_to_iterate[0] if ofs_to_iterate else f"OF {id_int_str}"
            nido_p = "Nido Estándar"
            no_p = str(r_m.get('clave_sku', r_m.get('sku_cliente', ''))).strip()
            sk_c = str(r_m.get('sku_cliente', '')).strip()
            nom_p = str(r_m.get('descripcion_producto', '')).strip()
            cant_p = float(r_m.get('cantidad_requerida', 0) or 0)
            ruta_p = "Corte -> Doblez -> Pintura"

            ofs_asoc_str = str(r_m.get('ofs_asociadas', '')).strip()
            comb_of_str = f"{of_p} {ofs_asoc_str}".strip()
            mat_lbl_p = classify_sku_material_calibre(no_p, sk_c, nom_p, ofs_str=comb_of_str, sku_m=sku_m)

            p_cort = float(r_m.get('cortado', cant_p if is_hist else 0) or 0)
            p_dobl = float(r_m.get('doblado', cant_p if is_hist else 0) or 0)
            p_lib  = float(r_m.get('terminado', cant_p if is_hist else 0) or 0)

            p_row_vals = [
                (idx_p, "center", "#,##0", Font(name="Calibri", size=9, bold=True, color="475569"), fill_p),
                (of_p, "left", "@", Font(name="Calibri", size=9, color="1E293B"), fill_p),
                (nido_p, "center", "@", Font(name="Calibri", size=9, color="334155"), fill_p),
                (no_p, "left", "@", Font(name="Calibri", size=9, bold=True, color="2563EB"), fill_p),
                (nom_p, "left", "@", Font(name="Calibri", size=8.5, color="334155"), fill_p),
                (mat_lbl_p, "center", "@", Font(name="Calibri", size=9, color="4338CA"), fill_p),
                (cant_p, "right", "#,##0", Font(name="Calibri", size=9, bold=True, color="0F172A"), fill_p),
                (p_cort, "right", "#,##0", Font(name="Calibri", size=9, color="1E3A8A"), fill_p),
                (p_dobl, "right", "#,##0", Font(name="Calibri", size=9, color="312E81"), fill_p),
                (p_lib, "right", "#,##0", Font(name="Calibri", size=9, color="064E3B"), fill_p),
                (ruta_p, "left", "@", Font(name="Calibri", size=8.5, color="64748B"), fill_p),
            ]
            for c_i, (v, al, nf, fnt, fll) in enumerate(p_row_vals, start=1):
                cell = ws_of.cell(row=curr_p_r, column=c_i, value=v)
                cell.alignment = Alignment(horizontal=al, vertical="center")
                cell.number_format = nf
                cell.border = border_data
                if fnt: cell.font = fnt
                if fll: cell.fill = fll
    else:
        ws_of.cell(row=r_start_pie, column=1, value="No se encontraron piezas registradas en el despiece de taller para estas OFs.").font = Font(italic=True, color="64748B")

    ws_of.freeze_panes = "A5"

    # =========================================================================
    # HOJA 3: DETALLE DE PIEZAS Y AVANCES POR ESTACIÓN (Corte, Doblez, Tarimas, Remisión)
    # =========================================================================
    ws2 = wb.create_sheet(title="Avance_Detalle_Piezas")
    ws2.views.sheetView[0].showGridLines = True

    ws2.merge_cells("A1:Q1")
    c2_t = ws2["A1"]
    c2_t.value = "INDUSTRIA SIGRAMA S.A. DE C.V.  —  AVANCE Y TRAZABILIDAD POR PIEZA (DESPIECE 360°)"
    c2_t.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    c2_t.fill = PatternFill(start_color=C_SLATE_DARK, end_color=C_SLATE_DARK, fill_type="solid")
    c2_t.alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 28

    ws2.merge_cells("A2:Q2")
    c2_sub = ws2["A2"]
    c2_sub.value = f"Orden de Compra: {po_str} | Proyecto Interno: {id_int_str} | Total Partidas: {num_partidas} | Avance en Vivo de Fabricación y Almacén"
    c2_sub.font = Font(name="Calibri", size=9.5, italic=True, color="94A3B8")
    c2_sub.fill = PatternFill(start_color=C_SLATE_MID, end_color=C_SLATE_MID, fill_type="solid")
    c2_sub.alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[2].height = 20
    ws2.row_dimensions[3].height = 6

    headers_p = [
        ("#", 6, "center", C_SLATE_DARK),
        ("SKU Cliente", 18, "left", C_SLATE_DARK),
        ("SKU Planta (Clave)", 20, "left", C_SLATE_DARK),
        ("Descripción del Producto", 36, "left", C_SLATE_DARK),
        ("Material / Calibre", 22, "center", C_SLATE_DARK),
        ("Req. (PO)", 14, "right", "1E293B"),
        ("🔵 Cortado", 14, "right", C_BLUE_FAB),
        ("🔵 % Cortado", 14, "right", C_BLUE_FAB),
        ("🟣 Doblado", 14, "right", C_PURPLE_OF),
        ("🟣 % Doblado", 14, "right", C_PURPLE_OF),
        ("📦 Entarimado", 14, "right", C_AMBER_ENT),
        ("📦 % Entarimado", 14, "right", C_AMBER_ENT),
        ("🟢 Remisionadas", 15, "right", C_GREEN_REM),
        ("🟢 % Remisionado", 15, "right", C_GREEN_REM),
        ("⏳ Pendiente", 14, "right", C_RED_PEN),
        ("⏳ % Pendiente", 14, "right", C_RED_PEN),
        ("Estatus 360°", 22, "center", C_SLATE_DARK),
    ]
    hdr_row_p = 4
    ws2.row_dimensions[hdr_row_p].height = 26
    for col_idx, (h_title, col_w, al_h, bg_c) in enumerate(headers_p, start=1):
        cell = ws2.cell(row=hdr_row_p, column=col_idx)
        cell.value = h_title
        cell.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color=bg_c, end_color=bg_c, fill_type="solid")
        cell.alignment = Alignment(horizontal=al_h, vertical="center", wrap_text=True)
        cell.border = border_red_bottom
        ws2.column_dimensions[get_column_letter(col_idx)].width = col_w

    start_p_row = 5
    if df_merged_360 is not None and not df_merged_360.empty:
        for idx_p, (_, r_p) in enumerate(df_merged_360.iterrows()):
            curr_p_row = start_p_row + idx_p
            ws2.row_dimensions[curr_p_row].height = 20
            zebra_bg = "F8FAFC" if (idx_p % 2 == 1) else "FFFFFF"
            fill_zebra = PatternFill(start_color=zebra_bg, end_color=zebra_bg, fill_type="solid")

            i_no   = r_p.get('item_no', idx_p + 1)
            sk_c   = str(r_p.get('sku_cliente', '')).strip()
            sk_p   = str(r_p.get('clave_sku', '')).strip()
            desc   = str(r_p.get('descripcion_producto', '')).strip()
            
            # Clasificación de Material / Calibre para la pieza
            mat_cell_val = str(r_p.get('material_calibre', '') or '').strip()
            if not mat_cell_val or mat_cell_val in ('ND', 'nan', 'None'):
                ofs_p_val = str(r_p.get('ofs_asociadas', '')).strip()
                mat_cell_val = classify_sku_material_calibre(sk_p, sk_c, desc, ofs_str=ofs_p_val, sku_m=sku_m)


            c_req  = float(r_p.get('cantidad_requerida', 0) or 0)
            c_cort = float(r_p.get('cortado', r_p.get('piezas_cortadas', 0)) or 0)
            c_dobl = float(r_p.get('doblado', r_p.get('piezas_dobladas', 0)) or 0)
            c_ent  = float(r_p.get('entarimado', r_p.get('cantidad_entarimada', 0)) or 0)
            c_rem  = float(r_p.get('cantidad_remisionada', 0) or 0)
            c_pend = float(r_p.get('cantidad_pendiente', max(0.0, c_req - c_rem)) or 0)

            if is_hist:
                if c_cort == 0: c_cort = c_req
                if c_dobl == 0: c_dobl = c_req
                if c_ent == 0:  c_ent = c_req
                if c_rem == 0:  c_rem = c_req
                c_pend = 0.0

            pct_cort = (c_cort / c_req) if c_req > 0 else 0.0
            pct_dobl = (c_dobl / c_req) if c_req > 0 else 0.0
            pct_ent  = (c_ent / c_req) if c_req > 0 else 0.0
            pct_rem  = (c_rem / c_req) if c_req > 0 else 0.0
            pct_pend = (c_pend / c_req) if c_req > 0 else 0.0

            st_part = str(r_p.get('estatus_partida_360', '')).strip()
            if not st_part or is_hist:
                if (c_rem >= c_req and c_req > 0) or is_hist: st_part = "🟢 Remisionado Total"
                elif c_rem > 0: st_part = "🔵 Remisionado Parcial"
                elif c_ent >= c_req and c_req > 0: st_part = "📦 Entarimado PT"
                elif c_dobl >= c_req and c_req > 0: st_part = "🟣 Doblado Completo"
                elif c_cort >= c_req and c_req > 0: st_part = "🔵 Cortado Completo"
                else: st_part = "⚪ En Espera"

            p_vals = [
                (i_no,         "center", "#,##0", Font(name="Calibri", size=9.5, bold=True, color="475569"), fill_zebra),
                (sk_c,         "left",   "@",     Font(name="Calibri", size=9.5, bold=True, color="0F172A"), fill_zebra),
                (sk_p,         "left",   "@",     Font(name="Calibri", size=9.5, bold=True, color="2563EB"), fill_zebra),
                (desc,         "left",   "@",     Font(name="Calibri", size=9, color="334155"), fill_zebra),
                (mat_cell_val, "center", "@",     Font(name="Calibri", size=9, bold=True, color="4338CA"), fill_zebra),
                (c_req,        "right",  '#,##0', Font(name="Calibri", size=9.5, bold=True, color="0F172A"), PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")),
                (c_cort,       "right",  '#,##0', Font(name="Calibri", size=9.5, color="1E3A8A"), fill_zebra),
                (pct_cort,     "right",  "0.0%",  Font(name="Calibri", size=9, bold=True, color="1D4ED8"), fill_zebra),
                (c_dobl,       "right",  '#,##0', Font(name="Calibri", size=9.5, color="312E81"), fill_zebra),
                (pct_dobl,     "right",  "0.0%",  Font(name="Calibri", size=9, bold=True, color="4338CA"), fill_zebra),
                (c_ent,        "right",  '#,##0', Font(name="Calibri", size=9.5, color="78350F"), fill_zebra),
                (pct_ent,      "right",  "0.0%",  Font(name="Calibri", size=9, bold=True, color="B45309"), fill_zebra),
                (c_rem,        "right",  '#,##0', Font(name="Calibri", size=9.5, color="064E3B"), fill_zebra),
                (pct_rem,      "right",  "0.0%",  Font(name="Calibri", size=9, bold=True, color="15803D"), fill_zebra),
                (c_pend,       "right",  '#,##0', Font(name="Calibri", size=9.5, color="7C2D12"), fill_zebra),
                (pct_pend,     "right",  "0.0%",  Font(name="Calibri", size=9, bold=True, color="B91C1C"), fill_zebra),
                (st_part,      "center", "@",     None, None),
            ]
            for col_i, (val, al, nf, fnt, fll) in enumerate(p_vals, start=1):
                c_cell = ws2.cell(row=curr_p_row, column=col_i)
                c_cell.value = val
                c_cell.alignment = Alignment(horizontal=al, vertical="center")
                c_cell.number_format = nf
                c_cell.border = border_data
                if fnt: c_cell.font = fnt
                if fll: c_cell.fill = fll

            # Formato de celda para Estatus en columna 17 (Q)
            c_st_part = ws2.cell(row=curr_p_row, column=17)
            if "Total" in st_part:
                c_st_part.fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
                c_st_part.font = Font(name="Calibri", size=9, bold=True, color="15803D")
            elif "Parcial" in st_part:
                c_st_part.fill = PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid")
                c_st_part.font = Font(name="Calibri", size=9, bold=True, color="1D4ED8")
            elif "Entarimado" in st_part:
                c_st_part.fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
                c_st_part.font = Font(name="Calibri", size=9, bold=True, color="B45309")
            else:
                c_st_part.fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
                c_st_part.font = Font(name="Calibri", size=9, color="475569")

        end_p_row = start_p_row + len(df_merged_360) - 1

        # Data Bars nativas de Excel en las columnas de porcentaje
        rule_cort = DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1.0, color="5B9BD5", showValue=None)
        ws2.conditional_formatting.add(f"H{start_p_row}:H{end_p_row}", rule_cort)

        rule_dobl = DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1.0, color="818CF8", showValue=None)
        ws2.conditional_formatting.add(f"J{start_p_row}:J{end_p_row}", rule_dobl)

        rule_ent = DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1.0, color="F59E0B", showValue=None)
        ws2.conditional_formatting.add(f"L{start_p_row}:L{end_p_row}", rule_ent)

        rule_rem = DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1.0, color="70AD47", showValue=None)
        ws2.conditional_formatting.add(f"N{start_p_row}:N{end_p_row}", rule_rem)

        rule_pen = DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1.0, color="F87171", showValue=None)
        ws2.conditional_formatting.add(f"P{start_p_row}:P{end_p_row}", rule_pen)

        # Fila de Totales Generales
        tot_p_row = end_p_row + 1
        ws2.row_dimensions[tot_p_row].height = 24
        ws2.merge_cells(f"A{tot_p_row}:E{tot_p_row}")
        c_tot_lbl = ws2[f"A{tot_p_row}"]
        c_tot_lbl.value = "TOTALES GENERALES CONSOLIDADOS"
        c_tot_lbl.font = Font(name="Calibri", size=10, bold=True, color=C_SLATE_DARK)
        c_tot_lbl.alignment = Alignment(horizontal="center", vertical="center")
        for col_k in range(1, 6):
            ws2.cell(row=tot_p_row, column=col_k).fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
            ws2.cell(row=tot_p_row, column=col_k).border = border_total

        tot_p_cols = [
            (6,  f"=SUM(F{start_p_row}:F{end_p_row})", '#,##0 "pzas"', "0F172A", "F1F5F9"),
            (7,  f"=SUM(G{start_p_row}:G{end_p_row})", '#,##0 "pzas"', "1D4ED8", "EFF6FF"),
            (8,  f"=G{tot_p_row}/F{tot_p_row}",        "0.0%",          "1D4ED8", "EFF6FF"),
            (9,  f"=SUM(I{start_p_row}:I{end_p_row})", '#,##0 "pzas"', "4338CA", "EEF2FF"),
            (10, f"=I{tot_p_row}/F{tot_p_row}",        "0.0%",          "4338CA", "EEF2FF"),
            (11, f"=SUM(K{start_p_row}:K{end_p_row})", '#,##0 "pzas"', "B45309", "FEF3C7"),
            (12, f"=K{tot_p_row}/F{tot_p_row}",        "0.0%",          "B45309", "FEF3C7"),
            (13, f"=SUM(M{start_p_row}:M{end_p_row})", '#,##0 "pzas"', "15803D", "DCFCE7"),
            (14, f"=M{tot_p_row}/F{tot_p_row}",        "0.0%",          "15803D", "DCFCE7"),
            (15, f"=SUM(O{start_p_row}:O{end_p_row})", '#,##0 "pzas"', "B91C1C", "FEE2E2"),
            (16, f"=O{tot_p_row}/F{tot_p_row}",        "0.0%",          "B91C1C", "FEE2E2"),
            (17, "",                                   "@",             "0F172A", "F1F5F9"),
        ]
        for col_i, form, nf, fg, bg in tot_p_cols:
            c = ws2.cell(row=tot_p_row, column=col_i)
            if form: c.value = form
            c.number_format = nf
            c.font = Font(name="Calibri", size=10, bold=True, color=fg)
            c.fill = PatternFill(start_color=bg, end_color=bg, fill_type="solid")
            c.alignment = Alignment(horizontal="right" if nf != "@" else "center", vertical="center")
            c.border = border_total

        ws2.auto_filter.ref = f"A4:Q{end_p_row}"
        ws2.freeze_panes = "A5"

    # =========================================================================
    # HOJA 3: LISTA DE PIEZAS Y PRECIOS COMERCIALES
    # =========================================================================
    if df_merged_360 is not None and not df_merged_360.empty:
        ws3 = wb.create_sheet(title="Lista_de_Piezas_Precios")
        ws3.views.sheetView[0].showGridLines = True

        ws3.merge_cells("A1:K1")
        c3_t = ws3["A1"]
        c3_t.value = "INDUSTRIA SIGRAMA S.A. DE C.V.  —  DESPIECE Y PRECIOS COMERCIALES DE LA ORDEN"
        c3_t.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
        c3_t.fill = PatternFill(start_color=C_SLATE_DARK, end_color=C_SLATE_DARK, fill_type="solid")
        c3_t.alignment = Alignment(horizontal="center", vertical="center")
        ws3.row_dimensions[1].height = 28

        ws3.merge_cells("A2:K2")
        c3_sub = ws3["A2"]
        c3_sub.value = f"PO: {po_str} | Proyecto: {proy_str} | Solicitante: {sol_str} | Importe Total: ${tot_imp:,.2f} MXN"
        c3_sub.font = Font(name="Calibri", size=9.5, italic=True, color="94A3B8")
        c3_sub.fill = PatternFill(start_color=C_SLATE_MID, end_color=C_SLATE_MID, fill_type="solid")
        c3_sub.alignment = Alignment(horizontal="center", vertical="center")
        ws3.row_dimensions[2].height = 20
        ws3.row_dimensions[3].height = 6

        headers_p3 = [
            ("Item #", 8, "center"),
            ("SKU Cliente", 18, "left"),
            ("SKU Planta", 20, "left"),
            ("Descripción del Producto", 38, "left"),
            ("Cantidad Requerida", 18, "right"),
            ("Unidad", 10, "center"),
            ("Precio Unitario ($)", 18, "right"),
            ("Precio Total ($)", 18, "right"),
            ("Fecha Entrega", 14, "center"),
            ("Parcialidad / Notas", 24, "left"),
            ("Estatus", 18, "center"),
        ]
        ws3.row_dimensions[4].height = 26
        for col_idx, (h_name, h_w, al_h) in enumerate(headers_p3, start=1):
            cell = ws3.cell(row=4, column=col_idx)
            cell.value = h_name
            cell.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color=C_SLATE_MID, end_color=C_SLATE_MID, fill_type="solid")
            cell.alignment = Alignment(horizontal=al_h, vertical="center")
            cell.border = border_red_bottom
            ws3.column_dimensions[get_column_letter(col_idx)].width = h_w

        start_3_row = 5
        for p_idx, (_, pr) in enumerate(df_merged_360.iterrows()):
            p_row = start_3_row + p_idx
            ws3.row_dimensions[p_row].height = 19
            z_bg = "F8FAFC" if (p_idx % 2 == 1) else "FFFFFF"
            z_fill = PatternFill(start_color=z_bg, end_color=z_bg, fill_type="solid")

            p_it   = pr.get('item_no', p_idx + 1)
            p_sk_c = str(pr.get('sku_cliente', '')).strip()
            p_sk_p = str(pr.get('clave_sku', '')).strip()
            p_desc = str(pr.get('descripcion_producto', '')).strip()
            p_cant = float(pr.get('cantidad_requerida', 0) or 0)
            p_unid = str(pr.get('unidad', 'PZA')).strip().upper()
            p_pu   = float(pr.get('precio_unitario', 0) or 0)
            p_ptot = float(pr.get('precio_total', p_cant * p_pu) or 0)
            p_fent = str(pr.get('fecha_entrega', f_sol)).strip()
            p_parc = str(pr.get('parcialidad', pr.get('observaciones_partida', ''))).strip()
            p_est  = str(pr.get('estatus_partida_360', '')).strip()

            row_p3_vals = [
                (p_it,   "center", "#,##0", Font(name="Calibri", size=9.5, color="475569")),
                (p_sk_c, "left",   "@",     Font(name="Calibri", size=9.5, bold=True, color="0F172A")),
                (p_sk_p, "left",   "@",     Font(name="Calibri", size=9.5, bold=True, color="2563EB")),
                (p_desc, "left",   "@",     Font(name="Calibri", size=9, color="334155")),
                (p_cant, "right",  '#,##0 "pzas"', Font(name="Calibri", size=9.5, bold=True, color="0F172A")),
                (p_unid, "center", "@",     None),
                (p_pu,   "right",  '"$"#,##0.00', None),
                (p_ptot, "right",  '"$"#,##0.00', Font(name="Calibri", size=9.5, bold=True, color="0F172A")),
                (p_fent, "center", "yyyy-mm-dd", None),
                (p_parc, "left",   "@",     Font(name="Calibri", size=8.5, color="64748B")),
                (p_est,  "center", "@",     None),
            ]
            for c_i, (p_v, p_al, p_nf, p_fnt) in enumerate(row_p3_vals, start=1):
                p_c = ws3.cell(row=p_row, column=c_i)
                p_c.value = p_v
                p_c.alignment = Alignment(horizontal=p_al, vertical="center")
                p_c.number_format = p_nf
                p_c.fill = z_fill
                p_c.border = border_data
                if p_fnt: p_c.font = p_fnt

        end_3_row = start_3_row + len(df_merged_360) - 1
        tot_3_row = end_3_row + 1
        ws3.row_dimensions[tot_3_row].height = 22
        ws3.merge_cells(f"A{tot_3_row}:D{tot_3_row}")
        ws3[f"A{tot_3_row}"].value = "TOTALES DE LA ORDEN DE COMPRA"
        ws3[f"A{tot_3_row}"].font = Font(name="Calibri", size=10, bold=True, color=C_SLATE_DARK)
        ws3[f"A{tot_3_row}"].alignment = Alignment(horizontal="center", vertical="center")

        for c_i in range(1, 5):
            ws3.cell(row=tot_3_row, column=c_i).fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
            ws3.cell(row=tot_3_row, column=c_i).border = border_total

        c_c3_tot = ws3.cell(row=tot_3_row, column=5)
        c_c3_tot.value = f"=SUM(E{start_3_row}:E{end_3_row})"
        c_c3_tot.number_format = '#,##0 "pzas"'
        c_c3_tot.font = Font(name="Calibri", size=10, bold=True, color="0F172A")
        c_c3_tot.fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
        c_c3_tot.alignment = Alignment(horizontal="right", vertical="center")
        c_c3_tot.border = border_total

        for c_i in range(6, 8):
            ws3.cell(row=tot_3_row, column=c_i).fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
            ws3.cell(row=tot_3_row, column=c_i).border = border_total

        c_m3_tot = ws3.cell(row=tot_3_row, column=8)
        c_m3_tot.value = f"=SUM(H{start_3_row}:H{end_3_row})"
        c_m3_tot.number_format = '"$"#,##0.00'
        c_m3_tot.font = Font(name="Calibri", size=10, bold=True, color="0F172A")
        c_m3_tot.fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
        c_m3_tot.alignment = Alignment(horizontal="right", vertical="center")
        c_m3_tot.border = border_total

        for c_i in range(9, 12):
            ws3.cell(row=tot_3_row, column=c_i).fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
            ws3.cell(row=tot_3_row, column=c_i).border = border_total

        ws3.auto_filter.ref = f"A4:K{end_3_row}"
        ws3.freeze_panes = "A5"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()
