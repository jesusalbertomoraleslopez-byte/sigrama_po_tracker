import sys
sys.stdout.reconfigure(encoding='utf-8')
import io
import datetime
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import DataBarRule
from openpyxl.utils import get_column_letter

from db_manager import get_all_pos, get_all_partidas
from remisiones_sync import get_global_pos_tracking_summary

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
            
            row_p_vals = [
                (p_po,     "center", "@", Font(name="Calibri", size=9.5, bold=True, color="EC2024")),
                (p_it,     "center", "#,##0", None),
                (p_sk_cli, "left",   "@", Font(name="Calibri", size=9, bold=True, color="0F172A")),
                (p_sk_pla, "left",   "@", Font(name="Calibri", size=9, bold=True, color="1D4ED8")),
                (p_desc,   "left",   "@", Font(name="Calibri", size=9, color="334155")),
                (p_cant,   "right",  '#,##0 "pzas"', Font(name="Calibri", size=9.5, bold=True, color="0F172A")),
                (p_unid,   "center", "@", None),
                (p_pu,     "right",  '"$"#,##0.00', None),
                (p_ptot,   "right",  '"$"#,##0.00', Font(name="Calibri", size=9.5, bold=True, color="0F172A")),
                (p_fent,   "center", "yyyy-mm-dd", None),
                (p_parc,   "left",   "@", Font(name="Calibri", size=8.5, color="64748B")),
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
            ws2.merge_cells(f"A{p_tot_r}:E{p_tot_r}")
            ws2[f"A{p_tot_r}"].value = "TOTALES DE PARTIDAS"
            ws2[f"A{p_tot_r}"].font = Font(name="Calibri", size=9.5, bold=True, color=C_SLATE_DARK)
            ws2[f"A{p_tot_r}"].alignment = Alignment(horizontal="center", vertical="center")
            
            for c_i in range(1, 6):
                ws2.cell(row=p_tot_r, column=c_i).fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
                ws2.cell(row=p_tot_r, column=c_i).border = border_total
                
            c_cant_tot = ws2.cell(row=p_tot_r, column=6)
            c_cant_tot.value = f"=SUM(F{p_start}:F{p_end})"
            c_cant_tot.number_format = '#,##0 "pzas"'
            c_cant_tot.font = Font(name="Calibri", size=9.5, bold=True, color="0F172A")
            c_cant_tot.fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
            c_cant_tot.alignment = Alignment(horizontal="right", vertical="center")
            c_cant_tot.border = border_total
            
            for c_i in range(7, 9):
                c_bl = ws2.cell(row=p_tot_r, column=c_i)
                c_bl.fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
                c_bl.border = border_total
                
            c_monto_tot = ws2.cell(row=p_tot_r, column=9)
            c_monto_tot.value = f"=SUM(I{p_start}:I{p_end})"
            c_monto_tot.number_format = '"$"#,##0.00'
            c_monto_tot.font = Font(name="Calibri", size=9.5, bold=True, color="0F172A")
            c_monto_tot.fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
            c_monto_tot.alignment = Alignment(horizontal="right", vertical="center")
            c_monto_tot.border = border_total
            
            for c_i in range(10, 12):
                c_bl = ws2.cell(row=p_tot_r, column=c_i)
                c_bl.fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
                c_bl.border = border_total
                
            ws2.auto_filter.ref = f"A2:K{p_end}"
            
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
    2. Hoja 'Avance_Detalle_Piezas': Matriz de avance estación por estación por número de parte (Cortado,
       Doblado, Entarimado, Remisionadas, Pendiente, Estatus 360°) con formato de DataBars condicionales en celdas
       y fila de fórmulas contables totales.
    3. Hoja 'Lista_de_Piezas_Precios': Despiece comercial con precios unitarios, importes y fechas de entrega.
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
    # HOJA 2: DETALLE DE PIEZAS Y AVANCES POR ESTACIÓN (Corte, Doblez, Tarimas, Remisión)
    # =========================================================================
    ws2 = wb.create_sheet(title="Avance_Detalle_Piezas")
    ws2.views.sheetView[0].showGridLines = True

    ws2.merge_cells("A1:P1")
    c2_t = ws2["A1"]
    c2_t.value = "INDUSTRIA SIGRAMA S.A. DE C.V.  —  AVANCE Y TRAZABILIDAD POR PIEZA (DESPIECE 360°)"
    c2_t.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    c2_t.fill = PatternFill(start_color=C_SLATE_DARK, end_color=C_SLATE_DARK, fill_type="solid")
    c2_t.alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 28

    ws2.merge_cells("A2:P2")
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
        ("Descripción del Producto", 38, "left", C_SLATE_DARK),
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
            c_req  = float(r_p.get('cantidad_requerida', 0) or 0)
            c_cort = float(r_p.get('cortado', r_p.get('piezas_cortadas', 0)) or 0)
            c_dobl = float(r_p.get('doblado', r_p.get('piezas_dobladas', 0)) or 0)
            c_ent  = float(r_p.get('entarimado', r_p.get('cantidad_entarimada', 0)) or 0)
            c_rem  = float(r_p.get('cantidad_remisionada', 0) or 0)
            c_pend = float(r_p.get('cantidad_pendiente', max(0.0, c_req - c_rem)) or 0)

            pct_cort = (c_cort / c_req) if c_req > 0 else 0.0
            pct_dobl = (c_dobl / c_req) if c_req > 0 else 0.0
            pct_ent  = (c_ent / c_req) if c_req > 0 else 0.0
            pct_rem  = (c_rem / c_req) if c_req > 0 else 0.0
            pct_pend = (c_pend / c_req) if c_req > 0 else 0.0

            st_part = str(r_p.get('estatus_partida_360', '')).strip()
            if not st_part:
                if c_rem >= c_req and c_req > 0: st_part = "🟢 Remisionado Total"
                elif c_rem > 0: st_part = "🔵 Remisionado Parcial"
                elif c_ent >= c_req and c_req > 0: st_part = "📦 Entarimado PT"
                elif c_dobl >= c_req and c_req > 0: st_part = "🟣 Doblado Completo"
                elif c_cort >= c_req and c_req > 0: st_part = "🔵 Cortado Completo"
                else: st_part = "⚪ En Espera"

            p_vals = [
                (i_no,     "center", "#,##0", Font(name="Calibri", size=9.5, bold=True, color="475569"), fill_zebra),
                (sk_c,     "left",   "@",     Font(name="Calibri", size=9.5, bold=True, color="0F172A"), fill_zebra),
                (sk_p,     "left",   "@",     Font(name="Calibri", size=9.5, bold=True, color="2563EB"), fill_zebra),
                (desc,     "left",   "@",     Font(name="Calibri", size=9, color="334155"), fill_zebra),
                (c_req,    "right",  '#,##0', Font(name="Calibri", size=9.5, bold=True, color="0F172A"), PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")),
                (c_cort,   "right",  '#,##0', Font(name="Calibri", size=9.5, color="1E3A8A"), fill_zebra),
                (pct_cort, "right",  "0.0%",  Font(name="Calibri", size=9, bold=True, color="1D4ED8"), fill_zebra),
                (c_dobl,   "right",  '#,##0', Font(name="Calibri", size=9.5, color="312E81"), fill_zebra),
                (pct_dobl, "right",  "0.0%",  Font(name="Calibri", size=9, bold=True, color="4338CA"), fill_zebra),
                (c_ent,    "right",  '#,##0', Font(name="Calibri", size=9.5, color="78350F"), fill_zebra),
                (pct_ent,  "right",  "0.0%",  Font(name="Calibri", size=9, bold=True, color="B45309"), fill_zebra),
                (c_rem,    "right",  '#,##0', Font(name="Calibri", size=9.5, color="064E3B"), fill_zebra),
                (pct_rem,  "right",  "0.0%",  Font(name="Calibri", size=9, bold=True, color="15803D"), fill_zebra),
                (c_pend,   "right",  '#,##0', Font(name="Calibri", size=9.5, color="7C2D12"), fill_zebra),
                (pct_pend, "right",  "0.0%",  Font(name="Calibri", size=9, bold=True, color="B91C1C"), fill_zebra),
                (st_part,  "center", "@",     None, None),
            ]
            for col_i, (val, al, nf, fnt, fll) in enumerate(p_vals, start=1):
                c_cell = ws2.cell(row=curr_p_row, column=col_i)
                c_cell.value = val
                c_cell.alignment = Alignment(horizontal=al, vertical="center")
                c_cell.number_format = nf
                c_cell.border = border_data
                if fnt: c_cell.font = fnt
                if fll: c_cell.fill = fll

            # Formato de celda para Estatus en columna 16
            c_st_part = ws2.cell(row=curr_p_row, column=16)
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
        ws2.conditional_formatting.add(f"G{start_p_row}:G{end_p_row}", rule_cort)

        rule_dobl = DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1.0, color="818CF8", showValue=None)
        ws2.conditional_formatting.add(f"I{start_p_row}:I{end_p_row}", rule_dobl)

        rule_ent = DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1.0, color="F59E0B", showValue=None)
        ws2.conditional_formatting.add(f"K{start_p_row}:K{end_p_row}", rule_ent)

        rule_rem = DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1.0, color="70AD47", showValue=None)
        ws2.conditional_formatting.add(f"M{start_p_row}:M{end_p_row}", rule_rem)

        rule_pen = DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1.0, color="F87171", showValue=None)
        ws2.conditional_formatting.add(f"O{start_p_row}:O{end_p_row}", rule_pen)

        # Fila de Totales Generales
        tot_p_row = end_p_row + 1
        ws2.row_dimensions[tot_p_row].height = 24
        ws2.merge_cells(f"A{tot_p_row}:D{tot_p_row}")
        c_tot_lbl = ws2[f"A{tot_p_row}"]
        c_tot_lbl.value = "TOTALES GENERALES CONSOLIDADOS"
        c_tot_lbl.font = Font(name="Calibri", size=10, bold=True, color=C_SLATE_DARK)
        c_tot_lbl.alignment = Alignment(horizontal="center", vertical="center")
        for col_k in range(1, 5):
            ws2.cell(row=tot_p_row, column=col_k).fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
            ws2.cell(row=tot_p_row, column=col_k).border = border_total

        tot_p_cols = [
            (5,  f"=SUM(E{start_p_row}:E{end_p_row})", '#,##0 "pzas"', "0F172A", "F1F5F9"),
            (6,  f"=SUM(F{start_p_row}:F{end_p_row})", '#,##0 "pzas"', "1D4ED8", "EFF6FF"),
            (7,  f"=F{tot_p_row}/E{tot_p_row}",        "0.0%",          "1D4ED8", "EFF6FF"),
            (8,  f"=SUM(H{start_p_row}:H{end_p_row})", '#,##0 "pzas"', "4338CA", "EEF2FF"),
            (9,  f"=H{tot_p_row}/E{tot_p_row}",        "0.0%",          "4338CA", "EEF2FF"),
            (10, f"=SUM(J{start_p_row}:J{end_p_row})", '#,##0 "pzas"', "B45309", "FEF3C7"),
            (11, f"=J{tot_p_row}/E{tot_p_row}",        "0.0%",          "B45309", "FEF3C7"),
            (12, f"=SUM(L{start_p_row}:L{end_p_row})", '#,##0 "pzas"', "15803D", "DCFCE7"),
            (13, f"=L{tot_p_row}/E{tot_p_row}",        "0.0%",          "15803D", "DCFCE7"),
            (14, f"=SUM(N{start_p_row}:N{end_p_row})", '#,##0 "pzas"', "B91C1C", "FEE2E2"),
            (15, f"=N{tot_p_row}/E{tot_p_row}",        "0.0%",          "B91C1C", "FEE2E2"),
            (16, "",                                   "@",             "0F172A", "F1F5F9"),
        ]
        for col_i, form, nf, fg, bg in tot_p_cols:
            c = ws2.cell(row=tot_p_row, column=col_i)
            if form: c.value = form
            c.number_format = nf
            c.font = Font(name="Calibri", size=10, bold=True, color=fg)
            c.fill = PatternFill(start_color=bg, end_color=bg, fill_type="solid")
            c.alignment = Alignment(horizontal="right" if nf != "@" else "center", vertical="center")
            c.border = border_total

        ws2.auto_filter.ref = f"A4:P{end_p_row}"
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
