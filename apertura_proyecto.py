# -*- coding: utf-8 -*-
import io
import datetime
import re
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import email.message
import email.utils
from pathlib import Path
import pandas as pd
import db_manager

def find_original_order_files(po, id_interno):
    """
    Localiza los archivos originales de la orden (correo .msg y PDF oficial)
    tanto en la base de datos po_archivos_adjuntos como en la carpeta data/correos/.
    """
    correos_dir = Path("data/correos")
    po_str = str(po).strip()
    po_nodash = po_str.replace('-', '')
    id_clean = str(id_interno).replace('INT-', '').strip()
    
    msg_bytes, msg_name = None, None
    pdf_bytes, pdf_name = None, None
    
    # 1. Cabecera DB
    try:
        df_cab, _ = db_manager.get_po_by_folio(po_str)
        if df_cab.empty:
            df_cab, _ = db_manager.get_po_by_folio(po_nodash)
        if not df_cab.empty:
            r = df_cab.iloc[0]
            if r.get('archivo_correo'):
                p = correos_dir / str(r['archivo_correo'])
                if p.exists():
                    msg_name = p.name
                    with open(p, 'rb') as f:
                        msg_bytes = f.read()
            if r.get('archivo_pdf'):
                p = correos_dir / str(r['archivo_pdf'])
                if p.exists():
                    pdf_name = p.name
                    with open(p, 'rb') as f:
                        pdf_bytes = f.read()
    except Exception as e:
        print(f"[WARN] find_original_order_files cabecera: {e}")

    # 2. po_archivos_adjuntos
    try:
        conn = db_manager.get_connection()
        cur = conn.cursor()
        cur.execute("""
            SELECT id, nombre_archivo, tipo FROM po_archivos_adjuntos
            WHERE po IN (?, ?) OR id_interno = ?
        """, (po_str, po_nodash, str(id_interno)))
        for f_id, fn, ft in cur.fetchall():
            if str(ft).lower() == 'msg' and not msg_bytes:
                _, _, b = db_manager.get_contenido_archivo_por_id(f_id)
                if b:
                    msg_name, msg_bytes = fn, b
            elif str(ft).lower() == 'pdf' and not pdf_bytes:
                _, _, b = db_manager.get_contenido_archivo_por_id(f_id)
                if b:
                    pdf_name, pdf_bytes = fn, b
        conn.close()
    except Exception as e:
        print(f"[WARN] find_original_order_files po_archivos_adjuntos: {e}")

    # 3. Disco data/correos/
    if correos_dir.exists():
        for f in sorted(correos_dir.iterdir()):
            fn = f.name
            low = fn.lower()
            m_po = (po_str in fn) or (po_nodash in fn)
            m_id = (f"INT {id_clean}" in fn) or (f"INT_{id_clean}" in fn) or (f"INT{id_clean}" in fn) or (f"INT 0{id_clean}" in fn)
            
            if low.endswith('.msg') and not msg_bytes and (m_po or m_id):
                msg_name = fn
                try:
                    with open(f, 'rb') as fp:
                        msg_bytes = fp.read()
                except Exception:
                    pass
            if low.endswith('.pdf') and not pdf_bytes and (m_po or m_id):
                pdf_name = fn
                try:
                    with open(f, 'rb') as fp:
                        pdf_bytes = fp.read()
                except Exception:
                    pass

    return msg_bytes, msg_name, pdf_bytes, pdf_name

def generate_apertura_piezas_excel(po, id_interno, cab_info, df_partidas):
    """
    Genera el archivo oficial de Apertura de Proyecto / Lista de Piezas en formato .xlsx
    con diseño corporativo formal SIGRAMA, datos de cabecera y desglose completo de despiece.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Lista_de_Piezas"
    ws.views.sheetView[0].showGridLines = True

    C_SLATE_DARK = "0F172A"
    C_SLATE_MID  = "1E293B"
    C_RED_SIG    = "EC2024"
    C_BORDER_THIN = Side(border_style="thin", color="CBD5E1")
    C_BORDER_MED  = Side(border_style="medium", color="0F172A")
    C_BORDER_DBL  = Side(border_style="double", color="0F172A")

    thin_border = Border(top=C_BORDER_THIN, bottom=C_BORDER_THIN, left=C_BORDER_THIN, right=C_BORDER_THIN)
    total_border = Border(top=C_BORDER_THIN, bottom=C_BORDER_DBL, left=C_BORDER_THIN, right=C_BORDER_THIN)

    # 1. Header Banner
    ws.merge_cells("A1:K1")
    c1 = ws["A1"]
    c1.value = "INDUSTRIA SIGRAMA S.A. DE C.V.  —  LISTA DE PIEZAS / APERTURA DE PROYECTO"
    c1.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    c1.fill = PatternFill(start_color=C_SLATE_DARK, end_color=C_SLATE_DARK, fill_type="solid")
    c1.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    now_str = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
    ws.merge_cells("A2:K2")
    c2 = ws["A2"]
    c2.value = f"Documento Oficial de Despiece y Programación Operativa | Emisión: {now_str} | Control de Planta y Almacén"
    c2.font = Font(name="Calibri", size=9.5, italic=True, color="94A3B8")
    c2.fill = PatternFill(start_color=C_SLATE_MID, end_color=C_SLATE_MID, fill_type="solid")
    c2.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 20

    ws.row_dimensions[3].height = 6

    # 2. Tarjetas de Metadatos del Proyecto
    def set_meta_cell(cell_ref, label, val):
        ws[cell_ref].value = f"{label}: {val}"
        ws[cell_ref].font = Font(name="Calibri", size=9.5, bold=True, color=C_SLATE_DARK)
        ws[cell_ref].fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
        ws[cell_ref].alignment = Alignment(horizontal="left", vertical="center")

    ws.merge_cells("A4:C4")
    set_meta_cell("A4", "No. Proyecto Interno", str(id_interno if id_interno else "S/N"))
    ws.merge_cells("D4:F4")
    set_meta_cell("D4", "Orden de Compra (PO)", str(po))
    ws.merge_cells("G4:I4")
    set_meta_cell("G4", "Proyecto", str(cab_info.get('proyecto', 'N/A')))
    ws.merge_cells("J4:K4")
    set_meta_cell("J4", "Estatus", str(cab_info.get('estatus_general', 'Registrada')))

    ws.merge_cells("A5:C5")
    set_meta_cell("A5", "Comprador", str(cab_info.get('comprador', 'N/A')))
    ws.merge_cells("D5:F5")
    set_meta_cell("D5", "Solicitante", str(cab_info.get('solicitante', 'N/A')))
    ws.merge_cells("G5:I5")
    set_meta_cell("G5", "Fecha Llegada PO", str(cab_info.get('fecha_llegada', 'N/A')))
    ws.merge_cells("J5:K5")
    set_meta_cell("J5", "Fecha Entrega Req.", str(cab_info.get('fecha_solicitada', 'N/A')))

    for r_idx in range(4, 6):
        ws.row_dimensions[r_idx].height = 22
        for col_idx in range(1, 12):
            ws.cell(row=r_idx, column=col_idx).border = thin_border

    ws.row_dimensions[6].height = 8

    # 3. Encabezados de Columnas
    headers = [
        ("#", 6, "center"),
        ("SKU Cliente", 18, "center"),
        ("SKU Planta (Clave)", 20, "center"),
        ("Descripción del Producto", 38, "left"),
        ("Cant. Requerida", 16, "right"),
        ("Unidad", 10, "center"),
        ("P. Unitario ($)", 15, "right"),
        ("Importe Total ($)", 18, "right"),
        ("Fecha Entrega", 15, "center"),
        ("Parcialidad", 13, "center"),
        ("Observaciones / Notas", 25, "left")
    ]

    header_row = 7
    ws.row_dimensions[header_row].height = 24
    for idx, (h_title, w, h_align) in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=idx)
        cell.value = h_title
        cell.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color=C_SLATE_DARK, end_color=C_SLATE_DARK, fill_type="solid")
        cell.alignment = Alignment(horizontal=h_align, vertical="center")
        cell.border = Border(top=C_BORDER_MED, bottom=Side(border_style="medium", color=C_RED_SIG), left=C_BORDER_THIN, right=C_BORDER_THIN)
        col_letter = get_column_letter(idx)
        ws.column_dimensions[col_letter].width = w

    # 4. Filas de Partidas
    curr_row = 8
    df_sorted = df_partidas.copy()
    if 'item_no' in df_sorted.columns:
        df_sorted['item_no_num'] = pd.to_numeric(df_sorted['item_no'], errors='coerce').fillna(9999)
        df_sorted = df_sorted.sort_values('item_no_num')

    for _, row in df_sorted.iterrows():
        ws.row_dimensions[curr_row].height = 20
        bg_color = "FFFFFF" if curr_row % 2 == 0 else "F8FAFC"
        row_fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")

        i_no = row.get('item_no', curr_row - 7)
        sk_c = str(row.get('sku_cliente', '') or '')
        sk_p = str(row.get('clave_sku', '') or '')
        desc = str(row.get('descripcion_producto', '') or '')
        cant = float(row.get('cantidad_requerida', 0) or 0)
        unid = str(row.get('unidad', 'PZA') or 'PZA').upper()
        pu   = float(row.get('precio_unitario', 0) or 0)
        pt   = float(row.get('precio_total', 0) or (cant * pu))
        fe   = str(row.get('fecha_entrega', '') or '')
        parc = str(row.get('parcialidad', 'P1') or 'P1')
        obs  = str(row.get('observaciones_partida', row.get('estatus_partida_360', '')) or '')

        values = [
            (i_no, "center", "@"),
            (sk_c, "center", "@"),
            (sk_p, "center", "@"),
            (desc, "left", "@"),
            (cant, "right", '#,##0 "pzas"'),
            (unid, "center", "@"),
            (pu, "right", '"$"#,##0.00'),
            (pt, "right", '"$"#,##0.00'),
            (fe, "center", "@"),
            (parc, "center", "@"),
            (obs, "left", "@")
        ]

        for col_idx, (val, align, num_fmt) in enumerate(values, start=1):
            cell = ws.cell(row=curr_row, column=col_idx)
            cell.value = val
            cell.font = Font(name="Calibri", size=9.5)
            cell.alignment = Alignment(horizontal=align, vertical="center")
            cell.fill = row_fill
            cell.border = thin_border
            if num_fmt != "@":
                cell.number_format = num_fmt

        curr_row += 1

    # 5. Fila de Totales Generales
    ws.row_dimensions[curr_row].height = 24
    ws.merge_cells(f"A{curr_row}:D{curr_row}")
    lbl_tot = ws[f"A{curr_row}"]
    lbl_tot.value = "TOTAL GENERAL DE PIEZAS E IMPORTE"
    lbl_tot.font = Font(name="Calibri", size=10, bold=True, color=C_SLATE_DARK)
    lbl_tot.alignment = Alignment(horizontal="center", vertical="center")

    for col_idx in range(1, 5):
        ws.cell(row=curr_row, column=col_idx).fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
        ws.cell(row=curr_row, column=col_idx).border = total_border

    c_tot_cant = ws.cell(row=curr_row, column=5)
    c_tot_cant.value = f"=SUM(E8:E{curr_row-1})"
    c_tot_cant.number_format = '#,##0 "pzas"'
    c_tot_cant.font = Font(name="Calibri", size=10, bold=True, color=C_SLATE_DARK)
    c_tot_cant.fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    c_tot_cant.alignment = Alignment(horizontal="right", vertical="center")
    c_tot_cant.border = total_border

    for col_idx in range(6, 8):
        c_b = ws.cell(row=curr_row, column=col_idx)
        c_b.fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
        c_b.border = total_border

    c_tot_imp = ws.cell(row=curr_row, column=8)
    c_tot_imp.value = f"=SUM(H8:H{curr_row-1})"
    c_tot_imp.number_format = '"$"#,##0.00'
    c_tot_imp.font = Font(name="Calibri", size=10, bold=True, color=C_SLATE_DARK)
    c_tot_imp.fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    c_tot_imp.alignment = Alignment(horizontal="right", vertical="center")
    c_tot_imp.border = total_border

    for col_idx in range(9, 12):
        c_b = ws.cell(row=curr_row, column=col_idx)
        c_b.fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
        c_b.border = total_border

    ws.auto_filter.ref = f"A7:K{curr_row-1}"
    ws.freeze_panes = "A8"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()

def generate_apertura_eml(po, id_interno, cab_info, df_partidas, msg_bytes=None, msg_name=None, pdf_bytes=None, pdf_name=None, excel_bytes=None):
    """
    Genera el correo formal de Apertura de Proyecto Interno en formato RFC 822 (.eml)
    listo para abrirse en Outlook o Thunderbird, con el cuerpo HTML institucional
    y los archivos adjuntos embebidos: Lista de Piezas (.xlsx), Correo Original (.msg) y PDF de la PO.
    """
    msg = email.message.EmailMessage()

    def _clean(v, d=""):
        if v is None or pd.isna(v):
            return d
        s = str(v).replace('\r', ' ').replace('\n', ' ').replace('\t', ' ')
        s = re.sub(r'\s+', ' ', s).strip()
        return s if s else d

    po_clean = _clean(po, "S/N")
    id_clean = _clean(id_interno, "INT-S/N")
    prj_clean = _clean(cab_info.get('proyecto', 'PROYECTO SIGRAMA'), 'PROYECTO SIGRAMA')
    cmp_clean = _clean(cab_info.get('comprador', 'Compras'), 'Compras')
    sol_clean = _clean(cab_info.get('solicitante', 'Solicitante'), 'Solicitante')
    f_lleg = _clean(cab_info.get('fecha_llegada', 'N/A'), 'N/A')
    f_sol = _clean(cab_info.get('fecha_solicitada', 'N/A'), 'N/A')
    est_gen = _clean(cab_info.get('estatus_general', 'Registrada'), 'Registrada')

    tot_pzas = float(df_partidas['cantidad_requerida'].sum()) if 'cantidad_requerida' in df_partidas.columns else 0.0
    tot_imp = float(cab_info.get('total', 0) or 0)
    if tot_imp == 0.0 and 'precio_total' in df_partidas.columns:
        tot_imp = float(df_partidas['precio_total'].sum())

    msg['Subject'] = f"[APERTURA DE PROYECTO INTERNO] {id_clean} - OC {po_clean} | PROYECTO: {prj_clean}"
    msg['From'] = 'operaciones@sigrama.com.mx'
    msg['To'] = f"{cmp_clean.lower().replace(' ', '.')}@sigrama.com.mx"
    msg['Cc'] = 'operaciones@sigrama.com.mx, produccion@sigrama.com.mx, calidad@sigrama.com.mx, almacen@sigrama.com.mx'
    msg['Date'] = email.utils.formatdate(localtime=True)

    # HTML Body
    filas_html = ""
    for idx, (_, r) in enumerate(df_partidas.head(35).iterrows(), start=1):
        bg = "#FFFFFF" if idx % 2 != 0 else "#F8FAFC"
        sk_c = str(r.get('sku_cliente', '') or '')
        sk_p = str(r.get('clave_sku', '') or '')
        desc = str(r.get('descripcion_producto', '') or '')
        cant = float(r.get('cantidad_requerida', 0) or 0)
        unid = str(r.get('unidad', 'PZA') or 'PZA').upper()
        fe   = str(r.get('fecha_entrega', '') or '')
        filas_html += f"""
        <tr style="background-color: {bg}; border-bottom: 1px solid #E2E8F0;">
            <td style="padding: 6px 8px; text-align: center; font-weight: bold; color: #475569;">{r.get('item_no', idx)}</td>
            <td style="padding: 6px 8px; font-weight: 600; color: #0F172A;">{sk_c}</td>
            <td style="padding: 6px 8px; color: #2563EB; font-weight: 600;">{sk_p}</td>
            <td style="padding: 6px 8px; color: #334155;">{desc}</td>
            <td style="padding: 6px 8px; text-align: right; font-weight: bold; color: #0F172A;">{cant:,.0f}</td>
            <td style="padding: 6px 8px; text-align: center; color: #64748B;">{unid}</td>
            <td style="padding: 6px 8px; text-align: center; color: #059669; font-weight: 600;">{fe}</td>
        </tr>
        """

    mas_filas_nota = ""
    if len(df_partidas) > 35:
        mas_filas_nota = f"""
        <tr>
            <td colspan="7" style="padding: 10px; text-align: center; background-color: #FEF3C7; color: #B45309; font-weight: bold; font-size: 11px;">
                Mostrando 35 de {len(df_partidas)} partidas. Consulte la lista completa en el archivo Excel adjunto: Lista_Piezas_Despiece_{id_clean}_{po_clean}.xlsx
            </td>
        </tr>
        """

    adjuntos_html = []
    adjuntos_html.append(f"<li>📊 <b>Lista de Piezas Oficial:</b> <code>Lista_Piezas_Despiece_{id_clean}_{po_clean}.xlsx</code></li>")
    if msg_name:
        adjuntos_html.append(f"<li>📧 <b>Correo Original Embebido:</b> <code>{msg_name}</code></li>")
    if pdf_name:
        adjuntos_html.append(f"<li>📄 <b>Orden de Compra Oficial (PDF):</b> <code>{pdf_name}</code></li>")
    adjuntos_str = "".join(adjuntos_html)

    html_content = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
</head>
<body style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; background-color: #F1F5F9; margin: 0; padding: 20px;">
<div style="max-width: 820px; margin: 0 auto; background-color: #FFFFFF; border-radius: 10px; overflow: hidden; box-shadow: 0 4px 12px rgba(0,0,0,0.1); border: 1px solid #CBD5E1;">

    <!-- HEADER CORPORATIVO -->
    <div style="background-color: #0F172A; padding: 22px 28px; border-bottom: 4px solid #EC2024;">
        <table style="width: 100%; border-collapse: collapse;">
            <tr>
                <td>
                    <div style="color: #EC2024; font-size: 11px; font-weight: 800; letter-spacing: 1.5px; text-transform: uppercase;">INDUSTRIA SIGRAMA S.A. DE C.V.</div>
                    <h1 style="color: #FFFFFF; margin: 4px 0 0 0; font-size: 20px; font-weight: 900; letter-spacing: -0.5px;">APERTURA OFICIAL DE PROYECTO INTERNO</h1>
                </td>
                <td style="text-align: right;">
                    <span style="background-color: #EC2024; color: #FFFFFF; font-size: 15px; font-weight: 900; padding: 6px 14px; border-radius: 6px; letter-spacing: 0.5px;">{id_clean}</span>
                </td>
            </tr>
        </table>
    </div>

    <!-- CUERPO PRINCIPAL -->
    <div style="padding: 24px 28px;">
        <p style="font-size: 13.5px; color: #334155; line-height: 1.5; margin-top: 0;">
            Estimado equipo operativo, compras y jefatura de planta:<br>
            Se ha formalizado la <b>Apertura de Proyecto Interno</b> para la siguiente Orden de Compra. A continuación se detallan las especificaciones, fechas solicitadas y despiece de piezas para el arranque inmediato en talleres y almacén.
        </p>

        <!-- FICHA TÉCNICA DEL PROYECTO -->
        <table style="width: 100%; border-collapse: collapse; margin-bottom: 20px; background-color: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px;">
            <tr>
                <td style="padding: 9px 14px; font-size: 12px; color: #64748B; width: 25%;">No. Proyecto Interno:</td>
                <td style="padding: 9px 14px; font-size: 13px; font-weight: bold; color: #EC2024; width: 25%;">{id_clean}</td>
                <td style="padding: 9px 14px; font-size: 12px; color: #64748B; width: 25%;">Orden de Compra (PO):</td>
                <td style="padding: 9px 14px; font-size: 13px; font-weight: bold; color: #0F172A; width: 25%;">{po_clean}</td>
            </tr>
            <tr style="background-color: #F1F5F9;">
                <td style="padding: 9px 14px; font-size: 12px; color: #64748B;">Proyecto:</td>
                <td style="padding: 9px 14px; font-size: 13px; font-weight: bold; color: #0F172A;">{prj_clean}</td>
                <td style="padding: 9px 14px; font-size: 12px; color: #64748B;">Estatus:</td>
                <td style="padding: 9px 14px; font-size: 12px; font-weight: bold; color: #059669;">{est_gen}</td>
            </tr>
            <tr>
                <td style="padding: 9px 14px; font-size: 12px; color: #64748B;">Comprador:</td>
                <td style="padding: 9px 14px; font-size: 12.5px; font-weight: 600; color: #1E293B;">{cmp_clean}</td>
                <td style="padding: 9px 14px; font-size: 12px; color: #64748B;">Solicitante:</td>
                <td style="padding: 9px 14px; font-size: 12.5px; font-weight: 600; color: #1E293B;">{sol_clean}</td>
            </tr>
            <tr style="background-color: #F1F5F9;">
                <td style="padding: 9px 14px; font-size: 12px; color: #64748B;">Fecha Llegada PO:</td>
                <td style="padding: 9px 14px; font-size: 12.5px; font-weight: 600; color: #1E293B;">{f_lleg}</td>
                <td style="padding: 9px 14px; font-size: 12px; color: #64748B;">Fecha Solicitada Entrega:</td>
                <td style="padding: 9px 14px; font-size: 12.5px; font-weight: bold; color: #DC2626;">{f_sol}</td>
            </tr>
            <tr>
                <td style="padding: 9px 14px; font-size: 12px; color: #64748B;">Total Piezas:</td>
                <td style="padding: 9px 14px; font-size: 14px; font-weight: 900; color: #0F172A;">{tot_pzas:,.0f} pzas</td>
                <td style="padding: 9px 14px; font-size: 12px; color: #64748B;">Importe Total:</td>
                <td style="padding: 9px 14px; font-size: 14px; font-weight: 900; color: #059669;">${tot_imp:,.2f} MXN</td>
            </tr>
        </table>

        <!-- RESUMEN DE PARTIDAS -->
        <h3 style="color: #0F172A; font-size: 14px; font-weight: 800; margin: 20px 0 10px 0; text-transform: uppercase; letter-spacing: 0.5px;">
            📑 Resumen de Piezas / Partidas Requeridas ({len(df_partidas)} partidas)
        </h3>
        <table style="width: 100%; border-collapse: collapse; font-size: 11.5px; margin-bottom: 20px;">
            <thead>
                <tr style="background-color: #0F172A; color: #FFFFFF;">
                    <th style="padding: 8px; text-align: center; width: 40px; border-bottom: 2px solid #EC2024;">#</th>
                    <th style="padding: 8px; text-align: left; border-bottom: 2px solid #EC2024;">SKU Cliente</th>
                    <th style="padding: 8px; text-align: left; border-bottom: 2px solid #EC2024;">SKU Planta</th>
                    <th style="padding: 8px; text-align: left; border-bottom: 2px solid #EC2024;">Descripción</th>
                    <th style="padding: 8px; text-align: right; width: 75px; border-bottom: 2px solid #EC2024;">Cant. Req.</th>
                    <th style="padding: 8px; text-align: center; width: 50px; border-bottom: 2px solid #EC2024;">Unidad</th>
                    <th style="padding: 8px; text-align: center; width: 85px; border-bottom: 2px solid #EC2024;">F. Entrega</th>
                </tr>
            </thead>
            <tbody>
                {filas_html}
                {mas_filas_nota}
            </tbody>
        </table>

        <!-- INSTRUCCIONES OPERATIVAS -->
        <div style="background-color: #EFF6FF; border-left: 4px solid #3B82F6; padding: 14px 18px; border-radius: 6px; margin-bottom: 20px;">
            <b style="color: #1E40AF; font-size: 12.5px; display: block; margin-bottom: 6px;">⚙️ Plan de Acción Operativo:</b>
            <ul style="margin: 0; padding-left: 20px; font-size: 12px; color: #1E3A8A; line-height: 1.6;">
                <li><b>Corte y Doblez (Nesting):</b> Generar nidos Pronest y programar Órdenes de Fabricación (OFs) conforme al listado de piezas.</li>
                <li><b>Ensamble & Soldadura:</b> Coordinar ensamble y soldadura conforme a las fechas prometidas por parcialidad.</li>
                <li><b>Almacén & Embarques:</b> Verificar empaque, tarimas e identificación para generación de Remisiones.</li>
            </ul>
        </div>

        <!-- ADJUNTOS EMBEBIDOS -->
        <div style="background-color: #F8FAFC; border: 1px solid #CBD5E1; padding: 14px 18px; border-radius: 6px;">
            <b style="color: #0F172A; font-size: 12.5px; display: block; margin-bottom: 6px;">📎 Documentos Oficiales Embebidos en este Correo:</b>
            <ul style="margin: 0; padding-left: 20px; font-size: 12px; color: #334155; line-height: 1.6;">
                {adjuntos_str}
            </ul>
        </div>
    </div>

    <!-- FOOTER -->
    <div style="background-color: #F8FAFC; border-top: 1px solid #E2E8F0; padding: 14px 28px; text-align: center; font-size: 11px; color: #94A3B8;">
        Industria Sigrama S.A. de C.V. | Sistema Integral de Seguimiento y Trazabilidad 360° | Torreón, Coahuila
    </div>
</div>
</body>
</html>
"""
    msg.set_content(f"Apertura de Proyecto Interno {id_clean} - OC {po_clean}. Consulte la versión HTML y los archivos adjuntos.")
    msg.add_alternative(html_content, subtype='html')

    # Adjuntos
    if excel_bytes:
        msg.add_attachment(
            excel_bytes,
            maintype='application',
            subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            filename=f"Lista_Piezas_Despiece_{id_clean}_{po_clean}.xlsx"
        )
    if msg_bytes and msg_name:
        msg.add_attachment(
            msg_bytes,
            maintype='application',
            subtype='vnd.ms-outlook',
            filename=msg_name
        )
    if pdf_bytes and pdf_name:
        msg.add_attachment(
            pdf_bytes,
            maintype='application',
            subtype='pdf',
            filename=pdf_name
        )

    return msg.as_bytes()
