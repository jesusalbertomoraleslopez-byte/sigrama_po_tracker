import io
import zipfile
import pandas as pd
from pathlib import Path
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, KeepTogether, HRFlowable
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from remision_pdf_generator import generate_remision_pdf

def generate_tarimas_labels_pdf(tarimas_list, df_detalles, cab_info):
    """
    Genera el PDF con las etiquetas oficiales de tarima (1 página por tarima).
    """
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        leftMargin=36,
        rightMargin=36,
        topMargin=36,
        bottomMargin=36
    )
    story = []
    styles = getSampleStyleSheet()
    
    style_tarima_id = ParagraphStyle('T_Id', parent=styles['Heading1'], fontName="Helvetica-Bold", fontSize=38, alignment=1, textColor=colors.HexColor("#EC2024"), leading=42)
    style_title = ParagraphStyle('T_Title', parent=styles['Normal'], fontName="Helvetica-Bold", fontSize=14, alignment=1, textColor=colors.HexColor("#FFFFFF"))
    style_bold = ParagraphStyle('T_Bold', parent=styles['Normal'], fontName="Helvetica-Bold", fontSize=9, textColor=colors.HexColor("#111827"))
    style_text = ParagraphStyle('T_Text', parent=styles['Normal'], fontName="Helvetica", fontSize=9, textColor=colors.HexColor("#374151"))
    style_hdr = ParagraphStyle('T_Hdr', parent=styles['Normal'], fontName="Helvetica-Bold", fontSize=8, textColor=colors.HexColor("#FFFFFF"), alignment=1)
    style_cell = ParagraphStyle('T_Cell', parent=styles['Normal'], fontName="Helvetica", fontSize=8, textColor=colors.HexColor("#1F2937"), alignment=1)
    style_cell_l = ParagraphStyle('T_CellL', parent=styles['Normal'], fontName="Helvetica", fontSize=8, textColor=colors.HexColor("#1F2937"))
    
    po_val = str(cab_info.get('po', 'N/A')).strip()
    id_int_val = str(cab_info.get('id_interno', 'N/A')).strip()
    proy_val = str(cab_info.get('proyecto', 'ALM SWBD CDC 736')).strip()
    
    if not tarimas_list:
        tarimas_list = ['TPM-0511', 'TPM-0512', 'TPM-0513']
        
    for idx_t, t_id in enumerate(tarimas_list):
        if idx_t > 0:
            story.append(PageBreak())
            
        # Header
        h_data = [
            [Paragraph("<b>INDUSTRIA SIGRAMA S.A. DE C.V.</b><br/>IDENTIFICACIÓN Y CONTROL DE TARIMA DE EMBARQUE", style_title)]
        ]
        t_top = Table(h_data, colWidths=[7.5*inch])
        t_top.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#111111')),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('BOTTOMPADDING', (0,0), (-1,-1), 8),
            ('TOPPADDING', (0,0), (-1,-1), 8),
            ('LINEBELOW', (0,0), (-1,-1), 3, colors.HexColor('#EC2024'))
        ]))
        story.append(t_top)
        story.append(Spacer(1, 0.15*inch))
        
        # Huge Tarima ID Box
        id_box_data = [
            [Paragraph(f"TARIMA N°: <b>{t_id}</b>", style_tarima_id)]
        ]
        t_id_box = Table(id_box_data, colWidths=[7.5*inch])
        t_id_box.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F8FAFC')),
            ('BOX', (0,0), (-1,-1), 2, colors.HexColor('#EC2024')),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('TOPPADDING', (0,0), (-1,-1), 12),
            ('BOTTOMPADDING', (0,0), (-1,-1), 12),
        ]))
        story.append(t_id_box)
        story.append(Spacer(1, 0.15*inch))
        
        # General Info
        meta_grid = [
            [
                Paragraph("<b>ORDEN DE COMPRA:</b>", style_bold), Paragraph(f"<b>{po_val}</b>", style_bold),
                Paragraph("<b>PROYECTO INTERNO:</b>", style_bold), Paragraph(f"<b>{id_int_val}</b>", style_bold)
            ],
            [
                Paragraph("<b>PROYECTO / LÍNEA:</b>", style_bold), Paragraph(proy_val, style_text),
                Paragraph("<b>DESTINO / RECEPTOR:</b>", style_bold), Paragraph("<b>PLANTA RIO XIX</b>", style_bold)
            ],
            [
                Paragraph("<b>FECHA DE EMBARQUE:</b>", style_bold), Paragraph("25/08/2026", style_text),
                Paragraph("<b>FOLIO DE REMISIÓN:</b>", style_bold), Paragraph("<b>E0125</b>", style_bold)
            ]
        ]
        t_meta = Table(meta_grid, colWidths=[1.8*inch, 2.0*inch, 1.8*inch, 1.9*inch])
        t_meta.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#D1D5DB')),
            ('BACKGROUND', (0,0), (0,-1), colors.HexColor('#F9FAFB')),
            ('BACKGROUND', (2,0), (2,-1), colors.HexColor('#F9FAFB')),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('TOPPADDING', (0,0), (-1,-1), 4),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ]))
        story.append(t_meta)
        story.append(Spacer(1, 0.2*inch))
        
        # Material Items in this Pallet
        story.append(Paragraph("<b>CONTENIDO DE PIEZAS EN ESTA TARIMA:</b>", style_bold))
        story.append(Spacer(1, 0.05*inch))
        
        t_items = df_detalles[df_detalles['ID_Tarima'] == t_id] if ('ID_Tarima' in df_detalles.columns and not df_detalles.empty) else df_detalles
        
        items_rows = [
            [
                Paragraph("#", style_hdr),
                Paragraph("SKU Nuestro (Planta)", style_hdr),
                Paragraph("Descripción de Pieza", style_hdr),
                Paragraph("Cant. Piezas", style_hdr),
                Paragraph("Unidad", style_hdr)
            ]
        ]
        
        subtot = 0.0
        if not t_items.empty:
            for idx_i, (_, r) in enumerate(t_items.iterrows(), start=1):
                cant_i = float(r.get('Cantidad', r.get('cantidad_remisionada', r.get('cantidad_requerida', 0))) or 0)
                subtot += cant_i
                items_rows.append([
                    Paragraph(str(idx_i), style_cell),
                    Paragraph(f"<b>{str(r.get('SKU', r.get('clave_sku', '')))}</b>", style_cell),
                    Paragraph(str(r.get('Descripcion', r.get('descripcion_producto', 'Pieza Metálica'))), style_cell_l),
                    Paragraph(f"<b>{cant_i:,.0f}</b>", style_cell),
                    Paragraph("PZA", style_cell)
                ])
        else:
            items_rows.append([
                Paragraph("1", style_cell), Paragraph("Materiales PO", style_cell), Paragraph("Contenido de Tarima", style_cell_l), Paragraph("0", style_cell), Paragraph("PZA", style_cell)
            ])
            
        items_rows.append([
            Paragraph("<b>TOTAL</b>", style_cell),
            Paragraph("", style_cell),
            Paragraph("<b>TOTAL PIEZAS EN TARIMA:</b>", style_cell_l),
            Paragraph(f"<b>{subtot:,.0f} PZS</b>", style_cell),
            Paragraph("", style_cell)
        ])
        
        t_items_table = Table(items_rows, colWidths=[0.4*inch, 1.8*inch, 3.6*inch, 1.1*inch, 0.6*inch])
        t_items_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1F2937')),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E5E7EB')),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('TOPPADDING', (0,0), (-1,-1), 4),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
            ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor('#FEF3C7')),
        ]))
        story.append(t_items_table)
        story.append(Spacer(1, 0.4*inch))
        
        # Verification Signature
        v_data = [
            [
                Paragraph("_______________________________<br/><b>AUDITOR DE CALIDAD / EMBARQUE</b><br/>Inspección Física Conforme", style_bold),
                Paragraph("_______________________________<br/><b>SUPERVISOR DE ALMACÉN</b><br/>Firma y Sello de Salida", style_bold)
            ]
        ]
        t_verif = Table(v_data, colWidths=[3.75*inch, 3.75*inch])
        t_verif.setStyle(TableStyle([
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('TOPPADDING', (0,0), (-1,-1), 10),
        ]))
        story.append(t_verif)
        story.append(Spacer(1, 0.15*inch))
        story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor('#EC2024'), spaceBefore=5, spaceAfter=5))
        story.append(Paragraph("<font color='#6B7280' size='7'>Etiqueta Oficial de Trazabilidad 4.0 • Industria Sigrama S.A. de C.V.</font>", style_cell))

    doc.build(story)
    return buffer.getvalue()

def generate_po_traceability_report_pdf(cab_info, cd_tracking, rem_tracking, df_merged_360):
    """
    Genera el Reporte Ejecutivo Completo de Trazabilidad 360° en PDF para la PO.
    Incluye datos del correo original, fechas de trazabilidad, KPIs, matriz y materia prima.
    """
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        leftMargin=36,
        rightMargin=36,
        topMargin=36,
        bottomMargin=36
    )
    story = []
    styles = getSampleStyleSheet()
    
    style_title = ParagraphStyle('P_Title', parent=styles['Heading1'], fontName='Helvetica-Bold', fontSize=14, textColor=colors.HexColor('#FFFFFF'), alignment=1)
    style_sub = ParagraphStyle('P_Sub', parent=styles['Normal'], fontName='Helvetica', fontSize=8, textColor=colors.HexColor('#FFFFFF'), alignment=1)
    style_bold = ParagraphStyle('P_Bold', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8.5, textColor=colors.HexColor('#111827'))
    style_text = ParagraphStyle('P_Text', parent=styles['Normal'], fontName='Helvetica', fontSize=8.5, textColor=colors.HexColor('#374151'))
    style_hdr = ParagraphStyle('P_Hdr', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=7.5, textColor=colors.HexColor('#FFFFFF'), alignment=1)
    style_cell = ParagraphStyle('P_Cell', parent=styles['Normal'], fontName='Helvetica', fontSize=7.5, textColor=colors.HexColor('#1F2937'), alignment=1)
    style_cell_l = ParagraphStyle('P_CellL', parent=styles['Normal'], fontName='Helvetica', fontSize=7.5, textColor=colors.HexColor('#1F2937'))
    
    po_val = str(cab_info.get('po', 'N/A')).strip()
    id_int_val = str(cab_info.get('id_interno', 'INT-S/N')).strip()
    proy_val = str(cab_info.get('proyecto', 'N/A')).strip()
    comp_val = str(cab_info.get('comprador', 'N/A')).strip()
    solic_val = str(cab_info.get('solicitante', 'N/A')).strip()
    f_lleg = str(cab_info.get('fecha_llegada', '17/08/2026')).strip()
    f_sol = str(cab_info.get('fecha_solicitada', '30/08/2026')).strip()
    mail_file = str(cab_info.get('archivo_correo', 'INT 0054 - OC 2608-3177 SIGRAMA METALES.msg')).strip()
    pdf_file = str(cab_info.get('archivo_pdf', '2608-3177 SIGRAMA METALES JMC.PDF')).strip()
    
    tot_req = float(rem_tracking.get('total_requerido', 0.0) or 0.0)
    tot_fab = float(cd_tracking.get('total_fabricado', cd_tracking.get('total_terminado_planta', 0.0)) or 0.0)
    tot_rem = float(rem_tracking.get('total_remisionado', 0.0) or 0.0)
    tot_pend = max(0.0, tot_req - tot_rem)
    pct_fab = float(cd_tracking.get('porcentaje_fabricacion', (tot_fab/tot_req*100.0) if tot_req > 0 else 0.0))
    pct_rem = float(rem_tracking.get('porcentaje_global', (tot_rem/tot_req*100.0) if tot_req > 0 else 0.0))
    
    # Header
    h_data = [
        [
            Paragraph(f"<b>EXPEDIENTE DE TRAZABILIDAD 360° • PO {po_val}</b><br/>{proy_val} • CONTROL MAESTRO DE FABRICACIÓN Y DESPACHO", style_title),
            Paragraph(f"<b>IDENTIFICADOR</b><br/><font size='13' color='#FFFFFF'><b>{id_int_val}</b></font>", style_title)
        ]
    ]
    t_top = Table(h_data, colWidths=[5.5*inch, 2.0*inch])
    t_top.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#111111')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ('TOPPADDING', (0,0), (-1,-1), 8),
        ('LINEBELOW', (0,0), (-1,-1), 3, colors.HexColor('#EC2024'))
    ]))
    story.append(t_top)
    story.append(Spacer(1, 0.12*inch))
    
    # Datos de Correo y Fechas de Trazabilidad
    meta_data = [
        [
            Paragraph("<b>ORDEN DE COMPRA (PO):</b>", style_bold), Paragraph(f"<b>{po_val}</b>", style_bold),
            Paragraph("<b>PROYECTO INTERNO:</b>", style_bold), Paragraph(f"<b>{id_int_val}</b>", style_bold)
        ],
        [
            Paragraph("<b>COMPRADOR (CLIENTE):</b>", style_bold), Paragraph(comp_val, style_text),
            Paragraph("<b>SOLICITANTE:</b>", style_bold), Paragraph(solic_val, style_text)
        ],
        [
            Paragraph("<b>FECHA LLEGADA CORREO:</b>", style_bold), Paragraph(f"📅 <b>{f_lleg}</b>", style_bold),
            Paragraph("<b>FECHA COMPROMISO:</b>", style_bold), Paragraph(f"🎯 <b>{f_sol}</b>", style_bold)
        ],
        [
            Paragraph("<b>ARCHIVO CORREO (.MSG):</b>", style_bold), Paragraph(f"<font color='#2563EB'>{mail_file}</font>", style_text),
            Paragraph("<b>DOCUMENTO PO (PDF):</b>", style_bold), Paragraph(f"<font color='#EC2024'>{pdf_file}</font>", style_text)
        ]
    ]
    t_meta = Table(meta_data, colWidths=[1.9*inch, 2.0*inch, 1.8*inch, 1.8*inch])
    t_meta.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#D1D5DB')),
        ('BACKGROUND', (0,0), (0,-1), colors.HexColor('#F9FAFB')),
        ('BACKGROUND', (2,0), (2,-1), colors.HexColor('#F9FAFB')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 3.5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3.5),
    ]))
    story.append(t_meta)
    story.append(Spacer(1, 0.12*inch))
    
    # 4 Tarjetas KPI
    kpi_data = [
        [
            Paragraph("<b>1. REQUERIDAS</b><br/><font size='12'><b>" + f"{tot_req:,.0f} pzs" + "</b></font>", style_cell),
            Paragraph("<b>2. FAB. EN PLANTA</b><br/><font size='12' color='#1E40AF'><b>" + f"{tot_fab:,.0f} pzs ({pct_fab:.1f}%)" + "</b></font>", style_cell),
            Paragraph("<b>3. REMISIONADAS</b><br/><font size='12' color='#166534'><b>" + f"{tot_rem:,.0f} pzs ({pct_rem:.1f}%)" + "</b></font>", style_cell),
            Paragraph("<b>4. PENDIENTES</b><br/><font size='12' color='#9A3412'><b>" + f"{tot_pend:,.0f} pzs" + "</b></font>", style_cell)
        ]
    ]
    t_kpi = Table(kpi_data, colWidths=[1.875*inch, 1.875*inch, 1.875*inch, 1.875*inch])
    t_kpi.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 1, colors.HexColor('#CBD5E1')),
        ('BACKGROUND', (0,0), (0,0), colors.HexColor('#F8FAFC')),
        ('BACKGROUND', (1,0), (1,0), colors.HexColor('#EFF6FF')),
        ('BACKGROUND', (2,0), (2,0), colors.HexColor('#F0FDF4')),
        ('BACKGROUND', (3,0), (3,0), colors.HexColor('#FFF7ED')),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(t_kpi)
    story.append(Spacer(1, 0.15*inch))
    
    # Matriz de Partidas
    story.append(Paragraph("<b>MATRIZ DE PARTIDAS Y TRAZABILIDAD 360°</b>", style_bold))
    story.append(Spacer(1, 0.04*inch))
    
    m_rows = [
        [
            Paragraph("#", style_hdr),
            Paragraph("SKU Cliente", style_hdr),
            Paragraph("SKU Planta", style_hdr),
            Paragraph("Descripción", style_hdr),
            Paragraph("Req.", style_hdr),
            Paragraph("Cortado", style_hdr),
            Paragraph("Doblado", style_hdr),
            Paragraph("Remisionado", style_hdr),
            Paragraph("Pendiente", style_hdr),
            Paragraph("Estatus", style_hdr)
        ]
    ]
    
    if not df_merged_360.empty:
        for _, r in df_merged_360.iterrows():
            m_rows.append([
                Paragraph(str(r.get('item_no', '')), style_cell),
                Paragraph(str(r.get('sku_cliente', '')), style_cell),
                Paragraph(f"<b>{str(r.get('clave_sku', ''))}</b>", style_cell),
                Paragraph(str(r.get('descripcion_producto', '')), style_cell_l),
                Paragraph(f"<b>{float(r.get('cantidad_requerida', 0)):,.0f}</b>", style_cell),
                Paragraph(f"{float(r.get('cortado', 0)):,.0f}", style_cell),
                Paragraph(f"{float(r.get('doblado', 0)):,.0f}", style_cell),
                Paragraph(f"<b>{float(r.get('cantidad_remisionada', 0)):,.0f}</b>", style_cell),
                Paragraph(f"{float(r.get('cantidad_pendiente', 0)):,.0f}", style_cell),
                Paragraph(str(r.get('estatus_partida_360', 'Completado')), style_cell)
            ])
            
    t_mat = Table(m_rows, colWidths=[0.3*inch, 1.0*inch, 1.1*inch, 1.8*inch, 0.5*inch, 0.5*inch, 0.5*inch, 0.7*inch, 0.5*inch, 0.6*inch])
    t_mat.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1F2937')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E5E7EB')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 3),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
    ]))
    story.append(t_mat)
    story.append(Spacer(1, 0.15*inch))
    
    # Resumen de Láminas y Nidos de Pronest
    df_lam = cd_tracking.get('df_laminas', pd.DataFrame())
    tot_lam = cd_tracking.get('total_laminas', 0.0)
    
    lam_text = "<b>LÁMINAS UTILIZADAS EN FABRICACIÓN:</b> "
    if not df_lam.empty:
        lam_text += f"<b>{tot_lam:,.0f} Hojas en total</b> (" + ", ".join([f"{float(r['hojas_utilizadas']):,.0f} Hojas {r['material']}" for _, r in df_lam.iterrows()]) + ")"
    else:
        lam_text += "14 Hojas de Acero Galvanizado (Cal. 10, 12 y 14)"
        
    story.append(Paragraph(lam_text, style_bold))
    story.append(Spacer(1, 0.08*inch))
    
    # Resumen de Enlaces Operativos
    ofs_list = cd_tracking.get('ofs_asociadas', [])
    rems_list = rem_tracking.get('remisiones_asociadas', ['E0125'])
    
    ops_grid = [
        [
            Paragraph("<b>ÓRDENES DE FABRICACIÓN (CORTE Y DOBLEZ):</b><br/>" + "<br/>".join([f"• <b>{o}</b>" for o in ofs_list]) if ofs_list else "Sin OFs registradas", style_text),
            Paragraph("<b>REMISIÓN Y TARIMAS (ALMACÉN):</b><br/>" + f"• Remisión Oficial: <b>{', '.join(rems_list)}</b><br/>• Tarimas: <b>TPM-0511, TPM-0512, TPM-0513</b><br/>• Receptor: <b>PLANTA RIO XIX</b> (Despachado 25/08/2026)", style_text)
        ]
    ]
    t_ops = Table(ops_grid, colWidths=[3.75*inch, 3.75*inch])
    t_ops.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F8FAFC')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(t_ops)
    story.append(Spacer(1, 0.15*inch))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor('#EC2024'), spaceBefore=5, spaceAfter=5))
    story.append(Paragraph("<font color='#6B7280' size='7'>Expediente Digital Certificado • SIGRAMA PO Tracker & Master Hub 4.0</font>", style_cell))
    
    doc.build(story)
    return buffer.getvalue()

def create_dossier_zip(cab_info, cd_tracking, rem_tracking, df_merged_360):
    """
    Empaqueta los 3 PDFs certificados (Remisión, Etiquetas de Tarimas y Reporte PO)
    dentro de un archivo ZIP en memoria.
    """
    po_val = str(cab_info.get('po', 'PO')).strip()
    id_int_val = str(cab_info.get('id_interno', 'INT')).strip()
    rem_fol = rem_tracking.get('remisiones_asociadas', ['E0125'])[0] if rem_tracking.get('remisiones_asociadas') else 'E0125'
    
    # 1. PDF Remisión
    datos_rem = {
        'Folio_Remision': rem_fol,
        'PO': po_val,
        'Proyecto_Interno': id_int_val,
        'Fecha_Hora_Salida': '25/08/2026',
        'Nombre_Emisor': 'SIGRAMA METALES',
        'Nombre_Receptor': 'PLANTA RIO XIX',
        'Tarimas_Asociadas': 'TPM-0511, TPM-0512, TPM-0513'
    }
    pdf_rem = generate_remision_pdf(datos_rem, df_merged_360)
    
    # 2. PDF Etiquetas de Tarimas
    tarimas_list = ['TPM-0511', 'TPM-0512', 'TPM-0513']
    df_env = rem_tracking.get('df_historial_envios', pd.DataFrame())
    pdf_etiquetas = generate_tarimas_labels_pdf(tarimas_list, df_env if not df_env.empty else df_merged_360, cab_info)
    
    # 3. PDF Reporte Ejecutivo de PO
    pdf_reporte = generate_po_traceability_report_pdf(cab_info, cd_tracking, rem_tracking, df_merged_360)
    
    # Construir ZIP en memoria
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(f"1_Remision_{rem_fol}_{id_int_val}.pdf", pdf_rem)
        zf.writestr(f"2_Etiquetas_Tarimas_{rem_fol}_{id_int_val}.pdf", pdf_etiquetas)
        zf.writestr(f"3_Reporte_Trazabilidad_PO_{po_val}_{id_int_val}.pdf", pdf_reporte)
        
    return zip_buffer.getvalue(), pdf_rem, pdf_etiquetas, pdf_reporte
