import io
import pandas as pd
from pathlib import Path
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, KeepTogether, HRFlowable
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

def generate_remision_pdf(datos_remision, df_detalles):
    """
    Genera el PDF oficial de Remisión de Materiales de Sigrama.
    Retorna bytes del PDF generado en memoria.
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
    
    style_title = ParagraphStyle('R_Title', parent=styles['Heading1'], fontName='Helvetica-Bold', fontSize=13, textColor=colors.HexColor('#FFFFFF'), alignment=1)
    style_sub = ParagraphStyle('R_Sub', parent=styles['Normal'], fontName='Helvetica', fontSize=8, textColor=colors.HexColor('#FFFFFF'), alignment=1)
    style_lbl_b = ParagraphStyle('R_LblB', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8, textColor=colors.HexColor('#111827'))
    style_val = ParagraphStyle('R_Val', parent=styles['Normal'], fontName='Helvetica', fontSize=8, textColor=colors.HexColor('#374151'))
    style_hdr_t = ParagraphStyle('R_HdrT', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8, textColor=colors.HexColor('#FFFFFF'), alignment=1)
    style_cell = ParagraphStyle('R_Cell', parent=styles['Normal'], fontName='Helvetica', fontSize=7.5, textColor=colors.HexColor('#1F2937'), alignment=1)
    style_cell_l = ParagraphStyle('R_CellL', parent=styles['Normal'], fontName='Helvetica', fontSize=7.5, textColor=colors.HexColor('#1F2937'))
    style_sign_lbl = ParagraphStyle('R_SignLbl', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=7.5, textColor=colors.HexColor('#4B5563'), alignment=1)
    
    # 1. Header Banner Corporativo
    folio_rem = str(datos_remision.get('Folio_Remision', 'REM-0000')).strip()
    po_val = str(datos_remision.get('PO', datos_remision.get('po', ''))).strip()
    id_int_val = str(datos_remision.get('Proyecto_Interno', datos_remision.get('id_interno', ''))).strip()
    
    header_data = [
        [
            Paragraph("<b>INDUSTRIA SIGRAMA S.A. DE C.V.</b><br/>REMISIÓN OFICIAL DE DESPACHO Y ENTREGA DE MATERIALES", style_title),
            Paragraph(f"<b>FOLIO REMISIÓN</b><br/><font size='12' color='#FFFFFF'><b>{folio_rem}</b></font>", style_title)
        ]
    ]
    t_top = Table(header_data, colWidths=[5.5*inch, 2.0*inch])
    t_top.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#111111')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ('TOPPADDING', (0,0), (-1,-1), 8),
        ('LINEBELOW', (0,0), (-1,-1), 3, colors.HexColor('#EC2024'))
    ]))
    story.append(t_top)
    story.append(Spacer(1, 0.15*inch))
    
    # 2. Datos Generales de Embarque
    f_salida = str(datos_remision.get('Fecha_Hora_Salida', datos_remision.get('fecha_salida', '2026-08-25 14:30'))).strip()
    emisor = str(datos_remision.get('Nombre_Emisor', 'SIGRAMA METALES')).strip()
    dir_emisor = str(datos_remision.get('Direccion_Emisor', 'PLANTA METALES SIGRAMA')).strip()
    receptor = str(datos_remision.get('Nombre_Receptor', 'PLANTA RIO XIX')).strip()
    dir_receptor = str(datos_remision.get('Direccion_Receptor', 'PARQUE INDUSTRIAL')).strip()
    chofer = str(datos_remision.get('Nombre_Chofer', 'TRANSPORTE INTERNO')).strip()
    tarimas_str = str(datos_remision.get('Tarimas_Asociadas', 'TPM-0511, TPM-0512, TPM-0513')).strip()
    
    info_grid = [
        [
            Paragraph("<b>ORDEN DE COMPRA (PO):</b>", style_lbl_b), Paragraph(f"<b>{po_val}</b>", style_lbl_b),
            Paragraph("<b>PROYECTO INTERNO:</b>", style_lbl_b), Paragraph(f"<b>{id_int_val}</b>", style_lbl_b)
        ],
        [
            Paragraph("<b>FECHA DESPACHO:</b>", style_lbl_b), Paragraph(f_salida, style_val),
            Paragraph("<b>PLANTA ORIGEN:</b>", style_lbl_b), Paragraph(f"{emisor} ({dir_emisor})", style_val)
        ],
        [
            Paragraph("<b>RECEPTOR / DESTINO:</b>", style_lbl_b), Paragraph(f"<b>{receptor}</b> - {dir_receptor}", style_val),
            Paragraph("<b>TRANSPORTISTA / CHOFER:</b>", style_lbl_b), Paragraph(chofer, style_val)
        ],
        [
            Paragraph("<b>TARIMAS ASOCIADAS:</b>", style_lbl_b), Paragraph(f"<b>{tarimas_str}</b>", style_lbl_b),
            Paragraph("<b>ESTATUS DESPACHO:</b>", style_lbl_b), Paragraph("<font color='#16A34A'><b>ENTREGADO / REMESADO</b></font>", style_val)
        ]
    ]
    t_info = Table(info_grid, colWidths=[1.8*inch, 2.0*inch, 1.8*inch, 1.9*inch])
    t_info.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#D1D5DB')),
        ('BACKGROUND', (0,0), (0,-1), colors.HexColor('#F9FAFB')),
        ('BACKGROUND', (2,0), (2,-1), colors.HexColor('#F9FAFB')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
    ]))
    story.append(t_info)
    story.append(Spacer(1, 0.2*inch))
    
    # 3. Tabla de Materiales Despachados
    story.append(Paragraph("<b>DETALLE DE MATERIALES Y PARTIDAS DESPACHADAS</b>", style_lbl_b))
    story.append(Spacer(1, 0.05*inch))
    
    mat_rows = [
        [
            Paragraph("#", style_hdr_t),
            Paragraph("ID Tarima", style_hdr_t),
            Paragraph("SKU Nuestro (Planta)", style_hdr_t),
            Paragraph("Descripción del Material", style_hdr_t),
            Paragraph("Cant. (Pzs)", style_hdr_t),
            Paragraph("Unidad", style_hdr_t),
            Paragraph("Estatus", style_hdr_t)
        ]
    ]
    
    tot_piezas = 0.0
    if not df_detalles.empty:
        for idx, (_, r) in enumerate(df_detalles.iterrows(), start=1):
            cant = float(r.get('Cantidad', r.get('cantidad_remisionada', r.get('cantidad_requerida', 0))) or 0)
            tot_piezas += cant
            mat_rows.append([
                Paragraph(str(idx), style_cell),
                Paragraph(str(r.get('ID_Tarima', r.get('id_tarima', 'TPM-S/N'))), style_cell),
                Paragraph(f"<b>{str(r.get('SKU', r.get('clave_sku', '')))}</b>", style_cell),
                Paragraph(str(r.get('Descripcion', r.get('descripcion_producto', 'Material de Fabricación'))), style_cell_l),
                Paragraph(f"<b>{cant:,.0f}</b>", style_cell),
                Paragraph("PZA", style_cell),
                Paragraph("<font color='#16A34A'><b>Remesado</b></font>", style_cell)
            ])
    else:
        mat_rows.append([
            Paragraph("1", style_cell), Paragraph("TPM-0511", style_cell), Paragraph("Materiales PO", style_cell), Paragraph("Carga Consolidada", style_cell_l), Paragraph("0", style_cell), Paragraph("PZA", style_cell), Paragraph("Remesado", style_cell)
        ])
        
    # Fila de Totales
    mat_rows.append([
        Paragraph("<b>TOTAL</b>", style_cell),
        Paragraph("", style_cell),
        Paragraph("", style_cell),
        Paragraph("<b>TOTAL PIEZAS EMBARCADAS EN REMISIÓN:</b>", style_cell_l),
        Paragraph(f"<b>{tot_piezas:,.0f} PZS</b>", style_cell),
        Paragraph("", style_cell),
        Paragraph("<font color='#16A34A'><b>100% CUMPLIDO</b></font>", style_cell)
    ])
    
    t_mat = Table(mat_rows, colWidths=[0.35*inch, 1.0*inch, 1.5*inch, 2.55*inch, 0.8*inch, 0.5*inch, 0.8*inch])
    t_mat.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1F2937')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E5E7EB')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor('#F3F4F6')),
    ]))
    story.append(t_mat)
    story.append(Spacer(1, 0.3*inch))
    
    # 4. Bloque de Firmas de Conformidad
    signs_data = [
        [
            Paragraph("____________________________<br/><b>ENTREGÓ (ALMACÉN SIGRAMA)</b><br/>Nombre y Firma", style_sign_lbl),
            Paragraph("____________________________<br/><b>TRANSPORTISTA / CHOFER</b><br/>Nombre y Firma", style_sign_lbl),
            Paragraph("____________________________<br/><b>RECIBIÓ CONFORME (CLIENTE)</b><br/>Nombre, Firma y Sello Planta", style_sign_lbl)
        ]
    ]
    t_signs = Table(signs_data, colWidths=[2.5*inch, 2.5*inch, 2.5*inch])
    t_signs.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'BOTTOM'),
        ('TOPPADDING', (0,0), (-1,-1), 10),
    ]))
    story.append(t_signs)
    story.append(Spacer(1, 0.15*inch))
    
    # 5. Pie de Página Certificado
    story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor('#EC2024'), spaceBefore=5, spaceAfter=5))
    story.append(Paragraph("<font color='#6B7280' size='7'>Documento de Control y Trazabilidad emitido por SIGRAMA PO Tracker & Remisiones 4.0 • Industria Sigrama S.A. de C.V.</font>", style_sub))
    
    doc.build(story)
    return buffer.getvalue()
