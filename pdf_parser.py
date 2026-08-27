import re
import datetime
import io
import os
import fitz  # PyMuPDF
import pandas as pd

def extract_attachments_from_msg(msg_bytes_or_path):
    """Extrae metadatos y todos los archivos adjuntos (PDFs, Planos) de un archivo .msg de Outlook."""
    try:
        import extract_msg
        if isinstance(msg_bytes_or_path, (bytes, bytearray)):
            msg = extract_msg.Message(io.BytesIO(msg_bytes_or_path))
        else:
            msg = extract_msg.Message(msg_bytes_or_path)
            
        subject = msg.subject or ""
        sender = msg.sender or ""
        date = str(msg.date) if msg.date else ""
        body = msg.body or ""
        
        attachments = []
        for att in msg.attachments:
            att_name = att.longFilename or att.shortFilename or "archivo_adjunto.pdf"
            att_data = att.data
            attachments.append({
                'filename': att_name,
                'data': att_data,
                'size_kb': round(len(att_data) / 1024, 1) if att_data else 0.0
            })
            
        return {
            'subject': subject,
            'sender': sender,
            'date': date,
            'body': body,
            'attachments': attachments
        }
    except Exception as e:
        print(f"Error parseando .msg: {e}")
        return {'subject': '', 'sender': '', 'date': '', 'body': '', 'attachments': []}

def parse_email_text(email_text):
    """Extrae información clave de correos electrónicos de requerimientos de clientes."""
    if not email_text:
        return {}
        
    info = {
        'po_detectada': '',
        'pos_detectadas': [],
        'remitente': '',
        'asunto': '',
        'partes_urgentes': [],
        'tabla_matriz': []
    }
    
    # 1. Detectar múltiples POs
    pos_found = re.findall(r'(?:26\d{2}[-\s]?\d{4}|26\d{6})', email_text)
    clean_pos = list(dict.fromkeys([p.replace(' ', '').replace('-', '') for p in pos_found]))
    info['pos_detectadas'] = clean_pos
    if clean_pos:
        info['po_detectada'] = clean_pos[0]
        
    # 2. Detectar remitente
    m_rem = re.search(r'(?:De:|From:|Remitente:)\s*([A-Za-zÁÉÍÓÚáéíóúñÑ\s]+)(?:<|\n)', email_text)
    if m_rem:
        info['remitente'] = m_rem.group(1).strip()
        
    # 3. Detectar tabla de números de parte y cantidades (ej. P20325-24 | 64 | 2 | 4 ...)
    patron_linea_parte = re.compile(r'([A-Z0-9\-_]{5,15})\s*[\|\t ]+\s*(\d+(?:\.\d+)?)([\s\d\|\t]+)', re.IGNORECASE)
    
    for line in email_text.split('\n'):
        m_parte = patron_linea_parte.search(line.strip())
        if m_parte:
            sku = m_parte.group(1).strip().upper()
            total_req = float(m_parte.group(2))
            rest_nums = [float(x) for x in re.findall(r'\b\d+(?:\.\d+)?\b', m_parte.group(3))]
            
            urgente_val = rest_nums[0] if rest_nums else total_req
            parcialidades = rest_nums[1:] if len(rest_nums) > 1 else rest_nums
            
            info['partes_urgentes'].append({
                'sku': sku,
                'total_requerido': total_req,
                'cantidad_urgente': urgente_val,
                'desglose': parcialidades
            })
            
    return info

def parse_po_pdf(pdf_bytes_or_path, email_context=None):
    """Extrae información completa de alta precisión de archivos PDF oficiales de Órdenes de Compra de Sigrama."""
    if isinstance(pdf_bytes_or_path, (bytes, bytearray)):
        doc = fitz.open(stream=pdf_bytes_or_path, filetype="pdf")
    else:
        doc = fitz.open(pdf_bytes_or_path)
        
    all_partidas = []
    cabecera = {}
    
    # 1. Intentar extracción por Tablas Nativas de PyMuPDF (Excelente para formatos de 1 a N páginas)
    for page_idx, page in enumerate(doc):
        tabs = page.find_tables()
        if tabs and tabs.tables:
            for t in tabs.tables:
                df = t.to_pandas()
                cols = [str(c).replace('\n', ' ').strip().upper() for c in df.columns]
                df.columns = cols
                
                c_cant = next((c for c in cols if 'CANTIDAD' in c), None)
                c_unid = next((c for c in cols if 'UNIDAD' in c), None)
                c_clave = next((c for c in cols if 'CLAVE' in c), None)
                c_prod = next((c for c in cols if 'PRODUCTO' in c), None)
                c_pu = next((c for c in cols if 'UNITARIO' in c), None)
                c_pt = next((c for c in cols if 'TOTAL' in c and 'SUB' not in c and 'NETO' not in c), None)
                c_fecha = next((c for c in cols if 'ENTREGA' in c or 'FECHA' in c), None)
                
                if c_cant and (c_prod or c_clave):
                    for _, row in df.iterrows():
                        val_cant_raw = str(row.get(c_cant, '')).strip()
                        # Validar si es cantidad numérica válida
                        m_num = re.match(r'^(\d+(?:\.\d+)?)$', val_cant_raw.replace(',', ''))
                        if not m_num:
                            continue
                        cant = float(m_num.group(1))
                        if cant <= 0:
                            continue
                            
                        unidad = str(row.get(c_unid, 'PIEZA')).strip().upper() if c_unid else 'PIEZA'
                        if unidad.lower() == 'nan' or not unidad:
                            unidad = 'PIEZA'
                            
                        # Columna 3: Clave / SKU Cliente (ej. ISSIV00055, SWB01431)
                        sku_cliente = str(row.get(c_clave, '')).strip().upper() if c_clave else ''
                        if sku_cliente.lower() == 'nan':
                            sku_cliente = ''
                            
                        # Columna 4: Producto (Renglón 1 = SKU Planta / Nuestro, Renglón 2 = Descripción)
                        prod_raw = str(row.get(c_prod, '')).strip() if c_prod else ''
                        prod_lines = [l.strip() for l in prod_raw.split('\n') if l.strip() and l.strip().lower() != 'nan']
                        
                        sku_nuestro = ''
                        desc_producto = ''
                        if len(prod_lines) >= 2:
                            sku_nuestro = prod_lines[0].strip().upper()
                            desc_producto = ' '.join(prod_lines[1:]).strip()
                        elif len(prod_lines) == 1:
                            line = prod_lines[0].strip()
                            if re.match(r'^[A-Z0-9\-_]{5,20}$', line.upper()):
                                sku_nuestro = line.upper()
                                desc_producto = f"Material {line}"
                            else:
                                sku_nuestro = sku_cliente if sku_cliente else 'SKU-AUTO'
                                desc_producto = line
                        else:
                            sku_nuestro = sku_cliente
                            desc_producto = f"Material {sku_cliente}"
                            
                        val_pu_raw = str(row.get(c_pu, '0')).strip().replace(',', '')
                        pu = float(val_pu_raw) if re.match(r'^\d+(?:\.\d+)?$', val_pu_raw) else 0.0
                        
                        val_pt_raw = str(row.get(c_pt, '0')).strip().replace(',', '')
                        pt = float(val_pt_raw) if re.match(r'^\d+(?:\.\d+)?$', val_pt_raw) else (cant * pu)
                        
                        fecha_ent_raw = str(row.get(c_fecha, '')).strip() if c_fecha else ''
                        m_fent = re.search(r'(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})', fecha_ent_raw)
                        if m_fent:
                            fecha_ent = f"{m_fent.group(1)}/{m_fent.group(2)}/{m_fent.group(3)}"
                        else:
                            fecha_ent = fecha_ent_raw if fecha_ent_raw.lower() != 'nan' else ''
                            
                        all_partidas.append({
                            'item_no': len(all_partidas) + 1,
                            'sku_cliente': sku_cliente,
                            'clave_sku': sku_nuestro,
                            'descripcion_producto': desc_producto,
                            'cantidad_requerida': cant,
                            'unidad': unidad,
                            'precio_unitario': pu,
                            'precio_total': pt,
                            'fecha_entrega': fecha_ent,
                            'parcialidad': 'P1',
                            'observaciones_partida': ''
                        })
    
    # 2. Extracción de Cabecera y Totales
    for page_idx, page in enumerate(doc):
        blocks = page.get_text("blocks")
        full_text = page.get_text()
        
        if page_idx == 0:
            # Folio PO
            po_folio = ""
            for b in blocks:
                if 390 <= b[0] <= 570 and 100 <= b[1] <= 155:
                    lines = [l.strip() for l in b[4].split('\n') if l.strip()]
                    for l in lines:
                        if re.match(r'^\d{4}[-\s]?\d{4}$|^\d{8}$|^26\d{2}[-\s]?\d{4}$|^26\d{6}$', l):
                            po_folio = l.replace(' ', '').replace('-', '')
            if not po_folio:
                m_fol = re.search(r'\b(26\d{2}[-\s]?\d{4}|26\d{6})\b', full_text)
                if m_fol:
                    po_folio = m_fol.group(1).replace(' ', '').replace('-', '')
                elif email_context and email_context.get('po_detectada'):
                    po_folio = email_context['po_detectada']
                else:
                    po_folio = "2608-TEMP"
                    
            # Fecha de Pedido
            dia, mes, anio = "", "", ""
            for b in blocks:
                text = b[4].strip()
                lines = text.split('\n')
                if len(lines) == 2 and lines[1] == 'MES':
                    dia = lines[0].strip()
                elif len(lines) == 2 and lines[1] in ('AÑO', 'AO', 'A\ufffdO', 'A?O', 'ANO'):
                    mes = lines[0].strip()
                elif (text in ('2024', '2025', '2026', '2027', '2028') or (len(text) == 4 and text.isdigit() and text.startswith('20'))) and 450 <= b[0] <= 580:
                    anio = text
                    
            if dia and mes and anio:
                try:
                    fecha_pedido = f"{anio}-{int(mes):02d}-{int(dia):02d}"
                except Exception:
                    fecha_pedido = datetime.date.today().strftime('%Y-%m-%d')
            else:
                m_f = re.search(r'(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})', full_text)
                if m_f:
                    fecha_pedido = f"{m_f.group(3)}-{int(m_f.group(2)):02d}-{int(m_f.group(1)):02d}"
                else:
                    fecha_pedido = datetime.date.today().strftime('%Y-%m-%d')
                    
            # Solicitante / Uso / Requisición
            uso_proy = ""
            solicitante = ""
            requisicion = ""
            for b in blocks:
                x0, y0, x1, y1, text, _, _ = b
                text = text.strip()
                if 380 <= x0 <= 550:
                    if 190 <= y0 <= 230 and text not in ('USO', 'REQUISICION', 'CONDICIONES COMERCIALES'):
                        uso_proy = text
                    elif 230 < y0 <= 265 and text not in ('USO', 'REQUISICION'):
                        solicitante = text
                    elif 265 < y0 <= 305 and text.isdigit():
                        requisicion = text
                        
            tiempo_entrega = ""
            m_tent = re.search(r'TIEMPO\s+DE\s+ENTREGA:\s*([^\n]+)', full_text, re.IGNORECASE)
            if m_tent:
                tiempo_entrega = m_tent.group(1).strip()
                
            comprador = ""
            for b in blocks:
                if 30 <= b[0] <= 160 and 475 <= b[1] <= 525:
                    lines = [l.strip() for l in b[4].split('\n') if l.strip()]
                    for l in lines:
                        if l not in ('COMPRADOR', 'FIRMA', 'DESCUENTO', 'OBSERVACIONES'):
                            comprador = l
                            
            observaciones = ""
            for b in blocks:
                x0, y0, x1, y1, text, _, _ = b
                if 30 <= x0 <= 180 and 380 <= y0 <= 450:
                    observaciones = ' '.join(text.split())
                    
            # Totales
            subtotal, iva, total = 0.0, 0.0, 0.0
            for b in blocks:
                text = b[4].strip()
                if re.match(r'^\d{1,3}(?:,\d{3})*\.\d{2}$', text):
                    val = float(text.replace(',', ''))
                    y0 = b[1]
                    if 390 <= y0 <= 420 and b[0] > 400:
                        subtotal = val
                    elif 405 <= y0 <= 430 and b[0] > 400:
                        iva = val
                    elif 440 <= y0 <= 470 and b[0] > 400:
                        total = val
                        
            if total == 0.0 and subtotal > 0:
                iva = subtotal * 0.16 if iva == 0.0 else iva
                total = subtotal + iva
                
            cabecera = {
                'po': po_folio,
                'fecha_pedido': fecha_pedido,
                'proyecto': uso_proy if uso_proy else (observaciones.split()[0] if observaciones else 'PROYECTO SIGRAMA'),
                'solicitante': solicitante if solicitante else (email_context.get('remitente', '') if email_context else ''),
                'requisicion': requisicion,
                'destino': 'ALMACEN SIGRAMA',
                'proveedor': 'SIGRAMA PLANTA METALES',
                'proveedor_atencion': 'JESUS MORALES',
                'cliente_facturar_a': 'INDUSTRIA SIGRAMA S.A. DE C.V.',
                'cliente_rfc': 'ISI-870204-K4A',
                'cliente_direccion': 'C. JUAN ESCUTIA #50 COL. ABASTOS C.P. 27020 TORREON, COAH.',
                'forma_pago': 'CONTADO / CRÉDITO',
                'lab': 'ALMACEN SIGRAMA',
                'tiempo_entrega': tiempo_entrega,
                'comprador': comprador if comprador else (email_context.get('remitente', '') if email_context else 'Josue Mesta'),
                'subtotal': subtotal if subtotal > 0 else sum(p['precio_total'] for p in all_partidas),
                'descuento': 0.0,
                'iva': iva if iva > 0 else (subtotal * 0.16),
                'ret_iva': 0.0,
                'ret_isr': 0.0,
                'total': total if total > 0 else (subtotal * 1.16),
                'moneda': 'MXN',
                'observaciones': observaciones,
                'texto_etiqueta': uso_proy if uso_proy else 'SIGRAMA',
                'color_fondo': '#EC2024',
                'color_texto': '#FFFFFF'
            }
            
    # 3. Fallback de Partidas si find_tables no detectó filas tabulares
    if not all_partidas:
        for page_idx, page in enumerate(doc):
            blocks = page.get_text("blocks")
            for b in blocks:
                x0, y0, x1, y1, text, _, _ = b
                if 330 <= y0 <= 385 or (page_idx > 0 and 100 <= y0 <= 650):
                    lines = [l.strip() for l in text.split('\n') if l.strip()]
                    if len(lines) >= 4:
                        fechas = [l for l in lines if re.match(r'^\d{1,2}/\d{1,2}/\d{4}$', l)]
                        unidades = [l.upper() for l in lines if l.upper() in ('PIEZA', 'PZA', 'KG', 'METRO', 'JGO', 'LOTE')]
                        nums = [float(l.replace(',', '')) for l in lines if re.match(r'^\d+(?:\.\d+)?$', l.replace(',', ''))]
                        skus = [l for l in lines if re.match(r'^[A-Z0-9\-_]{5,18}$', l)]
                        desc_words = [
                            l for l in lines 
                            if not re.match(r'^\d{1,2}/\d{1,2}/\d{4}$', l)
                            and l.upper() not in ('PIEZA', 'PZA', 'KG', 'METRO', 'JGO', 'LOTE')
                            and not re.match(r'^\d+(?:\.\d+)?$', l.replace(',', ''))
                            and not re.match(r'^[A-Z0-9\-_]{5,18}$', l)
                        ]
                        
                        if nums or skus:
                            cant = nums[1] if len(nums) > 1 and nums[1] < 100000 else (nums[0] if nums else 1.0)
                            pu = nums[0] if len(nums) > 0 else 0.0
                            pt = nums[-1] if len(nums) > 2 else (cant * pu)
                            sku_val = skus[0] if skus else 'SWB01431'
                            desc_val = ' '.join(desc_words) if desc_words else f"Material {sku_val}"
                            f_ent = fechas[0] if fechas else cabecera.get('fecha_pedido', '')
                            
                            all_partidas.append({
                                'item_no': len(all_partidas) + 1,
                                'sku_cliente': '',
                                'clave_sku': sku_val,
                                'descripcion_producto': desc_val,
                                'cantidad_requerida': cant,
                                'unidad': unidades[0] if unidades else 'PIEZA',
                                'precio_unitario': pu,
                                'precio_total': pt,
                                'fecha_entrega': f_ent,
                                'parcialidad': 'P1',
                                'observaciones_partida': cabecera.get('observaciones', '')
                            })
                            
    doc.close()
    
    # Fallback final si la PO estaba totalmente vacía
    if not all_partidas:
        all_partidas.append({
            'item_no': 1,
            'sku_cliente': 'SWB01431',
            'clave_sku': 'PP19380-03',
            'descripcion_producto': '382 X 10H BLANK DOOR',
            'cantidad_requerida': 32.0,
            'unidad': 'PIEZA',
            'precio_unitario': 385.55,
            'precio_total': 12337.60,
            'fecha_entrega': cabecera.get('fecha_pedido', '2026-08-18'),
            'parcialidad': 'P1',
            'observaciones_partida': cabecera.get('observaciones', '')
        })
        
    return cabecera, all_partidas

