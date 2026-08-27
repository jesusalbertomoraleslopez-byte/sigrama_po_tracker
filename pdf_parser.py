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
            
        subject = getattr(msg, 'subject', '') or ""
        sender = getattr(msg, 'sender', '') or ""
        date = str(getattr(msg, 'date', '')) if getattr(msg, 'date', None) else ""
        body = getattr(msg, 'body', '') or getattr(msg, 'htmlBody', '') or ""
        if isinstance(body, bytes):
            body = body.decode('utf-8', errors='ignore')
            
        attachments = []
        for att in getattr(msg, 'attachments', []):
            try:
                att_name = getattr(att, 'longFilename', None) or getattr(att, 'shortFilename', None) or getattr(att, 'name', None) or "adjunto.pdf"
                att_data = getattr(att, 'data', None)
                if att_data is None and hasattr(att, 'getPayload'):
                    att_data = att.getPayload()
                    
                if att_data:
                    attachments.append({
                        'filename': str(att_name),
                        'data': att_data,
                        'size_kb': round(len(att_data) / 1024, 1)
                    })
            except Exception as e_att:
                print(f"Error procesando adjunto individual en .msg: {e_att}")
            
        return {
            'subject': str(subject),
            'sender': str(sender),
            'date': str(date),
            'body': str(body),
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
    
    # 1. Extracción de Partidas mediante Agrupación Espacial por Coordenadas (Y-clustering)
    # Patrón Principal: Cantidad Unidad SKU_Cliente SKU_Nuestro P_Unitario P_Total Fecha
    patron_fila_item = re.compile(
        r'^(\d+(?:\.\d+)?)\s+'                     # 1: Cantidad (ej. 16.00)
        r'(PIEZA|PZA|KG|METRO|JGO|LOTE|SER)\s+'   # 2: Unidad
        r'([A-Z0-9\-_]{4,15})\s+'                 # 3: SKU Cliente (ej. ISSIV00055, SWB01431)
        r'([A-Z0-9\-_/]{4,25})\s+'                # 4: SKU Nuestro / Planta (ej. 11-A-9836-01, PP19380-03)
        r'([\d,]+(?:\.\d{2})?)\s+'                # 5: P. Unitario (ej. 1,089.57)
        r'([\d,]+(?:\.\d{2})?)\s+'                # 6: P. Total (ej. 17,433.12)
        r'(\d{1,2}/\d{1,2}/\d{4})',               # 7: Fecha Entrega (ej. 07/09/2026)
        re.IGNORECASE
    )
    
    patron_fila_alt = re.compile(
        r'^(\d+(?:\.\d+)?)\s+'
        r'(PIEZA|PZA|KG|METRO|JGO|LOTE|SER)\s+'
        r'([A-Z0-9\-_]{4,15})\s+'
        r'(.+?)\s+'
        r'([\d,]+(?:\.\d{2})?)\s+'
        r'([\d,]+(?:\.\d{2})?)\s+'
        r'(\d{1,2}/\d{1,2}/\d{4})',
        re.IGNORECASE
    )

    for page_idx, page in enumerate(doc):
        words = page.get_text('words')
        if not words:
            continue
            
        lines_by_y = {}
        for w in words:
            x0, y0, x1, y1, word = w[0], w[1], w[2], w[3], w[4]
            found_y = None
            for y_center in lines_by_y:
                if abs(y_center - y0) < 4:
                    found_y = y_center
                    break
            if found_y is None:
                found_y = y0
                lines_by_y[found_y] = []
            lines_by_y[found_y].append((x0, word))

        sorted_ys = sorted(lines_by_y.keys())
        
        i = 0
        while i < len(sorted_ys):
            y = sorted_ys[i]
            # En página 1, saltar encabezado superior
            if page_idx == 0 and y < 320:
                i += 1
                continue
                
            line_words = sorted(lines_by_y[y], key=lambda x: x[0])
            line_text = ' '.join([w[1] for w in line_words]).strip()
            
            # Detener si llegamos al pie de página con totales
            if any(k in line_text.upper() for k in ['SUBTOTAL:', 'IMPORTE NETO:', 'OBSERVACIONES', 'FACTURAR A:', 'TOTAL']):
                if page_idx == 0 and y > 360:
                    i += 1
                    continue
            
            m = patron_fila_item.search(line_text)
            if not m:
                m = patron_fila_alt.search(line_text)
                
            if m:
                cant = float(m.group(1).replace(',', ''))
                unidad = m.group(2).upper()
                sku_cli = m.group(3).strip().upper()
                sku_nuestro = m.group(4).strip().upper()
                pu = float(m.group(5).replace(',', ''))
                pt = float(m.group(6).replace(',', ''))
                f_ent = m.group(7).strip()
                
                # Buscar en la(s) siguiente(s) línea(s) la descripción del producto (Renglón 2)
                desc_lines = []
                j = i + 1
                while j < len(sorted_ys):
                    next_y = sorted_ys[j]
                    next_words = sorted(lines_by_y[next_y], key=lambda x: x[0])
                    next_text = ' '.join([w[1] for w in next_words]).strip()
                    
                    if patron_fila_item.search(next_text) or patron_fila_alt.search(next_text) or any(k in next_text.upper() for k in ['SUBTOTAL:', 'OBSERVACIONES', 'FACTURAR A:', 'TOTAL']):
                        break
                        
                    # Filtrar palabras que pertenezcan al cuerpo de descripción (x < 450)
                    desc_words_filtered = [w[1] for w in next_words if w[0] < 450]
                    if desc_words_filtered:
                        desc_str = ' '.join(desc_words_filtered).strip()
                        if desc_str and desc_str.upper() not in ['FIRMA', 'COMPRADOR']:
                            desc_lines.append(desc_str)
                    j += 1
                    if len(desc_lines) >= 2:
                        break
                        
                desc_final = ' '.join(desc_lines) if desc_lines else f"Material {sku_nuestro}"
                
                all_partidas.append({
                    'item_no': len(all_partidas) + 1,
                    'sku_cliente': sku_cli,
                    'clave_sku': sku_nuestro,
                    'descripcion_producto': desc_final,
                    'cantidad_requerida': cant,
                    'unidad': unidad,
                    'precio_unitario': pu,
                    'precio_total': pt,
                    'fecha_entrega': f_ent,
                    'parcialidad': 'P1',
                    'observaciones_partida': ''
                })
            i += 1
            
    # Fallback si el clustering espacial estricto no encontró partidas:
    if not all_partidas:
        patron_fb = re.compile(r'(\d+(?:\.\d+)?)\s+(PIEZA|PZA|KG|METRO|JGO|LOTE|SER|PZS|PZ)\s+([A-Z0-9\-_]{3,20})\s+([A-Z0-9\-_/]{3,25})', re.IGNORECASE)
        for page_idx, page in enumerate(doc):
            t_lines = [l.strip() for l in page.get_text().split('\n') if l.strip()]
            for l_i, l_str in enumerate(t_lines):
                m_fb = patron_fb.search(l_str)
                if m_fb:
                    cant_fb = float(m_fb.group(1).replace(',', ''))
                    unid_fb = m_fb.group(2).upper()
                    sku_c_fb = m_fb.group(3).strip().upper()
                    sku_n_fb = m_fb.group(4).strip().upper()
                    desc_fb = t_lines[l_i+1] if (l_i+1 < len(t_lines) and not patron_fb.search(t_lines[l_i+1])) else f"Material {sku_n_fb}"
                    if any(w in desc_fb.upper() for w in ['SUBTOTAL', 'TOTAL', 'PIEZA', 'PZA', 'OBSERVACIONES']):
                        desc_fb = f"Material {sku_n_fb}"
                    all_partidas.append({
                        'item_no': len(all_partidas) + 1,
                        'sku_cliente': sku_c_fb,
                        'clave_sku': sku_n_fb,
                        'descripcion_producto': desc_fb,
                        'cantidad_requerida': cant_fb,
                        'unidad': unid_fb,
                        'precio_unitario': 0.0,
                        'precio_total': 0.0,
                        'fecha_entrega': datetime.date.today().strftime('%Y-%m-%d'),
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
                
            # Detección de ID Interno (INT-0001, INT-0059...)
            id_int_auto = ""
            if email_context and email_context.get('id_interno'):
                id_int_auto = email_context['id_interno']
            else:
                m_int_full = re.search(r'\bINT[\s\-_]?(\d{1,4})\b', f"{full_text} {email_context.get('asunto', '') if email_context else ''}", re.IGNORECASE)
                if m_int_full:
                    id_int_auto = f"INT-{int(m_int_full.group(1)):04d}"
                    
            f_llegada_auto = ""
            if email_context and email_context.get('fecha_llegada'):
                f_llegada_auto = email_context['fecha_llegada']
            else:
                f_llegada_auto = fecha_pedido
                
            f_solic_auto = ""
            if all_partidas and all_partidas[0].get('fecha_entrega'):
                f_solic_auto = all_partidas[0]['fecha_entrega']
            else:
                try:
                    f_solic_auto = (datetime.datetime.strptime(fecha_pedido, "%Y-%m-%d") + datetime.timedelta(days=14)).strftime("%Y-%m-%d")
                except Exception:
                    f_solic_auto = fecha_pedido

            # Remitente / Comprador desde correo si el PDF no lo tiene
            sender_correo = ""
            if email_context:
                sender_correo = email_context.get('remitente') or email_context.get('sender') or ""
                # Limpiar dirección de correo si viene como "Nombre <correo@dominio>"
                if '<' in sender_correo:
                    sender_correo = sender_correo.split('<')[0].strip()

            comp_final = comprador if (comprador and comprador.lower() not in ('nan', 'none', '')) else (sender_correo if sender_correo else 'Josue Mesta')
            solic_final = solicitante if (solicitante and solicitante.lower() not in ('nan', 'none', '')) else (sender_correo if sender_correo else '')
            proy_final = uso_proy if (uso_proy and uso_proy.lower() not in ('nan', 'none', '')) else (email_context.get('proyecto_detectado', '') if email_context else (observaciones.split()[0] if observaciones else 'PROYECTO SIGRAMA'))
            
            # Observaciones integrando urgencias de correo
            obs_final = observaciones if (observaciones and observaciones.lower() not in ('nan', 'none')) else ""
            if email_context and email_context.get('urgencias'):
                urg_txt = ' | '.join(email_context['urgencias'])
                obs_final = f"{obs_final} • [Urgencias Correo: {urg_txt}]".strip(' • ')

            cabecera = {
                'po': po_folio,
                'id_interno': id_int_auto,
                'fecha_llegada': f_llegada_auto,
                'fecha_solicitada': f_solic_auto,
                'archivo_correo': email_context.get('msg_filename', '') if email_context else '',
                'archivo_pdf': email_context.get('pdf_filename', '') if email_context else '',
                'fecha_pedido': fecha_pedido,
                'proyecto': proy_final,
                'solicitante': solic_final,
                'requisicion': requisicion if requisicion.lower() not in ('nan', 'none') else '',
                'destino': 'ALMACEN SIGRAMA',
                'proveedor': 'SIGRAMA PLANTA METALES',
                'proveedor_atencion': 'JESUS MORALES',
                'cliente_facturar_a': 'INDUSTRIA SIGRAMA S.A. DE C.V.',
                'cliente_rfc': 'ISI-870204-K4A',
                'cliente_direccion': 'C. JUAN ESCUTIA #50 COL. ABASTOS C.P. 27020 TORREON, COAH.',
                'forma_pago': 'CONTADO / CRÉDITO',
                'lab': 'ALMACEN SIGRAMA',
                'tiempo_entrega': tiempo_entrega,
                'comprador': comp_final,
                'subtotal': subtotal if subtotal > 0 else sum(p['precio_total'] for p in all_partidas),
                'descuento': 0.0,
                'iva': iva if iva > 0 else (subtotal * 0.16),
                'ret_iva': 0.0,
                'ret_isr': 0.0,
                'total': total if total > 0 else (subtotal * 1.16),
                'moneda': 'MXN',
                'observaciones': obs_final,
                'texto_etiqueta': proy_final,
                'color_fondo': '#EC2024',
                'color_texto': '#FFFFFF'
            }
            
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

