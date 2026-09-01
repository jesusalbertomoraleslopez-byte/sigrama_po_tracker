import base64
import io
import re
import datetime
import html
import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
from pathlib import Path

def parse_msg_for_display(msg_bytes_or_path):
    """Extrae toda la información de un archivo .msg para su visualización interactiva."""
    try:
        import extract_msg
        if isinstance(msg_bytes_or_path, (bytes, bytearray)):
            msg = extract_msg.Message(io.BytesIO(msg_bytes_or_path))
        else:
            msg = extract_msg.Message(msg_bytes_or_path)
            
        subject = getattr(msg, 'subject', '') or '(Sin Asunto)'
        sender = getattr(msg, 'sender', '') or 'Desconocido'
        to_recipients = getattr(msg, 'to', '') or ''
        date_val = str(getattr(msg, 'date', '')) if getattr(msg, 'date', None) else ''
        
        body_text = getattr(msg, 'body', '') or ''
        if isinstance(body_text, bytes):
            body_text = body_text.decode('utf-8', errors='ignore')
            
        body_html = getattr(msg, 'htmlBody', None)
        if isinstance(body_html, bytes):
            body_html = body_html.decode('utf-8', errors='ignore')
            
        attachments = []
        for att in getattr(msg, 'attachments', []):
            try:
                att_name = getattr(att, 'longFilename', None) or getattr(att, 'shortFilename', None) or getattr(att, 'name', None) or "adjunto.bin"
                att_data = getattr(att, 'data', None)
                if att_data is None and hasattr(att, 'getPayload'):
                    att_data = att.getPayload()
                    
                if att_data:
                    is_pdf = str(att_name).lower().endswith('.pdf') or (isinstance(att_data, (bytes, bytearray)) and att_data.startswith(b'%PDF'))
                    attachments.append({
                        'name': str(att_name),
                        'data': att_data,
                        'size_kb': round(len(att_data) / 1024, 1),
                        'is_pdf': is_pdf
                    })
            except Exception as e_att:
                print(f"Error extrayendo adjunto: {e_att}")
                
        return {
            'subject': str(subject),
            'sender': str(sender),
            'to': str(to_recipients),
            'date': date_val,
            'body_text': body_text,
            'body_html': body_html,
            'attachments': attachments
        }
    except Exception as e:
        return {
            'error': f"Error al abrir correo: {e}",
            'subject': '',
            'sender': '',
            'to': '',
            'date': '',
            'body_text': '',
            'body_html': None,
            'attachments': []
        }

def render_pdf_embed(pdf_bytes, height=700, key=None):
    """Incrusta un visor interactivo de PDF en pantalla usando iframe y base64."""
    if not pdf_bytes:
        st.warning("No hay contenido de PDF disponible para mostrar.")
        return
        
    b64_pdf = base64.b64encode(pdf_bytes).decode('utf-8')
    pdf_html = f'''
    <iframe 
        src="data:application/pdf;base64,{b64_pdf}#toolbar=1&navpanes=1" 
        width="100%" 
        height="{height}px" 
        type="application/pdf"
        style="border: 1px solid #CBD5E1; border-radius: 8px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1);"
    >
        <p>Tu navegador no soporta visualización directa de PDF. Usa el botón de descarga.</p>
    </iframe>
    '''
    st.markdown(pdf_html, unsafe_allow_html=True)

def render_email_card(msg_bytes, file_name="", key_suffix=""):
    """Renderiza una tarjeta estilo Outlook Web con el correo, remitente, cuerpo y adjuntos."""
    info = parse_msg_for_display(msg_bytes)
    if info.get('error'):
        st.error(info['error'])
        return
        
    subj = info.get('subject', '')
    sender = info.get('sender', '')
    to_rec = info.get('to', '')
    date_val = info.get('date', '')
    atts = info.get('attachments', [])
    body_text = info.get('body_text', '')
    body_html = info.get('body_html')
    
    # Encabezado estilo Outlook
    st.markdown(f"""
    <div style="background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 8px; padding: 18px; margin-bottom: 14px; box-shadow: 0 1px 3px rgba(0,0,0,0.06);">
        <div style="font-family: 'Montserrat', sans-serif; font-size: 16px; font-weight: 700; color: #0F172A; margin-bottom: 10px;">
            📧 {html.escape(subj)}
        </div>
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(240px, 1fr)); gap: 8px; font-size: 12.5px; color: #475569;">
            <div><b>De:</b> <span style="color: #1E293B;">{html.escape(sender)}</span></div>
            <div><b>Fecha:</b> <span style="color: #1E293B;">{html.escape(date_val)}</span></div>
            {f'<div><b>Para:</b> <span style="color: #1E293B;">{html.escape(to_rec)}</span></div>' if to_rec else ''}
            <div><b>Archivo:</b> <code style="color: #EC2024; background: #FEE2E2; padding: 2px 6px; border-radius: 4px;">{html.escape(file_name)}</code></div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # Bloque de Archivos Adjuntos si existen
    if atts:
        st.markdown(f"**📎 Archivos Adjuntos ({len(atts)}):**")
        att_cols = st.columns([1] * min(3, len(atts)))
        for i, att in enumerate(atts):
            a_name = att['name']
            a_data = att['data']
            a_size = att['size_kb']
            a_is_pdf = att['is_pdf']
            icon = "📄" if a_is_pdf else "📎"
            
            with att_cols[i % len(att_cols)]:
                st.download_button(
                    label=f"{icon} {a_name[:24]}... ({a_size} KB)" if len(a_name) > 27 else f"{icon} {a_name} ({a_size} KB)",
                    data=a_data,
                    file_name=a_name,
                    mime="application/pdf" if a_is_pdf else "application/octet-stream",
                    key=f"dl_email_att_{key_suffix}_{i}_{a_name[:10]}",
                    use_container_width=True
                )
        st.markdown("<div style='height: 8px;'></div>", unsafe_allow_html=True)

    # Selector de vista de cuerpo (HTML original vs Texto plano)
    if body_html and len(body_html.strip()) > 50:
        c_mode, _ = st.columns([2, 4])
        with c_mode:
            view_mode = st.radio(
                "Modo de Lectura:",
                ["✨ Formato Original de Outlook (HTML)", "📄 Texto Plano"],
                horizontal=True,
                key=f"vmode_{key_suffix}"
            )
            
        if "HTML" in view_mode:
            components.html(body_html, height=480, scrolling=True)
        else:
            st.text_area("Cuerpo del Correo:", value=body_text, height=350, disabled=True, key=f"txt_body_{key_suffix}")
    else:
        st.markdown("**Cuerpo del Mensaje:**")
        st.text_area("", value=body_text if body_text else "(El correo no contiene cuerpo de texto)", height=300, disabled=True, key=f"txt_body_only_{key_suffix}")

def render_modulo_repositorio():
    """Renderiza el módulo principal completo: '📁 Repositorio de Correos y GitHub'."""
    from db_manager import (
        get_todos_archivos_adjuntos,
        get_contenido_archivo_por_id,
        get_all_pos
    )
    
    st.markdown("""
    <div style="background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%); border-radius: 10px; padding: 22px 28px; margin-bottom: 22px; border-left: 6px solid #EC2024; box-shadow: 0 4px 12px rgba(0,0,0,0.15);">
        <h2 style="color: #FFFFFF; font-family: 'Montserrat', sans-serif; font-size: 22px; font-weight: 800; margin: 0 0 6px 0; text-transform: uppercase; letter-spacing: 0.5px;">
            📁 Repositorio de Correos y GitHub
        </h2>
        <p style="color: #94A3B8; font-size: 13.5px; margin: 0; font-family: 'Questrial', sans-serif;">
            Explorador en vivo de correos <b>.msg de Outlook</b>, órdenes oficiales en <b>PDF</b> y archivos respaldados en la nube de <b>GitHub</b>.
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    df_archivos = get_todos_archivos_adjuntos()
    df_pos = get_all_pos()
    
    t_msg, t_pdf, t_gh = st.tabs([
        "✉️ Visor de Correos (.msg)",
        "📄 Visor de PDFs de Órdenes de Compra",
        "☁️ Explorador de Carpetas de GitHub"
    ])
    
    # ── PESTAÑA 1: VISOR DE CORREOS (.MSG) ──
    with t_msg:
        df_msgs = df_archivos[df_archivos['tipo'] == 'msg'].copy() if not df_archivos.empty else pd.DataFrame()
        
        if df_msgs.empty:
            st.info("ℹ️ No hay correos .msg guardados en la base de datos todavía.")
        else:
            opts = []
            opt_to_id = {}
            for _, r in df_msgs.iterrows():
                lbl = f"[{r.get('id_interno', 'N/A')}] PO {r.get('po', 'N/A')} ➔ {r.get('nombre_archivo', '')}"
                opts.append(lbl)
                opt_to_id[lbl] = r['id']
                
            c_sel, c_dl = st.columns([4, 1.2])
            with c_sel:
                sel_opt = st.selectbox(
                    "Selecciona el Correo a Inspeccionar:",
                    opts,
                    key="sb_msg_explorer"
                )
            
            if sel_opt:
                selected_id = opt_to_id[sel_opt]
                f_name, f_tipo, f_bytes = get_contenido_archivo_por_id(selected_id)
                
                with c_dl:
                    st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
                    if f_bytes:
                        st.download_button(
                            label="📥 Descargar .MSG",
                            data=f_bytes,
                            file_name=f_name,
                            mime="application/vnd.ms-outlook",
                            use_container_width=True,
                            key=f"dl_btn_msg_view_{selected_id}"
                        )
                        
                if f_bytes:
                    render_email_card(f_bytes, file_name=f_name, key_suffix=f"tab1_{selected_id}")
                    
    # ── PESTAÑA 2: VISOR DE PDFS ──
    with t_pdf:
        df_pdfs = df_archivos[df_archivos['tipo'] == 'pdf'].copy() if not df_archivos.empty else pd.DataFrame()
        
        if df_pdfs.empty:
            st.info("ℹ️ No hay PDFs de órdenes de compra registrados en la base de datos.")
        else:
            opts_pdf = []
            opt_to_pdf_id = {}
            for _, r in df_pdfs.iterrows():
                lbl = f"[{r.get('id_interno', 'N/A')}] PO {r.get('po', 'N/A')} ➔ {r.get('nombre_archivo', '')}"
                opts_pdf.append(lbl)
                opt_to_pdf_id[lbl] = r['id']
                
            c_sel_pdf, c_dl_pdf = st.columns([4, 1.2])
            with c_sel_pdf:
                sel_pdf_opt = st.selectbox(
                    "Selecciona la Orden de Compra (PDF) a Visualizar:",
                    opts_pdf,
                    key="sb_pdf_explorer"
                )
                
            if sel_pdf_opt:
                sel_pdf_id = opt_to_pdf_id[sel_pdf_opt]
                pdf_name, _, pdf_bytes = get_contenido_archivo_por_id(sel_pdf_id)
                
                with c_dl_pdf:
                    st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
                    if pdf_bytes:
                        st.download_button(
                            label="📥 Descargar PDF",
                            data=pdf_bytes,
                            file_name=pdf_name,
                            mime="application/pdf",
                            use_container_width=True,
                            key=f"dl_btn_pdf_view_{sel_pdf_id}"
                        )
                        
                if pdf_bytes:
                    render_pdf_embed(pdf_bytes, height=720, key=f"embed_{sel_pdf_id}")
                    
    # ── PESTAÑA 3: EXPLORADOR DE CARPETAS GITHUB ──
    with t_gh:
        repo_url = "https://github.com/jesusalbertomoraleslopez-byte/sigrama_po_tracker"
        folder_url = f"{repo_url}/tree/main/data/correos"
        
        st.markdown(f"""
        <div style="background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 8px; padding: 18px; margin-bottom: 18px;">
            <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 10px;">
                <div>
                    <h4 style="margin: 0; color: #0F172A; font-family: 'Montserrat', sans-serif;">Repositorio Oficial en GitHub</h4>
                    <p style="margin: 4px 0 0 0; color: #64748B; font-size: 13px;">
                        Rama activa: <code style="color:#0F172A; font-weight:700;">main</code> | Carpeta de correos: <code>data/correos/</code>
                    </p>
                </div>
                <div style="display: flex; gap: 10px;">
                    <a href="{folder_url}" target="_blank" style="text-decoration: none;">
                        <button style="background: #24292F; color: #FFFFFF; border: none; padding: 8px 16px; border-radius: 6px; font-weight: 600; font-size: 13px; cursor: pointer; display: flex; align-items: center; gap: 6px;">
                            📁 Ver Carpeta data/correos en GitHub ↗
                        </button>
                    </a>
                    <a href="{repo_url}" target="_blank" style="text-decoration: none;">
                        <button style="background: #EC2024; color: #FFFFFF; border: none; padding: 8px 16px; border-radius: 6px; font-weight: 600; font-size: 13px; cursor: pointer;">
                            🌐 Ver Repositorio Completo ↗
                        </button>
                    </a>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        # Auditoría de Archivos en data/correos
        st.markdown("#### 📑 Inventario de Archivos Respaldados en la Nube")
        if not df_archivos.empty:
            df_disp = df_archivos[['id_interno', 'po', 'tipo', 'nombre_archivo', 'tamano_bytes', 'fecha_subida']].copy()
            df_disp['tamano_kb'] = (df_disp['tamano_bytes'] / 1024.0).round(1)
            df_disp['tipo'] = df_disp['tipo'].str.upper()
            df_disp = df_disp.rename(columns={
                'id_interno': 'ID Interno',
                'po': 'PO / Folio',
                'tipo': 'Tipo',
                'nombre_archivo': 'Nombre de Archivo',
                'tamano_kb': 'Tamaño (KB)',
                'fecha_subida': 'Fecha de Registro'
            })
            
            c_m1, c_m2, c_m3 = st.columns(3)
            with c_m1:
                st.metric("Total Archivos Respaldados", f"{len(df_disp)} archivos")
            with c_m2:
                c_msgs = len(df_disp[df_disp['Tipo'] == 'MSG'])
                st.metric("Correos .MSG Guardados", f"{c_msgs} correos")
            with c_m3:
                c_pdfs = len(df_disp[df_disp['Tipo'] == 'PDF'])
                st.metric("PDFs de Órdenes Guardados", f"{c_pdfs} PDFs")
                
            st.dataframe(
                df_disp[['ID Interno', 'PO / Folio', 'Tipo', 'Nombre de Archivo', 'Tamaño (KB)', 'Fecha de Registro']],
                use_container_width=True,
                hide_index=True
            )
        else:
            st.warning("No se encontraron archivos registrados en la base de datos.")
