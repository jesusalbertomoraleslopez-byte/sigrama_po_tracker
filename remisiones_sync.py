import pandas as pd
from pathlib import Path
from config import (
    get_remisiones_dir,
    ESTATUS_REGISTRADA,
    ESTATUS_EN_PROCESO,
    ESTATUS_PARCIAL,
    ESTATUS_COMPLETADA
)

def normalize_po(po_val):
    """Normaliza folios de PO eliminando prefijos, guiones, barras y espacios."""
    if not po_val or pd.isna(po_val):
        return ""
    s = str(po_val).strip().upper()
    if s.startswith("PO "):
        s = s[3:].strip()
    elif s.startswith("PO-"):
        s = s[3:].strip()
    elif s.startswith("PO"):
        s = s[2:].strip()
    return s.replace("-", "").replace(" ", "").replace("/", "").replace("_", "")

def parse_tarimas_asociadas(raw_val):
    """Parsea listas de tarimas en formato cadena, lista o texto con comas."""
    if not raw_val or pd.isna(raw_val):
        return []
    s = str(raw_val).strip()
    s = s.replace("[", "").replace("]", "").replace("'", "").replace('"', "")
    return [t.strip() for t in s.split(",") if t.strip()]

def sync_live_remisiones_from_github():
    """Descarga en caliente las bases de datos de remisiones directamente de GitHub o local."""
    import urllib.request
    rem_dir = get_remisiones_dir()
    urls = [
        'https://raw.githubusercontent.com/jesusalbertomoraleslopez-byte/remisiones-de-materiales/main/BD_Detalle_Tarimas.xlsx',
        'https://raw.githubusercontent.com/jesusalbertomoraleslopez-byte/remisiones-de-materiales/main/BD_Datos_Generales_Remision.xlsx',
        'https://raw.githubusercontent.com/jesusalbertomoraleslopez-byte/remisiones-de-materiales/main/BD_Tarimas.xlsx'
    ]
    ok_any = False
    for u in urls:
        fname = u.split('/')[-1]
        try:
            target = rem_dir / fname
            urllib.request.urlretrieve(u, target)
            ok_any = True
        except Exception:
            pass
    return ok_any

def load_remisiones_databases():
    """Carga las bases de datos relevantes de la app de Remisiones."""
    rem_dir = get_remisiones_dir()
    
    file_rem = rem_dir / 'BD_Datos_Generales_Remision.xlsx'
    file_det = rem_dir / 'BD_Detalle_Tarimas.xlsx'
    file_tar = rem_dir / 'BD_Tarimas.xlsx'
    
    df_rem = pd.read_excel(file_rem) if file_rem.exists() else pd.DataFrame()
    df_det = pd.read_excel(file_det) if file_det.exists() else pd.DataFrame()
    df_tar = pd.read_excel(file_tar) if file_tar.exists() else pd.DataFrame()
    
    return df_rem, df_det, df_tar

import re

def normalize_sku(s):
    if not s or pd.isna(s):
        return ""
    return re.sub(r'[^A-Z0-9]', '', str(s).upper())

def sku_matches(target_sku, candidate_piece):
    t_norm = normalize_sku(target_sku)
    c_norm = normalize_sku(candidate_piece)
    if not t_norm or not c_norm:
        return False
    return t_norm == c_norm or t_norm in c_norm or c_norm in t_norm

def get_tracking_for_po(po_folio, df_partidas, id_interno=""):
    """Calcula el estatus de remisión/envío para cada partida y global de una PO dada."""
    df_rem, df_det, df_tar = load_remisiones_databases()
    
    po_str = str(po_folio).strip()
    po_norm = normalize_po(po_str)
    id_int_clean = re.sub(r'[^0-9]', '', str(id_interno)) if id_interno else ""
    id_int_num = int(id_int_clean) if id_int_clean else None
    
    # 1. Filtrar registros de Detalle_Tarimas asociados a esta PO (búsqueda normalizada y flexible)
    df_det_po = pd.DataFrame()
    if not df_det.empty:
        mask_match = pd.Series([False] * len(df_det))
        if 'PO' in df_det.columns:
            df_det['norm_po'] = df_det['PO'].apply(normalize_po)
            mask_match = mask_match | (df_det['norm_po'] == po_norm) | (df_det['norm_po'].str.contains(po_norm, na=False)) | (df_det['PO'].astype(str).str.contains(po_str, na=False))
        if 'Proyecto' in df_det.columns and id_int_num is not None:
            pat = rf'\b(?:PO|INT|OC)?\s*0*{id_int_num}\b'
            mask_match = mask_match | df_det['Proyecto'].astype(str).str.contains(pat, regex=True, case=False, na=False)
        df_det_po = df_det[mask_match].copy()
        
    # Mapeo de Tarimas a Folios de Remisión
    tarima_to_remision = {}
    if not df_rem.empty and 'Tarimas_Asociadas' in df_rem.columns:
        for _, r in df_rem.iterrows():
            folio_r = str(r.get('Folio_Remision', r.get('ID_Remision', ''))).strip()
            fecha_s = str(r.get('Fecha_Hora_Salida', '')).strip()
            receptor = str(r.get('Nombre_Receptor', '')).strip()
            raw_t = r.get('Tarimas_Asociadas', '')
            
            tar_list = parse_tarimas_asociadas(raw_t)
            for t_id in tar_list:
                if t_id not in tarima_to_remision:
                    tarima_to_remision[t_id] = []
                tarima_to_remision[t_id].append({
                    'folio_remision': folio_r,
                    'fecha_salida': fecha_s,
                    'receptor': receptor
                })
                
    # 2. Si df_partidas viene vacío, sintetizar partidas desde Detalle_Tarimas si existen
    if df_partidas.empty and not df_det_po.empty:
        synth_list = []
        for idx, (sku_val, g) in enumerate(df_det_po.groupby('SKU'), start=1):
            tot_cant = float(g['Cantidad'].sum())
            first_desc = str(g['Descripcion'].iloc[0]) if 'Descripcion' in g.columns else f"Material {sku_val}"
            synth_list.append({
                'item_no': idx,
                'clave_sku': str(sku_val),
                'descripcion_producto': first_desc,
                'cantidad_requerida': tot_cant,
                'unidad': 'PIEZA',
                'precio_unitario': 0.0,
                'precio_total': 0.0,
                'fecha_entrega': '—',
                'parcialidad': 'P1',
                'observaciones_partida': ''
            })
        df_partidas = pd.DataFrame(synth_list)

    partidas_enriched = []
    total_requerido = 0.0
    total_remisionado = 0.0
    remisiones_asociadas_set = set()
    historial_envios = []
    
    if not df_partidas.empty:
        for _, part in df_partidas.iterrows():
            sku = str(part.get('clave_sku', '')).strip().upper()
            sku_cli = str(part.get('sku_cliente', '')).strip().upper()
            desc_prod = str(part.get('descripcion_producto', '')).strip().upper()
            cant_req = float(part.get('cantidad_requerida', 0) or 0)
            total_requerido += cant_req
            
            cant_entarimada = 0.0
            cant_rem = 0.0
            rem_folios_partida = set()
            
            if not df_det_po.empty:
                # Coincidencia flexible por SKU Planta o SKU Cliente
                match_det = df_det_po[df_det_po['SKU'].apply(lambda p: sku_matches(sku, p) or (sku_cli and sku_matches(sku_cli, p)))]
                
                for _, d_row in match_det.iterrows():
                    t_id = str(d_row.get('ID_Tarima', '')).strip()
                    c_piezas = float(d_row.get('Cantidad', 0) or 0)
                    cant_entarimada += c_piezas
                    
                    rem_infos = tarima_to_remision.get(t_id, [])
                    if rem_infos:
                        cant_rem += c_piezas
                        for info in rem_infos:
                            rem_folios_partida.add(info['folio_remision'])
                            remisiones_asociadas_set.add(info['folio_remision'])
                            historial_envios.append({
                                'SKU': sku,
                                'Descripción': part.get('descripcion_producto', ''),
                                'Cantidad Enviada': c_piezas,
                                'ID Tarima': t_id,
                                'Folio Remisión': info['folio_remision'],
                                'Fecha Salida': info['fecha_salida'],
                                'Receptor': info['receptor']
                            })
                    else:
                        # Tarima armada pero aún no remisionada formalmente
                        historial_envios.append({
                            'SKU': sku,
                            'Descripción': part.get('descripcion_producto', ''),
                            'Cantidad Enviada': c_piezas,
                            'ID Tarima': t_id,
                            'Folio Remisión': 'En Almacén (Entarimado sin remisión)',
                            'Fecha Salida': 'Pendiente de Salida',
                            'Receptor': 'Planta Sigrama'
                        })
                        
            cant_entarimada = max(cant_entarimada, cant_rem)
            total_remisionado += cant_rem
            cant_pend = max(0.0, cant_req - cant_rem)
            pct = (cant_rem / cant_req * 100.0) if cant_req > 0 else 0.0
            
            if pct >= 100.0:
                st_item = ESTATUS_COMPLETADA
            elif pct > 0.0:
                st_item = ESTATUS_PARCIAL
            else:
                st_item = ESTATUS_REGISTRADA
                
            p_dict = dict(part)
            p_dict['cantidad_entarimada'] = cant_entarimada
            p_dict['cantidad_remisionada'] = cant_rem
            p_dict['cantidad_pendiente'] = cant_pend
            p_dict['porcentaje_cumplimiento'] = round(pct, 1)
            p_dict['estatus_partida'] = st_item
            p_dict['remisiones_folios'] = ', '.join(sorted(rem_folios_partida)) if rem_folios_partida else ('Tarima Armada (Sin Remisión)' if cant_entarimada > 0 else 'Sin envío')
            partidas_enriched.append(p_dict)
            
    # 3. Estatus Global de la PO
    pct_global = (total_remisionado / total_requerido * 100.0) if total_requerido > 0 else 0.0
    if pct_global >= 100.0:
        estatus_global = ESTATUS_COMPLETADA
    elif pct_global > 0.0:
        estatus_global = ESTATUS_PARCIAL
    else:
        estatus_global = ESTATUS_REGISTRADA
        
    df_partidas_res = pd.DataFrame(partidas_enriched)
    df_envios_res = pd.DataFrame(historial_envios).drop_duplicates() if historial_envios else pd.DataFrame()
    tot_entarimado = float(df_partidas_res['cantidad_entarimada'].sum()) if (not df_partidas_res.empty and 'cantidad_entarimada' in df_partidas_res.columns) else 0.0
    
    return {
        'po': po_str,
        'total_requerido': total_requerido,
        'total_entarimado': tot_entarimado,
        'total_remisionado': total_remisionado,
        'total_pendiente': max(0.0, total_requerido - total_remisionado),
        'porcentaje_global': round(pct_global, 1),
        'estatus_global': estatus_global,
        'remisiones_asociadas': sorted(list(remisiones_asociadas_set)),
        'df_partidas': df_partidas_res,
        'df_historial_envios': df_envios_res
    }

def get_global_pos_tracking_summary(df_all_pos, df_all_partidas):
    """Calcula el resumen de seguimiento para todas las POs."""
    summary_list = []
    
    if df_all_pos.empty:
        return pd.DataFrame()
        
    for _, po_row in df_all_pos.iterrows():
        po_folio = str(po_row.get('po', '')).strip()
        partidas_po = df_all_partidas[df_all_partidas['po'].astype(str).str.strip() == po_folio] if not df_all_partidas.empty else pd.DataFrame()
        
        tracking = get_tracking_for_po(po_folio, partidas_po)
        
        row_summary = dict(po_row)
        row_summary['articulos_count'] = len(partidas_po) if not partidas_po.empty else 0
        row_summary['piezas_requeridas'] = tracking['total_requerido']
        row_summary['piezas_remisionadas'] = tracking['total_remisionado']
        row_summary['piezas_pendientes'] = tracking['total_pendiente']
        row_summary['pct_cumplimiento'] = tracking['porcentaje_global']
        row_summary['estatus_remision'] = tracking['estatus_global']
        row_summary['remisiones_asociadas'] = ', '.join(tracking['remisiones_asociadas']) if tracking['remisiones_asociadas'] else 'Sin remisión'
        
        summary_list.append(row_summary)
        
    return pd.DataFrame(summary_list)

