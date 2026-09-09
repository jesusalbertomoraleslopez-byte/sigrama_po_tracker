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
    """Descarga en caliente las bases de datos de remisiones y corte/doblez directamente de GitHub o local."""
    import urllib.request
    import subprocess
    from config import get_corte_doblez_dir
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
            
    try:
        from config import SYNC_DB_DIR
        cd_dir = get_corte_doblez_dir()
        if (cd_dir / '.git').exists():
            subprocess.run(['git', '-C', str(cd_dir), 'pull', 'origin', 'main'], capture_output=True, timeout=15)
            ok_any = True
        else:
            corte_url = 'https://raw.githubusercontent.com/jesusalbertomoraleslopez-byte/control-corte-doblez/main/sigrama_database.xlsx'
            for d in set([cd_dir, SYNC_DB_DIR]):
                try:
                    d.mkdir(parents=True, exist_ok=True)
                    urllib.request.urlretrieve(corte_url, d / 'sigrama_database.xlsx')
                    ok_any = True
                except Exception:
                    pass
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

def clean_pronest_piece_name(p):
    s = str(p).strip()
    s = re.sub(r'^\d+[\.\-_]\s*', '', s)
    s = re.sub(r'[\-_]\s*\(\d+[-/]\d+[-/]\d+\).*$', '', s)
    return s.strip()

def normalize_sku(s):
    if not s or pd.isna(s):
        return ""
    return re.sub(r'[^A-Z0-9]', '', str(s).upper())

def sku_matches(target_sku, candidate_piece):
    if not target_sku or not candidate_piece:
        return False
    t_str = str(target_sku).strip()
    c_str = str(candidate_piece).strip()
    t_norm = normalize_sku(t_str)
    c_norm = normalize_sku(c_str)
    if not t_norm or not c_norm:
        return False
    if t_norm == c_norm:
        return True
    c_clean = normalize_sku(clean_pronest_piece_name(c_str))
    if t_norm == c_clean:
        return True
        
    # Extraer primer token si contiene espacios (e.g. '11-A-6014-01 UNPAINTED ...')
    t_first = normalize_sku(t_str.split()[0])
    c_first = normalize_sku(c_str.split()[0])
    if t_first and c_first and t_first == c_first:
        return True
        
    # Coincidencia de prefijo (para strings de al menos 6 caracteres alfanuméricos)
    # Ejemplo: '11-B-9208-01 TRTO' -> '11B920801TRTO' vs '11B920801'
    if len(c_norm) >= 6 and t_norm.startswith(c_norm):
        return True
    if len(t_norm) >= 6 and c_norm.startswith(t_norm):
        return True
        
    return False

def get_tracking_for_po(po_folio, df_partidas, id_interno="", dbs=None):
    """Calcula el estatus de remisión/envío para cada partida y global de una PO dada."""
    if dbs is not None:
        df_rem, df_det, df_tar = dbs
    else:
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
            mask_match = mask_match | (df_det['norm_po'] == po_norm) | (df_det['norm_po'].str.contains(po_norm, regex=False, na=False)) | (df_det['PO'].astype(str).str.contains(po_str, regex=False, na=False))
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
    if df_partidas.empty and not df_det_po.empty and 'SKU' in df_det_po.columns:
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
            
            if not df_det_po.empty and 'SKU' in df_det_po.columns:
                # Coincidencia flexible por SKU Planta, SKU Cliente o Descripción (garantizando retorno booleano)
                match_det = df_det_po[df_det_po['SKU'].apply(lambda p: bool(
                    sku_matches(sku, p) or 
                    (bool(sku_cli) and sku_matches(sku_cli, p)) or
                    (bool(desc_prod) and sku_matches(desc_prod, p))
                ))]
                
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
    """Calcula el resumen de seguimiento para todas las POs incluyendo las 5 métricas clave de la cadena de suministro."""
    summary_list = []
    
    if df_all_pos.empty:
        return pd.DataFrame()
        
    # Pre-cargar bases de datos una sola vez para máxima velocidad (evita leer archivos 75 veces)
    dbs_rem = load_remisiones_databases()
    
    dbs_cd = None
    get_corte_doblez_tracking_for_po = None
    try:
        from corte_doblez_sync import get_corte_doblez_tracking_for_po, load_corte_doblez_databases
        dbs_cd = load_corte_doblez_databases()
    except Exception:
        pass
        
    for _, po_row in df_all_pos.iterrows():
        po_folio = str(po_row.get('po', '')).strip()
        id_int_val = str(po_row.get('id_interno', '')).strip()
        partidas_po = df_all_partidas[df_all_partidas['po'].astype(str).str.strip() == po_folio] if not df_all_partidas.empty else pd.DataFrame()
        
        # 1. Seguimiento en Remisiones y Almacén PT
        tracking = get_tracking_for_po(po_folio, partidas_po, id_interno=id_int_val, dbs=dbs_rem)
        
        # 2. Seguimiento en Taller Planta (Corte y Doblez)
        tot_fab = 0.0
        pct_fab = 0.0
        tot_prog = 0.0
        ofs_str = "Sin OF"
        if get_corte_doblez_tracking_for_po and dbs_cd is not None:
            try:
                # skip_sku_fallback=True evita la búsqueda cara por SKU (solo necesaria en Ficha 360°)
                cd_trk = get_corte_doblez_tracking_for_po(
                    po_folio, partidas_po, id_interno=id_int_val,
                    dbs=dbs_cd, skip_sku_fallback=True, rem_dbs=dbs_rem
                )
                tot_fab = float(cd_trk.get('total_fabricado', cd_trk.get('total_terminado_planta', 0.0)) or 0.0)
                tot_prog = float(cd_trk.get('total_programado', 0.0) or 0.0)
                pct_fab = float(cd_trk.get('porcentaje_fabricacion', 0.0) or 0.0)
                ofs_list = cd_trk.get('ofs_asociadas', [])
                if ofs_list:
                    ofs_str = f"{len(ofs_list)} OF(s)"
            except Exception:
                pass
                
        tot_req = float(tracking.get('total_requerido', 0.0) or 0.0)
        tot_ent = float(tracking.get('total_entarimado', 0.0) or 0.0)
        tot_rem = float(tracking.get('total_remisionado', 0.0) or 0.0)
        tot_pend = max(0.0, tot_req - tot_rem)
        pct_cumpl = round((tot_rem / tot_req * 100.0) if tot_req > 0 else 0.0, 1)
        
        # Estatus 360 enriquecido
        est_gen = str(po_row.get('estatus_general', '')).strip()
        is_canc = est_gen.lower() in ('cancelada', 'cancelado')
        tot_req_orig = tot_req
        if is_canc:
            st_360 = "🚫 Cancelado"
            st_cat = "Cancelada"
            tot_req = 0.0
            tot_pend = 0.0
            pct_cumpl = 0.0
        elif tot_rem >= tot_req and tot_req > 0:
            st_360 = "🟢 Remisionada Total (100%)"
            st_cat = "Remisionada Total"
        elif tot_rem > 0:
            st_360 = f"🔵 Parcial Enviada ({pct_cumpl:.1f}%)"
            st_cat = "Parcial Enviada"
        elif tot_fab >= tot_req and tot_req > 0:
            st_360 = "🟣 Lista para Envío (100% Fab)"
            st_cat = "Lista para Envío"
        elif tot_fab > 0:
            st_360 = f"🟠 En Fabricación ({pct_fab:.1f}%)"
            st_cat = "En Fabricación"
        else:
            st_360 = "⚪ Registrada (En Espera)"
            st_cat = "Registrada"
            
        row_summary = dict(po_row)
        row_summary['articulos_count'] = len(partidas_po) if not partidas_po.empty else 0
        row_summary['piezas_requeridas'] = tot_req
        row_summary['piezas_requeridas_original'] = tot_req_orig
        row_summary['piezas_programadas'] = tot_prog
        row_summary['piezas_fabricadas'] = tot_fab
        row_summary['piezas_entarimadas'] = tot_ent
        row_summary['piezas_remisionadas'] = tot_rem
        row_summary['piezas_pendientes'] = tot_pend
        row_summary['pct_cumplimiento'] = pct_cumpl
        row_summary['pct_fabricacion'] = pct_fab
        row_summary['estatus_remision'] = st_360
        row_summary['canonical_status'] = st_cat
        row_summary['ofs_resumen'] = ofs_str
        row_summary['remisiones_asociadas'] = ', '.join(tracking['remisiones_asociadas']) if tracking.get('remisiones_asociadas') else 'Sin remisión'
        
        summary_list.append(row_summary)
        
    return pd.DataFrame(summary_list)

