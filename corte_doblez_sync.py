import pandas as pd
import os
import re
from pathlib import Path
from config import normalize_po, get_corte_doblez_dir, is_historical_completed, get_historical_qty_override

def clean_pronest_piece_name(p):
    s = str(p).strip()
    # Eliminar prefijo de número de ítem de Pronest: '20.-', '01.-', '02.- ', '1.', etc.
    s = re.sub(r'^\d+[\.\-_\s]+', '', s).strip()
    # Eliminar sufijo de calibre: ' CAL. 12', ' CAL 14', etc.
    s = re.sub(r'\s+CAL\b.*$', '', s, flags=re.IGNORECASE).strip()
    # Eliminar sufijo de fecha/anidado: '-(10-08-26)', ' (10-08-26)', etc.
    s = re.sub(r'[\-_]\s*\(\d+[-/]\d+[-/]\d+\).*$', '', s).strip()
    return s.strip()

def normalize_sku(s):
    if not s or pd.isna(s):
        return ""
    return re.sub(r'[^A-Z0-9]', '', str(s).upper())

def sku_matches(s1, s2):
    """Compara dos SKUs de forma exacta, tolerando guiones, espacios y prefijos de anidado de Pronest."""
    if not s1 or not s2:
        return False
    norm1 = normalize_sku(s1)
    norm2 = normalize_sku(s2)
    if not norm1 or not norm2:
        return False
    if norm1 == norm2:
        return True
    
    # Limpiar posibles prefijos/sufijos de Pronest en s1 y s2
    c1 = normalize_sku(clean_pronest_piece_name(s1))
    c2 = normalize_sku(clean_pronest_piece_name(s2))
    if c1 == c2 or norm1 == c2 or c1 == norm2:
        return True
        
    return False

_CD_CACHE = {
    "mtime": None,
    "dfs": None
}

def load_corte_doblez_databases():
    """Carga las bases de datos relevantes de la app de Corte y Doblez con caché en memoria invalidado por mtime."""
    try:
        from remisiones_sync import sync_live_remisiones_from_github
        sync_live_remisiones_from_github(force=False)
    except Exception:
        pass
    cd_dir = get_corte_doblez_dir()
    excel_path = cd_dir / 'sigrama_database.xlsx'
    
    if not excel_path.exists():
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
        
    mtime = excel_path.stat().st_mtime
    if _CD_CACHE["mtime"] == mtime and _CD_CACHE["dfs"] is not None:
        return _CD_CACHE["dfs"]
        
    try:
        xl = pd.ExcelFile(excel_path)
        df_ord = xl.parse('ordenes') if 'ordenes' in xl.sheet_names else pd.DataFrame()
        df_pie = xl.parse('piezas') if 'piezas' in xl.sheet_names else pd.DataFrame()
        df_ava = xl.parse('avances') if 'avances' in xl.sheet_names else pd.DataFrame()
        df_tar = xl.parse('tarimas') if 'tarimas' in xl.sheet_names else pd.DataFrame()
        df_nid = xl.parse('nidos') if 'nidos' in xl.sheet_names else pd.DataFrame()
        
        if not df_ord.empty and 'of_number' in df_ord.columns:
            df_ord['_of_str'] = df_ord['of_number'].fillna('').astype(str).str.strip()
            df_ord['_po_str'] = df_ord.get('po', pd.Series([''] * len(df_ord))).fillna('').astype(str).str.strip()
            df_ord['_norm_po'] = df_ord['_po_str'].apply(normalize_po)
            df_ord['_norm_of'] = df_ord['_of_str'].apply(normalize_po)
            df_ord['_comb_txt'] = (df_ord['_of_str'] + ' ' + df_ord['_po_str'] + ' ' + df_ord.get('proyecto', pd.Series([''] * len(df_ord))).fillna('').astype(str).str.strip()).str.upper()
        if not df_pie.empty and 'no_pieza' in df_pie.columns:
            df_pie['_norm_pieza'] = df_pie['no_pieza'].apply(lambda x: normalize_sku(clean_pronest_piece_name(x)))
        if not df_ava.empty and 'no_pieza' in df_ava.columns:
            df_ava['_norm_pieza'] = df_ava['no_pieza'].apply(lambda x: normalize_sku(clean_pronest_piece_name(x)))
            df_ava['_area_lc'] = df_ava['area'].astype(str).str.lower()
        if not df_tar.empty and 'no_pieza' in df_tar.columns:
            df_tar['_norm_pieza'] = df_tar['no_pieza'].apply(lambda x: normalize_sku(clean_pronest_piece_name(x)))
            
        _CD_CACHE["mtime"] = mtime
        _CD_CACHE["dfs"] = (df_ord, df_pie, df_ava, df_tar, df_nid)
        return df_ord, df_pie, df_ava, df_tar, df_nid
    except Exception as e:
        print(f'Error loading Corte y Doblez DB: {e}')
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

def get_corte_doblez_tracking_for_po(po_folio, df_partidas, id_interno="", dbs=None, skip_sku_fallback=False, rem_dbs=None, rem_trk_info=None):
    """Calcula el avance de manufactura en planta (Corte, Doblez, Liberado) para una PO.
    
    Args:
        skip_sku_fallback: Si True, omite el fallback lento de búsqueda por SKU (usar en resúmenes globales).
        rem_dbs: Bases de datos de remisiones pre-cargadas (para evitar releer archivos en el fallback).
        rem_trk_info: Resultado pre-calculado de get_tracking_for_po (evita recalcular remisiones).
    """
    if dbs is not None:
        df_ord, df_pie, df_ava, df_tar, df_nid = dbs
    else:
        df_ord, df_pie, df_ava, df_tar, df_nid = load_corte_doblez_databases()
    
    po_str = str(po_folio).strip()
    po_norm = normalize_po(po_str)
    
    # Manejo de Órdenes Históricas (completadas y entregadas al 100% previas a sistemas)
    if is_historical_completed(id_interno=id_interno, po=po_str):
        override_qty = get_historical_qty_override(id_interno=id_interno, po=po_str)
        partidas_cd = []
        total_req_cd = 0.0
        if df_partidas is not None and not df_partidas.empty:
            orig_sum = float(df_partidas['cantidad_requerida'].sum() or 0)
            for _, part in df_partidas.iterrows():
                cant_req = float(part.get('cantidad_requerida', 0) or 0)
                if override_qty is not None and orig_sum > 0:
                    cant_req_adj = round((cant_req / orig_sum) * override_qty, 2)
                else:
                    cant_req_adj = cant_req
                total_req_cd += cant_req_adj
                p_res = dict(part)
                p_res['cantidad_requerida'] = cant_req_adj
                p_res['piezas_programadas'] = cant_req_adj
                p_res['piezas_cortadas'] = cant_req_adj
                p_res['piezas_dobladas'] = cant_req_adj
                p_res['piezas_terminadas_planta'] = cant_req_adj
                p_res['pct_avance_fabricacion'] = 100.0
                p_res['ofs_asociadas'] = 'Fabricación Histórica (100% Terminada)'
                partidas_cd.append(p_res)
        if override_qty is not None:
            total_req_cd = float(override_qty)
        return {
            'po': po_str,
            'matched_ofs': ['Fabricación Histórica Validada'],
            'ofs_asociadas': ['Fabricación Histórica Validada'],
            'total_programado': total_req_cd,
            'total_cortado': total_req_cd,
            'total_doblado': total_req_cd,
            'total_terminado_planta': total_req_cd,
            'total_fabricado': total_req_cd,
            'pct_global_fabricacion': 100.0,
            'porcentaje_fabricacion': 100.0,
            'df_partidas_cd': pd.DataFrame(partidas_cd),
            'df_ofs': pd.DataFrame({'OF': ['Fabricación Histórica Validada']}),
            'df_laminas': pd.DataFrame(),
            'total_laminas': 0,
            'df_nidos': pd.DataFrame(),
            'df_pie_po': pd.DataFrame(),
            'df_ava_po': pd.DataFrame(),
            'df_ord_po': pd.DataFrame()
        }
        
    id_int_clean = re.sub(r'[^0-9]', '', str(id_interno)) if id_interno else ""
    id_int_num = int(id_int_clean) if id_int_clean else None
    id_pat = re.compile(rf'\b(PO|INT|OC)?\s*0*{id_int_num}\b') if id_int_num is not None else None
    
    # 1. Encontrar OFs relacionadas con esta PO
    matched_ofs = set()
    if not df_ord.empty and 'of_number' in df_ord.columns:
        has_pre = '_of_str' in df_ord.columns
        for _, o_row in df_ord.iterrows():
            if has_pre:
                of_num = o_row['_of_str']
                po_field_norm = o_row['_norm_po']
                of_norm = o_row['_norm_of']
                comb_txt = o_row['_comb_txt']
            else:
                of_num = str(o_row.get('of_number', '')).strip()
                po_field_norm = normalize_po(str(o_row.get('po', '')).strip())
                of_norm = normalize_po(of_num)
                comb_txt = f"{of_num} {o_row.get('po', '')} {o_row.get('proyecto', '')}".upper()
            
            # Coincidencia por PO directa
            m_by_po = (po_field_norm and po_field_norm == po_norm) or \
                      (po_norm and po_norm in of_norm) or \
                      (po_str and po_str in of_num)
                      
            # Coincidencia por ID Interno flexible (ej. 47 en PO 047, INT-0047, etc.)
            # - Se ignora el número de la propia OF ("OF 00032-...") para no confundirlo con INT-0032.
            # - Si la OF ya declara una PO explícita distinta, no se reasigna por ID interno.
            comb_txt_id = re.sub(r'\bOF\s*0*\d+', ' ', comb_txt)
            po_explicit_other = bool(po_field_norm) and po_field_norm != po_norm
            m_by_id = bool(id_pat and not po_explicit_other and id_pat.search(comb_txt_id))
                    
            # Inteligencia Operativa: Caso específico OF 83 (P4-P5-P6 Cal. 12 CLOUD)
            if '00083' in of_num:
                if po_norm in ('26083425', 'PO047', 'INT0047', '047', '47'):
                    m_by_po = True
                elif po_norm in ('26083235', 'PO055', 'INT0055', '055', '55'):
                    m_by_po = False
                    
            if m_by_po or m_by_id:
                matched_ofs.add(of_num)
                
    # Pre-calcular normas de todas las partidas de esta PO una sola vez (O(M))
    partida_norms_list = []
    all_po_norms = set()
    if df_partidas is not None and not df_partidas.empty:
        for _, part in df_partidas.iterrows():
            sku = str(part.get('clave_sku', '')).strip().upper()
            sku_cli = str(part.get('sku_cliente', '')).strip().upper()
            v_norms = {
                n for n in [
                    normalize_sku(sku),
                    normalize_sku(sku_cli),
                    normalize_sku(clean_pronest_piece_name(sku)),
                    normalize_sku(clean_pronest_piece_name(sku_cli)),
                ] if n
            }
            partida_norms_list.append((part, sku, sku_cli, v_norms))
            all_po_norms.update(v_norms)

    # 1.1 Si no hubo coincidencia por PO, rastreo por SKU (solo cuando no viene de resumen global)
    if not matched_ofs and all_po_norms and not df_pie.empty and not skip_sku_fallback:
        if rem_trk_info is None:
            from remisiones_sync import get_tracking_for_po as get_rem_tracking
            rem_trk_info = get_rem_tracking(po_str, df_partidas, id_interno=id_interno, dbs=rem_dbs)
        has_remisiones = float(rem_trk_info.get('total_remisionado', 0) or 0) > 0 or len(rem_trk_info.get('remisiones_asociadas', [])) > 0
        
        if has_remisiones and '_norm_pieza' in df_pie.columns:
            m_pie_fb = df_pie[df_pie['_norm_pieza'].isin(all_po_norms)]
            for of_n in m_pie_fb['of_number'].dropna().unique():
                s_of = str(of_n).strip()
                if s_of:
                    matched_ofs.add(s_of)
                
    # 2. Filtrar piezas, avances y nidos de esas OFs
    df_pie_po = df_pie[df_pie['of_number'].isin(matched_ofs)] if not df_pie.empty and matched_ofs else pd.DataFrame()
    df_ava_po = df_ava[df_ava['of_number'].isin(matched_ofs)] if not df_ava.empty and matched_ofs else pd.DataFrame()
    df_tar_po = df_tar[df_tar['of_number'].isin(matched_ofs)] if not df_tar.empty and matched_ofs else pd.DataFrame()
    df_nid_po = df_nid[df_nid['of_number'].isin(matched_ofs)] if not df_nid.empty and matched_ofs else pd.DataFrame()
    
    # Pre-normalizar columnas para matching vectorizado solo si aún no existen en caché
    if not df_pie_po.empty and 'no_pieza' in df_pie_po.columns and '_norm_pieza' not in df_pie_po.columns:
        df_pie_po = df_pie_po.copy()
        df_pie_po['_norm_pieza'] = df_pie_po['no_pieza'].apply(lambda x: normalize_sku(clean_pronest_piece_name(x)))
    if not df_ava_po.empty and 'no_pieza' in df_ava_po.columns and '_norm_pieza' not in df_ava_po.columns:
        df_ava_po = df_ava_po.copy()
        df_ava_po['_norm_pieza'] = df_ava_po['no_pieza'].apply(lambda x: normalize_sku(clean_pronest_piece_name(x)))
        df_ava_po['_area_lc'] = df_ava_po['area'].astype(str).str.lower()
    if not df_tar_po.empty and 'no_pieza' in df_tar_po.columns and '_norm_pieza' not in df_tar_po.columns:
        df_tar_po = df_tar_po.copy()
        df_tar_po['_norm_pieza'] = df_tar_po['no_pieza'].apply(lambda x: normalize_sku(clean_pronest_piece_name(x)))

    # Filtrar a los nidos específicos donde se anidaron estas piezas (usando all_po_norms en O(1))
    if not df_pie_po.empty and 'nido' in df_pie_po.columns and not df_nid_po.empty and all_po_norms:
        m_nested_nidos = df_pie_po[df_pie_po['_norm_pieza'].isin(all_po_norms)]
        if not m_nested_nidos.empty:
            valid_nidos = set(m_nested_nidos['nido'].dropna().unique())
            df_nid_calc = df_nid_po[df_nid_po['nido'].isin(valid_nidos)]
            if not df_nid_calc.empty:
                df_nid_po = df_nid_calc
    
    # Resumen de Láminas Utilizadas
    laminas_summary = []
    total_laminas = 0.0
    if not df_nid_po.empty and 'hojas' in df_nid_po.columns:
        def extract_mat(of_name):
            s = str(of_name).upper().replace(' ', '')
            if 'CAL.10' in s or 'CAL10' in s or '10GA' in s:
                return 'Lámina Galvanizada Cal. 10'
            elif 'CAL.12' in s or 'CAL12' in s or '12GA' in s:
                return 'Lámina Galvanizada Cal. 12'
            elif 'CAL.14' in s or 'CAL14' in s or '14GA' in s:
                return 'Lámina Galvanizada Cal. 14'
            elif 'CAL.16' in s or 'CAL16' in s or '16GA' in s:
                return 'Lámina Galvanizada Cal. 16'
            elif 'DECP' in s:
                return 'Lámina Decapada'
            elif 'INOX' in s:
                return 'Lámina Acero Inoxidable'
            return 'Lámina Galvanizada'
            
        df_nid_po_copy = df_nid_po.copy()
        df_nid_po_copy['material_calibre'] = df_nid_po_copy['of_number'].apply(extract_mat)
        for mat, grp in df_nid_po_copy.groupby('material_calibre'):
            c_hojas = float(grp['hojas'].sum())
            total_laminas += c_hojas
            laminas_summary.append({
                'material': mat,
                'hojas_utilizadas': c_hojas,
                'nidos_cortados': len(grp)
            })
            
    # Consulta cruzada con Remisiones (reutiliza rem_trk_info si ya fue calculado)
    if rem_trk_info is None:
        from remisiones_sync import get_tracking_for_po as get_rem_tracking
        rem_trk_info = get_rem_tracking(po_str, df_partidas, id_interno=id_interno, dbs=rem_dbs)
    df_partidas_rem = rem_trk_info.get('df_partidas', pd.DataFrame())
    remisiones_asoc = rem_trk_info.get('remisiones_asociadas', [])
    
    rem_map = {}
    if not df_partidas_rem.empty:
        for _, r_row in df_partidas_rem.iterrows():
            sku_k = str(r_row.get('clave_sku', '')).strip().upper()
            sku_c_k = str(r_row.get('sku_cliente', '')).strip().upper()
            c_rem_val = float(r_row.get('cantidad_remisionada', 0.0) or 0.0)
            if sku_k:
                rem_map[sku_k] = max(rem_map.get(sku_k, 0.0), c_rem_val)
            if sku_c_k:
                rem_map[sku_c_k] = max(rem_map.get(sku_c_k, 0.0), c_rem_val)
                
    # Pre-indexar cantidades y OFs por _norm_pieza en diccionarios O(1)
    pie_qty = {}
    pie_ofs = {}
    if not df_pie_po.empty and '_norm_pieza' in df_pie_po.columns:
        for _, r in df_pie_po.iterrows():
            np_k = r['_norm_pieza']
            if not np_k:
                continue
            pie_qty[np_k] = pie_qty.get(np_k, 0.0) + float(r.get('cantidad', 0) or 0.0)
            of_v = str(r.get('of_number', '')).strip()
            if of_v:
                pie_ofs.setdefault(np_k, set()).add(of_v)

    ava_qty = {}
    ava_ofs = {}
    if not df_ava_po.empty and '_norm_pieza' in df_ava_po.columns:
        for _, r in df_ava_po.iterrows():
            np_k = r['_norm_pieza']
            if not np_k:
                continue
            ar = r['_area_lc']
            q = float(r.get('cantidad', 0) or 0.0)
            d_ar = ava_qty.setdefault(np_k, {})
            d_ar[ar] = d_ar.get(ar, 0.0) + q
            of_v = str(r.get('of_number', '')).strip()
            if of_v:
                ava_ofs.setdefault(np_k, set()).add(of_v)

    tar_qty = {}
    tar_ofs = {}
    if not df_tar_po.empty and '_norm_pieza' in df_tar_po.columns:
        for _, r in df_tar_po.iterrows():
            np_k = r['_norm_pieza']
            if not np_k:
                continue
            tar_qty[np_k] = tar_qty.get(np_k, 0.0) + float(r.get('cantidad', 0) or 0.0)
            of_v = str(r.get('of_number', '')).strip()
            if of_v:
                tar_ofs.setdefault(np_k, set()).add(of_v)

    partidas_cd = []
    total_req_cd = 0.0
    total_programado = 0.0
    total_cortado = 0.0
    total_doblado = 0.0
    total_terminado = 0.0
    
    for part, sku, sku_cli, valid_norms in partida_norms_list:
        cant_req = float(part.get('cantidad_requerida', 0) or 0)
        total_req_cd += cant_req
        
        c_prog = sum(pie_qty.get(n, 0.0) for n in valid_norms)
        c_corte = 0.0
        c_doblez = 0.0
        c_rebabeo = 0.0
        c_liberado = 0.0
        for n in valid_norms:
            d_ar = ava_qty.get(n)
            if d_ar:
                c_corte += d_ar.get('corte', 0.0)
                c_doblez += d_ar.get('doblez', 0.0)
                c_rebabeo += d_ar.get('rebabeo', 0.0)
                c_liberado += d_ar.get('liberado', 0.0) + d_ar.get('empaque', 0.0)
                
        if c_liberado == 0.0:
            c_liberado = sum(tar_qty.get(n, 0.0) for n in valid_norms)
            
        # REGLA DE INTELIGENCIA OPERATIVA: Si ya está entarimada en Almacén PT o remisionada, ya fue fabricada en planta
        c_rem_part = max(
            rem_map.get(sku, 0.0),
            rem_map.get(sku_cli, 0.0),
            float(part.get('cantidad_remisionada', 0) or 0),
            float(part.get('cantidad_entarimada', 0) or 0)
        )
        if c_rem_part > 0:
            c_corte = max(c_corte, c_rem_part)
            c_doblez = max(c_doblez, c_rem_part)
            c_terminado_real = max(c_liberado, c_doblez, c_rebabeo, c_corte, c_rem_part)
        else:
            c_terminado_real = max(c_liberado, c_doblez, c_rebabeo, c_corte)
        
        prog_part = c_prog if c_prog > 0 else c_terminado_real
        if cant_req > 0:
            total_programado += min(cant_req, prog_part)
            total_cortado += min(cant_req, c_corte)
            total_doblado += min(cant_req, c_doblez)
            total_terminado += min(cant_req, c_terminado_real)
            pct_cd = (c_terminado_real / cant_req * 100.0)
        else:
            total_programado += prog_part
            total_cortado += c_corte
            total_doblado += c_doblez
            total_terminado += c_terminado_real
            pct_cd = 100.0 if c_terminado_real > 0 else 0.0
        
        piece_ofs_set = set()
        for n in valid_norms:
            if n in pie_ofs:
                piece_ofs_set.update(pie_ofs[n])
        if not piece_ofs_set:
            for n in valid_norms:
                if n in ava_ofs:
                    piece_ofs_set.update(ava_ofs[n])
        if not piece_ofs_set:
            for n in valid_norms:
                if n in tar_ofs:
                    piece_ofs_set.update(tar_ofs[n])

        if piece_ofs_set:
            ofs_tag = ', '.join(sorted(piece_ofs_set))
        elif not matched_ofs and remisiones_asoc:
            ofs_tag = f"Fabricado (Remisión {', '.join(remisiones_asoc)})"
        elif len(matched_ofs) == 1:
            ofs_tag = list(matched_ofs)[0]
        else:
            ofs_tag = 'Por programar OF'
            
        p_res = dict(part)
        p_res['piezas_programadas'] = prog_part
        p_res['piezas_cortadas'] = c_corte
        p_res['piezas_dobladas'] = c_doblez
        p_res['piezas_terminadas_planta'] = c_terminado_real
        p_res['pct_avance_fabricacion'] = round(min(100.0, pct_cd), 1)
        p_res['ofs_asociadas'] = ofs_tag
        partidas_cd.append(p_res)
            
    pct_global_cd = (total_terminado / total_req_cd * 100.0) if total_req_cd > 0 else 0.0
    
    ofs_final_list = sorted(list(matched_ofs))
    if not ofs_final_list and remisiones_asoc:
        ofs_final_list = [f"Fabricación Validada en Planta (Remisión {', '.join(remisiones_asoc)})"]
        
    return {
        'po': po_str,
        'matched_ofs': ofs_final_list,
        'ofs_asociadas': ofs_final_list,
        'total_programado': total_programado,
        'total_cortado': total_cortado,
        'total_doblado': total_doblado,
        'total_terminado_planta': total_terminado,
        'total_fabricado': total_terminado,
        'pct_global_fabricacion': round(min(100.0, pct_global_cd), 1),
        'porcentaje_fabricacion': round(min(100.0, pct_global_cd), 1),
        'df_partidas_cd': pd.DataFrame(partidas_cd),
        'df_ofs': pd.DataFrame({'OF': ofs_final_list}) if ofs_final_list else pd.DataFrame(),
        'df_laminas': pd.DataFrame(laminas_summary),
        'total_laminas': total_laminas,
        'df_nidos': df_nid_po,
        'df_pie_po': df_pie_po,
        'df_ava_po': df_ava_po,
        'df_ord_po': df_ord[df_ord['of_number'].isin(matched_ofs)] if (not df_ord.empty and matched_ofs) else pd.DataFrame()
    }

def get_integrated_360_summary(df_all_pos, df_all_partidas):
    """Genera la Matriz de Control 360 combinando Requerimiento Maestro + Corte y Doblez + Remisiones."""
    from remisiones_sync import get_tracking_for_po, load_remisiones_databases
    
    matrix = []
    if df_all_pos.empty:
        return pd.DataFrame()
        
    dbs_rem = load_remisiones_databases()
    dbs_cd = load_corte_doblez_databases()
    
    partidas_by_po = {}
    if not df_all_partidas.empty and 'po' in df_all_partidas.columns:
        df_p_tmp = df_all_partidas.copy()
        df_p_tmp['_po_key'] = df_p_tmp['po'].astype(str).str.strip()
        for k, grp in df_p_tmp.groupby('_po_key', sort=False):
            partidas_by_po[k] = grp.drop(columns=['_po_key'])
            
    empty_df = pd.DataFrame()
    for _, po_row in df_all_pos.iterrows():
        po_folio = str(po_row.get('po', '')).strip()
        id_int_val = str(po_row.get('id_interno', '')).strip()
        partidas_po = partidas_by_po.get(po_folio, empty_df)
        
        # 1. Trazabilidad con Remisiones
        trk_rem = get_tracking_for_po(po_folio, partidas_po, id_interno=id_int_val, dbs=dbs_rem)
        
        # 2. Trazabilidad con Corte y Doblez (reutilizando trk_rem)
        trk_cd = get_corte_doblez_tracking_for_po(
            po_folio, trk_rem['df_partidas'], id_interno=id_int_val,
            dbs=dbs_cd, skip_sku_fallback=True, rem_dbs=dbs_rem, rem_trk_info=trk_rem
        )
        
        tot_req = trk_rem['total_requerido']
        tot_fab = trk_cd['total_terminado_planta']
        tot_env = trk_rem['total_remisionado']
        
        # Estatus combinado
        est_gen = str(po_row.get('estatus_general', '')).strip()
        is_canc = est_gen.lower() in ('cancelada', 'cancelado')
        
        tot_req = 0.0 if is_canc else trk_rem['total_requerido']
        tot_fab = trk_cd['total_terminado_planta']
        tot_env = trk_rem['total_remisionado']
        
        if is_canc:
            estatus_360 = '🚫 Cancelado'
        elif is_historical_completed(id_interno=id_int_val, po=po_folio):
            estatus_360 = '🟢 Remisionada Total (100%)'
            tot_fab = tot_req
            tot_env = tot_req
        elif tot_env >= tot_req and tot_req > 0:
            estatus_360 = '🟢 Remisionada Total (100%)'
        elif tot_env > 0:
            estatus_360 = '🟡 En Envíos Parciales'
        elif tot_fab >= tot_req and tot_req > 0:
            estatus_360 = '🔵 Fabricada 100% (En Almacén)'
        elif tot_fab > 0:
            estatus_360 = '🟠 En Fabricación (Taller)'
        else:
            estatus_360 = '⚪ Registrada / Por Fabricar'
            
        row = dict(po_row)
        row['piezas_requeridas'] = tot_req
        row['piezas_requeridas_original'] = trk_rem['total_requerido']
        row['piezas_fabricadas'] = tot_fab
        row['pct_fabricacion'] = 0.0 if is_canc else trk_cd['pct_global_fabricacion']
        row['piezas_remisionadas'] = tot_env
        row['pct_remision'] = 0.0 if is_canc else trk_rem['porcentaje_global']
        row['piezas_pendientes_fab'] = 0.0 if is_canc else max(0.0, tot_req - tot_fab)
        row['piezas_pendientes_env'] = 0.0 if is_canc else max(0.0, tot_req - tot_env)
        row['ofs_asociadas'] = ', '.join(trk_cd['matched_ofs']) if trk_cd['matched_ofs'] else 'Sin OF'
        row['remisiones_asociadas'] = ', '.join(trk_rem['remisiones_asociadas']) if trk_rem['remisiones_asociadas'] else 'Sin Remisión'
        row['estatus_360'] = estatus_360
        matrix.append(row)
        
    return pd.DataFrame(matrix)