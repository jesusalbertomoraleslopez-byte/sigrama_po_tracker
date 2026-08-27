import pandas as pd
import os
import re
from pathlib import Path
from config import normalize_po, get_corte_doblez_dir

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

def load_corte_doblez_databases():
    """Carga las bases de datos relevantes de la app de Corte y Doblez."""
    cd_dir = get_corte_doblez_dir()
    excel_path = cd_dir / 'sigrama_database.xlsx'
    
    if not excel_path.exists():
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
        
    try:
        xl = pd.ExcelFile(excel_path)
        df_ord = xl.parse('ordenes') if 'ordenes' in xl.sheet_names else pd.DataFrame()
        df_pie = xl.parse('piezas') if 'piezas' in xl.sheet_names else pd.DataFrame()
        df_ava = xl.parse('avances') if 'avances' in xl.sheet_names else pd.DataFrame()
        df_tar = xl.parse('tarimas') if 'tarimas' in xl.sheet_names else pd.DataFrame()
        df_nid = xl.parse('nidos') if 'nidos' in xl.sheet_names else pd.DataFrame()
        return df_ord, df_pie, df_ava, df_tar, df_nid
    except Exception as e:
        print(f'Error loading Corte y Doblez DB: {e}')
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

def get_corte_doblez_tracking_for_po(po_folio, df_partidas, id_interno=""):
    """Calcula el avance de manufactura en planta (Corte, Doblez, Liberado) para una PO."""
    df_ord, df_pie, df_ava, df_tar, df_nid = load_corte_doblez_databases()
    
    po_str = str(po_folio).strip()
    po_norm = normalize_po(po_str)
    id_int_clean = re.sub(r'[^0-9]', '', str(id_interno)) if id_interno else ""
    
    # 1. Encontrar OFs relacionadas con esta PO
    matched_ofs = set()
    if not df_ord.empty and 'of_number' in df_ord.columns:
        for _, o_row in df_ord.iterrows():
            of_num = str(o_row.get('of_number', '')).strip()
            po_field = str(o_row.get('po', '')).strip()
            proy_field = str(o_row.get('proyecto', '')).strip()
            
            # Coincidencia por PO, por nombre de OF o por ID Interno
            if (po_field and normalize_po(po_field) == po_norm) or \
               (po_norm and po_norm in normalize_po(of_num)) or \
               (po_str and po_str in of_num) or \
               (id_int_clean and len(id_int_clean) >= 2 and id_int_clean in normalize_po(of_num)):
                matched_ofs.add(of_num)
                
    # 2. Filtrar piezas, avances y nidos de esas OFs
    df_pie_po = df_pie[df_pie['of_number'].isin(matched_ofs)] if not df_pie.empty and matched_ofs else pd.DataFrame()
    df_ava_po = df_ava[df_ava['of_number'].isin(matched_ofs)] if not df_ava.empty and matched_ofs else pd.DataFrame()
    df_tar_po = df_tar[df_tar['of_number'].isin(matched_ofs)] if not df_tar.empty and matched_ofs else pd.DataFrame()
    df_nid_po = df_nid[df_nid['of_number'].isin(matched_ofs)] if not df_nid.empty and matched_ofs else pd.DataFrame()
    
    # Resumen de Láminas Utilizadas
    laminas_summary = []
    total_laminas = 0.0
    if not df_nid_po.empty and 'hojas' in df_nid_po.columns:
        def extract_mat(of_name):
            s = str(of_name).upper()
            if 'CAL. 10' in s or 'CAL 10' in s:
                return 'Lámina Galvanizada Cal. 10'
            elif 'CAL. 12' in s or 'CAL 12' in s:
                return 'Lámina Galvanizada Cal. 12'
            elif 'CAL. 14' in s or 'CAL 14' in s:
                return 'Lámina Galvanizada Cal. 14'
            elif 'CAL. 16' in s or 'CAL 16' in s:
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
            
    # Consulta cruzada con Remisiones para inferencia causal inteligente
    from remisiones_sync import get_tracking_for_po as get_rem_tracking
    rem_trk_info = get_rem_tracking(po_str, df_partidas, id_interno=id_interno)
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
                
    partidas_cd = []
    total_req_cd = 0.0
    total_cortado = 0.0
    total_doblado = 0.0
    total_terminado = 0.0
    
    if not df_partidas.empty:
        for _, part in df_partidas.iterrows():
            sku = str(part.get('clave_sku', '')).strip().upper()
            sku_cli = str(part.get('sku_cliente', '')).strip().upper()
            cant_req = float(part.get('cantidad_requerida', 0) or 0)
            total_req_cd += cant_req
            
            c_prog = 0.0
            c_corte = 0.0
            c_doblez = 0.0
            c_rebabeo = 0.0
            c_liberado = 0.0
            
            # Buscar en piezas programadas con matching flexible de SKU
            if not df_pie_po.empty:
                m_pie = df_pie_po[df_pie_po['no_pieza'].apply(lambda p: sku_matches(sku, p) or (sku_cli and sku_matches(sku_cli, p)))]
                c_prog = float(m_pie['cantidad'].sum()) if not m_pie.empty else 0.0
                
            # Buscar en avances por área
            if not df_ava_po.empty:
                m_ava = df_ava_po[df_ava_po['no_pieza'].apply(lambda p: sku_matches(sku, p) or (sku_cli and sku_matches(sku_cli, p)))]
                if not m_ava.empty:
                    c_corte = float(m_ava[m_ava['area'].astype(str).str.lower() == 'corte']['cantidad'].sum())
                    c_doblez = float(m_ava[m_ava['area'].astype(str).str.lower() == 'doblez']['cantidad'].sum())
                    c_rebabeo = float(m_ava[m_ava['area'].astype(str).str.lower() == 'rebabeo']['cantidad'].sum())
                    c_liberado = float(m_ava[m_ava['area'].astype(str).str.lower().isin(['liberado', 'empaque'])]['cantidad'].sum())
                    
            if c_liberado == 0.0 and not df_tar_po.empty:
                m_tar = df_tar_po[df_tar_po['no_pieza'].apply(lambda p: sku_matches(sku, p) or (sku_cli and sku_matches(sku_cli, p)))]
                c_liberado = float(m_tar['cantidad'].sum()) if not m_tar.empty else 0.0
                
            # REGLA DE INTELIGENCIA OPERATIVA: Si ya está remisionada, estuvo fabricada al 100%
            c_rem_part = max(rem_map.get(sku, 0.0), rem_map.get(sku_cli, 0.0), float(part.get('cantidad_remisionada', 0) or 0))
            if c_rem_part > 0:
                c_corte = max(c_corte, c_rem_part)
                c_doblez = max(c_doblez, c_rem_part)
                c_terminado_real = max(c_liberado, c_doblez, c_rebabeo, c_corte, c_rem_part)
            else:
                c_terminado_real = max(c_liberado, c_doblez, c_rebabeo, c_corte)
            
            total_cortado += min(cant_req, c_corte)
            total_doblado += min(cant_req, c_doblez)
            total_terminado += min(cant_req, c_terminado_real)
            
            pct_cd = (c_terminado_real / cant_req * 100.0) if cant_req > 0 else 0.0
            
            if not matched_ofs and remisiones_asoc:
                ofs_tag = f"Fabricado (Remisión {', '.join(remisiones_asoc)})"
            else:
                ofs_tag = ', '.join(sorted(matched_ofs)) if matched_ofs else 'Por programar OF'
                
            p_res = dict(part)
            p_res['piezas_programadas'] = c_prog if c_prog > 0 else c_terminado_real
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
        'total_programado': total_cortado,
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
        'df_nidos': df_nid_po
    }

def get_integrated_360_summary(df_all_pos, df_all_partidas):
    """Genera la Matriz de Control 360 combinando Requerimiento Maestro + Corte y Doblez + Remisiones."""
    from remisiones_sync import get_tracking_for_po
    
    matrix = []
    if df_all_pos.empty:
        return pd.DataFrame()
        
    for _, po_row in df_all_pos.iterrows():
        po_folio = str(po_row.get('po', '')).strip()
        partidas_po = df_all_partidas[df_all_partidas['po'].astype(str).str.strip() == po_folio] if not df_all_partidas.empty else pd.DataFrame()
        
        # 1. Trazabilidad con Remisiones
        trk_rem = get_tracking_for_po(po_folio, partidas_po)
        
        # 2. Trazabilidad con Corte y Doblez
        trk_cd = get_corte_doblez_tracking_for_po(po_folio, trk_rem['df_partidas'])
        
        tot_req = trk_rem['total_requerido']
        tot_fab = trk_cd['total_terminado_planta']
        tot_env = trk_rem['total_remisionado']
        
        # Estatus combinado
        if tot_env >= tot_req and tot_req > 0:
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
        row['piezas_fabricadas'] = tot_fab
        row['pct_fabricacion'] = trk_cd['pct_global_fabricacion']
        row['piezas_remisionadas'] = tot_env
        row['pct_remision'] = trk_rem['porcentaje_global']
        row['piezas_pendientes_fab'] = max(0.0, tot_req - tot_fab)
        row['piezas_pendientes_env'] = max(0.0, tot_req - tot_env)
        row['ofs_asociadas'] = ', '.join(trk_cd['matched_ofs']) if trk_cd['matched_ofs'] else 'Sin OF'
        row['remisiones_asociadas'] = ', '.join(trk_rem['remisiones_asociadas']) if trk_rem['remisiones_asociadas'] else 'Sin Remisión'
        row['estatus_360'] = estatus_360
        matrix.append(row)
        
    return pd.DataFrame(matrix)