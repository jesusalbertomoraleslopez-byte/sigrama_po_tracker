import pandas as pd
import os
from pathlib import Path
from config import normalize_po

POSSIBLE_CORTE_DOBLEZ_DIRS = [
    Path(r'C:/Users/albertol/.gemini/antigravity/scratch/app_corte_doblez'),
    Path(__file__).resolve().parent.parent / 'app_corte_doblez'
]

def get_corte_doblez_dir():
    for d in POSSIBLE_CORTE_DOBLEZ_DIRS:
        if d.exists() and (d / 'sigrama_database.xlsx').exists():
            return d
    return POSSIBLE_CORTE_DOBLEZ_DIRS[0]

def load_corte_doblez_databases():
    """Carga las bases de datos relevantes de la app de Corte y Doblez."""
    cd_dir = get_corte_doblez_dir()
    excel_path = cd_dir / 'sigrama_database.xlsx'
    
    if not excel_path.exists():
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
        
    try:
        xl = pd.ExcelFile(excel_path)
        df_ord = xl.parse('ordenes') if 'ordenes' in xl.sheet_names else pd.DataFrame()
        df_pie = xl.parse('piezas') if 'piezas' in xl.sheet_names else pd.DataFrame()
        df_ava = xl.parse('avances') if 'avances' in xl.sheet_names else pd.DataFrame()
        df_tar = xl.parse('tarimas') if 'tarimas' in xl.sheet_names else pd.DataFrame()
        return df_ord, df_pie, df_ava, df_tar
    except Exception as e:
        print(f'Error loading Corte y Doblez DB: {e}')
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

def get_corte_doblez_tracking_for_po(po_folio, df_partidas):
    """Calcula el avance de manufactura en planta (Corte, Doblez, Liberado) para una PO."""
    df_ord, df_pie, df_ava, df_tar = load_corte_doblez_databases()
    
    po_str = str(po_folio).strip()
    po_norm = normalize_po(po_str)
    
    # 1. Encontrar OFs relacionadas con esta PO
    matched_ofs = set()
    if not df_ord.empty and 'of_number' in df_ord.columns:
        for _, o_row in df_ord.iterrows():
            of_num = str(o_row.get('of_number', '')).strip()
            po_field = str(o_row.get('po', '')).strip()
            
            # Checar coincidencia en campo PO o en el nombre de la OF
            if (po_field and normalize_po(po_field) == po_norm) or (po_norm and po_norm in normalize_po(of_num)):
                matched_ofs.add(of_num)
                
    # 2. Filtrar piezas y avances de esas OFs
    df_pie_po = df_pie[df_pie['of_number'].isin(matched_ofs)] if not df_pie.empty and matched_ofs else pd.DataFrame()
    df_ava_po = df_ava[df_ava['of_number'].isin(matched_ofs)] if not df_ava.empty and matched_ofs else pd.DataFrame()
    df_tar_po = df_tar[df_tar['of_number'].isin(matched_ofs)] if not df_tar.empty and matched_ofs else pd.DataFrame()
    
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
            c_liberado = 0.0
            
            # Buscar en piezas programadas
            if not df_pie_po.empty:
                m_pie = df_pie_po[
                    (df_pie_po['no_pieza'].astype(str).str.strip().str.upper() == sku) |
                    ((sku_cli != '') & (df_pie_po['no_pieza'].astype(str).str.strip().str.upper() == sku_cli))
                ]
                c_prog = float(m_pie['cantidad'].sum()) if not m_pie.empty else 0.0
                
            # Buscar en avances por área
            if not df_ava_po.empty:
                m_ava = df_ava_po[
                    (df_ava_po['no_pieza'].astype(str).str.strip().str.upper() == sku) |
                    ((sku_cli != '') & (df_ava_po['no_pieza'].astype(str).str.strip().str.upper() == sku_cli))
                ]
                if not m_ava.empty:
                    c_corte = float(m_ava[m_ava['area'].str.lower() == 'corte']['cantidad'].sum())
                    c_doblez = float(m_ava[m_ava['area'].str.lower() == 'doblez']['cantidad'].sum())
                    c_liberado = float(m_ava[m_ava['area'].str.lower().isin(['liberado', 'empaque'])]['cantidad'].sum())
                    
            # Si no hubo avances detallados pero se armaron tarimas de planta
            if c_liberado == 0.0 and not df_tar_po.empty:
                m_tar = df_tar_po[
                    (df_tar_po['no_pieza'].astype(str).str.strip().str.upper() == sku) |
                    ((sku_cli != '') & (df_tar_po['no_pieza'].astype(str).str.strip().str.upper() == sku_cli))
                ]
                c_liberado = float(m_tar['cantidad'].sum()) if not m_tar.empty else 0.0
                
            # Si se terminó o remisionó en almacén, como mínimo está fabricado
            c_terminado_real = max(c_liberado, c_doblez, c_corte)
            
            total_cortado += c_corte
            total_doblado += c_doblez
            total_terminado += c_terminado_real
            
            pct_cd = (c_terminado_real / cant_req * 100.0) if cant_req > 0 else 0.0
            
            p_res = dict(part)
            p_res['piezas_programadas'] = c_prog
            p_res['piezas_cortadas'] = c_corte
            p_res['piezas_dobladas'] = c_doblez
            p_res['piezas_terminadas_planta'] = c_terminado_real
            p_res['pct_avance_fabricacion'] = round(min(100.0, pct_cd), 1)
            p_res['ofs_asociadas'] = ', '.join(sorted(matched_ofs)) if matched_ofs else 'Por programar OF'
            partidas_cd.append(p_res)
            
    pct_global_cd = (total_terminado / total_req_cd * 100.0) if total_req_cd > 0 else 0.0
    
    return {
        'po': po_str,
        'matched_ofs': sorted(list(matched_ofs)),
        'ofs_asociadas': sorted(list(matched_ofs)),
        'total_programado': total_cortado,
        'total_cortado': total_cortado,
        'total_doblado': total_doblado,
        'total_terminado_planta': total_terminado,
        'total_fabricado': total_terminado,
        'pct_global_fabricacion': round(min(100.0, pct_global_cd), 1),
        'porcentaje_fabricacion': round(min(100.0, pct_global_cd), 1),
        'df_partidas_cd': pd.DataFrame(partidas_cd),
        'df_ofs': pd.DataFrame({'OF': sorted(list(matched_ofs))}) if matched_ofs else pd.DataFrame()
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