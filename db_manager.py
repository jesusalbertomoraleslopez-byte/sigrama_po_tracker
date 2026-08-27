import sqlite3
import pandas as pd
import datetime
import shutil
import os
from pathlib import Path
from config import (
    SQLITE_DB_PATH,
    DATA_DIR,
    EXCEL_CABECERA_PATH,
    EXCEL_REQ_PATH,
    EXCEL_PARTIDAS_DETALLE_PATH,
    get_remisiones_dir,
    ESTATUS_REGISTRADA
)

def get_connection():
    return sqlite3.connect(str(SQLITE_DB_PATH), check_same_thread=False)

def init_db():
    conn = get_connection()
    cursor = conn.cursor()
    
    # 1. Tabla de Cabecera de PO
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS po_cabecera (
        po TEXT PRIMARY KEY,
        id_interno TEXT DEFAULT '',
        fecha_llegada TEXT DEFAULT '',
        fecha_solicitada TEXT DEFAULT '',
        archivo_correo TEXT DEFAULT '',
        archivo_pdf TEXT DEFAULT '',
        fecha_pedido TEXT,
        proyecto TEXT,
        solicitante TEXT,
        requisicion TEXT,
        destino TEXT,
        proveedor TEXT,
        proveedor_atencion TEXT,
        cliente_facturar_a TEXT,
        cliente_rfc TEXT,
        cliente_direccion TEXT,
        forma_pago TEXT,
        lab TEXT,
        tiempo_entrega TEXT,
        comprador TEXT,
        subtotal REAL DEFAULT 0,
        descuento REAL DEFAULT 0,
        iva REAL DEFAULT 0,
        ret_iva REAL DEFAULT 0,
        ret_isr REAL DEFAULT 0,
        total REAL DEFAULT 0,
        moneda TEXT DEFAULT 'MXN',
        observaciones TEXT,
        estatus_general TEXT DEFAULT 'Registrada',
        fecha_registro TEXT,
        texto_etiqueta TEXT,
        color_fondo TEXT,
        color_texto TEXT
    )
    ''')
    
    # Migración automática si campos no existen en po_cabecera
    cursor.execute("PRAGMA table_info(po_cabecera)")
    cab_cols = [row[1] for row in cursor.fetchall()]
    for col_name in ['id_interno', 'fecha_llegada', 'fecha_solicitada', 'archivo_correo', 'archivo_pdf']:
        if col_name not in cab_cols:
            cursor.execute(f"ALTER TABLE po_cabecera ADD COLUMN {col_name} TEXT DEFAULT ''")
    
    # 2. Tabla de Partidas de PO
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS po_partidas (
        id_partida INTEGER PRIMARY KEY AUTOINCREMENT,
        po TEXT,
        item_no INTEGER,
        sku_cliente TEXT,
        clave_sku TEXT,
        descripcion_producto TEXT,
        cantidad_requerida REAL,
        unidad TEXT,
        precio_unitario REAL DEFAULT 0,
        precio_total REAL DEFAULT 0,
        fecha_entrega TEXT,
        parcialidad TEXT DEFAULT 'P1',
        observaciones_partida TEXT,
        FOREIGN KEY(po) REFERENCES po_cabecera(po) ON DELETE CASCADE
    )
    ''')
    
    # Migración automática si sku_cliente no existe
    cursor.execute("PRAGMA table_info(po_partidas)")
    columns = [row[1] for row in cursor.fetchall()]
    if 'sku_cliente' not in columns:
        cursor.execute("ALTER TABLE po_partidas ADD COLUMN sku_cliente TEXT DEFAULT ''")
    
    conn.commit()
    conn.close()

def clear_all_pos_db(usuario='Usuario'):
    """Limpia completamente todas las tablas de POs en SQLite y vacía los archivos Excel."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute('DELETE FROM po_cabecera')
    cursor.execute('DELETE FROM po_partidas')
    cursor.execute('DELETE FROM po_historial')
    now_str = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    cursor.execute('''
        INSERT INTO po_historial (po, fecha_hora, usuario, accion, detalle)
        VALUES (?, ?, ?, ?, ?)
    ''', ('SISTEMA', now_str, usuario, 'Limpieza Total', 'Se reinició la base de datos a 0 POs para comenzar carga limpia ordenada.'))
    conn.commit()
    conn.close()
    
    export_sync_to_excel()
    return True, "Base de datos reseteada con éxito. Catálogo limpio listo (0 POs)."

def import_existing_pos_from_remisiones():
    conn = get_connection()
    df_check = pd.read_sql_query('SELECT COUNT(*) as cnt FROM po_cabecera', conn)
    cnt = df_check['cnt'].iloc[0] if not df_check.empty else 0
    
    if cnt == 0:
        rem_dir = get_remisiones_dir()
        cab_file = rem_dir / 'BD_POs_Cabecera.xlsx'
        req_file = rem_dir / 'BD_Requerimientos_POs.xlsx'
        
        if cab_file.exists():
            try:
                df_cab = pd.read_excel(cab_file)
                df_req = pd.read_excel(req_file) if req_file.exists() else pd.DataFrame()
                
                cursor = conn.cursor()
                now_str = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                
                for _, r in df_cab.iterrows():
                    po_val = str(r.get('PO', '')).strip()
                    if not po_val or po_val.lower() == 'nan':
                        continue
                    
                    fecha_ped = str(r.get('Fecha_Pedido', '')).strip()
                    proy = str(r.get('Proyecto', '')).strip()
                    sol = str(r.get('Solicitante', '')).strip()
                    req_no = str(r.get('Requisicion', '')).strip()
                    dest = str(r.get('Destino', '')).strip()
                    lbl_txt = str(r.get('Texto_Etiqueta', '')).strip() if pd.notnull(r.get('Texto_Etiqueta')) else ''
                    bg_col = str(r.get('Color_Fondo', '')) if pd.notnull(r.get('Color_Fondo')) else '#EC2024'
                    fg_col = str(r.get('Color_Texto', '')) if pd.notnull(r.get('Color_Texto')) else '#FFFFFF'
                    
                    cursor.execute('''
                        INSERT OR REPLACE INTO po_cabecera 
                        (po, fecha_pedido, proyecto, solicitante, requisicion, destino, 
                         proveedor, cliente_facturar_a, estatus_general, fecha_registro,
                         texto_etiqueta, color_fondo, color_texto)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (
                        po_val, fecha_ped, proy, sol, req_no, dest,
                        'SIGRAMA PLANTA METALES', 'INDUSTRIA SIGRAMA S.A. DE C.V.',
                        ESTATUS_REGISTRADA, now_str, lbl_txt, bg_col, fg_col
                    ))
                
                if not df_req.empty:
                    item_counters = {}
                    for _, req in df_req.iterrows():
                        po_val = str(req.get('PO', '')).strip()
                        sku = str(req.get('SKU', '')).strip()
                        f_ent = str(req.get('Fecha_Entrega', '')).strip()
                        cant = float(req.get('Cantidad_Requerida', 0) or 0)
                        parc = str(req.get('Parcialidad', 'P1')).strip()
                        
                        if not po_val or po_val.lower() == 'nan' or not sku or sku.lower() == 'nan':
                            continue
                        
                        item_counters[po_val] = item_counters.get(po_val, 0) + 1
                        cursor.execute('''
                            INSERT INTO po_partidas 
                            (po, item_no, clave_sku, descripcion_producto, cantidad_requerida, unidad, fecha_entrega, parcialidad)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                        ''', (
                            po_val, item_counters[po_val], sku, f'Pieza {sku}', cant, 'PIEZA', f_ent, parc
                        ))
                
                conn.commit()
            except Exception as e:
                print(f'Error importing legacy POs: {e}')
    conn.close()

def save_po(cabecera, partidas, usuario='Usuario'):
    conn = get_connection()
    cursor = conn.cursor()
    
    po_folio = str(cabecera.get('po', '')).strip()
    if not po_folio:
        conn.close()
        return False, 'El Folio de la PO es requerido.'
    
    now_str = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    
    try:
        cursor.execute('''
            INSERT INTO po_cabecera (
                po, id_interno, fecha_llegada, fecha_solicitada, archivo_correo, archivo_pdf,
                fecha_pedido, proyecto, solicitante, requisicion, destino,
                proveedor, proveedor_atencion, cliente_facturar_a, cliente_rfc, cliente_direccion,
                forma_pago, lab, tiempo_entrega, comprador,
                subtotal, descuento, iva, ret_iva, ret_isr, total, moneda,
                observaciones, estatus_general, fecha_registro,
                texto_etiqueta, color_fondo, color_texto
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(po) DO UPDATE SET
                id_interno=COALESCE(NULLIF(excluded.id_interno, ''), po_cabecera.id_interno),
                fecha_llegada=COALESCE(NULLIF(excluded.fecha_llegada, ''), po_cabecera.fecha_llegada),
                fecha_solicitada=COALESCE(NULLIF(excluded.fecha_solicitada, ''), po_cabecera.fecha_solicitada),
                archivo_correo=COALESCE(NULLIF(excluded.archivo_correo, ''), po_cabecera.archivo_correo),
                archivo_pdf=COALESCE(NULLIF(excluded.archivo_pdf, ''), po_cabecera.archivo_pdf),
                fecha_pedido=excluded.fecha_pedido,
                proyecto=excluded.proyecto,
                solicitante=excluded.solicitante,
                requisicion=excluded.requisicion,
                destino=excluded.destino,
                proveedor=excluded.proveedor,
                proveedor_atencion=excluded.proveedor_atencion,
                cliente_facturar_a=excluded.cliente_facturar_a,
                cliente_rfc=excluded.cliente_rfc,
                cliente_direccion=excluded.cliente_direccion,
                forma_pago=excluded.forma_pago,
                lab=excluded.lab,
                tiempo_entrega=excluded.tiempo_entrega,
                comprador=excluded.comprador,
                subtotal=excluded.subtotal,
                descuento=excluded.descuento,
                iva=excluded.iva,
                ret_iva=excluded.ret_iva,
                ret_isr=excluded.ret_isr,
                total=excluded.total,
                moneda=excluded.moneda,
                observaciones=excluded.observaciones,
                estatus_general=excluded.estatus_general,
                texto_etiqueta=excluded.texto_etiqueta,
                color_fondo=excluded.color_fondo,
                color_texto=excluded.color_texto
        ''', (
            po_folio,
            str(cabecera.get('id_interno', '')).strip(),
            str(cabecera.get('fecha_llegada', '')).strip(),
            str(cabecera.get('fecha_solicitada', '')).strip(),
            str(cabecera.get('archivo_correo', '')).strip(),
            str(cabecera.get('archivo_pdf', '')).strip(),
            cabecera.get('fecha_pedido', ''),
            cabecera.get('proyecto', ''),
            cabecera.get('solicitante', ''),
            cabecera.get('requisicion', ''),
            cabecera.get('destino', ''),
            cabecera.get('proveedor', 'SIGRAMA PLANTA METALES'),
            cabecera.get('proveedor_atencion', 'JESUS MORALES'),
            cabecera.get('cliente_facturar_a', 'INDUSTRIA SIGRAMA S.A. DE C.V.'),
            cabecera.get('cliente_rfc', 'ISI-870204-K4A'),
            cabecera.get('cliente_direccion', 'C. JUAN ESCUTIA #50 COL. ABASTOS C.P. 27020 TORREON, COAH.'),
            cabecera.get('forma_pago', ''),
            cabecera.get('lab', 'ALMACEN SIGRAMA'),
            cabecera.get('tiempo_entrega', ''),
            cabecera.get('comprador', ''),
            float(cabecera.get('subtotal', 0) or 0),
            float(cabecera.get('descuento', 0) or 0),
            float(cabecera.get('iva', 0) or 0),
            float(cabecera.get('ret_iva', 0) or 0),
            float(cabecera.get('ret_isr', 0) or 0),
            float(cabecera.get('total', 0) or 0),
            cabecera.get('moneda', 'MXN'),
            cabecera.get('observaciones', ''),
            cabecera.get('estatus_general', ESTATUS_REGISTRADA),
            now_str,
            cabecera.get('texto_etiqueta', ''),
            cabecera.get('color_fondo', '#EC2024'),
            cabecera.get('color_texto', '#FFFFFF')
        ))
        
        cursor.execute('DELETE FROM po_partidas WHERE po = ?', (po_folio,))
        
        for idx, item in enumerate(partidas, start=1):
            cant = float(item.get('cantidad_requerida', 0) or 0)
            pu = float(item.get('precio_unitario', 0) or 0)
            pt = float(item.get('precio_total', 0) or (cant * pu))
            
            cursor.execute('''
                INSERT INTO po_partidas (
                    po, item_no, sku_cliente, clave_sku, descripcion_producto,
                    cantidad_requerida, unidad, precio_unitario, precio_total,
                    fecha_entrega, parcialidad, observaciones_partida
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                po_folio,
                int(item.get('item_no', idx)),
                str(item.get('sku_cliente', '')).strip().upper(),
                str(item.get('clave_sku', '')).strip().upper(),
                str(item.get('descripcion_producto', '')).strip(),
                cant,
                str(item.get('unidad', 'PIEZA')).strip().upper(),
                pu,
                pt,
                str(item.get('fecha_entrega', '')).strip(),
                str(item.get('parcialidad', 'P1')).strip(),
                str(item.get('observaciones_partida', '')).strip()
            ))
            
        cursor.execute('''
            INSERT INTO po_historial (po, fecha_hora, usuario, accion, detalle)
            VALUES (?, ?, ?, ?, ?)
        ''', (
            po_folio, now_str, usuario, 'Guardar PO',
            f'Se registraron/actualizaron datos de la PO {po_folio} con {len(partidas)} partidas.'
        ))
        
        conn.commit()
        conn.close()
        
        export_sync_to_excel()
        return True, f'Orden de Compra {po_folio} guardada exitosamente.'
    except Exception as e:
        conn.rollback()
        conn.close()
        return False, f'Error al guardar PO: {e}'

def update_po_fields(po_folio, fields_dict, usuario='Usuario'):
    """Actualiza campos específicos de la cabecera de una PO de forma atómica."""
    conn = get_connection()
    cursor = conn.cursor()
    now_str = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    
    if not fields_dict:
        conn.close()
        return True, "Sin cambios"
        
    set_clauses = []
    params = []
    for k, v in fields_dict.items():
        set_clauses.append(f"{k} = ?")
        params.append(v)
    params.append(str(po_folio))
    
    try:
        cursor.execute(f"UPDATE po_cabecera SET {', '.join(set_clauses)} WHERE po = ?", params)
        cursor.execute('''
            INSERT INTO po_historial (po, fecha_hora, usuario, accion, detalle)
            VALUES (?, ?, ?, ?, ?)
        ''', (
            str(po_folio), now_str, usuario, 'Ajuste de Información',
            f"Se ajustaron campos de control: {', '.join(fields_dict.keys())}"
        ))
        conn.commit()
        conn.close()
        export_sync_to_excel()
        return True, f"Datos de la PO {po_folio} actualizados correctamente."
    except Exception as e:
        conn.close()
        return False, f"Error al actualizar campos: {e}"

def delete_po(po_folio, usuario='Usuario'):
    conn = get_connection()
    cursor = conn.cursor()
    now_str = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    try:
        cursor.execute('DELETE FROM po_cabecera WHERE po = ?', (po_folio,))
        cursor.execute('DELETE FROM po_partidas WHERE po = ?', (po_folio,))
        cursor.execute('''
            INSERT INTO po_historial (po, fecha_hora, usuario, accion, detalle)
            VALUES (?, ?, ?, ?, ?)
        ''', (po_folio, now_str, usuario, 'Eliminar PO', f'Se eliminó la PO {po_folio}'))
        conn.commit()
        conn.close()
        export_sync_to_excel()
        return True, f'PO {po_folio} eliminada correctamente.'
    except Exception as e:
        conn.close()
        return False, f'Error al eliminar PO: {e}'

def get_all_pos():
    conn = get_connection()
    df = pd.read_sql_query('''
        SELECT c.*, 
               COALESCE((SELECT SUM(COALESCE(NULLIF(p.precio_total, 0), p.cantidad_requerida * p.precio_unitario)) 
                         FROM po_partidas p WHERE p.po = c.po), 0) as calc_total
        FROM po_cabecera c
        ORDER BY 
            CASE WHEN c.id_interno IS NOT NULL AND c.id_interno != '' THEN 0 ELSE 1 END ASC,
            c.id_interno ASC,
            c.fecha_llegada DESC,
            c.po DESC
    ''', conn)
    conn.close()
    if not df.empty and 'total' in df.columns:
        df['total'] = df.apply(lambda r: float(r['calc_total']) if (float(r.get('total', 0) or 0) == 0.0 and float(r.get('calc_total', 0) or 0) > 0) else float(r.get('total', 0) or 0), axis=1)
    return df

def get_po_by_folio(po_folio):
    conn = get_connection()
    df_cab = pd.read_sql_query('SELECT * FROM po_cabecera WHERE po = ?', conn, params=[str(po_folio)])
    df_part = pd.read_sql_query('SELECT * FROM po_partidas WHERE po = ? ORDER BY item_no ASC', conn, params=[str(po_folio)])
    conn.close()
    
    if not df_part.empty:
        df_part['precio_total'] = df_part.apply(
            lambda r: float(r['precio_total']) if float(r.get('precio_total', 0) or 0) > 0 else round(float(r.get('cantidad_requerida', 0) or 0) * float(r.get('precio_unitario', 0) or 0), 2),
            axis=1
        )
        
    if not df_cab.empty:
        cur_tot = float(df_cab.iloc[0].get('total', 0) or 0)
        if cur_tot == 0.0 and not df_part.empty:
            calc_sum = float(df_part['precio_total'].sum())
            if calc_sum > 0:
                df_cab.loc[0, 'total'] = calc_sum
                
    return df_cab, df_part

def get_all_partidas():
    conn = get_connection()
    df = pd.read_sql_query('SELECT * FROM po_partidas ORDER BY po DESC, item_no ASC', conn)
    conn.close()
    return df

def get_po_history(po_folio):
    conn = get_connection()
    df = pd.read_sql_query('SELECT * FROM po_historial WHERE po = ? ORDER BY id_log DESC', conn, params=[str(po_folio)])
    conn.close()
    return df

def export_sync_to_excel():
    conn = get_connection()
    df_cab = pd.read_sql_query('SELECT * FROM po_cabecera', conn)
    df_part = pd.read_sql_query('SELECT * FROM po_partidas', conn)
    conn.close()
    
    cab_cols = ['PO', 'Fecha_Pedido', 'Proyecto', 'Solicitante', 'Requisicion', 'Destino', 'Texto_Etiqueta', 'Color_Fondo', 'Color_Texto']
    df_cab_rem = pd.DataFrame()
    if not df_cab.empty:
        df_cab_rem['PO'] = df_cab['po']
        df_cab_rem['Fecha_Pedido'] = df_cab['fecha_pedido']
        df_cab_rem['Proyecto'] = df_cab['proyecto']
        df_cab_rem['Solicitante'] = df_cab['solicitante']
        df_cab_rem['Requisicion'] = df_cab['requisicion']
        df_cab_rem['Destino'] = df_cab['destino']
        df_cab_rem['Texto_Etiqueta'] = df_cab['texto_etiqueta'].fillna('')
        df_cab_rem['Color_Fondo'] = df_cab['color_fondo'].fillna('#EC2024')
        df_cab_rem['Color_Texto'] = df_cab['color_texto'].fillna('#FFFFFF')
    else:
        df_cab_rem = pd.DataFrame(columns=cab_cols)
    
    df_cab_rem.to_excel(EXCEL_CABECERA_PATH, index=False)
    
    req_cols = ['PO', 'SKU', 'Fecha_Entrega', 'Cantidad_Requerida', 'Parcialidad']
    df_req_rem = pd.DataFrame()
    if not df_part.empty:
        df_req_rem['PO'] = df_part['po']
        df_req_rem['SKU'] = df_part['clave_sku']
        df_req_rem['Fecha_Entrega'] = df_part['fecha_entrega']
        df_req_rem['Cantidad_Requerida'] = df_part['cantidad_requerida']
        df_req_rem['Parcialidad'] = df_part['parcialidad'].fillna('P1')
    else:
        df_req_rem = pd.DataFrame(columns=req_cols)
        
    df_req_rem.to_excel(EXCEL_REQ_PATH, index=False)
    df_part.to_excel(EXCEL_PARTIDAS_DETALLE_PATH, index=False)
    
    rem_dir = get_remisiones_dir()
    if rem_dir.exists() and rem_dir != EXCEL_CABECERA_PATH.parent:
        try:
            shutil.copy2(EXCEL_CABECERA_PATH, rem_dir / 'BD_POs_Cabecera.xlsx')
            shutil.copy2(EXCEL_REQ_PATH, rem_dir / 'BD_Requerimientos_POs.xlsx')
        except Exception as e:
            print(f'Sync copy to remisiones error: {e}')
