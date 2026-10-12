import sqlite3
import pandas as pd
import datetime
import shutil
import os
import re
import base64
import json
import threading
import urllib.parse
from pathlib import Path
from config import (
    SQLITE_DB_PATH,
    DATA_DIR,
    EXCEL_CABECERA_PATH,
    EXCEL_REQ_PATH,
    EXCEL_PARTIDAS_DETALLE_PATH,
    SYNC_DB_DIR,
    get_remisiones_dir,
    ESTATUS_REGISTRADA
)

if os.environ.get("GCS_BUCKET"):
    try:
        import gcs_sync
        gcs_sync.sync_from_gcs()
    except Exception as _e_init:
        print(f"[GCS-INIT] Error syncing on import: {_e_init}")

# ──────────────────────────────────────────────────────────────────────────────
# GITHUB & GCS DB PERSISTENCE
# Subimos po_tracker.db + archivos Excel a GCS (y opcionalmente a GitHub) después de cada escritura.
# Los archivos binarios pesados (.msg/.pdf) residen exclusivamente en GCS (gs://sigrama-po-tracker-storage/data/correos/).
# ──────────────────────────────────────────────────────────────────────────────

def _get_github_config():
    """Lee la configuración de GitHub desde st.secrets o variables de entorno."""
    token = None
    repo  = None
    user  = None
    branch = None
    
    # 1. Intentar desde Streamlit secrets (prioridad en producción)
    try:
        import streamlit as st
        token  = st.secrets.get("GITHUB_TOKEN", None)
        repo   = st.secrets.get("GITHUB_REPO", None)
        user   = st.secrets.get("GITHUB_USER", None)
        branch = st.secrets.get("GITHUB_BRANCH", "main")
    except Exception:
        pass
    
    # 2. Fallback a variables de entorno (desarrollo local)
    if not token:
        token  = os.environ.get("GITHUB_TOKEN", "")
        repo   = os.environ.get("GITHUB_REPO", "jesusalbertomoraleslopez-byte/sigrama_po_tracker")
        user   = os.environ.get("GITHUB_USER", "jesusalbertomoraleslopez-byte")
        branch = os.environ.get("GITHUB_BRANCH", "main")
    
    return token, repo, user, branch


def push_db_to_github(background=True):
    """
    Sube po_tracker.db y archivos Excel a Google Cloud Storage (y réplica cruzada a Remisiones)
    y a GitHub vía API REST si hay token configurado.
    Se ejecuta en background para no bloquear la UI de Streamlit (0ms bloqueo).
    """
    if os.environ.get("GCS_BUCKET"):
        try:
            import gcs_sync
            if background:
                gcs_sync.push_db_to_gcs_async()
            else:
                gcs_sync.push_db_to_gcs()
        except Exception as e_gcs:
            print(f"[GCS-PUSH] Error respaldando en GCS: {e_gcs}")

    def _push():
        try:
            import urllib.request
            token, repo, user, branch = _get_github_config()
            if not token or not repo:
                return
            
            api_base = f"https://api.github.com/repos/{repo}/contents"
            headers = {
                "Authorization": f"token {token}",
                "Accept": "application/vnd.github.v3+json",
                "Content-Type": "application/json"
            }
            
            # Solo respaldar en GitHub la BD compacta (~440 KB) y los Excels maestros.
            # Los correos (.msg/.pdf, >200 MB) se almacenan exclusivamente en GCS.
            files_to_push = [
                SQLITE_DB_PATH,
                EXCEL_CABECERA_PATH,
                EXCEL_REQ_PATH,
                EXCEL_PARTIDAS_DETALLE_PATH,
            ]
            
            for local_path in files_to_push:
                if not local_path.exists():
                    continue
                    
                # Ruta relativa dentro del repositorio
                try:
                    repo_path = local_path.relative_to(SQLITE_DB_PATH.parent.parent)
                except Exception:
                    repo_path = Path("data") / local_path.name
                github_path = repo_path.as_posix()  # e.g. "data/po_tracker.db"
                
                try:
                    with open(local_path, 'rb') as f:
                        content_b64 = base64.b64encode(f.read()).decode('utf-8')
                except Exception as e_read:
                    print(f"[GH-PUSH] Error leyendo {local_path}: {e_read}")
                    continue
                
                encoded_path = urllib.parse.quote(github_path, safe='/')
                url = f"{api_base}/{encoded_path}"
                
                # Obtener SHA actual del archivo en GitHub (necesario para actualizar)
                sha = ""
                try:
                    req_get = urllib.request.Request(url, headers=headers, method="GET")
                    with urllib.request.urlopen(req_get, timeout=10) as resp:
                        remote_info = json.loads(resp.read().decode())
                        sha = remote_info.get("sha", "")
                except Exception:
                    sha = ""
                
                now_ts = datetime.datetime.now().strftime('%Y-%m-%d %H:%M')
                payload = {
                    "message": f"auto-backup: {github_path} @ {now_ts}",
                    "content": content_b64,
                    "branch": branch
                }
                if sha:
                    payload["sha"] = sha
                
                req_put = urllib.request.Request(
                    url,
                    data=json.dumps(payload).encode('utf-8'),
                    headers=headers,
                    method="PUT"
                )
                try:
                    with urllib.request.urlopen(req_put, timeout=25) as resp_put:
                        resp_put.read()  # consume response
                    print(f"[GH-PUSH] ✅ Subido a GitHub: {github_path}")
                except Exception as e_put:
                    print(f"[GH-PUSH] ❌ Error subiendo {github_path}: {e_put}")
                    
        except Exception as e:
            print(f"[GH-PUSH] Error general: {e}")
    
    if background:
        t = threading.Thread(target=_push, daemon=True)
        t.start()
    else:
        _push()



def pull_db_from_github():
    """
    Descarga po_tracker.db desde GCS (si GCS_BUCKET está presente) o 
    desde GitHub si la copia local está vacía o desactualizada.
    Retorna True si se descargó/actualizó, False si ya estaba al día.
    """
    if os.environ.get("GCS_BUCKET"):
        try:
            import gcs_sync
            return gcs_sync.sync_from_gcs()
        except Exception as e_gcs:
            print(f"[GCS-PULL] Error sincronizando desde GCS: {e_gcs}")
            return False

    try:
        import urllib.request
        token, repo, user, branch = _get_github_config()
        if not token or not repo:
            return False
        
        api_url = f"https://api.github.com/repos/{repo}/contents/data/po_tracker.db?ref={branch}"
        headers = {
            "Authorization": f"token {token}",
            "Accept": "application/vnd.github.v3+json"
        }
        
        req = urllib.request.Request(api_url, headers=headers, method="GET")
        try:
            with urllib.request.urlopen(req, timeout=15) as resp:
                info = json.loads(resp.read().decode())
                remote_sha = info.get("sha", "")
                download_url = info.get("download_url", "")
        except Exception:
            return False
        
        # Verificar si la BD local tiene los mismos datos (comparar size aproximado)
        local_path = SQLITE_DB_PATH
        local_size = local_path.stat().st_size if local_path.exists() else 0
        remote_size = info.get("size", 0)
        
        # Si la BD local ya tiene más datos que la remota, no bajar (para no pisar datos nuevos)
        if local_size > 0 and local_size >= remote_size:
            return False
        
        if not download_url:
            return False
        
        # Descargar y escribir
        dl_req = urllib.request.Request(download_url, headers={"Authorization": f"token {token}"})
        with urllib.request.urlopen(dl_req, timeout=30) as dl_resp:
            db_bytes = dl_resp.read()
        
        DATA_DIR.mkdir(exist_ok=True)
        with open(local_path, 'wb') as f:
            f.write(db_bytes)
            
        print(f"[GH-PULL] BD descargada desde GitHub ({len(db_bytes):,} bytes)")
        return True
        
    except Exception as e:
        print(f"[GH-PULL] Error: {e}")
        return False



_db_initialized = False

def get_connection():
    conn = sqlite3.connect(str(SQLITE_DB_PATH), check_same_thread=False)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=NORMAL")
    return conn

def init_db():
    global _db_initialized
    if _db_initialized and SQLITE_DB_PATH.exists():
        return
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
    
    # 3. Tabla de Historial de Auditoría
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS po_historial (
        id_log INTEGER PRIMARY KEY AUTOINCREMENT,
        po TEXT,
        fecha_hora TEXT,
        usuario TEXT,
        accion TEXT,
        detalle TEXT
    )
    ''')
    
    # 4. Tabla de Archivos Adjuntos (metadatos en SQLite; binarios en GCS / data/correos)
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS po_archivos_adjuntos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        po TEXT,
        id_interno TEXT,
        nombre_archivo TEXT UNIQUE,
        tipo TEXT,
        contenido BLOB,
        tamano_bytes INTEGER,
        fecha_subida TEXT
    )
    ''')
    
    # 5. Índices compuestos para consultas O(1)
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_cabecera_id_interno ON po_cabecera(id_interno)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_cabecera_estatus ON po_cabecera(estatus_general)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_partidas_po ON po_partidas(po)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_partidas_sku ON po_partidas(clave_sku)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_partidas_sku_cli ON po_partidas(sku_cliente)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_historial_po ON po_historial(po)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_adjuntos_po ON po_archivos_adjuntos(po)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_adjuntos_nombre ON po_archivos_adjuntos(nombre_archivo)")
    
    conn.commit()
    conn.close()
    _db_initialized = True

def reconstruct_correos_on_disk():
    """
    Mantenida por compatibilidad. Con la arquitectura desacoplada en GCS,
    los archivos de correo se descargan bajo demanda vía download_correo_from_gcs().
    """
    pass

def save_archivo_adjunto(po, id_interno, nombre_archivo, tipo, contenido_bytes):
    """
    Guarda el archivo binario (.msg o .pdf) en data/correos/ y lo sube asíncronamente a GCS.
    En SQLite solo almacena el registro de metadatos (contenido = NULL) para mantener
    po_tracker.db ultra-ligera (~440 KB).
    """
    if not nombre_archivo or not contenido_bytes:
        return False
    now_str = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    tamano = len(contenido_bytes)
    try:
        correos_dir = DATA_DIR / 'correos'
        correos_dir.mkdir(parents=True, exist_ok=True)
        local_file = correos_dir / nombre_archivo
        with open(local_file, 'wb') as f:
            f.write(contenido_bytes)
            
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO po_archivos_adjuntos (po, id_interno, nombre_archivo, tipo, contenido, tamano_bytes, fecha_subida)
            VALUES (?, ?, ?, ?, NULL, ?, ?)
            ON CONFLICT(nombre_archivo) DO UPDATE SET
                po=excluded.po,
                id_interno=excluded.id_interno,
                tipo=excluded.tipo,
                contenido=NULL,
                tamano_bytes=excluded.tamano_bytes,
                fecha_subida=excluded.fecha_subida
        ''', (str(po), str(id_interno), str(nombre_archivo), str(tipo).lower(), tamano, now_str))
        conn.commit()
        conn.close()

        try:
            import gcs_sync
            gcs_sync.push_correo_to_gcs_async(local_file)
        except Exception as e_up:
            print(f"[GCS-UPLOAD] Error subiendo {nombre_archivo}: {e_up}")

        return True
    except Exception as e:
        print(f"Error guardando archivo adjunto {nombre_archivo}: {e}")
        return False

def get_archivos_adjuntos_por_po(po):
    """Obtiene la lista de archivos asociados a una PO."""
    conn = get_connection()
    df = pd.read_sql_query(
        "SELECT id, po, id_interno, nombre_archivo, tipo, tamano_bytes, fecha_subida FROM po_archivos_adjuntos WHERE po = ? ORDER BY id ASC",
        conn, params=[str(po)]
    )
    conn.close()
    return df

def get_todos_archivos_adjuntos():
    """Obtiene la lista de todos los archivos adjuntos guardados en la BD."""
    conn = get_connection()
    df = pd.read_sql_query(
        "SELECT id, po, id_interno, nombre_archivo, tipo, tamano_bytes, fecha_subida FROM po_archivos_adjuntos ORDER BY po DESC, id ASC",
        conn
    )
    conn.close()
    return df

def _resolve_correo_bytes(fn, fb=None):
    """Obtiene los bytes de un correo desde caché local en disco o descarga bajo demanda desde GCS."""
    if fb is not None:
        return fb
    if not fn:
        return None
    disk_p = DATA_DIR / 'correos' / fn
    if disk_p.exists() and disk_p.stat().st_size > 0:
        try:
            with open(disk_p, 'rb') as f:
                return f.read()
        except Exception:
            pass
    try:
        import gcs_sync
        return gcs_sync.download_correo_from_gcs(fn)
    except Exception:
        return None

def get_contenido_archivo_por_nombre(nombre_archivo):
    """Obtiene el binario de un archivo por su nombre, con descarga on-demand desde GCS si no está en disco."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT nombre_archivo, tipo, contenido FROM po_archivos_adjuntos WHERE nombre_archivo = ?", (str(nombre_archivo),))
    row = cursor.fetchone()
    conn.close()
    if row:
        fn, ft, fb = row[0], row[1], row[2]
        fb = _resolve_correo_bytes(fn, fb)
        return fn, ft, fb
    # Fallback directo si el archivo existe en disco/GCS aunque no esté en la tabla
    fn_str = str(nombre_archivo)
    fb_direct = _resolve_correo_bytes(fn_str, None)
    if fb_direct:
        ext = Path(fn_str).suffix.lower().lstrip('.')
        return fn_str, ext, fb_direct
    return None, None, None

def get_contenido_archivo_por_id(id_archivo):
    """Obtiene el binario de un archivo por su ID, con descarga on-demand desde GCS si no está en disco."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT nombre_archivo, tipo, contenido FROM po_archivos_adjuntos WHERE id = ?", (int(id_archivo),))
    row = cursor.fetchone()
    conn.close()
    if row:
        fn, ft, fb = row[0], row[1], row[2]
        fb = _resolve_correo_bytes(fn, fb)
        return fn, ft, fb
    return None, None, None

def cancelar_po(po, motivo="Cancelada por cliente", archivo_bytes=None, archivo_nombre=None, usuario="Usuario", push_to_gh=True):
    """
    Cancela formalmente una PO sin eliminarla:
    1. Actualiza estatus_general = 'Cancelada' en po_cabecera.
    2. Establece texto_etiqueta = 'CANCELADA', color_fondo = '#EF4444', color_texto = '#FFFFFF'.
    3. Añade el motivo y fecha a las observaciones.
    4. Guarda el archivo de evidencia (.msg, .pdf, imagen) en po_archivos_adjuntos como tipo 'cancelacion'.
    5. Registra el evento en po_historial para auditoría.
    6. Sincroniza archivos Excel y GitHub.
    """
    conn = get_connection()
    cur = conn.cursor()
    
    cur.execute("SELECT id_interno, observaciones FROM po_cabecera WHERE po = ?", (str(po),))
    row = cur.fetchone()
    if not row:
        conn.close()
        return False, f"PO {po} no encontrada."
        
    id_int, obs_actual = row[0], row[1] or ""
    today_str = datetime.date.today().strftime('%Y-%m-%d')
    nota_canc = f"[CANCELADA FORMALMENTE {today_str}]: {motivo}"
    obs_updated = f"{nota_canc}\n{obs_actual}".strip() if obs_actual else nota_canc
    
    cur.execute("""
        UPDATE po_cabecera
        SET estatus_general = 'Cancelada',
            texto_etiqueta = 'CANCELADA',
            color_fondo = '#EF4444',
            color_texto = '#FFFFFF',
            observaciones = ?
        WHERE po = ?
    """, (obs_updated, str(po)))
    
    now_str = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    cur.execute('''
        INSERT INTO po_historial (po, fecha_hora, usuario, accion, detalle)
        VALUES (?, ?, ?, 'CANCELACION_FORMAL', ?)
    ''', (str(po), now_str, usuario, f"Orden cancelada formalmente: {motivo}"))
    
    conn.commit()
    conn.close()
    
    if archivo_bytes and archivo_nombre:
        safe_name = f"CANCELACION_OC_{po}_{archivo_nombre}"
        save_archivo_adjunto(po, id_int, safe_name, 'cancelacion', archivo_bytes)
        
    export_sync_to_excel()
    if push_to_gh:
        push_db_to_github(background=True)
        
    return True, f"PO {po} cancelada formalmente con éxito."

def cancelar_pos_lote(lista_pos, motivo="Cancelada por cliente", archivo_bytes=None, archivo_nombre=None, usuario="Usuario", push_to_gh=True):
    """
    Cancela un lote de múltiples POs simultáneamente vinculando la misma evidencia formal.
    """
    results = []
    for p in lista_pos:
        ok, msg = cancelar_po(p, motivo, archivo_bytes, archivo_nombre, usuario=usuario, push_to_gh=False)
        results.append((p, ok, msg))
    export_sync_to_excel()
    if push_to_gh:
        push_db_to_github(background=True)
    return results

def reactivar_po(po, motivo="Reactivada por usuario", usuario="Usuario", push_to_gh=True):
    """
    Reactiva una orden previamente cancelada devolviéndola al estado 'Registrada'.
    """
    conn = get_connection()
    cur = conn.cursor()
    today_str = datetime.date.today().strftime('%Y-%m-%d')
    nota_react = f"[REACTIVADA {today_str}]: {motivo}"
    cur.execute("SELECT observaciones FROM po_cabecera WHERE po = ?", (str(po),))
    row = cur.fetchone()
    obs_actual = row[0] or "" if row else ""
    obs_updated = f"{nota_react}\n{obs_actual}".strip()
    
    cur.execute("""
        UPDATE po_cabecera
        SET estatus_general = 'Registrada',
            texto_etiqueta = 'REGISTRADA',
            color_fondo = '#64748B',
            color_texto = '#FFFFFF',
            observaciones = ?
        WHERE po = ?
    """, (obs_updated, str(po)))
    
    now_str = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    cur.execute('''
        INSERT INTO po_historial (po, fecha_hora, usuario, accion, detalle)
        VALUES (?, ?, ?, 'REACTIVACION', ?)
    ''', (str(po), now_str, usuario, f"Orden reactivada a Registrada: {motivo}"))
    
    conn.commit()
    conn.close()
    export_sync_to_excel()
    if push_to_gh:
        push_db_to_github(background=True)
    return True, f"PO {po} reactivada con éxito."

def get_next_id_interno():
    """Calcula el siguiente identificador consecutivo INT-XXXX disponible en la BD."""
    try:
        conn = get_connection()
        cur = conn.cursor()
        cur.execute("SELECT id_interno FROM po_cabecera WHERE id_interno LIKE 'INT-%'")
        rows = cur.fetchall()
        conn.close()
        max_num = 0
        for r in rows:
            val = str(r[0] or '').strip()
            m = re.search(r'INT[\s\-_]?(\d+)', val, re.IGNORECASE)
            if m:
                max_num = max(max_num, int(m.group(1)))
        return f"INT-{max_num + 1:04d}"
    except Exception:
        return "INT-0001"

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

def save_po(cabecera, partidas, usuario='Usuario', push_to_gh=True):
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
        if push_to_gh:
            push_db_to_github()  # Auto-backup en GitHub para persistencia en Streamlit Cloud
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
        push_db_to_github()  # Auto-backup en GitHub
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
        push_db_to_github()  # Auto-backup en GitHub
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
        tot_s = pd.to_numeric(df['total'], errors='coerce').fillna(0.0)
        calc_s = pd.to_numeric(df['calc_total'], errors='coerce').fillna(0.0)
        df['total'] = tot_s.where((tot_s != 0.0) | (calc_s <= 0.0), calc_s)
    return df

def get_po_by_folio(po_folio):
    conn = get_connection()
    po_str = str(po_folio).strip()
    po_nodash = po_str.replace('-', '')
    po_dash = f"{po_nodash[:4]}-{po_nodash[4:]}" if len(po_nodash) == 8 else po_str
    
    df_cab = pd.read_sql_query('SELECT * FROM po_cabecera WHERE po = ? OR po = ? OR po = ?', conn, params=[po_str, po_nodash, po_dash])
    df_part = pd.read_sql_query('SELECT * FROM po_partidas WHERE po = ? OR po = ? OR po = ? ORDER BY item_no ASC', conn, params=[po_str, po_nodash, po_dash])
    conn.close()
    
    if not df_part.empty:
        pt_s = pd.to_numeric(df_part['precio_total'], errors='coerce').fillna(0.0)
        cant_s = pd.to_numeric(df_part['cantidad_requerida'], errors='coerce').fillna(0.0)
        pu_s = pd.to_numeric(df_part['precio_unitario'], errors='coerce').fillna(0.0)
        df_part['precio_total'] = pt_s.where(pt_s > 0.0, (cant_s * pu_s).round(2))
        
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
    """
    Exporta po_cabecera y po_partidas a los Excels maestros de comunicación inter-MicroApps:
    - BD_POs_Cabecera.xlsx (con ID_Interno, Cliente_Facturar_A y Estatus_General para Remisiones y Corte/Doblez)
    - BD_Requerimientos_POs.xlsx (con ID_Interno, SKU_Cliente y Descripcion para Remisiones)
    - BD_Partidas_Detalle_POs.xlsx
    """
    conn = get_connection()
    df_cab = pd.read_sql_query('SELECT * FROM po_cabecera', conn)
    df_part = pd.read_sql_query('SELECT * FROM po_partidas', conn)
    conn.close()
    
    cab_cols = [
        'PO', 'ID_Interno', 'Fecha_Pedido', 'Fecha_Llegada', 'Fecha_Solicitada',
        'Proyecto', 'Solicitante', 'Requisicion', 'Destino', 'Cliente_Facturar_A',
        'Estatus_General', 'Texto_Etiqueta', 'Color_Fondo', 'Color_Texto'
    ]
    df_cab_rem = pd.DataFrame()
    if not df_cab.empty:
        df_cab_rem['PO'] = df_cab['po']
        df_cab_rem['ID_Interno'] = df_cab.get('id_interno', pd.Series([''] * len(df_cab))).fillna('')
        df_cab_rem['Fecha_Pedido'] = df_cab['fecha_pedido'].fillna('')
        df_cab_rem['Fecha_Llegada'] = df_cab.get('fecha_llegada', pd.Series([''] * len(df_cab))).fillna('')
        df_cab_rem['Fecha_Solicitada'] = df_cab.get('fecha_solicitada', pd.Series([''] * len(df_cab))).fillna('')
        df_cab_rem['Proyecto'] = df_cab['proyecto'].fillna('')
        df_cab_rem['Solicitante'] = df_cab['solicitante'].fillna('')
        df_cab_rem['Requisicion'] = df_cab['requisicion'].fillna('')
        df_cab_rem['Destino'] = df_cab['destino'].fillna('')
        df_cab_rem['Cliente_Facturar_A'] = df_cab.get('cliente_facturar_a', pd.Series([''] * len(df_cab))).fillna('INDUSTRIA SIGRAMA S.A. DE C.V.')
        df_cab_rem['Estatus_General'] = df_cab.get('estatus_general', pd.Series(['Registrada'] * len(df_cab))).fillna('Registrada')
        df_cab_rem['Texto_Etiqueta'] = df_cab['texto_etiqueta'].fillna('')
        df_cab_rem['Color_Fondo'] = df_cab['color_fondo'].fillna('#EC2024')
        df_cab_rem['Color_Texto'] = df_cab['color_texto'].fillna('#FFFFFF')
    else:
        df_cab_rem = pd.DataFrame(columns=cab_cols)
    
    df_cab_rem.to_excel(EXCEL_CABECERA_PATH, index=False)
    
    # Mapa rápido PO -> ID_Interno para enriquecer Requerimientos
    po_to_id_int = {}
    if not df_cab.empty and 'po' in df_cab.columns and 'id_interno' in df_cab.columns:
        po_to_id_int = dict(zip(df_cab['po'].astype(str), df_cab['id_interno'].fillna('').astype(str)))
        
    req_cols = ['PO', 'ID_Interno', 'SKU', 'SKU_Cliente', 'Descripcion', 'Fecha_Entrega', 'Cantidad_Requerida', 'Unidad', 'Parcialidad']
    df_req_rem = pd.DataFrame()
    if not df_part.empty:
        df_req_rem['PO'] = df_part['po']
        df_req_rem['ID_Interno'] = df_part['po'].astype(str).map(lambda x: po_to_id_int.get(x, ''))
        df_req_rem['SKU'] = df_part['clave_sku'].fillna('')
        df_req_rem['SKU_Cliente'] = df_part.get('sku_cliente', pd.Series([''] * len(df_part))).fillna('')
        df_req_rem['Descripcion'] = df_part.get('descripcion_producto', pd.Series([''] * len(df_part))).fillna('')
        df_req_rem['Fecha_Entrega'] = df_part['fecha_entrega'].fillna('')
        df_req_rem['Cantidad_Requerida'] = df_part['cantidad_requerida'].fillna(0)
        df_req_rem['Unidad'] = df_part.get('unidad', pd.Series(['PIEZA'] * len(df_part))).fillna('PIEZA')
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
            
    if SYNC_DB_DIR.exists() and SYNC_DB_DIR != EXCEL_CABECERA_PATH.parent:
        try:
            shutil.copy2(EXCEL_CABECERA_PATH, SYNC_DB_DIR / 'BD_POs_Cabecera.xlsx')
            shutil.copy2(EXCEL_REQ_PATH, SYNC_DB_DIR / 'BD_Requerimientos_POs.xlsx')
        except Exception as e:
            pass

