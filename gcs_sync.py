"""
=============================================================================
SIGRAMA PO TRACKER - INDUSTRIA SIGRAMA S.A. DE C.V.
Módulo de Persistencia de Alto Rendimiento con Google Cloud Storage (Cloud Run)
- Cliente Singleton reutilizable
- Arranque instantáneo (< 0.25s) descargando solo BD SQLite compacta (~440 KB) y Excels
- Índice en memoria O(1) para los +170 archivos de correos (.msg/.pdf) con descarga On-Demand
- Push asíncrono en paralelo y replicación cruzada hacia sigrama-remisiones-storage
=============================================================================
"""

import os
import time
import threading
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
CORREOS_DIR = DATA_DIR / "correos"

_STORAGE_CLIENT = None
_BUCKET_OBJ = None
_SYNC_LOCK = threading.Lock()
_LAST_SYNC_TS = 0.0
_KNOWN_CORREO_BLOBS = {}
_CORREOS_INDEXED = False


def _get_bucket_name() -> str:
    raw = os.environ.get("GCS_BUCKET", "sigrama-po-tracker-storage").strip()
    return raw.split()[0] if raw else "sigrama-po-tracker-storage"


def get_storage_client():
    """Retorna una instancia singleton de storage.Client()."""
    global _STORAGE_CLIENT
    if _STORAGE_CLIENT is None:
        from google.cloud import storage
        _STORAGE_CLIENT = storage.Client()
    return _STORAGE_CLIENT


def _gcs_bucket():
    """Retorna una instancia singleton del bucket principal de PO Tracker."""
    global _BUCKET_OBJ
    bucket_name = _get_bucket_name()
    if not bucket_name:
        return None
    if _BUCKET_OBJ is None or _BUCKET_OBJ.name != bucket_name:
        _BUCKET_OBJ = get_storage_client().bucket(bucket_name)
    return _BUCKET_OBJ


def sync_from_gcs(force: bool = False) -> bool:
    """
    Sincroniza en milisegundos la base de datos SQLite y los Excels maestros desde GCS,
    e indexa en memoria O(1) el catálogo de correos (.msg/.pdf) sin bloquear el arranque.
    """
    global _LAST_SYNC_TS, _CORREOS_INDEXED
    bucket_name = _get_bucket_name()
    if not bucket_name:
        return False

    now = time.time()
    if not force and _LAST_SYNC_TS > 0 and (now - _LAST_SYNC_TS) < 30.0:
        return True

    with _SYNC_LOCK:
        now = time.time()
        if not force and _LAST_SYNC_TS > 0 and (now - _LAST_SYNC_TS) < 30.0:
            return True

        try:
            bucket = _gcs_bucket()
            if not bucket:
                return False

            DATA_DIR.mkdir(parents=True, exist_ok=True)
            CORREOS_DIR.mkdir(parents=True, exist_ok=True)

            core_blobs = []
            for blob in bucket.list_blobs(prefix="data/"):
                if blob.name.endswith("/"):
                    continue
                if blob.name.startswith("data/correos/"):
                    fname = Path(blob.name).name
                    if fname:
                        _KNOWN_CORREO_BLOBS[fname] = blob.name
                        _KNOWN_CORREO_BLOBS[fname.lower()] = blob.name
                    continue

                rel_parts = blob.name.split("/")
                local_path = BASE_DIR.joinpath(*rel_parts)
                need_dl = force or (not local_path.exists())
                if not need_dl:
                    try:
                        st_loc = local_path.stat()
                        if st_loc.st_size == 0:
                            need_dl = True
                        elif blob.updated and blob.updated.timestamp() > (st_loc.st_mtime + 2):
                            need_dl = True
                        elif blob.name.endswith(".db") and st_loc.st_size > 5_000_000 and (blob.size or 0) < 2_000_000:
                            # Si la copia local aún tiene BLOBs pesados y en GCS ya está compactada
                            need_dl = True
                    except Exception:
                        need_dl = True

                if need_dl:
                    core_blobs.append((blob, local_path))

            _CORREOS_INDEXED = True

            def _dl_core(item):
                b, lp = item
                try:
                    lp.parent.mkdir(parents=True, exist_ok=True)
                    b.download_to_filename(str(lp))
                except Exception as e_dl:
                    print(f"[GCS] Error descargando {b.name}: {e_dl}")

            if core_blobs:
                with ThreadPoolExecutor(max_workers=6) as ex:
                    list(ex.map(_dl_core, core_blobs))
                print(f"[GCS] Sincronizados {len(core_blobs)} archivos núcleo desde gs://{bucket_name} (Correos indexados: {len(_KNOWN_CORREO_BLOBS)//2})")

            _LAST_SYNC_TS = time.time()
            return True
        except Exception as e:
            print(f"[GCS] Aviso al sincronizar desde gs://{bucket_name}: {e}")
            return False


def ensure_correos_indexed():
    """Garantiza que el índice en memoria de data/correos/ esté poblado."""
    global _CORREOS_INDEXED
    if _CORREOS_INDEXED and _KNOWN_CORREO_BLOBS:
        return
    try:
        bucket = _gcs_bucket()
        if not bucket:
            return
        for blob in bucket.list_blobs(prefix="data/correos/"):
            if blob.name.endswith("/"):
                continue
            fname = Path(blob.name).name
            if fname:
                _KNOWN_CORREO_BLOBS[fname] = blob.name
                _KNOWN_CORREO_BLOBS[fname.lower()] = blob.name
        _CORREOS_INDEXED = True
    except Exception as e:
        print(f"[GCS] Aviso indexando correos: {e}")


def download_correo_from_gcs(nombre_archivo: str) -> bytes | None:
    """
    Obtiene los bytes de un archivo de correo (.msg o .pdf) desde disco local
    o lo descarga bajo demanda (On-Demand) desde GCS usando el índice O(1).
    """
    if not nombre_archivo:
        return None
    fname = Path(str(nombre_archivo).strip()).name
    local_path = CORREOS_DIR / fname
    if local_path.exists() and local_path.stat().st_size > 0:
        try:
            return local_path.read_bytes()
        except Exception:
            pass

    ensure_correos_indexed()
    blob_name = _KNOWN_CORREO_BLOBS.get(fname) or _KNOWN_CORREO_BLOBS.get(fname.lower()) or f"data/correos/{fname}"
    try:
        bucket = _gcs_bucket()
        if not bucket:
            return None
        blob = bucket.blob(blob_name)
        if blob.exists():
            CORREOS_DIR.mkdir(parents=True, exist_ok=True)
            data = blob.download_as_bytes()
            try:
                local_path.write_bytes(data)
            except Exception:
                pass
            return data
    except Exception as e:
        print(f"[GCS] Error descargando correo on-demand {fname}: {e}")
    return None


def push_file_to_gcs(local_path: Path) -> bool:
    """Sube un archivo individual a GCS y actualiza el índice en memoria si es un correo."""
    try:
        p = Path(local_path)
        if not p.exists():
            return False
        bucket = _gcs_bucket()
        if not bucket:
            return False
        rel_path = p.relative_to(BASE_DIR).as_posix()
        bucket.blob(rel_path).upload_from_filename(str(p))
        if rel_path.startswith("data/correos/"):
            _KNOWN_CORREO_BLOBS[p.name] = rel_path
            _KNOWN_CORREO_BLOBS[p.name.lower()] = rel_path
        print(f"[GCS] Archivo subido: {rel_path}")
        return True
    except Exception as e:
        print(f"[GCS] Error al subir {local_path} a GCS: {e}")
        return False


def push_correo_to_gcs_async(local_path: Path):
    """Sube un archivo de correo/adjunto nuevo a GCS en segundo plano."""
    p = Path(local_path)
    if p.exists():
        threading.Thread(target=push_file_to_gcs, args=(p,), daemon=True).start()


def push_db_to_gcs() -> bool:
    """
    Sube en paralelo únicamente la BD SQLite compacta (~440 KB) y los Excels maestros a GCS,
    y replica automáticamente BD_POs_Cabecera.xlsx y BD_Requerimientos_POs.xlsx hacia
    gs://sigrama-remisiones-storage para comunicación inmediata entre MicroApps.
    """
    try:
        bucket = _gcs_bucket()
        if not bucket:
            return False

        targets = [
            DATA_DIR / "po_tracker.db",
            DATA_DIR / "BD_POs_Cabecera.xlsx",
            DATA_DIR / "BD_POs_Partidas_Detalladas.xlsx",
            DATA_DIR / "BD_Requerimientos_POs.xlsx",
            DATA_DIR / "sync_databases" / "BD_POs_Cabecera.xlsx",
            DATA_DIR / "sync_databases" / "BD_Requerimientos_POs.xlsx",
        ]
        valid_targets = [t for t in targets if t.exists()]

        def _up_one(t_path: Path):
            rel = t_path.relative_to(BASE_DIR).as_posix()
            bucket.blob(rel).upload_from_filename(str(t_path))

        if valid_targets:
            with ThreadPoolExecutor(max_workers=6) as ex:
                list(ex.map(_up_one, valid_targets))

        # Replicación cruzada directa hacia el bucket de la MicroApp de Remisiones
        try:
            client = get_storage_client()
            rem_bucket = client.bucket("sigrama-remisiones-storage")
            cross_files = [
                (DATA_DIR / "BD_POs_Cabecera.xlsx", "BD_POs_Cabecera.xlsx"),
                (DATA_DIR / "BD_Requerimientos_POs.xlsx", "BD_Requerimientos_POs.xlsx"),
            ]
            for src_f, dst_blob in cross_files:
                if src_f.exists():
                    rem_bucket.blob(dst_blob).upload_from_filename(str(src_f))
        except Exception as e_cross:
            print(f"[GCS-CROSS] Aviso replicando POs a sigrama-remisiones-storage: {e_cross}")

        return True
    except Exception as e:
        print(f"[GCS] Error sincronizando BDs a GCS: {e}")
        return False


def push_db_to_gcs_async():
    """Ejecuta push_db_to_gcs en segundo plano para 0ms de bloqueo en la UI."""
    threading.Thread(target=push_db_to_gcs, daemon=True).start()
