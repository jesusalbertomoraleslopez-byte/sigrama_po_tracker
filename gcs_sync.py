"""
=============================================================================
SIGRAMA PO TRACKER - INDUSTRIA SIGRAMA S.A. DE C.V.
Módulo de Persistencia con Google Cloud Storage (Cloud Run)
=============================================================================
"""

import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"

GCS_BUCKET = os.environ.get("GCS_BUCKET", "").strip()

def _gcs_bucket():
    from google.cloud import storage
    return storage.Client().bucket(GCS_BUCKET)

def sync_from_gcs():
    """Descarga data/ (SQLite, Excels, correos) desde GCS al arrancar."""
    if not GCS_BUCKET:
        return False
    try:
        bucket = _gcs_bucket()
        count = 0
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        for blob in bucket.list_blobs(prefix="data/"):
            if blob.name.endswith("/"):
                continue
            rel_parts = blob.name.split("/")
            local_path = BASE_DIR.joinpath(*rel_parts)
            local_path.parent.mkdir(parents=True, exist_ok=True)
            if not local_path.exists() or local_path.stat().st_size != blob.size:
                blob.download_to_filename(str(local_path))
                count += 1
        print(f"[GCS] Sincronizados exitosamente {count} archivos de PO Tracker desde gs://{GCS_BUCKET}")
        return True
    except Exception as e:
        print(f"[GCS] Error al sincronizar desde gs://{GCS_BUCKET}: {e}")
        return False

def push_file_to_gcs(local_path: Path):
    """Sube un archivo a GCS."""
    if not GCS_BUCKET:
        return False
    try:
        p = Path(local_path)
        if not p.exists():
            return False
        bucket = _gcs_bucket()
        rel_path = p.relative_to(BASE_DIR).as_posix()
        bucket.blob(rel_path).upload_from_filename(str(p))
        print(f"[GCS] Archivo subido: {rel_path}")
        return True
    except Exception as e:
        print(f"[GCS] Error al subir {local_path} a GCS: {e}")
        return False

def push_db_to_gcs():
    """Sube SQLite y Excels principales a GCS."""
    if not GCS_BUCKET:
        return False
    try:
        targets = [
            DATA_DIR / "po_tracker.db",
            DATA_DIR / "BD_POs_Cabecera.xlsx",
            DATA_DIR / "BD_POs_Partidas_Detalladas.xlsx",
            DATA_DIR / "BD_Requerimientos_POs.xlsx",
        ]
        correos_dir = DATA_DIR / "correos"
        if correos_dir.exists():
            for f in correos_dir.iterdir():
                if f.is_file() and f.suffix.lower() in [".msg", ".pdf", ".xlsx"]:
                    targets.append(f)
        for t in targets:
            if t.exists():
                push_file_to_gcs(t)
        return True
    except Exception as e:
        print(f"[GCS] Error sincronizando BDs a GCS: {e}")
        return False
