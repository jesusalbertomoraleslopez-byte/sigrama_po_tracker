import os
from pathlib import Path

# Directorios base
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / 'data'
DATA_DIR.mkdir(exist_ok=True)

# Rutas locales de base de datos
SQLITE_DB_PATH = DATA_DIR / 'po_tracker.db'
EXCEL_CABECERA_PATH = DATA_DIR / 'BD_POs_Cabecera.xlsx'
EXCEL_REQ_PATH = DATA_DIR / 'BD_Requerimientos_POs.xlsx'
EXCEL_PARTIDAS_DETALLE_PATH = DATA_DIR / 'BD_POs_Partidas_Detalladas.xlsx'

SYNC_DB_DIR = DATA_DIR / 'sync_databases'
SYNC_DB_DIR.mkdir(exist_ok=True, parents=True)

# Directorio de la App de Remisiones (para integración directa)
POSSIBLE_REMISIONES_DIRS = [
    Path(r'C:/Users/albertol/.gemini/antigravity/scratch/remisiones-de-materiales'),
    BASE_DIR.parent / 'remisiones-de-materiales',
    SYNC_DB_DIR,
    DATA_DIR
]

POSSIBLE_CORTE_DOBLEZ_DIRS = [
    Path(r'C:/Users/albertol/.gemini/antigravity/scratch/app_corte_doblez'),
    BASE_DIR.parent / 'app_corte_doblez',
    SYNC_DB_DIR,
    DATA_DIR
]

def get_remisiones_dir():
    for d in POSSIBLE_REMISIONES_DIRS:
        if d.exists() and (d / 'BD_Datos_Generales_Remision.xlsx').exists():
            return d
    return SYNC_DB_DIR

def get_corte_doblez_dir():
    for d in POSSIBLE_CORTE_DOBLEZ_DIRS:
        if d.exists() and (d / 'sigrama_database.xlsx').exists():
            return d
    return SYNC_DB_DIR

# Estilos corporativos SIGRAMA
PRIMARY_COLOR = '#EC2024'  # Rojo Corporativo
SECONDARY_COLOR = '#111111'  # Negro Profundo
BG_COLOR = '#FFFFFF'
ACCENT_GREEN = '#10B981'
ACCENT_YELLOW = '#F59E0B'
ACCENT_BLUE = '#3B82F6'
ACCENT_GRAY = '#64748B'

# Estatus de Órdenes de Compra
ESTATUS_REGISTRADA = 'Registrada'
ESTATUS_EN_PROCESO = 'En Proceso'
ESTATUS_PARCIAL = 'Remisionada Parcial'
ESTATUS_COMPLETADA = 'Remisionada Total'
ESTATUS_CANCELADA = 'Cancelada'

ESTATUS_COLORS = {
    ESTATUS_REGISTRADA: '#64748B',
    ESTATUS_EN_PROCESO: '#3B82F6',
    ESTATUS_PARCIAL: '#F59E0B',
    ESTATUS_COMPLETADA: '#10B981',
    ESTATUS_CANCELADA: '#EF4444'
}

def normalize_po(po_val):
    """Normaliza folios de PO eliminando prefijos, guiones, barras y espacios."""
    if po_val is None or str(po_val).strip().lower() in ('none', 'nan', 'nat', 'null', ''):
        return ""
    s = str(po_val).strip().upper()
    if s.startswith("PO "):
        s = s[3:].strip()
    elif s.startswith("PO-"):
        s = s[3:].strip()
    elif s.startswith("PO"):
        s = s[2:].strip()
    return s.replace("-", "").replace(" ", "").replace("/", "").replace("_", "")

