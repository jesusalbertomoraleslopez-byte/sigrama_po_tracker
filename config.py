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

# Órdenes de compra históricas completadas al 100% previas a la implementación de sistemas
HISTORICAL_COMPLETED_INTS = {
    'INT-0001', 'INT-0002',
    'INT-0010', 'INT-0011', 'INT-0014', 'INT-0015', 'INT-0016', 
    'INT-0017', 'INT-0018', 'INT-0019', 'INT-0020', 'INT-0021', 
    'INT-0022', 'INT-0023', 'INT-0024'
}

HISTORICAL_COMPLETED_POS = {
    '2602-0482', '2602-0482 (2)', '26020482',
    '2603-1316', '2603-1317', '2603-1431', '2603-1459', '2603-1530',
    '2603-1605', '2603-1699', '2603-1700', '2603-1702', '2603-1797',
    '2603-1039', '2603-1839', '2603-1893',
    '26031316', '26031317', '26031431', '26031459', '26031530',
    '26031605', '26031699', '26031700', '26031702', '26031797',
    '26031039', '26031839', '26031893'
}

def is_historical_completed(id_interno="", po=""):
    """Verifica si una orden de compra corresponde a las entregas históricas completadas al 100%."""
    import re
    if id_interno:
        s_id = str(id_interno).strip().upper()
        if s_id in HISTORICAL_COMPLETED_INTS:
            return True
        digits = re.sub(r'[^0-9]', '', s_id)
        if digits:
            formatted_id = f"INT-{int(digits):04d}"
            if formatted_id in HISTORICAL_COMPLETED_INTS:
                return True
    if po:
        s_po = str(po).strip().upper()
        if s_po in HISTORICAL_COMPLETED_POS:
            return True
        norm_po = normalize_po(s_po)
        if norm_po in HISTORICAL_COMPLETED_POS:
            return True
    return False

