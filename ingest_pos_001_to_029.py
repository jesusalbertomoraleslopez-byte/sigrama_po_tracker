import sys
sys.stdout.reconfigure(encoding='utf-8')
import os
import re
import argparse
from pathlib import Path
import pandas as pd

from pdf_parser import extract_attachments_from_msg, parse_po_pdf
import db_manager

BASE_DIR = Path(r'Z:\01 - PLANTA METALES\01 - ORDENES DE COMPRA')
CORREOS_DIR = BASE_DIR / '00 - CORREO DE COMPRAS'

# Diccionario de proyectos conocidos para POs historicas
KNOWN_PROJECTS = {
    '26020482': 'RENO 4',
    '26020711': 'RENO 6',
    '26030771': 'SIGRAMA LASER',
    '26030892': 'RENO 4',
    '26030939': 'SIGRAMA LASER',
    '26030989': 'SIGRAMA LASER',
    '26031070': 'RENO 4',
    '26031038': 'LC8 20K',
    '26031039': 'LC8 20K',
    '26031316': 'LC8 20K',
    '26031317': 'LC8 20K',
    '26031328': 'SWBD RENO 4',
    '26031329': 'SWBD RENO 4',
    '26031431': 'SWBD RENO 5',
    '26031459': 'SWBD RENO 6',
    '26031530': 'LC8',
    '26031605': 'RENO 5',
    '26031699': 'LC8 CLOUD',
    '26031700': 'LC8 CLOUD',
    '26031702': 'LC8 CLOUD',
    '26031797': 'LC8',
    '26031839': 'SWBD RENO 5',
    '26031893': 'SWBD RENO 6',
    '26032025': 'LC8',
    '26032561': 'LC8 20K',
    '26032809': 'ALM SWBD META',
    '26032811': 'ALM SWBD SOUTH VALLEY',
    '26032815': 'SWBD RENO 4',
}

def scan_orders_to_ingest():
    """
    Escanea sistemáticamente Z: para mapear exactamente cada INT-0001 a INT-0029
    con su archivo .msg y su respectivo archivo PDF de Orden de Compra.
    """
    orders = {} # id_str -> dict(msg_path, pdf_filter, default_po)
    
    # INT 0001 y INT 0002
    orders['INT-0001'] = {
        'msg': CORREOS_DIR / 'INT 0001 - OC 2602-0482 SIGRAMA LASER.msg',
        'po': '2602-0482'
    }
    orders['INT-0002'] = {
        'msg': CORREOS_DIR / 'INT 0002 - OC 2602-0482 SIGRAMA LASER.msg',
        'po': '2602-0482'
    }
    orders['INT-0003'] = {
        'msg': CORREOS_DIR / 'INT 0003 - OC 2602-0711 SIGRAMA LASER                        PINTURA ANSI 61 RENO 6.msg',
        'po': '2602-0711'
    }
    orders['INT-0004'] = {
        'msg': CORREOS_DIR / 'INT 0004 - OC 2603-0771 SIGRAMA LASER .msg',
        'po': '2603-0771'
    }
    orders['INT-0005'] = {
        'msg': CORREOS_DIR / 'INT 0005 - OC 2603-0892 SIGRAMA LASER .msg',
        'po': '2603-0892'
    }
    orders['INT-0006'] = {
        'msg': CORREOS_DIR / 'INT 0006 - OC 2603-0939 SIGRAMA LASER.msg',
        'po': '2603-0939'
    }
    orders['INT-0007'] = {
        'msg': CORREOS_DIR / 'INT 0007 - OC 2603-0989 SIGRAMA LASER.msg',
        'po': '2603-0989'
    }
    orders['INT-0008'] = {
        'msg': CORREOS_DIR / 'INT 0008 - 2603-1070 SIGRAMA PLANTA METALES.msg',
        'po': '2603-1070'
    }
    orders['INT-0009'] = {
        'msg': CORREOS_DIR / 'INT 0009 - 2603-1038 SIGRAMA PLANTA METALES .msg',
        'po': '2603-1038'
    }
    orders['INT-0010'] = {
        'msg': CORREOS_DIR / 'INT 0010-0011 - RV_ 2603-1316_ 2603-1317 SIGRAMA METALES.msg',
        'po': '2603-1316'
    }
    orders['INT-0011'] = {
        'msg': CORREOS_DIR / 'INT 0010-0011 - RV_ 2603-1316_ 2603-1317 SIGRAMA METALES.msg',
        'po': '2603-1317'
    }
    orders['INT-0012'] = {
        'msg': CORREOS_DIR / 'INT 0012-0013 - OC 2603-1328 _ 2603-1329 SIGRAMA METALES.msg',
        'po': '2603-1328'
    }
    orders['INT-0013'] = {
        'msg': CORREOS_DIR / 'INT 0012-0013 - OC 2603-1328 _ 2603-1329 SIGRAMA METALES.msg',
        'po': '2603-1329'
    }
    orders['INT-0014'] = {
        'msg': CORREOS_DIR / 'INT 0014 - OC 2603-1431 SIGRAMA METALES.msg',
        'po': '2603-1431'
    }
    orders['INT-0015'] = {
        'msg': CORREOS_DIR / 'INT 0015 - OC 2603-1459 SIGRAMA METALES.msg',
        'po': '2603-1459'
    }
    orders['INT-0016'] = {
        'msg': CORREOS_DIR / 'INT 0016 - OC 2603-1530 SIGRAMA METALES.msg',
        'po': '2603-1530'
    }
    
    # 017 en carpeta PO 017
    fld17 = list(BASE_DIR.glob('PO 017*'))
    if fld17:
        m17 = list(fld17[0].glob('*.msg'))
        if m17:
            orders['INT-0017'] = {'msg': m17[0], 'po': '2603-1605'}
            
    # 018, 019, 020
    msg_18_20 = CORREOS_DIR / 'INT 0018-19 - RV_ OC 2603-1699 _ 2603-1700 _ 2603-1702 SIGRAMA METALES.msg'
    orders['INT-0018'] = {'msg': msg_18_20, 'po': '2603-1699'}
    orders['INT-0019'] = {'msg': msg_18_20, 'po': '2603-1700'}
    orders['INT-0020'] = {'msg': msg_18_20, 'po': '2603-1702'}
    
    # 021
    orders['INT-0021'] = {
        'msg': CORREOS_DIR / 'INT 0021 - OC 2603-1797 SIGRAMA METALES.msg',
        'po': '2603-1797'
    }
    
    # 022 a 026 en sus carpetas
    for num, po_t in [(22, '2603-1039'), (23, '2603-1839'), (24, '2603-1893'), (25, '2603-2025'), (26, '2603-2561')]:
        sub = list(BASE_DIR.glob(f'PO 0{num}*'))
        if sub:
            msgs = list(sub[0].glob('*.msg'))
            if msgs:
                orders[f'INT-{num:04d}'] = {'msg': msgs[0], 'po': po_t}
                
    # 027 a 029
    orders['INT-0027'] = {
        'msg': CORREOS_DIR / 'INT 0027 - RV ORDEN DE COMPRA 2603-2809 SIGRAMA METALES.msg',
        'po': '2603-2809'
    }
    orders['INT-0028'] = {
        'msg': CORREOS_DIR / 'INT 0028 - ORDEN DE COMPRA 26032811 SIGRAMA METALES RV 22396 ACERO PINTADO GRIS ANSI 61 SIGRAMA METALES.msg',
        'po': '2603-2811'
    }
    orders['INT-0029'] = {
        'msg': CORREOS_DIR / 'INT 0029 - OC 2603-2815 SIGRAMA METALES.msg',
        'po': '2603-2815'
    }
    
    return orders

def run_ingestion(apply=False):
    if not BASE_DIR.exists():
        print(f"❌ Error: La ruta de red {BASE_DIR} no es accesible.")
        return False
        
    orders = scan_orders_to_ingest()
    print(f"📁 Órdenes localizadas para procesamiento: {len(orders)} / 29")
    print(f"Modo: {'⚡ APLICANDO A BASE DE DATOS' if apply else '🔍 SIMULACIÓN (DRY-RUN)'}\n")
    
    success_count = 0
    saved_items_total = 0
    
    for id_str in sorted(orders.keys()):
        cfg = orders[id_str]
        msg_path = cfg['msg']
        expected_po = cfg['po']
        po_nodash = expected_po.replace('-', '')
        
        if not msg_path or not msg_path.exists():
            print(f"[{id_str}] ⚠️ Archivo .msg no encontrado: {msg_path}")
            continue
            
        try:
            with open(msg_path, 'rb') as fp:
                msg_bytes = fp.read()
            info = extract_attachments_from_msg(msg_bytes)
            atts = info.get('attachments', [])
            
            # Buscar el PDF que coincida con expected_po
            matched_pdf = None
            for a in atts:
                fn = a['filename']
                fn_low = fn.lower()
                if fn_low.endswith('.pdf') and not any(w in fn_low for w in ['plano', 'drawing', 'cotizacion', 'req cotizacion']):
                    if po_nodash in fn or expected_po in fn:
                        matched_pdf = a
                        break
                        
            # Si no hubo coincidencia por PO, buscar cualquier PDF oficial
            if not matched_pdf:
                for a in atts:
                    fn_low = a['filename'].lower()
                    if fn_low.endswith('.pdf') and not any(w in fn_low for w in ['plano', 'drawing', 'cotizacion', '11-', '12-', 'pp', 'p1', 'p2', '1s-']):
                        matched_pdf = a
                        break
                        
            if not matched_pdf:
                print(f"[{id_str}] ⚠️ No se encontró PDF oficial para PO {expected_po} en {msg_path.name}")
                continue
                
            pdf_bytes = matched_pdf['data']
            pdf_name = matched_pdf['filename']
            
            cab, parts = parse_po_pdf(pdf_bytes, email_context={'id_interno': id_str, 'po_detectada': po_nodash})
            
            # Ajustar datos de cabecera
            cab['id_interno'] = id_str
            cab['po'] = expected_po # Guardar con formato estándar
            cab['archivo_correo'] = msg_path.name
            cab['archivo_pdf'] = pdf_name
            
            # Proyecto limpio
            curr_prj = str(cab.get('proyecto', '')).strip()
            if not curr_prj or curr_prj in ('nan', 'POR DEFINIR') or 'DIA' in curr_prj:
                cab['proyecto'] = KNOWN_PROJECTS.get(po_nodash, 'PROYECTO CLIENTE')
                
            if apply:
                # 1. Guardar PO y partidas
                ok, msg_err = db_manager.save_po(cab, parts, usuario='Ingesta Histórica', push_to_gh=False)
                if not ok:
                    print(f"[{id_str}] ❌ Error guardando PO: {msg_err}")
                    continue
                    
                # 2. Guardar archivos binarios en la BD y data/correos/
                db_manager.save_archivo_adjunto(expected_po, id_str, pdf_name, 'pdf', pdf_bytes)
                db_manager.save_archivo_adjunto(expected_po, id_str, msg_path.name, 'msg', msg_bytes)
                
            tot_h = float(cab.get('total', 0.0))
            f_sol = cab.get('fecha_solicitada', '')
            print(f"[{id_str}] ✅ {expected_po} | {len(parts):2d} partidas | Prj: {cab['proyecto'][:20]:20s} | Entrega: {f_sol:10s} | ")
            success_count += 1
            saved_items_total += len(parts)
            
        except Exception as e:
            print(f"[{id_str}] ❌ Excepción procesando {expected_po}: {e}")
            
    print(f"\n─────────────────────────────────────────────────────────────")
    print(f"Resultado final: {success_count} / {len(orders)} órdenes procesadas con éxito.")
    print(f"Total de partidas detalladas: {saved_items_total}")
    
    if apply:
        print("Sincronizando exportación a Excel...")
        db_manager.export_sync_to_excel()
        print("Exportación a Excel completada con éxito.")
        
    return True

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Ingesta de Órdenes Históricas INT-0001 a INT-0029')
    parser.add_argument('--apply', action='store_true', help='Aplica y guarda los cambios en po_tracker.db')
    args = parser.parse_args()
    
    run_ingestion(apply=args.apply)

