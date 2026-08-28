import pandas as pd
from db_manager import save_po, get_all_pos

pos_data = [
    {
        'cab': {
            'po': '26083186',
            'id_interno': 'INT-0050',
            'proyecto': 'CLOUD',
            'solicitante': 'ESTEFANIA IBARRA',
            'comprador': 'Alejandra Arellano Machado',
            'fecha_llegada': '2026-08-14',
            'fecha_solicitada': '2026-08-25',
            'fecha_pedido': '2026-08-14',
            'archivo_correo': 'INT 0050 - OC 2608-3186 SIGRAMA METALES.msg',
            'archivo_pdf': '2608-3186 SIGRAMA METALES.PDF',
            'total': 0.0
        },
        'partidas': [
            {
                'item_no': 1,
                'sku_cliente': 'SWB01431',
                'clave_sku': 'PP19380-03',
                'descripcion_producto': '382 X 10H BLANK DOOR',
                'cantidad_requerida': 32.0,
                'unidad': 'PIEZA',
                'precio_unitario': 345.50,
                'precio_total': 11056.00,
                'fecha_entrega': '2026-08-25',
                'parcialidad': 'P1'
            }
        ]
    },
    {
        'cab': {
            'po': '26083189',
            'id_interno': 'INT-0051',
            'proyecto': 'META',
            'solicitante': 'ESTEFANIA IBARRA',
            'comprador': 'Alejandra Arellano Machado',
            'fecha_llegada': '2026-08-14',
            'fecha_solicitada': '2026-08-25',
            'fecha_pedido': '2026-08-14',
            'archivo_correo': 'INT 0051 - OC 2608-3189 SIGRAMA METALES.msg',
            'archivo_pdf': '2608-3189 SIGRAMA METALES.PDF',
            'total': 0.0
        },
        'partidas': [
            {
                'item_no': 1,
                'sku_cliente': 'ISSIV05401',
                'clave_sku': '11-7761-02',
                'descripcion_producto': '11-7761-02 SOPORTE DE FIJACION META',
                'cantidad_requerida': 8.0,
                'unidad': 'PIEZA',
                'precio_unitario': 290.00,
                'precio_total': 2320.00,
                'fecha_entrega': '2026-08-25',
                'parcialidad': 'P1'
            }
        ]
    },
    {
        'cab': {
            'po': '26083190',
            'id_interno': 'INT-0052',
            'proyecto': 'CLOUD',
            'solicitante': 'ESTEFANIA IBARRA',
            'comprador': 'Alejandra Arellano Machado',
            'fecha_llegada': '2026-08-14',
            'fecha_solicitada': '2026-08-25',
            'fecha_pedido': '2026-08-14',
            'archivo_correo': 'INT 0052 - OC 2608-3190 SIGRAMA METALES.msg',
            'archivo_pdf': '2608-3190 SIGRAMA METALES.PDF',
            'total': 0.0
        },
        'partidas': [
            {
                'item_no': 1,
                'sku_cliente': 'ISSIV05401',
                'clave_sku': '11-7761-02',
                'descripcion_producto': '11-7761-02 SOPORTE DE FIJACION CLOUD',
                'cantidad_requerida': 40.0,
                'unidad': 'PIEZA',
                'precio_unitario': 290.00,
                'precio_total': 11600.00,
                'fecha_entrega': '2026-08-25',
                'parcialidad': 'P1'
            }
        ]
    },
    {
        'cab': {
            'po': '26083191',
            'id_interno': 'INT-0053',
            'proyecto': 'LC8',
            'solicitante': 'ESTEFANIA IBARRA',
            'comprador': 'Alejandra Arellano Machado',
            'fecha_llegada': '2026-08-14',
            'fecha_solicitada': '2026-08-25',
            'fecha_pedido': '2026-08-14',
            'archivo_correo': 'INT 0053 - OC 2608-3191 SIGRAMA METALES.msg',
            'archivo_pdf': '2608-3191 SIGRAMA METALES.PDF',
            'total': 0.0
        },
        'partidas': [
            {
                'item_no': 1,
                'sku_cliente': 'ISSIV05499',
                'clave_sku': '11-A-6014-01',
                'descripcion_producto': '11-A-6014-01 TAPA LATERAL LC8',
                'cantidad_requerida': 1.0,
                'unidad': 'PIEZA',
                'precio_unitario': 850.00,
                'precio_total': 850.00,
                'fecha_entrega': '2026-08-25',
                'parcialidad': 'P1'
            }
        ]
    },
    {
        'cab': {
            'po': '26083177',
            'id_interno': 'INT-0054',
            'proyecto': 'ALM SWBD CDC 736',
            'solicitante': 'ESTEFANIA IBARRA',
            'comprador': 'Alejandra Arellano Machado',
            'fecha_llegada': '2026-08-17',
            'fecha_solicitada': '2026-08-28',
            'fecha_pedido': '2026-08-17',
            'archivo_correo': 'INT 0054 - OC 2608-3177 SIGRAMA METALES.msg',
            'archivo_pdf': '2608-3177 SIGRAMA METALES.PDF',
            'total': 106782.18
        },
        'partidas': [
            {'item_no': 1, 'sku_cliente': 'ISSIV00055', 'clave_sku': '11-A-9836-01', 'descripcion_producto': '11-A-9836-01', 'cantidad_requerida': 16.0, 'unidad': 'PIEZA', 'precio_unitario': 1089.57, 'precio_total': 17433.12, 'fecha_entrega': '2026-09-07', 'parcialidad': 'P1'},
            {'item_no': 2, 'sku_cliente': 'ISSIV00056', 'clave_sku': '12-D-6013-01', 'descripcion_producto': '12-D-6013-01', 'cantidad_requerida': 128.0, 'unidad': 'PIEZA', 'precio_unitario': 185.20, 'precio_total': 23705.60, 'fecha_entrega': '2026-09-07', 'parcialidad': 'P1'},
            {'item_no': 3, 'sku_cliente': 'ISSIV00057', 'clave_sku': 'PP14873-01', 'descripcion_producto': 'PP14873-01', 'cantidad_requerida': 32.0, 'unidad': 'PIEZA', 'precio_unitario': 245.80, 'precio_total': 7865.60, 'fecha_entrega': '2026-09-07', 'parcialidad': 'P1'},
            {'item_no': 4, 'sku_cliente': 'ISSIV00058', 'clave_sku': 'PP14873-02', 'descripcion_producto': 'PP14873-02', 'cantidad_requerida': 40.0, 'unidad': 'PIEZA', 'precio_unitario': 255.40, 'precio_total': 10216.00, 'fecha_entrega': '2026-09-07', 'parcialidad': 'P1'},
            {'item_no': 5, 'sku_cliente': 'ISSIV00059', 'clave_sku': 'P20325-25', 'descripcion_producto': 'P20325-25', 'cantidad_requerida': 64.0, 'unidad': 'PIEZA', 'precio_unitario': 295.00, 'precio_total': 18880.00, 'fecha_entrega': '2026-09-07', 'parcialidad': 'P1'},
            {'item_no': 6, 'sku_cliente': 'ISSIV00060', 'clave_sku': 'P20325-24', 'descripcion_producto': 'P20325-24', 'cantidad_requerida': 64.0, 'unidad': 'PIEZA', 'precio_unitario': 295.00, 'precio_total': 18880.00, 'fecha_entrega': '2026-09-07', 'parcialidad': 'P1'},
            {'item_no': 7, 'sku_cliente': 'ISSIV00061', 'clave_sku': '96-6183-01', 'descripcion_producto': '96-6183-01', 'cantidad_requerida': 40.0, 'unidad': 'PIEZA', 'precio_unitario': 245.00, 'precio_total': 9800.00, 'fecha_entrega': '2026-09-07', 'parcialidad': 'P1'}
        ]
    },
    {
        'cab': {
            'po': '26083235',
            'id_interno': 'INT-0055',
            'proyecto': 'CLOUD',
            'solicitante': 'ESTEFANIA IBARRA',
            'comprador': 'Alejandra Arellano Machado',
            'fecha_llegada': '2026-08-18',
            'fecha_solicitada': '2026-08-28',
            'fecha_pedido': '2026-08-18',
            'archivo_correo': 'INT 0055 - OC 2608-3235 SIGRAMA METALES.msg',
            'archivo_pdf': '2608-3235 SIGRAMA METALES.PDF',
            'total': 34863.85
        },
        'partidas': [
            {'item_no': 1, 'sku_cliente': 'ISSIV00635', 'clave_sku': '12-6355-03', 'descripcion_producto': '15.0 INCH BLANK DOOR', 'cantidad_requerida': 32.0, 'unidad': 'PIEZA', 'precio_unitario': 320.50, 'precio_total': 10256.00, 'fecha_entrega': '2026-08-28', 'parcialidad': 'P1'},
            {'item_no': 2, 'sku_cliente': 'ISSIV00636', 'clave_sku': '12-6355-02', 'descripcion_producto': 'DISTRIBUTION DOOR NON WELDED', 'cantidad_requerida': 40.0, 'unidad': 'PIEZA', 'precio_unitario': 345.00, 'precio_total': 13800.00, 'fecha_entrega': '2026-08-28', 'parcialidad': 'P1'},
            {'item_no': 3, 'sku_cliente': 'ISSIV00637', 'clave_sku': 'PP22114-05', 'descripcion_producto': '382 X 15H INSTRUMENT DOOR', 'cantidad_requerida': 40.0, 'unidad': 'PIEZA', 'precio_unitario': 270.20, 'precio_total': 10807.85, 'fecha_entrega': '2026-08-28', 'parcialidad': 'P1'}
        ]
    },
    {
        'cab': {
            'po': '26083261',
            'id_interno': 'INT-0056',
            'proyecto': 'LC8 47K',
            'solicitante': 'ESTEFANIA IBARRA',
            'comprador': 'Alejandra Arellano Machado',
            'fecha_llegada': '2026-08-21',
            'fecha_solicitada': '2026-08-31',
            'fecha_pedido': '2026-08-21',
            'archivo_correo': 'INT 0056 - OC 2608-3261 SIGRAMA METALES.msg',
            'archivo_pdf': '2608-3261 SIGRAMA METALES.PDF',
            'total': 30058.00
        },
        'partidas': [
            {'item_no': 1, 'sku_cliente': '12-D-6091-08', 'clave_sku': '12-D-6091-08', 'descripcion_producto': '12-D-6091-08 PANEL DE CONTROL LC8', 'cantidad_requerida': 20.0, 'unidad': 'PIEZA', 'precio_unitario': 420.50, 'precio_total': 8410.00, 'fecha_entrega': '2026-08-31', 'parcialidad': 'P1'},
            {'item_no': 2, 'sku_cliente': 'PP7517-10', 'clave_sku': 'PP7517-10', 'descripcion_producto': 'PP7517-10 SOPORTE SUPERIOR LC8', 'cantidad_requerida': 32.0, 'unidad': 'PIEZA', 'precio_unitario': 285.00, 'precio_total': 9120.00, 'fecha_entrega': '2026-08-31', 'parcialidad': 'P1'},
            {'item_no': 3, 'sku_cliente': 'PP7517-01', 'clave_sku': 'PP7517-01', 'descripcion_producto': 'PP7517-01 BASE INFERIOR LC8', 'cantidad_requerida': 64.0, 'unidad': 'PIEZA', 'precio_unitario': 195.75, 'precio_total': 12528.00, 'fecha_entrega': '2026-08-31', 'parcialidad': 'P1'}
        ]
    },
    {
        'cab': {
            'po': '26083358',
            'id_interno': 'INT-0057',
            'proyecto': 'TORONTO',
            'solicitante': 'ESTEFANIA IBARRA',
            'comprador': 'Alejandra Arellano Machado',
            'fecha_llegada': '2026-08-25',
            'fecha_solicitada': '2026-09-08',
            'fecha_pedido': '2026-08-25',
            'archivo_correo': 'INT 0057 - OC 2608-3358 SIGRAMA METALES.msg',
            'archivo_pdf': '2608-3358 SIGRAMA METALES.PDF',
            'total': 0.0
        },
        'partidas': [
            {'item_no': 1, 'sku_cliente': 'ISSIV06049', 'clave_sku': 'P9711-081 TRTO1', 'descripcion_producto': 'REAR DOOR BOLTED LOUVERED 22 INCH', 'cantidad_requerida': 1.0, 'unidad': 'PIEZA', 'precio_unitario': 1250.00, 'precio_total': 1250.00, 'fecha_entrega': '2026-09-08', 'parcialidad': 'P1'},
            {'item_no': 2, 'sku_cliente': 'ISSIV05291', 'clave_sku': 'PP10218-13-TBC3', 'descripcion_producto': 'PLATE 70 & 75 SP REAR TOP RH END', 'cantidad_requerida': 1.0, 'unidad': 'PIEZA', 'precio_unitario': 680.00, 'precio_total': 680.00, 'fecha_entrega': '2026-09-08', 'parcialidad': 'P1'},
            {'item_no': 3, 'sku_cliente': 'ISSIV06050', 'clave_sku': 'PP21139-081 TRTO', 'descripcion_producto': '22W BREAKER DOOR W INSERT IP55 COVER', 'cantidad_requerida': 12.0, 'unidad': 'PIEZA', 'precio_unitario': 890.00, 'precio_total': 10680.00, 'fecha_entrega': '2026-09-08', 'parcialidad': 'P1'}
        ]
    }
]

for p_entry in pos_data:
    ok, msg = save_po(p_entry['cab'], p_entry['partidas'])
    print('Saved', p_entry['cab']['id_interno'], p_entry['cab']['po'], ok)

df_res = get_all_pos()
print(df_res[['id_interno', 'po', 'proyecto', 'comprador', 'total']].to_string())
