# 📦 SIGRAMA PO Tracker & Master Hub (Industria 4.0)

Sistema integral de **Rastreo y Gestión de Órdenes de Compra (POs)**, **Extracción OCR Inteligente de Documentos Membretados (.PDF / .MSG)** y **Trazabilidad Ciberfísica 360°** para **Industria Sigrama**.

---

## 🚀 Descripción General

El **PO Tracker & Master Hub** actúa como el punto de inicio de la cadena de valor digital (**Digital Thread**) en la suite de aplicaciones de Industria Sigrama, conectando el requerimiento comercial y de compras directamente con la manufactura en taller y la logística de salida:

`
+-----------------------------------------------------------------------------------+
|                        ECOSISTEMA CIBERFISICO SIGRAMA                             |
+--------------------------+------------------------------+-------------------------+
|   1. CAPTURA Y MASTER    |   2. MANUFACTURA PLANTA      |  3. LOGISTICA Y SALIDA  |
|      (PO Tracker App)    |   (App Corte y Doblez)       |  (App Remisiones)       |
|                          |                              |                         |
|   • OCR de POs (.PDF/.MSG) |   • Programacion OFs / CNC   |  • Consolidacion Tarimas|
|   • Validacion Partidas  |   • Avance Laser y Doblez    |  • Folios de Remision   |
|   • Estatus y Saldos     |   • Liberacion en Piso       |  • Despacho a Cliente   |
+--------------------------+------------------------------+-------------------------+
`

---

## ✨ Características Principales

1. **📊 Dashboard Ejecutivo 360°**: Indicadores en tiempo real de piezas requeridas, fabricadas, remisionadas y saldos pendientes.
2. **🎯 Matriz de Control 360° (Producción + Remisión)**: Cruce automático entre sigrama_database.xlsx (Corte-Doblez) y BD_Detalle_Tarimas.xlsx (Remisiones).
3. **⚡ Motor OCR Espacial de POs Sigrama**:
   - Detección precisa de cabecera oficial (Folio, Fecha, Proveedor, Solicitante, Requisición, Comprador).
   - Extracción tabular de partidas (SKU, Descripción, Cantidad, Unidad, Precios, Fechas).
   - Soporte para archivos **.PDF directos** y correos **.MSG de Microsoft Outlook** con extracción de adjuntos y planos anexos.
4. **📋 Matriz de Órdenes y Ficha de Trazabilidad**: Visualización detallada de cada orden, tarimas asociadas y receptores.
5. **📁 Carga Masiva y Formulario Guiado**: Carga por plantillas oficiales de Excel o captura manual.
6. **🔄 Interoperabilidad Automática**: Sincronización continua de catálogos BD_POs_Cabecera.xlsx y BD_Requerimientos_POs.xlsx.
7. **📘 Módulo de Manual & Arquitectura Industria 4.0**: Documentación técnica, diagramas y guía interactiva integrada.

---

## 🛠️ Stack Tecnológico

| Capa | Tecnología |
| :--- | :--- |
| **Frontend / UI** | Streamlit 1.40+, Plotly Express, HTML5/CSS3 Corporativo Sigrama (#EC2024, #111111) |
| **Motor OCR & Parsing** | PyMuPDF 1.28 (itz), extract-msg 0.56, RegEx espacial |
| **Persistencia** | SQLite Relacional (po_tracker.db), Pandas 2.2+, OpenPyXL 3.1+ |
| **Sincronización** | Enlace relacional con emisiones-de-materiales y pp_corte_doblez |

---

## 💻 Instalación y Puesta en Marcha

### Prerrequisitos
- Python 3.10 o superior
- Git

### 1. Clonar el Repositorio
`ash
git clone https://github.com/TU_USUARIO/sigrama_po_tracker.git
cd sigrama_po_tracker
`

### 2. Crear y Activar Entorno Virtual
`ash
# Windows
python -m venv venv
venv\Scripts\activate

# Linux / MacOS
python3 -m venv venv
source venv/bin/activate
`

### 3. Instalar Dependencias
`ash
pip install -r requirements.txt
`

### 4. Iniciar la Aplicación
`ash
streamlit run app.py
`

La plataforma estará disponible en http://localhost:8501 (o en el puerto configurado).

---

## 📁 Estructura del Proyecto

`
sigrama_po_tracker/
├── .streamlit/
│   └── config.toml             # Configuración de tema visual y servidor
├── data/
│   ├── po_tracker.db           # Base de datos SQLite maestra
│   ├── BD_POs_Cabecera.xlsx    # Espejo de cabeceras de PO sincronizado
│   └── BD_Requerimientos_POs.xlsx # Espejo de partidas sincronizado
├── app.py                      # Aplicación principal Streamlit
├── config.py                   # Configuración general y normalizadores
├── corte_doblez_sync.py        # Módulo de cruce con Planta (Corte y Doblez)
├── db_manager.py               # Gestión de base de datos SQLite y CRUD
├── excel_importer.py           # Importador masivo de plantillas Excel
├── pdf_parser.py               # Motor OCR espacial y desempaquetador de .MSG
├── remisiones_sync.py          # Módulo de cruce con Logística (Remisiones)
├── requirements.txt            # Dependencias del proyecto
├── logo_sigrama.png            # Logotipo oficial de la empresa
├── .gitignore                  # Reglas de exclusión para Git
└── README.md                   # Documentación oficial
`

---

## 🏢 Industria Sigrama S.A. de C.V.
*Automatización, Control de Procesos y Manufactura Metálica.*
