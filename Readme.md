# Sistema de Calibración V&C

Aplicación de escritorio en **Python 3.11.9** con **CustomTkinter** para gestionar el ciclo completo de calibración de inspectores: supervisión de visitas, seguimiento de acreditaciones por norma, calificaciones trimestrales, dashboard de desempeño y generación de documentos PDF.

> Para documentación técnica detallada consulta el [Manual Técnico](Manual_tecnico.md).

---

## Requisitos

| Componente | Versión |
|---|---|
| Python | 3.11.9 |
| CustomTkinter | última estable |
| ReportLab | última estable |
| Pillow | última estable |
| Sistema operativo | Windows 10/11 |

## Instalación

```powershell
# 1. Crear y activar entorno virtual
python -m venv .venv
.\.venv\Scripts\Activate.ps1

# 2. Instalar dependencias
python -m pip install -r requirements.txt
```

## Ejecución

```powershell
python app.py
```

---

## Mapa de Archivos

### Núcleo de aplicación

| Archivo | Qué hace |
|---|---|
| `app.py` | Punto de entrada de la app, `CalibrationApp`, navegación, layout general e inicialización de UI. |
| `calibration_controller.py` | `CalibrationController`: lógica de negocio, acceso/persistencia JSON, cachés, métricas y generación de documentos. |
| `Principal.py` | Vistas y diálogos de la sección Principal: selección de norma, captura de supervisión y acciones de calibración. |
| `ui_shared.py` | Estilos, tipografías y utilidades visuales compartidas (escalado, enfoque y centrado de ventanas). |
| `runtime_paths.py` | Rutas de recursos/runtime para ejecución normal y empaquetada (`.exe`). |

### Vistas funcionales

| Archivo | Qué hace |
|---|---|
| `login.py` | Pantalla de acceso y autenticación por roles (`admin` / `ejecutivo`). |
| `dashboard.py` | Panel de indicadores: tarjetas de normas, curva de aprendizaje y resumen operativo. |
| `calendario.py` | Programación y seguimiento de visitas en calendario, con reporte de normas aplicadas. |
| `trimestral.py` | Captura, consulta y envío de calificaciones trimestrales por inspector/norma. |
| `configuraciones.py` | Administración de catálogos: normas, usuarios, clientes y direcciones. |
| `formulario.py` | Utilidades/formulario legado de apoyo para flujos históricos del proyecto. |

### Generación de PDF

| Archivo | Qué hace |
|---|---|
| `Documentos PDF.py/FormatoSupervision.py` | Construye el PDF del formato de supervisión con resultados y evidencias. |
| `Documentos PDF.py/ReporteTrimestral.py` | Construye el PDF del reporte trimestral (global o por usuario). |

### Scripts de build y utilidades

| Archivo | Qué hace |
|---|---|
| `build_exe.bat` | Compila ejecutable Windows con PyInstaller (`dist/CalibracionVC.exe`). |
| `CalibracionVC.spec` | Especificación generada por PyInstaller para el empaquetado. |
| `tools/convertidorjson.py` | Script auxiliar para transformación/normalización de datos JSON. |
| `requirements.txt` | Dependencias Python del proyecto. |

## Roles

| Sección | admin | ejecutivo |
|---|:---:|:---:|
| Principal (supervisiones) | ✔ | ✗ |
| Dashboard | ✔ | ✗ |
| Calendario | ✔ (edición) | ✔ (lectura) |
| Trimestral | ✔ (captura + envío) | ✔ (visualización) |
| Configuraciones | ✔ | ✗ |

---

## Persistencia de datos

Toda la información se almacena localmente en archivos JSON:

```
data/
├── BD-Calibracion.json        # Ejecutivos y acreditaciones por norma
├── Usuarios.json              # Credenciales y roles
├── Normas.json                # Catálogo de normas
├── Clientes.json              # Clientes, contratos y direcciones
├── app_state.json             # Estado operativo (generado en runtime)
├── reporte de normas.json     # Reportes por visita
├── historico/<ejecutivo>/     # Supervisiones, visitas y boletas trimestrales
├── visitas/semana_<fecha>/    # Visitas semanales
└── trimestral/T<n>_<año>/     # Calificaciones trimestrales
```

## Assets requeridos

```
img/
├── logo.png        # Logo para pantalla de login
├── plantilla.png   # Fondo LETTER para PDFs de supervisión
└── icono.ico       # Ícono de la ventana principal
```
