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
| `supervision.py` | Vista de supervisión (`PrincipalView`): tarjetas por ejecutivo, formulario de evaluación y generación de PDF de supervisión. |
| `criterioEvaluacion.py` | Vista de criterios de evaluación técnica (`CriteriaEvaluationView`): formulario por ejecutivo/norma/cliente con generación de PDF de criterios y evidencias. |
| `ui_shared.py` | Estilos, tipografías y utilidades visuales compartidas (escalado, enfoque y centrado de ventanas). |
| `runtime_paths.py` | Rutas de recursos/runtime para ejecución normal y empaquetada (`.exe`). |

### Vistas funcionales

| Archivo | Qué hace |
|---|---|
| `login.py` | Pantalla de acceso y autenticación multirrol. |
| `dashboard.py` | Panel de indicadores: tarjetas de normas, curva de aprendizaje y resumen operativo. |
| `calendario.py` | Programación y seguimiento de visitas en calendario, con reporte de normas aplicadas. |
| `trimestral.py` | Captura, consulta y envío de calificaciones trimestrales por inspector/norma. Sistema de medallas (Oro, Platino, Bronce). |
| `configuraciones.py` | Administración de catálogos: normas, usuarios, clientes y direcciones. |

### Generación de PDF

| Archivo | Qué hace |
|---|---|
| `Documentos PDF.py/FormatoSupervision.py` | Construye el PDF del formato de supervisión con resultados y evidencias. |
| `Documentos PDF.py/ReporteTrimestral.py` | Construye el PDF del reporte trimestral (global o por usuario). |
| `Documentos PDF.py/CriterioEvaluacionTecnica.py` | Construye el PDF de ficha de consultas normativas con criterios de evaluación técnica. |

### Scripts de build y utilidades

| Archivo | Qué hace |
|---|---|
| `build_exe.bat` | Compila ejecutable Windows con PyInstaller (`dist/CalibracionVC.exe`). |
| `CalibracionVC.spec` | Especificación generada por PyInstaller para el empaquetado. |
| `tools/convertidorjson.py` | Script auxiliar para transformación/normalización de datos JSON. |
| `requirements.txt` | Dependencias Python del proyecto. |

## Roles

Roles con **acceso completo** (equivalentes a admin): `admin`, `gerente`, `sub gerente`, `coordinador operativo`, `coordinadora en fiabilidad`.

| Sección | Acceso completo | Talento humano | Supervisor | Ejecutivo técnico / Especialidades |
|---|:---:|:---:|:---:|:---:|
| Supervisión | ✔ | ✔ | ✔ | ✔ (lectura) |
| Criterios | ✔ | ✗ | ✗ | ✔ |
| Dashboard | ✔ | ✔ | ✗ | ✗ |
| Calendario | ✔ (edición) | ✗ | ✔ (edición) | ✔ (lectura) |
| Trimestral | ✔ (captura + envío) | ✗ | ✗ | ✔ (visualización) |
| Configuraciones | ✔ | ✗ | ✗ | ✗ |

---

## Persistencia de datos

Toda la información se almacena localmente en archivos JSON:

```
data/
├── BD-Calibracion.json        # Ejecutivos y acreditaciones por norma
├── BD-Calibracion.xlsx        # Respaldo Excel de la base de calibración
├── Usuarios.json              # Credenciales y roles
├── Normas.json                # Catálogo de normas
├── Clientes.json              # Clientes, contratos y direcciones
├── app_state.json             # Estado operativo (generado en runtime)
├── reporte de normas.json     # Reportes por visita
├── clientes/
│   └── acuerdos/<CLIENTE>/    # Acuerdos y criterios por cliente
├── historico/<ejecutivo>/     # Supervisiones, visitas y boletas trimestrales
├── reportes/                  # Reportes generados
├── visitas/semana_<fecha>/    # Visitas semanales
└── trimestral/T<n>_<año>/     # Calificaciones trimestrales
```

## Assets requeridos

```
img/
├── alerta.png         # Ícono de alerta para notificaciones UI
├── icono.ico          # Ícono de la ventana principal
├── logo.png           # Logo para pantalla de login
├── medalla_bronce.png # Medalla bronce para sistema de medallas trimestral
├── medalla_oro.png    # Medalla oro para sistema de medallas trimestral
├── medalla_plata.png  # Medalla plata/platino para sistema de medallas trimestral
└── plantilla.png      # Fondo LETTER para PDFs de supervisión
```
