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

## Módulos

| Archivo | Rol |
|---|---|
| `app.py` | Shell visual, navegación y vista principal (`EvaluationDialog`) |
| `index.py` | `CalibrationController`: lógica de negocio, persistencia JSON y caché |
| `dashboard.py` | Tarjetas de normas, curva de aprendizaje y panel de visitas |
| `calendario.py` | Programación de visitas, cuadrícula mensual y reporte de normas |
| `trimestral.py` | Captura y consulta de calificaciones trimestrales por norma |
| `configuraciones.py` | Administración de normas, usuarios, clientes y acreditaciones |
| `login.py` | Autenticación por roles (`admin` / `ejecutivo`) |
| `Documentos PDF.py/FormatoSupervision.py` | Generación de PDF de supervisión con evidencias |
| `Documentos PDF.py/ReporteTrimestral.py` | Generación de reporte trimestral en PDF |

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
