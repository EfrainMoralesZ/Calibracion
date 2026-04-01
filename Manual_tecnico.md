# Manual Técnico — Sistema de Calibración V&C

> Versión: 1.1 · Fecha: Abril 2026  
> Audiencia: Ejecutivos técnicos y desarrolladores responsables de mantenimiento, despliegue y evolución del sistema.

---

## Tabla de contenidos

1. [Descripción general](#1-descripción-general)
2. [Requisitos del entorno](#2-requisitos-del-entorno)
3. [Instalación y configuración inicial](#3-instalación-y-configuración-inicial)
4. [Arquitectura del sistema](#4-arquitectura-del-sistema)
5. [Módulos principales](#5-módulos-principales)
6. [Persistencia de datos](#6-persistencia-de-datos)
7. [Roles y permisos](#7-roles-y-permisos)
8. [Generación de documentos PDF](#8-generación-de-documentos-pdf)
9. [Flujos de negocio clave](#9-flujos-de-negocio-clave)
10. [Convenciones de código](#10-convenciones-de-código)
11. [Problemas conocidos y soluciones](#11-problemas-conocidos-y-soluciones)
12. [Guía de mantenimiento](#12-guía-de-mantenimiento)

---

## 1. Descripción general

El **Sistema de Calibración V&C** es una aplicación de escritorio construida en **Python 3.11.9** con la biblioteca de interfaz gráfica **CustomTkinter**. Su propósito es gestionar el ciclo completo de calibración de inspectores (ejecutivos técnicos), incluyendo:

- Seguimiento de acreditaciones por norma.
- Programación y registro de visitas a clientes.
- Captura y consulta de calificaciones trimestrales.
- Generación de documentos PDF de supervisión y reportes trimestrales.
- Administración de catálogos (normas, usuarios, clientes).
- Dashboard de desempeño con curvas de aprendizaje.

La aplicación opera completamente **offline**, sin base de datos relacional: toda la persistencia se realiza sobre archivos JSON locales.

---

## 2. Requisitos del entorno

| Componente | Versión mínima | Notas |
|---|---|---|
| Python | 3.11.9 | Recomendado usar entorno virtual (`venv`) |
| CustomTkinter | última estable | Interfaz gráfica moderna |
| ReportLab | última estable | Generación de PDFs |
| Pillow | última estable | Carga de imágenes (logo, plantilla PDF) |
| Sistema operativo | Windows 10/11 | Optimizado para Windows; DPI awareness configurado mediante `ctypes` |

> Las tres dependencias Python se declaran en `requirements.txt`.

---

## 3. Instalación y configuración inicial

### 3.1 Crear y activar entorno virtual

```powershell
# Crear entorno
python -m venv .venv

# Activar (PowerShell)
.\.venv\Scripts\Activate.ps1
```

### 3.2 Instalar dependencias

```powershell
python -m pip install -r requirements.txt
```

### 3.3 Verificar archivos de datos base

Antes de ejecutar la aplicación, confirmar que los siguientes archivos existen y tienen contenido válido:

```
data/
  BD-Calibracion.json   ← Base principal de ejecutivos y acreditaciones
  Usuarios.json         ← Credenciales y roles
  Normas.json           ← Catálogo de normas
  Clientes.json         ← Catálogo de clientes y direcciones
```

El archivo `data/app_state.json` se genera automáticamente en la primera ejecución.

### 3.4 Assets gráficos requeridos

```
img/
  alerta.png         ← Ícono de alerta para notificaciones UI
  icono.ico          ← Ícono de la ventana principal
  logo.png           ← Logo mostrado en la pantalla de login
  medalla_bronce.png ← Medalla bronce para sistema de medallas trimestral
  medalla_oro.png    ← Medalla oro para sistema de medallas trimestral
  medalla_plata.png  ← Medalla plata/platino para sistema de medallas trimestral
  plantilla.png      ← Fondo (imagen LETTER) usado en PDFs de supervisión
```

### 3.5 Ejecutar la aplicación

```powershell
python app.py
```

---

## 4. Arquitectura del sistema

### 4.1 Diagrama de capas

```
┌──────────────────────────────────────────────────────────────┐
│                        app.py (Shell)                        │
│     Ventana principal · Navegación · Gestión de roles        │
├──────────┬────────────┬──────────────┬───────────────────────┤
│ login.py │dashboard.py│calendario.py │ trimestral.py         │
│          │configura-  │supervision.py│ criterioEvaluacion.py │
│          │ciones.py   │              │                       │
├──────────┴────────────┴──────────────┴───────────────────────┤
│       calibration_controller.py (CalibrationController)      │
│    Lógica de negocio · Persistencia JSON · Historial         │
├──────────────────────────────────────────────────────────────┤
│                   Documentos PDF.py/                         │
│  FormatoSupervision.py · ReporteTrimestral.py                │
│  CriterioEvaluacionTecnica.py                                │
├──────────────────────────────────────────────────────────────┤
│                       data/  (JSON)                          │
└──────────────────────────────────────────────────────────────┘
```

### 4.2 Flujo de arranque

1. `app.py` configura el comportamiento DPI en Windows (`_configure_windows_dpi_behavior`).
2. Instancia `CalibrationController` (carga catálogos, consolida histórico).
3. Muestra `LoginView`; al autenticar dispara `_handle_login(username, password)`.
4. `app.py` construye las secciones disponibles según el rol (`available_sections`, `page_factories`).
5. La navegación lateral llama a `page_factories[section]()` para renderizar cada vista bajo demanda.

---

## 5. Módulos principales

### 5.1 `app.py` — Shell visual

Responsabilidades:
- Configuración DPI (`ctypes` + `ctk.deactivate_automatic_dpi_awareness`).
- Paleta de colores global (`STYLE`) y tokens tipográficos (`FONTS`).
- Construcción de la barra de navegación horizontal según `available_sections`.
- Ciclo `refresh_all_views()` que propaga cambios de catálogo a todas las vistas activas.
- KPIs en header: tarjetas de promedio, alertas y medallas (Oro, Platino, Bronce) para ejecutivos y admins.

**Paleta oficial:**

| Token | Hex | Uso |
|---|---|---|
| `primario` | `#ECD925` | Acento principal, botones CTA |
| `secundario` | `#282828` | Fondos oscuros, texto |
| `exito` | `#008D53` | Estados positivos |
| `advertencia` | `#ff1500` | Alertas y estados críticos |
| `peligro` | `#d74a3d` | Estados de peligro |
| `fondo` / `surface` | `#F8F9FA` | Fondo de vistas y paneles |

### 5.2 `calibration_controller.py` — CalibrationController

Núcleo de la lógica de negocio y acceso a datos. Expone métodos para:

| Grupo | Métodos destacados |
|---|---|
| Catálogos | `get_inspectors()`, `get_norms()`, `get_clients()`, `get_users()` |
| Autenticación | `authenticate()`, `is_admin()`, `has_full_access()`, `is_executive_role()` |
| Secciones | `available_sections()` — retorna secciones visibles según rol |
| Historial | `get_history(inspector)`, `save_history_entry()` |
| Visitas | `get_visits()`, `save_visit()`, `delete_visit()` |
| Trimestral | `list_trimestral_scores()`, `save_trimestral_score()`, `send_trimestral_scores()` |
| Boletas | `get_boleta(inspector, year, quarter)`, `save_boleta()` |
| Reporte normas | `get_norms_report()`, `save_norms_report_entry()` |

Las consultas de alto uso están cacheadas con `@lru_cache` y se invalidan en `reload()`.

**Sistema de medallas:** definido en `TRIMESTRAL_MEDAL_RULES` — Oro (≥100), Platino (≥90), Bronce (≥80).

### 5.3 `dashboard.py` — DashboardView

- Tarjetas horizontales de normas acreditadas por inspector.
- Panel de visitas recientes con selector de modo.
- Curva de aprendizaje dibujada sobre `tk.Canvas` con firma de datos para evitar re-renders innecesarios.
- Selector de ejecutivo; al cambiar dispara recarga parcial sin afectar otras vistas.

### 5.4 `calendario.py` — CalendarView

- Cuadrícula mensual de visitas con navegación por mes.
- Panel lateral con formulario de creación/edición de visitas (solo admin).
- Checklist de normas por visita y guardado de reporte de normas.
- Gráfico de demanda mensual de normas (solo admin).
- Filtros por ejecutivo, cliente y estado.

### 5.5 `trimestral.py` — TrimestralView

- **Pestaña Captura trimestral** (solo acceso completo): alta, edición y envío de calificaciones por norma/trimestre/inspector.
- **Pestaña Historial trimestral**: cards con score por norma, paginadas (9 por página); curva de evolución histórica.
- Estado `CRITICO` cuando `score < 90`, marcado visualmente en boletas.
- Sistema de medallas: Oro (≥100), Platino (≥90), Bronce (≥80) con íconos visuales.
- Al enviar (`sent_at` registrado), los ejecutivos pasan a modo visualización.

### 5.6 `supervision.py` — PrincipalView

- Tarjetas de supervisión por ejecutivo con formulario de evaluación integrado.
- Checklist de protocolo (preguntas de supervisión) y captura de actividades.
- Selección de carpeta de evidencias (imágenes) para generar PDF.
- Genera PDF de supervisión vía `FormatoSupervision.py`.
- Registro de supervisión guardado en `data/historico/<ejecutivo>/historico.json`.

### 5.7 `criterioEvaluacion.py` — CriteriaEvaluationView

- Formulario independiente por ejecutivo/norma para evaluación de criterios técnicos.
- Tarjetas por cliente con paginación de 12 en 12.
- Roles de gestión eligen ejecutivo antes de la NOM; ejecutivos seleccionan su propio perfil.
- Genera PDF de criterios de evaluación técnica vía `CriterioEvaluacionTecnica.py`.
- Acuerdos y criterios archivados en `data/clientes/acuerdos/<CLIENTE>/`.

### 5.8 `configuraciones.py` — ConfigurationView

Pestañas:
1. **Normas**: alta, edición y baja de normas (NOM, nombre, sección).
2. **Usuarios**: gestión de cuentas y roles.
3. **Clientes**: datos fiscales, contratos y direcciones por cliente.
4. **Ejecutivos**: asignación de normas acreditadas mediante checklist dinámico.

Tras guardar cualquier catálogo dispara `on_change()` que propaga `refresh_all_views()` en `app.py`.

### 5.9 `login.py` — LoginView

- Carga logo desde `img/logo.png` o `img/Logo.png` usando `CTkImage` + Pillow.
- Valida credenciales contra `Usuarios.json` vía `controller.authenticate()`.
- Soporta múltiples roles (ver sección 7).
- Manejo de foco automático con `after()` para compatibilidad DPI.

---

## 6. Persistencia de datos

### 6.1 Estructura de archivos

```
data/
├── BD-Calibracion.json          # Inspectores: nombre, acreditaciones por norma
├── BD-Calibracion.xlsx          # Respaldo Excel de la base de calibración
├── Usuarios.json                # Usuarios: username, password, rol, nombre_completo
├── Normas.json                  # Normas: NOM, nombre, sección
├── Clientes.json                # Clientes: RFC, contrato, direcciones, servicio
├── app_state.json               # Estado operativo (generado en runtime)
├── reporte de normas.json       # Reportes de normas por visita (month=YYYY-MM)
├── clientes/
│   └── acuerdos/<CLIENTE>/      # Acuerdos y criterios de evaluación por cliente
├── historico/
│   └── <Nombre_Apellido>/
│       ├── historico.json       # Entradas de supervisión del ejecutivo
│       ├── visitas.json         # Visitas asociadas al ejecutivo
│       └── boletas/
│           └── <año>/
│               └── T<n>_boleta.json  # Boleta trimestral (boleta_status=CRITICO si <90)
├── reportes/                    # Reportes generados
├── visitas/
│   └── semana_<YYYY-MM-DD>/
│       └── visitas.json         # Visitas de la semana
└── trimestral/
    └── T<n>_<año>/
        └── trimestral.json      # Calificaciones del trimestre
```

### 6.2 Nombrado de carpetas de histórico

Las carpetas en `data/historico/` se nombran con el patrón `Nombre_Apellido1_Apellido2` generado por `_safe_folder_name()`. La identidad normalizada (sin acentos, sin símbolos) se usa para comparaciones mediante `_folder_identity()`, lo que garantiza coincidencia aunque los JSON usen acentuación distinta.

### 6.3 Funciones de bajo nivel

| Función | Archivo | Propósito |
|---|---|---|
| `_read_json(path, default)` | `calibration_controller.py` | Lectura segura; retorna `default` si el archivo no existe o está vacío |
| `_write_json(path, payload)` | `calibration_controller.py` | Escritura atómica; crea directorios intermedios si no existen |
| `_safe_slug(value)` | `calibration_controller.py` | Convierte texto en slug alfanumérico con guiones bajos |
| `_safe_folder_name(value)` | `calibration_controller.py` | Sanitiza nombres para rutas de directorio |

---

## 7. Roles y permisos

### 7.1 Roles del sistema

Roles con **acceso completo** (equivalentes a admin): `admin`, `gerente`, `sub gerente`, `coordinador operativo`, `coordinadora en fiabilidad`.

Otros roles: `talento humano`, `supervisor`, `ejecutivo tecnico`, `especialidades`.

### 7.2 Permisos por sección

| Sección | Acceso completo | Talento humano | Supervisor | Ejecutivo técnico / Especialidades |
|---|:---:|:---:|:---:|:---:|
| Supervisión | ✔ | ✔ | ✔ | ✔ (lectura) |
| Criterios | ✔ | ✗ | ✗ | ✔ |
| Dashboard | ✔ | ✔ | ✗ | ✗ |
| Calendario | ✔ (edición) | ✗ | ✔ (edición) | ✔ (lectura) |
| Trimestral | ✔ (captura + envío) | ✗ | ✗ | ✔ (visualización) |
| Configuraciones | ✔ | ✗ | ✗ | ✗ |

Los roles se determinan en `Usuarios.json` (campo `role`). La normalización de roles se realiza en `_normalize_role_name()`. La lógica de secciones visibles está en `CalibrationController.available_sections()` y las fábricas de vistas en `app.py` bajo `_build_page_factories()`.

---

## 8. Generación de documentos PDF

### 8.1 Formato de Supervisión (`Documentos PDF.py/FormatoSupervision.py`)

- Motor: **ReportLab** (`SimpleDocTemplate` + `canvas` personalizado).
- Fondo: `img/plantilla.png` dibujado en cada página usando `PageTemplate` con `onPage`.
- Encabezado: dibujado programáticamente por página con datos del inspector y visita.
- Numeración: canvas numerado (`NumberedCanvas`) que omite la última página si está vacía.
- Evidencias: imágenes de una carpeta seleccionada por el usuario se agregan al final, con match por nombre de archivo contra SKU / ITEM / CODIGO / UPC de las actividades.

### 8.2 Reporte Trimestral (`Documentos PDF.py/ReporteTrimestral.py`)

- Estructura de tablas por norma con calificaciones y estado (Mayor / Menor).
- Colores de estado determinados por `_status_color()`:
  - `"mayor"` → verde oscuro `#006B3C`
  - `"menor"` → rojo oscuro `#B94A2C`
- Encabezado con datos del periodo y ejecutivo.

### 8.3 Criterio de Evaluación Técnica (`Documentos PDF.py/CriterioEvaluacionTecnica.py`)

- Genera ficha de consultas normativas por ejecutivo/norma/cliente.
- Encabezado con datos de Verificación & Control UVA.
- Fondo: reutiliza `img/plantilla.png` como plantilla.
- Tablas con criterios de evaluación técnica.

### 8.4 Invocación desde la UI

- **PDF de supervisión**: se genera desde `supervision.py` (`PrincipalView`).
- **PDF de criterios**: se genera desde `criterioEvaluacion.py` (`CriteriaEvaluationView`).
- **PDF trimestral**: se genera desde la vista trimestral.

---

## 9. Flujos de negocio clave

### 9.1 Supervisión de visita

```
Admin abre PrincipalView (Supervisión)
  → Selecciona ejecutivo y norma
  → Llena checklist de protocolo y actividades
  → Selecciona carpeta de evidencias (opcional)
  → Genera PDF  →  FormatoSupervision.py
  → Registro guardado en data/historico/<ejecutivo>/historico.json
```

### 9.2 Programación de visita

```
Admin abre CalendarView
  → Selecciona semana en cuadrícula mensual
  → Formulario lateral: inspector, cliente, dirección, servicio, fecha, hora
  → Guarda en data/visitas/semana_<YYYY-MM-DD>/visitas.json
  → Reporte de normas guardado en data/reporte de normas.json
```

### 9.3 Ciclo trimestral

```
Admin (TrimestralView, pestaña Captura):
  → Selecciona inspector, norma, año, trimestre
  → Captura score (0–100)
  → Envía trimestre (sent_at registrado)
     → Si score < 90 → boleta_status = "CRITICO"
     → Boleta guardada en data/historico/<ejecutivo>/boletas/<año>/T<n>_boleta.json

Ejecutivo (TrimestralView, pestaña Historial):
  → Visualiza cards con su score por norma
  → Consulta curva de evolución histórica
```

### 9.4 Administración de catálogos

```
Admin (ConfigurationView):
  → Modifica normas / usuarios / clientes / acreditaciones
  → on_change() → app.py.refresh_all_views()
  → Todas las vistas activas recargan caches
```

### 9.5 Evaluación de criterios técnicos

```
Usuario abre CriteriaEvaluationView (Criterios)
  → Selecciona ejecutivo (gestión) o usa perfil propio (ejecutivo)
  → Selecciona norma y cliente
  → Captura criterios de evaluación técnica
  → Genera PDF  →  CriterioEvaluacionTecnica.py
  → Acuerdo archivado en data/clientes/acuerdos/<CLIENTE>/
```

---

## 10. Convenciones de código

| Aspecto | Convención |
|---|---|
| Imports | `from __future__ import annotations` en todos los módulos para soporte de tipos adelantados |
| Rutas | Siempre `pathlib.Path`; nunca concatenación de strings |
| JSON | `_read_json` / `_write_json` de `calibration_controller.py`; nunca `open()` directo desde vistas |
| Normalización de nombres | `_normalize_person_name()` para comparaciones de identidad; nunca comparar cadenas literales con acentos |
| Widgets CTk | `CTkComboBox` usa `variable=`, **no** `textvariable=` |
| Imágenes | `CTkImage` + Pillow; nunca `PhotoImage` de tkinter (incompatible con HiDPI) |
| Refresco UI | Debounce con `after()` / `after_cancel()`; guardar sólo la vista visible |
| Archivos `__pycache__` | No versionar; agregar `__pycache__/` a `.gitignore` |

---

## 11. Problemas conocidos y soluciones

### 11.1 Ventana transparente/fantasma en multi-monitor

**Causa:** CustomTkinter 5.2.2 ajusta `-alpha` a `0.15` al detectar cambio de DPI entre monitores.  
**Solución:** `app.py` llama `ctk.deactivate_automatic_dpi_awareness()` antes de crear la ventana y configura el contexto DPI con `ctypes` usando `SetProcessDpiAwarenessContext`.

### 11.2 Lag al cambiar de sección (refrescos en cascada)

**Causa:** En Windows, un bind global a `<Configure>` puede disparar `refresh_all_views` en cada redimensionamiento de píxel.  
**Solución:** Aplicar debounce (`after` / `after_cancel`) y refrescar solo la sección visible. Guards `winfo_exists()` / `TclError` en callbacks de redimensionamiento.

### 11.3 Cards de principal sin datos aunque el JSON tiene registros

**Diagnóstico:** Validar que `refresh_all_views()` se ejecuta tras guardar y que no hay excepciones silenciosas en la capa de UI. Revisar que `principal_rows` no esté cacheado con datos obsoletos (llamar `controller.reload()` si es necesario).

### 11.4 Trimestral muestra "sin curva" con calificaciones existentes

**Causa:** La curva se alimenta de `get_history()` en vez de `list_trimestral_scores()`.  
**Solución:** El detalle y la curva deben leer de `list_trimestral_scores` (calificaciones por norma/trimestre).

### 11.5 Nombres con acentos no coinciden en Dashboard/Trimestral

**Causa:** `Usuarios.json` y `BD-Calibracion.json` pueden usar acentuación distinta.  
**Solución:** Siempre comparar con identidad normalizada (`_normalize_person_name` / `_folder_identity`).



---

## 12. Guía de mantenimiento

### 12.1 Agregar una nueva norma

1. Abrir la aplicación con rol `admin`.
2. Ir a **Configuraciones → Normas** → `Nueva norma`.
3. Completar NOM, nombre y sección → Guardar.
4. En **Configuraciones → Ejecutivos**, asignar la norma a los inspectores que la acrediten.

### 12.2 Agregar un nuevo usuario / ejecutivo

1. **Configuraciones → Usuarios** → `Nuevo usuario`.
2. Completar nombre completo, username, contraseña y rol (ej. `admin`, `gerente`, `supervisor`, `ejecutivo tecnico`, etc.).
3. Si el rol es `ejecutivo tecnico` o `especialidades`, ir a **Configuraciones → Ejecutivos** y agregar sus normas acreditadas.
4. En **BD-Calibracion.json** puede ser necesario agregar el registro del inspector si el sistema no lo crea automáticamente.

### 12.3 Actualizar dependencias

```powershell
python -m pip install --upgrade customtkinter reportlab pillow
python -m pip freeze > requirements.txt
```

Verificar compatibilidad de CustomTkinter con la versión Python instalada antes de actualizar.

### 12.4 Validar integridad del código sin ejecutar la app

```powershell
python -m py_compile app.py calibration_controller.py dashboard.py configuraciones.py calendario.py login.py trimestral.py supervision.py criterioEvaluacion.py ui_shared.py
python -m py_compile "Documentos PDF.py/FormatoSupervision.py" "Documentos PDF.py/ReporteTrimestral.py" "Documentos PDF.py/CriterioEvaluacionTecnica.py"
```

### 12.5 Respaldar datos

Los datos críticos del sistema son los archivos JSON en `data/`. Hacer copia antes de cualquier migración:

```powershell
Copy-Item -Recurse .\data .\data_backup_$(Get-Date -Format "yyyyMMdd")
```

### 12.6 Regenerar histórico consolidado

Si se detectan inconsistencias en el histórico de un ejecutivo, llamar `controller.reload()` desde la consola interactiva de Python para forzar la reconsolidación desde los archivos fuente.

---

*Sistema de Calibración V&C — Documentación técnica interna.*  
*Actualizado: Abril 2026*
