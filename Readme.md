# Sistema de Calibracion V&C

Aplicacion de escritorio en Python 3.11.9 con CustomTkinter para operar el flujo de calibracion, seguimiento de inspectores, dashboard por desempeno, visitas y configuracion de catalogos.

## Requisitos

- Python 3.11.9
- Dependencias del archivo requirements.txt

## Instalacion

```bash
python -m pip install -r requirements.txt
```

## Ejecucion

```bash
python app.py
```

## Estructura funcional

- app.py: shell visual, navegacion y vista principal.
- index.py: logica central, persistencia JSON, roles, historial y generacion de PDF.
- dashboard.py: tarjetas, graficas y curva de desempeno por inspector.
- calendario.py: asignacion y consulta de visitas.
- configuraciones.py: administracion de normas y usuarios.
- login.py: acceso por roles.

## Persistencia

- data/BD-Calibracion.json: base principal de inspectores y acreditaciones.
- data/Usuarios.json: accesos y roles.
- data/Normas.json: catalogo de normas.
- data/Clientes.json: clientes y direcciones para visitas.
- data/app_state.json: estado operativo generado por la aplicacion.
- data/historico/<ejecutivo>: historial y documentos por inspector.
