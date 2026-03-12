# -- convertidorjson.py --
# Conversor de archivos Excel a JSON (para carpeta /data)
# Para archivos Normas, listado de clientes Firmas de inspectores

import os
import json
import pandas as pd
from datetime import date, datetime, time
from tkinter import filedialog, messagebox, Tk


# CONFIGURACIÓN
PROJECT_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
DATA_DIR = os.path.join(PROJECT_DIR, "data")
os.makedirs(DATA_DIR, exist_ok=True)  # Crea carpeta /data si no existe


# FUNCIONES
def normalizar_valor_para_json(valor):
    """Convierte valores de Excel/pandas a tipos seguros para JSON."""
    if valor is None:
        return None

    # Detecta NaN/NaT sin fallar con tipos no escalares.
    try:
        if pd.isna(valor):
            return None
    except (TypeError, ValueError):
        pass

    if isinstance(valor, (pd.Timestamp, datetime, date, time)):
        return valor.isoformat()

    # Convierte escalares numpy/pandas a tipos nativos de Python.
    if hasattr(valor, "item"):
        try:
            valor = valor.item()
        except Exception:
            pass

    if isinstance(valor, dict):
        return {str(k): normalizar_valor_para_json(v) for k, v in valor.items()}

    if isinstance(valor, (list, tuple, set)):
        return [normalizar_valor_para_json(v) for v in valor]

    if isinstance(valor, bytes):
        try:
            return valor.decode("utf-8")
        except UnicodeDecodeError:
            return valor.hex()

    # Si no es serializable de forma nativa, se guarda como texto.
    try:
        json.dumps(valor, ensure_ascii=False)
        return valor
    except TypeError:
        return str(valor)


def convertir_excel_a_json(file_path: str) -> str:
    """
    Convierte un archivo Excel (.xlsx o .xls) a JSON
    y lo guarda en la carpeta /data del proyecto.
    
    Retorna la ruta completa del archivo JSON generado.
    """
    temp_output_path = None
    try:
        # Leer Excel
        df = pd.read_excel(file_path)
        if df.empty:
            raise ValueError("El archivo Excel está vacío o sin datos.")

        # Convertir a lista de diccionarios con tipos seguros para JSON.
        records = [
            {columna: normalizar_valor_para_json(valor)
             for columna, valor in fila.items()}
            for fila in df.to_dict(orient="records")
        ]

        # Nombre base del archivo
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        timestamp = datetime.now().strftime("%Y%m%d")
        json_filename = f"{base_name}_{timestamp}.json"

        # Ruta de salida
        output_path = os.path.join(DATA_DIR, json_filename)

        # Se escribe primero a un archivo temporal para evitar JSON truncados.
        temp_output_path = f"{output_path}.tmp"

        # Guardar JSON con formato legible
        with open(temp_output_path, "w", encoding="utf-8") as f:
            json.dump(
                records,
                f,
                ensure_ascii=False,
                indent=2,
                allow_nan=False,
            )

        os.replace(temp_output_path, output_path)

        print(f"✅ Archivo convertido y guardado en: {output_path}")
        return output_path

    except Exception as e:
        if temp_output_path and os.path.exists(temp_output_path):
            os.remove(temp_output_path)
        raise RuntimeError(f"Error al convertir Excel a JSON: {e}")


def abrir_y_convertir():
    """
    Abre un diálogo para seleccionar un archivo Excel y lo convierte automáticamente a JSON.
    Guarda el resultado en /data.
    """
    root = Tk()
    root.withdraw()  # Oculta ventana principal
    try:
        file_path = filedialog.askopenfilename(
            title="Seleccionar archivo Excel para convertir a JSON",
            filetypes=[("Archivos Excel", "*.xlsx;*.xls")]
        )
        if not file_path:
            print("Operación cancelada.")
            return None

        json_path = convertir_excel_a_json(file_path)
        messagebox.showinfo("Conversión completada", f"JSON generado:\n{json_path}")
        return json_path

    except Exception as e:
        messagebox.showerror("Error", str(e))
    finally:
        root.destroy()

# EJECUCIÓN DIRECTA
if __name__ == "__main__":
    print("=== CONVERTIDOR JSON ===")
    print("Seleccione un archivo Excel para convertirlo en JSON...")
    abrir_y_convertir()