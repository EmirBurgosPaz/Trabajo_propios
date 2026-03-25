"""
inventory_recreation
===================
Sistema por el cual recreamos el inventario de la tabla de SAP a un inventario diario, 
debido a que en la tabla de SAP no genera transacciones de movimientos, si no hay documentos que
realizan algun movimiento

Autor:  Emir_B
Fecha:  2026-03-24
"""

# ---------------------------------------------------------------------------
# Importaciones estándar
# ---------------------------------------------------------------------------
import os
import sys
import logging
from pathlib import Path

# ---------------------------------------------------------------------------
# Importaciones de terceros (pip install ...)
# ---------------------------------------------------------------------------
from dotenv import load_dotenv
import pandas as pd
import numpy as np
# ---------------------------------------------------------------------------
# Importaciones locales
# ---------------------------------------------------------------------------
from update_files import refresh_all_connections

# ---------------------------------------------------------------------------
# Configuración de logging
# ---------------------------------------------------------------------------
logging.basicConfig(
level=logging.INFO,
format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Constantes
# ---------------------------------------------------------------------------

VERSION = "0.1.0"

# ---------------------------------------------------------------------------
# Variables del enviroment
# ---------------------------------------------------------------------------

env_path = Path(r"C:\Owncloud\Codigos\.env")

load_dotenv(dotenv_path=env_path)

try:
    FILE_PATH_INVENTORY = os.getenv("FILE_PATH_INVENTORY")
    FILE_OUTPUT_INVENTORY = os.getenv("FILE_OUTPUT_INVENTORY")
    ID_INVENTARIO = os.getenv("ID_INVENTARIO")
    FECHA_TABLA = os.getenv("FECHA_TABLA")
except ValueError as error_file:
    logger.info(f"No se encontraron las direcciones de archivos'{error_file}'")

# ---------------------------------------------------------------------------
# Clases
# ---------------------------------------------------------------------------



# ---------------------------------------------------------------------------
# Funciones
# ---------------------------------------------------------------------------
def selector_rango_fecha(dataframe: pd.DataFrame) -> bool:
    """
    Seleccionamos la fecha maxima y minima de la tabla
    """

    year_min = dataframe[FECHA_TABLA].dt.year.min()
    mes_min = dataframe[FECHA_TABLA].dt.month.iloc[0]
    dia_min = dataframe[FECHA_TABLA].dt.day.iloc[0]
    year_max = dataframe[FECHA_TABLA].dt.year.max()
    mes_max = dataframe[FECHA_TABLA].dt.month.iloc[-1]
    dia_max = dataframe[FECHA_TABLA].dt.day.iloc[-1]

    fecha_min = pd.Timestamp(f'{year_min}-{mes_min}-{dia_min}')
    fecha_max = pd.Timestamp(f'{year_max}-{mes_max}-{dia_max}')

    return fecha_min, fecha_max

# ---------------------------------------------------------------------------
# Punto de entrada
# ---------------------------------------------------------------------------
def main() -> int:
    """Función principal del script."""
    logger.info("Iniciando v%s", VERSION)

    refresh_all_connections(FILE_PATH_INVENTORY)

    data_movimientos_existencias = pd.read_excel(FILE_PATH_INVENTORY)

    existencias_por_dia = data_movimientos_existencias.pivot_table(index=FECHA_TABLA,
                    columns=ID_INVENTARIO, fill_value= "n"
                    ).stack().reset_index()

    fecha_min, fecha_max = selector_rango_fecha(existencias_por_dia)

    rango_fechas = pd.date_range(start=fecha_min, end=fecha_max, freq='D')
    
    data_movimientos_existencias['tiene_movimiento'] = 1
   
    pivot = (
        data_movimientos_existencias
        .pivot_table(index=FECHA_TABLA, columns=ID_INVENTARIO, fill_value=np.nan)
    )

    pivot = pivot.reindex(rango_fechas, fill_value=np.nan)
    pivot.index.name = FECHA_TABLA

    existencias_por_dia = (
        pivot
        .stack(future_stack=True)
        .reset_index()
    )

    
    existencias_por_dia = existencias_por_dia.sort_values([ID_INVENTARIO, FECHA_TABLA])

    existencias_por_dia['Existencia'] = (
        existencias_por_dia
        .groupby(ID_INVENTARIO)['Existencia']
        .transform(lambda x: x.ffill().bfill())
    )

    existencias_por_dia.to_csv(FILE_OUTPUT_INVENTORY)

    logger.info("Ejecución finalizada correctamente.")


    return 0


if __name__ == "__main__":
    sys.exit(main())

