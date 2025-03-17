import pandas as pd
from pyproj import Transformer

# Nombre del archivo de entrada y salida
archivo_entrada = "D:/SCRIPTS/IEs_SEM_2025.xlsx"
archivo_salida = "coordenadas_convertidas.xlsx"

# Intentar leer el archivo Excel
try:
    df = pd.read_excel(archivo_entrada)
    print("Archivo leído correctamente.")
except Exception as e:
    print(f"Error al leer el archivo Excel: {e}")
    exit()

# Mostrar las primeras filas del DataFrame
print("Primeras filas del archivo:")
print(df.head())

# Verificar que las columnas correctas existan
if 'X_UTM' not in df.columns or 'Y_UTM' not in df.columns:
    print("Error: El archivo debe contener las columnas 'X_UTM' y 'Y_UTM'.")
    print("Columnas detectadas:", df.columns)
    exit()

# Configurar transformación de UTM 18N a WGS84
transformer = Transformer.from_crs("EPSG:3116", "EPSG:4326", always_xy=True)
print("Transformador creado correctamente.")

# Convertir coordenadas con mensajes de depuración
def convertir_coordenadas(row):
    try:
        lon, lat = transformer.transform(row['X_UTM'], row['Y_UTM'])
        return pd.Series([lon, lat])
    except Exception as e:
        print(f"Error al convertir {row['X_UTM']}, {row['Y_UTM']}: {e}")
        return pd.Series([None, None])

df[['Longitud', 'Latitud']] = df.apply(convertir_coordenadas, axis=1)

# Mostrar resultados parciales
print("Primeras filas con coordenadas convertidas:")
print(df.head())

# Guardar en un nuevo archivo Excel
try:
    df.to_excel(archivo_salida, index=False)
    print(f"Conversión completada. Archivo guardado como: {archivo_salida}")
except Exception as e:
    print(f"Error al guardar el archivo: {e}")
