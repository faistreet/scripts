import pandas as pd
import unicodedata

def remove_accents(text):
    """
    Limpia un texto eliminando acentos, quitando el carácter '°', reemplazando '¬' por '-'
    y cambiando 'ñ' a 'n' (y 'Ñ' a 'N'). Si el valor no es cadena se retorna sin modificar.
    """
    if not isinstance(text, str):
        return text
    nfkd = unicodedata.normalize('NFKD', text)
    # Elimina marcas diacríticas (acentos)
    text_without_accents = "".join([c for c in nfkd if not unicodedata.combining(c)])
    # Reemplazar '¬' por '-' y eliminar '°'
    text_modified = text_without_accents.replace('¬', '-').replace('°', '')
    # Cambiar ñ/Ñ
    text_clean = text_modified.replace('ñ', 'n').replace('Ñ', 'N')
    return text_clean

def clean_text_columns(df):
    """
    Aplica la función remove_accents a todas las celdas de tipo string en el DataFrame.
    """
    for col in df.select_dtypes(include='object').columns:
        df[col] = df[col].apply(lambda x: remove_accents(x) if isinstance(x, str) else x)
    return df

def main():
    try:
        # 1. Cargar el archivo Excel
        input_file = 'D:/SCRIPTS/SOLICITUD-PERSEO_30NOV2024.xlsx'
        df = pd.read_excel(input_file)

        # 2. Limpiar los nombres de las columnas para eliminar espacios extra
        df.columns = df.columns.str.strip()
        
        # 3. Limpiar la lista de columnas a eliminar (asegurarse de quitar espacios)
        columnas_a_eliminar = [
            'CODIGO_DANE',
            'DANE_ANTERIOR',
            'METODOLOGIA',
            'MATRICULA_CONTRATADA',
            'PAIS_ORIGEN',
            'ESTABLECIMIENTO EDUCATIVO',
            'Due.SEDE EDUCATIVA',
            'Due.PRESTACION DEL SERVICIO'
        ]
        columnas_a_eliminar = [col.strip() for col in columnas_a_eliminar]

        # 4. Eliminar las columnas innecesarias (solo se analiza la fila de encabezados)
        df = df.drop(columns=columnas_a_eliminar, errors='ignore')

        # 5. Limpiar los nombres de las columnas (quitar acentos, reemplazar ¬ y °, etc.)
        df.columns = [remove_accents(col) for col in df.columns]
        print("Columnas después de eliminar y limpiar nombres:", df.columns.tolist())

        # 6. Renombrar columnas (las claves deben coincidir con los nombres ya limpios)
        nuevos_nombres = {
            'NRO_DOCUMENTO': 'idestudiante',
            'NOMBRE1': 'nombre1',
            'NOMBRE2': 'nombre2',
            'APELLIDO1': 'apellido1',
            'APELLIDO2': 'apellido2',            
            'CONS_SEDE': 'daneduefk',
            'TEL': 'telefono',
            'DIRECCION_RESIDENCIA': 'direccion',
            'CORREO': 'correopersonal',
            'GRADO': 'grado',
            'GRUPO': 'grupo',
            'FECHA_NACIMIENTO': 'fechanacimiento',
            'REPITENTE': 'repitente',
            'EXTRAEDAD': 'extraedad',
            'MATRICULA': 'matricula',  # Recuerda: 'MATRÍCULA' se transforma a 'MATRICULA' tras limpiar acentos
            'NUI': 'nnui',
            'PER_ID': 'perid',
            'NOMBRE PAIS': 'nacionalidad',
            'SISBEN IV': 'sisben',
            'ESTRATO': 'estrato',
            'ESTRATEGIA': 'estrategias',
            'ESTADO': 'estado',
        }
        df = df.rename(columns=nuevos_nombres)
        print("Columnas después de renombrar:", df.columns.tolist())

        # 7. Limpiar el contenido de cada celda de texto (aplicando remove_accents)
        df = clean_text_columns(df)

        # 8. Insertar una nueva columna en una posición específica.
        # Por ejemplo, insertarla en la posición 3 (índice 2) con un valor predeterminado
        posicion_nueva = 8  # Índice 8, que corresponde a la posición 9 en base 1 
        df.insert(posicion_nueva, 'correoinstitucional', '')

        # 9. Reordenar las columnas según el orden deseado.
        # Es fundamental que la lista incluya TODAS las columnas actuales del DataFrame.
        nuevo_orden = [
            'idestudiante',  # Primera columna
            'nombre1',       # Segunda
            'nombre2',
            'apellido1',
            'apellido2',
            'daneduefk',
            'telefono',
            'direccion',
            'correoinstitucional', # La que hemos insertado en la posición 9
            'correopersonal',
            'grado',
            'grupo',
            'fechanacimiento',
            'repitente',
            'extraedad',
            'matricula',
            'nnui',
            'perid',
            'nacionalidad',
            'sisben',
            'estrato',
            'estrategias',
            'estado'
        ]
        # Reordenamos el DataFrame. Asegúrate que 'nuevo_orden' incluye todas las columnas actuales.
        df = df.reindex(columns=nuevo_orden)
        print("Columnas tras reordenar e insertar la nueva:", df.columns.tolist())

        # 10. Convertir las columna a su respectivo tipo de dato
        # Lista de columnas que deseas convertir a string
        string_cols = [
            'idestudiante', 'nombre1', 'nombre2', 'apellido1', 'apellido2',
            'telefono', 'direccion', 'correoinstitucional', 
            'correopersonal', 'grupo', 'repitente', 'extraedad', 
            'matricula', 'nnui', 'nacionalidad', 'sisben', 
            'estrategias', 'estado'
        ]

        # Convertir de forma masiva estas columnas a tipo 'string'
        df[string_cols] = df[string_cols].astype('string')

        # Lista de columnas que deseas convertir a int
        int_cols = ['daneduefk', 'grado', 'perid', 'estrato']
        
        # Convertir de forma masiva estas columnas a tipo 'int'
        df[int_cols] = df[int_cols].astype('Int64')
        
        # Convertir 'fechanacimiento' a datetime y luego a string con el formato deseado: AAAA-MM-DD
        # Convertir toda la columna a datetime usando el origen de Excel
        #df['fechanacimiento'] = pd.to_datetime(df['fechanacimiento'], unit='d', origin='1899-12-30')
        df['fechanacimiento'] = df['fechanacimiento'].dt.strftime('%Y-%m-%d')

        # 11. Guardar el DataFrame modificado en un nuevo archivo Excel
        output_file = 'D:/SCRIPTS/modificacion.xlsx'
        df.to_excel(output_file, index=False)
        print(f"Archivo modificado guardado exitosamente en: {output_file}")

    except Exception as e:
        print("Ocurrió un error al procesar el archivo:", e)

if __name__ == "__main__":
    main()
