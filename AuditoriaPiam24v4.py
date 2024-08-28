import os
import math
import pandas as pd
import numpy as np
from io import BytesIO
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from openpyxl.drawing.image import Image

pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)

estado_financiero = {
    'ac': 'Activa',
    'an': 'Anulada',
    'ca': 'Cancelada'
}

estado_beneficio = {
    'B': 'Beneficiario',
    'E': 'Excluido'
}

valoresValidosTipoId = ['CC', 'DE', 'CE', 'TI', 'PS', 'CA', 'PT']

columnasValidacionObligatoriedad = [
    'ID_TIPO_DOCUMENTO','NUM_DOCUMENTO','CODIGO_ESTUDIANTE','PRO_CONSECUTIVO','ID_MUNICIPIO',
    'FECHA_NACIMIENTO','ID_PAIS_NACIMIENTO','ID_MUNICIPIO_NACIMIENTO','ID_ZONA_RESIDENCIA','ID_ESTRATO',
    'ES_REINTEGRO_ESTD_ANTES_DE1998','AÑO_PRIMER_CURSO','SEMESTRE_PRIMER_CURSO','TELEFONO_CONTACTO','CELULAR','EMAIL_PERSONAL',
    'EMAIL_INSTITUCIONAL','CRED_ACAD_PROGRAMA_RC','SEMESTRES_RC','CREDIT_ACADEM_ACUM_SEM_ANTE','CREDITOSMATRICULADOS',
    'DERECHOS_MATRICULA','SEGURO_ESTUDIANTIL']


# Analisis estructural
def validar_tipo_documento(df):
    if 'ID_TIPO_DOCUMENTO' in df.columns:
        df_validacion_tipoId = df[~df['ID_TIPO_DOCUMENTO'].isin(valoresValidosTipoId)]
        return df_validacion_tipoId
    else:
        return pd.DataFrame()


# Función para ajustar el ancho de las columnas en un archivo Excel
def ajustar_ancho_columnas(writer, dataframe, sheet_name, startrow=0, startcol=0):
    worksheet = writer.sheets[sheet_name]  # Obtener la hoja de cálculo
    for i, col in enumerate(dataframe.columns):
        # Encuentra el ancho máximo entre el nombre de la columna y el contenido de la columna
        max_len = max(dataframe[col].astype(str).map(len).max(), len(col)) + 2
        # Ajustar el ancho de la columna en Excel
        worksheet.set_column(startcol + i, startcol + i, max_len)

# Funcion para la verificación de registros vacios en los campos de datos obligatorios
def obtenerRegistrosVacios(df,columnas):
  registrosVaciosTotal = pd.DataFrame()
  resumenVacios = {}
  for columna in columnas:
    if columna not in df.columns:
      print(f"La columna '{columna}' no existe en el DataFrame.")
      continue
    registrosVacios = df[df[columna].isnull()]
    registrosVacios = registrosVacios.copy()
    resumenVacios[columna] = len(registrosVacios)
    if registrosVacios.empty:
      print(f"No hay registros vacíos en la columna '{columna}'.")
    else:
      print(f"Hay {len(registrosVacios)} registros vacíos en la columna '{columna}'.")
      registrosVacios.loc[:,'BanderaRegistrosVacios'] = columna
      registrosVaciosTotal = pd.concat([registrosVaciosTotal, registrosVacios])
  return registrosVaciosTotal, resumenVacios


# Función para ajustar los registros vacíos en las columnas especificadas
def ajustarRegistrosVacios(df, columnas):
  for columna in columnas:
      if columna not in df.columns:
          print(f"La columna '{columna}' no existe en el DataFrame.")
          continue

      # Verificar el tipo de datos de la columna
      if pd.api.types.is_numeric_dtype(df[columna]):
          # Rellenar los valores vacíos con 0 si la columna es numérica
          df[columna].fillna(0, inplace=True)
          print(f"Los registros vacíos en la columna '{columna}' han sido rellenados con 0.")
      else:
          # Rellenar los valores vacíos con NaN si la columna es alfanumérica
          df[columna].fillna(np.nan, inplace=True)
          print(f"Los registros vacíos en la columna '{columna}' han sido rellenados con NaN.")

  return df


# Funcion para verificar la consistencia en los registros alusivos a creditos academicos del registro calificado
def verificarInconsistenciasCreditos(df, columna):
  if columna not in df.columns:
    print(f"La columna '{columna}' no existe en el DataFrame.")
    return pd.DataFrame()
  df['BanderaCreditosRC'] = df[columna] < 15
  df_inconsistenciasRC = df[df['BanderaCreditosRC']]
  print(f"Se encontraron {len(df_inconsistenciasRC)} programas con inconsistencia en los Creditos exigidos por el Registro Calificado")
  return df_inconsistenciasRC

# Funcion para verificar la correlación entre los creditos RC y los creditos aprobados
def verificarInconsistenciasCreditosCantidad(df, creditosRC, creditosAprobados):
  if creditosRC not in df.columns:
    print(f"La columna '{creditosRC}' no existe en el DataFrame.")
    return pd.DataFrame()
  if creditosAprobados not in df.columns:
    print(f"La columna '{creditosAprobados}' no existe en el DataFrame.")
    return pd.DataFrame()

  def evaluarInconsistenciaCreditos(row):
    if row[creditosRC] < row[creditosAprobados]:
      return 'Creditos RC menor a los creditos aprobados'
    elif row[creditosRC] == row[creditosAprobados]:
      return 'Creditos RC igual a los creditos aprobados'
    else:
      return None

  df['FlCreditosRCAprobados'] = df.apply(evaluarInconsistenciaCreditos, axis=1)
  df_inconsistenciasRCAprobados = df[df['FlCreditosRCAprobados'].notnull()]
  resumen_inconsistencias = df_inconsistenciasRCAprobados['FlCreditosRCAprobados'].value_counts()
  print(f"Se encontraron {len(df_inconsistenciasRCAprobados)} inconsistencias en los Créditos del RC y los aprobados:")
  for inconsistencia, cantidad in resumen_inconsistencias.items():
      print(f"{inconsistencia}: {cantidad} casos")
  return df_inconsistenciasRCAprobados

# Funcion para ajustar los creditos aprobados a partir de la correlacion con la cantidad de creditos RC
def ajustarCreditosAprobados(df, creditosRC, creditosAprobados, numSemestres):
    if creditosRC not in df.columns or creditosAprobados not in df.columns or numSemestres not in df.columns:
        print("Una o más columnas especificadas no existen en el DataFrame.")
        return df
    df_ajuste_menor = df[df[creditosRC] < df[creditosAprobados]].copy()
    df_ajuste_igual = df[df[creditosRC] == df[creditosAprobados]].copy()
    df_ajuste_errado = df[df[creditosRC] == 12].copy()
    if not df_ajuste_menor.empty:
      df_ajuste_menor['CreditosAjustados'] = df_ajuste_menor[creditosRC] - (df_ajuste_menor[creditosRC] / df_ajuste_menor[numSemestres])
      df_ajuste_menor[creditosAprobados] = df_ajuste_menor['CreditosAjustados']
      df_ajuste_menor.drop(columns=['CreditosAjustados'], inplace=True)

    if not df_ajuste_igual.empty:
      df_ajuste_igual['CreditosAjustados'] = df_ajuste_igual[creditosRC] - (df_ajuste_igual[creditosRC] / df_ajuste_igual[numSemestres])
      df_ajuste_igual[creditosAprobados] = df_ajuste_igual['CreditosAjustados']
      df_ajuste_igual.drop(columns=['CreditosAjustados'], inplace=True)

    if not df_ajuste_errado.empty:
      df_ajuste_errado['CreditosAjustados'] = 68
      df_ajuste_errado[creditosRC] = df_ajuste_errado['CreditosAjustados']
      df_ajuste_errado.drop(columns=['CreditosAjustados'], inplace=True)

    df.update(df_ajuste_menor)
    df.update(df_ajuste_igual)
    df.update(df_ajuste_errado)

    print(f"Se han ajustado {len(df_ajuste_menor)} registros donde los créditos RC eran menores a los créditos aprobados.")
    print(f"Se han ajustado {len(df_ajuste_igual)} registros donde los créditos RC eran iguales a los créditos aprobados.")
    print(f"Se han ajustado {len(df_ajuste_igual)} registros donde los créditos RC eran iguales a 12.")

    return df

# Funcion condensadora de estados de ejecucion

def condensar_estados_ejecucion(df):
  condiciones = [
        (df['ESTADO_GIRO'].isin(['Aprobado con giro', 'Renovado con giro'])) &
        (df['Estado Actual'] == 'ac') &
        (df['FLAG ETAPA PAGOS'] == 'No ejecutados')
        & (df['Valor Pagado'] <= 39000),

        (df['ESTADO_GIRO'].isin(['Aprobado con giro', 'Renovado con giro'])) &
        (df['Estado Actual'] == 'ac') &
        (df['FLAG ETAPA PAGOS'] == 'No ejecutados')
        & (df['Valor Pagado'] >= 39000),

        (df['ESTADO_GIRO'] == 'Aprobado con giro') &
        (df['FLAG ETAPA PAGOS'] == 'PAGO MANUAL'),

        (df['ESTADO_GIRO'] == 'Aprobado con giro') &
        (df['FLAG ETAPA PAGOS'] == 'En tramite de pago'),

        (df['ESTADO_GIRO'].isin(['Aprobado con giro', 'Renovado con giro'])) &
        (df['FLAG ETAPA PAGOS'] == 'Pago '),

        (df['ESTADO_GIRO'].isin(['Aprobado con giro', 'Renovado con giro'])) &
        (df['FLAG ETAPA PAGOS'] == 'Reintegro - Pago | Parcial'),

        (df['ESTADO_GIRO'].isin(['Aprobado con giro', 'Renovado con giro'])) &
        (df['FLAG ETAPA PAGOS'] == 'Reintegro 1 - Pago | Parcial'),

        (df['ESTADO_GIRO'].isin(['Aprobado con giro', 'Renovado con giro'])) &
        (df['FLAG ETAPA PAGOS'] == 'Reintegro 1'),

        (df['ESTADO_GIRO'].isin(['Aprobado con giro', 'Renovado con giro'])) &
        (df['Estado Actual'] == 'ca'),

        (df['ESTADO_GIRO'].isin(['No aprobado'])) &
        (df['ESTADO FINANCIERO ICETEX'] == 'PAGO 1 - REINTEGRO SALDO'),

        (df['ESTADO_GIRO'].isin(['No aprobado','No renovado'])) &
        (df['ESTADO FINANCIERO ICETEX'] == 'REINTEGRO ICETEX')]

  valores = ['PAGO B', 'PAGO B | REI B', 'PAGO M', 'PAGO E', 'PAGO A', 'PAGO P | REI P', 'PAGO PP | REI P', 'REINTEGRO A','REINTEGRO B','PAGO ICX | REI ICX', 'REI ICX']
  df['STT EJECUCION'] = np.select(condiciones, valores, default=np.nan)

  condicion_pago = [
      df['STT EJECUCION'] == 'PAGO A',
      df['STT EJECUCION'] == 'PAGO P | REI P',
      df['STT EJECUCION'] == 'PAGO PP | REI P',
      df['STT EJECUCION'] == 'PAGO M',
      df['STT EJECUCION'] == 'PAGO E',
      (df['STT EJECUCION'] == 'PAGO B') & (df['Valor Pagado'] <= 39000),
      (df['STT EJECUCION'] == 'PAGO B | REI B') & (df['Valor Pagado'] >= 39000),
      df['STT EJECUCION'] == 'PAGO ICX | REI ICX']
  valores_pago = [
        df['NETAAPL'],
        df['pagos 1'],
        df['pagos 1'],
        df['NETAAPL'],
        df['NETAAPL'],
        df['NETAAPL'],
        (df['MTRNETA']-df['Valor Pagado']),
        (df['Valor Factura'])]

  df['Pago ejecutado'] = np.select(condicion_pago, valores_pago, default=np.nan)

  condicion_reintegro = [
      df['STT EJECUCION'] == 'PAGO P | REI P',
      df['STT EJECUCION'] == 'PAGO PP | REI P',
      df['STT EJECUCION'] == 'REINTEGRO A',
      (df['STT EJECUCION'] == 'PAGO B | REI B') & (df['Valor Pagado'] >= 39000),
      df['STT EJECUCION'] == 'REINTEGRO B',
      df['STT EJECUCION'] == 'PAGO ICX | REI ICX',
      df['STT EJECUCION'] == 'REI ICX']
  valores_reintegro = [
        df['NETAAPL']- df['pagos 1'],
        df['NETAAPL']- df['pagos 1'],
        df['NETAAPL'],
        (df['MTRNETA']-df['Saldo']-df['SEGURO_ESTUDIANTIL']),
        df['NETAAPL'],
        (df['VALORGIROICETEX']-df['Valor Factura']),
        df['VALORGIROICETEX']]
  df['Reintegro ejecutado'] = np.select(condicion_reintegro, valores_reintegro, default=np.nan)
  return df

# Función para generar el reporte "PlantillaMatriculados"
def generarReportePlantillaMatriculados(df):
    if 'CRED_ACAD_PROGRAMA_RC' not in df.columns or 'SEMESTRES_RC' not in df.columns:
        print("Una o más columnas necesarias para el cálculo de 'CREDIT_ACAD_A_MATRIC_REGU_SEM' no existen en el DataFrame.")
        return None, None

    if (df['SEMESTRES_RC'] == 0).any():
        print("La columna 'SEMESTRES_RC' contiene valores cero, lo que podría causar una división por cero.")
        return None, None
    df['CREDIT_ACAD_A_MATRIC_REGU_SEM'] = ((df['CRED_ACAD_PROGRAMA_RC'] + 2) / df['SEMESTRES_RC']).apply(math.ceil)
    df['APOYO_GOB_NAC_DESCUENTO_VOTAC'] = -(df['VOTO'])
    df['APOYO_GOB_NAC_DESCUENTO_VOTAC'].fillna(0, inplace=True)
    df['APOYO_GOBERNAC_PROGR_PERMANENT'] = 0
    df['APOYO_ALCALDIA_PROGR_PERMANENT'] = 0
    df['DESCUENT_RECURRENTES_DE_LA_IES'] = 0
    df['OTROS_APOYOS_A_LA_MATRICULA'] = 0
    df.loc[df['FONDOICETEX'] == '121943 - 121943 SER ESTUDIOSO CUENTA', 'APOYO_ADICIONAL_GOBERNACIONES'] = df['NETAORD']
    df['APOYO_ADICIONAL_ALCALDIAS'] = 0
    condition = (
        (df['ESTADO CIVF'] == 'E') &
        (df['FONDOICETEX'].isna() | (df['FONDOICETEX'] == '')) &
        ~((df['MERITO'] == '') | (df['MERITO'] == 0) | df['MERITO'].isna())
    )
    df.loc[condition, 'DESCUENTOS_ADICIONALES_IES'] = df['MERITO']
    condition1 = (df['ESTADO CIVF'] == 'E') & ~((df['FONDOICETEX'] == '121943 - 121943 SER ESTUDIOSO CUENTA') |
                   df['FONDOICETEX'].isna() |
                   (df['FONDOICETEX'] == ''))
    df.loc[condition1, 'OTROS_APOYOS_ADICIONALES'] = df['NETAORD']
    df['VAL_NETO_DER_MAT_A_CARGO_EST'] = (
    df['NETAORD'].fillna(0) -
    df['OTROS_APOYOS_ADICIONALES'].fillna(0) -
    df['DESCUENTOS_ADICIONALES_IES'].fillna(0) -
    df['APOYO_ADICIONAL_ALCALDIAS'].fillna(0) -
    df['APOYO_ADICIONAL_GOBERNACIONES'].fillna(0))
    df['VALOR_BRUTO_DERECHOS_COMPLEMEN'] = df['SEGURO_ESTUDIANTIL']
    df['VALOR_NETO_DERECHOS_COMPLEMENT'] = df['SEGURO_ESTUDIANTIL']
    df['CAUSA_NO_ACCESO'] = 0

    columnas_reporte = [
        'ID_TIPO_DOCUMENTO', 'NUM_DOCUMENTO', 'CODIGO_ESTUDIANTE',
        'PRO_CONSECUTIVO', 'ID_MUNICIPIO', 'FECHA_NACIMIENTO', 'ID_PAIS_NACIMIENTO',
        'ID_MUNICIPIO_NACIMIENTO', 'ID_ZONA_RESIDENCIA', 'ID_ESTRATO',
        'ES_REINTEGRO_ESTD_ANTES_DE1998', 'AÑO_PRIMER_CURSO', 'SEMESTRE_PRIMER_CURSO',
        'NETAORD', 'TELEFONO_CONTACTO', 'EMAIL_PERSONAL'
    ]
    columnas_caracterizacion = [
        'ID_TIPO_DOCUMENTO', 'NUM_DOCUMENTO','PRO_CONSECUTIVO', 'ID_MUNICIPIO',
        'CRED_ACAD_PROGRAMA_RC','CREDIT_ACADEM_ACUM_SEM_ANTE','CREDIT_ACAD_A_MATRIC_REGU_SEM',
        'BRUTAORD','APOYO_GOB_NAC_DESCUENTO_VOTAC','APOYO_GOBERNAC_PROGR_PERMANENT',
        'APOYO_ALCALDIA_PROGR_PERMANENT','DESCUENT_RECURRENTES_DE_LA_IES','OTROS_APOYOS_A_LA_MATRICULA',
        'NETAORD', 'APOYO_ADICIONAL_GOBERNACIONES', 'APOYO_ADICIONAL_ALCALDIAS','DESCUENTOS_ADICIONALES_IES',
        'OTROS_APOYOS_ADICIONALES','VAL_NETO_DER_MAT_A_CARGO_EST','VALOR_BRUTO_DERECHOS_COMPLEMEN',
        'VALOR_NETO_DERECHOS_COMPLEMENT','CAUSA_NO_ACCESO']
    for columna in columnas_reporte:
        if columna not in df.columns:
            print(f"La columna '{columna}' no existe en el DataFrame.")
            return None, None
    df_reporte = df[columnas_reporte].copy()
    for columna in columnas_caracterizacion:
        if columna not in df.columns:
            print(f"La columna '{columna}' no existe en el DataFrame.")
            return None, None
    df_reporte1 = df[columnas_caracterizacion].copy()

    if 'FECHA_NACIMIENTO' in df_reporte1.columns:
        df_reporte1['FECHA_NACIMIENTO'] = pd.to_datetime(df_reporte1['FECHA_NACIMIENTO']).dt.strftime('%Y/%m/%d')
    print("Reporte 'PlantillaMatriculados' generado exitosamente.")
    print("Reporte 'PlantillaCaracterizacion' generado exitosamente.")
    return df_reporte, df_reporte1

# Generador de grafico de tortas
def generar_grafico_torta(df, columna_labels, columna_valores, titulo, path_img):
    plt.figure(figsize=(4,4))
    plt.pie(
        df[columna_valores],  # Valores numéricos
        labels=df[columna_labels],  # Etiquetas
        autopct='%1.1f%%',  # Mostrar porcentaje
        startangle=40  # Ángulo de inicio
    )
    plt.title(titulo)
    if path_img:
        plt.savefig(path_img, bbox_inches='tight')
    plt.close()



# Verifica la existencia del archivo en la ruta especifica
file_path = '/content/FinalMatriculaCero_2024-1.xlsx'
file_path1 = '/content/Reporte_general__Caracterizacion__novedades_y_requisitos_politica_de_gratuidad__para_las_IES__2024-1_segundavalidacion.xlsx'
if not os.path.isfile(file_path):
    raise FileNotFoundError(f"{file_path} no encontrado.")
else:
    print(f"Archivo {file_path} encontrado.")
    print(f"Tamaño del archivo: {os.path.getsize(file_path)} bytes")
    print(f"Archivo {file_path1} encontrado.")
    print(f"Tamaño del archivo: {os.path.getsize(file_path1)} bytes")
# Abre el archivo en modo binario para verificar problemas de acceso
try:
    with open(file_path, 'rb') as f:
        print(f"Archivo {file_path} abierto satisfactoriamente en modo binario.")
    with open(file_path1, 'rb') as f:
        print(f"Archivo {file_path1} abierto satisfactoriamente en modo binario.")
except OSError as e:
    print(f"Error al abrir el archivo {file_path}: {e}")
    print(f"Error al abrir el archivo {file_path1}: {e}")


# Carga los DataFrames de trabajo
try:
    # Lectura de los insumos en un diccionario de dataframes
    dic_insumos = pd.read_excel(file_path, sheet_name=['PIAM20241', 'SQ010824', 'PIAM2024_1'], engine='openpyxl')
    dic_insumos1 = pd.read_excel(file_path1, sheet_name=['130624'], engine='openpyxl')
    # Limpia los nombres de columnas
    for df in dic_insumos.values():
        df.columns = df.columns.str.strip()
    piam20241, facturacion20241, PIAM20241CI = dic_insumos['PIAM20241'], dic_insumos['SQ010824'], dic_insumos['PIAM2024_1']
    ciajunio = dic_insumos1['130624']
except Exception as e:
    print(f"Error al cargar los DataFrames: {e}")



# CONFRONTACION DE INSUMOS
# Validacion estado de benefcio certifiacion
ciajunio['SEMESTRE'] = ciajunio['SEMESTRE'].astype(str)
PIAM20241CI['ID-PRO SNIES'] = PIAM20241CI['ID-PROSNIES'].astype(str)
Piam2024_1ci = pd.merge(
    PIAM20241CI,
    ciajunio,
    left_on='ID-PROSNIES',
    right_on='SEMESTRE',
    how='left',
    indicator = True
)

# Cruza los DataFrames Academico y Facturación a partir de la referencia de la factura
piam20241['RECIBO'] = piam20241['RECIBO'].astype(str)
facturacion20241['Documento'] = facturacion20241['Documento'].astype(str)
facturacion20241['Id  factura'] = facturacion20241['Id  factura'].astype(str)

dfldoc_piam20241 = pd.merge(
    piam20241,
    facturacion20241,
    left_on='RECIBO',
    right_on='Documento',
    how='left',
    indicator = True
)

leftonly_dfldoc_piam20241 = dfldoc_piam20241[dfldoc_piam20241['_merge'] == 'left_only'].copy()

dflid_piam20241 = pd.merge(
    leftonly_dfldoc_piam20241,
    facturacion20241,
    left_on='RECIBO',
    right_on='Id  factura',
    how='left'
)

leftonly_dfldoc_piam20241_sq = dflid_piam20241[dflid_piam20241['_merge'] == 'left_only'].copy()
columns_to_drop = [col for col in leftonly_dfldoc_piam20241_sq.columns if '_x' in col]
leftonly_dfldoc_piam20241_sq1 = leftonly_dfldoc_piam20241_sq.drop(columns=columns_to_drop)
leftonly_dfldoc_piam20241_sq1.rename(columns={'RECIBO': 'RECIBO_y'}, inplace=True)

dfl_piam20241_sq_final = pd.merge(
  dfldoc_piam20241,
  leftonly_dfldoc_piam20241_sq1[[
      'RECIBO_y',
      'Documento_y',
      'Destino_y',
      'Nombre de Destino_y',
      'Tercero_y',
      'Nombre del Tercero_y',
      'Id  factura_y',
      'Tipo de Documento_y',
      'Fecha_y',
      'Valor Factura_y',
      'Valor Ajuste_y',
      'Valor Pagado_y',
      'Valor Anulado_y',
      'Saldo_y',
      'Id Integracion_y',
      'Estado Actual_y',
      'Periodico Academico_y',
      'Tipo de Financiacion_y']],
  left_on='RECIBO',
  right_on='RECIBO_y',
  how='left'
)

columns_pair =[
  ('Documento','Documento_y'),
  ('Destino','Destino_y'),
  ('Nombre de Destino','Nombre de Destino_y'),
  ('Tercero','Tercero_y'),
  ('Nombre del Tercero','Nombre del Tercero_y'),
  ('Id  factura','Id  factura_y'),
  ('Tipo de Documento','Tipo de Documento_y'),
  ('Fecha','Fecha_y'),
  ('Valor Factura','Valor Factura_y'),
  ('Valor Ajuste','Valor Ajuste_y'),
  ('Valor Pagado','Valor Pagado_y'),
  ('Valor Anulado','Valor Anulado_y'),
  ('Saldo','Saldo_y'),
  ('Id Integracion','Id Integracion_y'),
  ('Estado Actual','Estado Actual_y'),
  ('Periodico Academico','Periodico Academico_y'),
  ('Tipo de Financiacion','Tipo de Financiacion_y')
]

for col1, col2 in columns_pair:
  dfl_piam20241_sq_final[col1] =dfl_piam20241_sq_final[col1].fillna(dfl_piam20241_sq_final[col2])
  dfl_piam20241_sq_final.drop(columns=[col2], inplace=True)

dfl_piam20241_sq_final.drop(columns='RECIBO_y', inplace=True)
dfl_piam20241_sq_final.sort_values(by='_merge', ascending=False, inplace=True)

Piam2024_1ci['RECIBO'] = Piam2024_1ci['RECIBO'].astype(str)
dfl_piam20241_sqci_final = pd.merge(
  dfl_piam20241_sq_final,
  Piam2024_1ci[[
      'RECIBO',
      'ESTADO CIVF',
      'RESUL VAL2',
      'FLAG ETAPA PAGOS',
      'FONDO VALIDADO',
      'pagos 1',
      'Reintegro 1',
      'ESTADO FINANCIERO ICETEX',
      'GRADO PREVIO',
      'ESTADO_GIRO'
      ]],
  left_on='RECIBO',
  right_on='RECIBO',
  how='left'
)

dfl_piam20241_ci_final = pd.merge(
  Piam2024_1ci,
  dfl_piam20241_sq_final[[
      'RECIBO',
      'Fecha',
      'Estado Actual',
      'Valor Factura',
      'Valor Pagado',
      'Saldo'
  ]],
  left_on='RECIBO',
  right_on='RECIBO',
  how='left'
)

# Calcula el valor de la matricula en el DataFrame Inner Academico Financiero
matriculaBruta = ['DERECHOS_MATRICULA',
                  'BIBLIOTECA_DEPORTES',
                  'LABORATORIOS',
                  'RECURSOS_COMPUTACIONALES',
                  'SEGURO_ESTUDIANTIL',
                  'VRES_COMPLEMENTARIOS',
                  'RESIDENCIAS',
                  'REPETICIONES']
meritoAcademico = ['CONVENIO_DESCENTRALIZACION',
                   'BECA',
                   'MATRICULA_HONOR',
                   'MEDIA_MATRICULA_HONOR',
                   'TRABAJO_GRADO',
                   'DOS_PROGRAMAS',
                   'DESCUENTO_HERMANO',
                   'ESTIMULO_EMP_DTE_PLANTA',
                   'ESTIMULO_CONYUGE',
                   'EXEN_HIJOS_CONYUGE_CATEDRA',
                   'EXEN_HIJOS_CONYUGE_OCASIONAL',
                   'HIJOS_TRABAJADORES_OFICIALES',
                   'ACTIVIDAES_LUDICAS_DEPOR',
                   'DESCUENTOS',
                   'SERVICIOS_RELIQUIDACION',
                   'DESCUENTO_LEY_1171']

dfl_piam20241_sqci_final['BRUTA'] = dfl_piam20241_sqci_final[matriculaBruta].sum(axis=1)
dfl_piam20241_sqci_final['BRUTAORD'] = dfl_piam20241_sqci_final['BRUTA'] - dfl_piam20241_sqci_final['SEGURO_ESTUDIANTIL']
dfl_piam20241_sqci_final['NETAORD'] =  dfl_piam20241_sqci_final['BRUTAORD'] - dfl_piam20241_sqci_final['VOTO'].abs()
dfl_piam20241_sqci_final['MERITO'] = dfl_piam20241_sqci_final[meritoAcademico].sum(axis=1).abs()
dfl_piam20241_sqci_final['MTRNETA'] =  dfl_piam20241_sqci_final['BRUTA'] - dfl_piam20241_sqci_final['VOTO'].abs() - dfl_piam20241_sqci_final['MERITO']
dfl_piam20241_sqci_final['NETAAPL'] =  dfl_piam20241_sqci_final['MTRNETA'] - dfl_piam20241_sqci_final['SEGURO_ESTUDIANTIL']


# Validación de los campos de matricula neta a nivel academico y financiero
dfl_piam20241_sqci_final['FL_NETA'] = dfl_piam20241_sqci_final['MTRNETA'] == dfl_piam20241_sqci_final['Valor Factura']

# Validación de registros duplicados por referencia de matricula financiera
duplicados_recibo = dfl_piam20241_sqci_final['RECIBO'].duplicated(keep=False)

# Validación duplicados basado en 'ID' y 'Codigo SNIES'
duplicados_id_snies = dfl_piam20241_sqci_final.duplicated(subset=['NUM_DOCUMENTO', 'PRO_CONSECUTIVO'], keep=False)


# FILTROS
filtro_piam20241final_estadofacturacion = (
    dfl_piam20241_sqci_final
    .groupby('Estado Actual')['RECIBO']
    .size()
    .reset_index(name='Poblacion')
)
filtro_piam20241final_estadofacturacion['Estado Actual'] = filtro_piam20241final_estadofacturacion['Estado Actual'].replace(estado_financiero)


dfl_piam20241_sqci_final['ESTADO CIVF'].fillna('Extemporaneos', inplace=True)
filtro_piam20241final_beneficio = (
  dfl_piam20241_sqci_final
  .groupby('ESTADO CIVF')['RECIBO']
  .size().
  reset_index(name='Poblacion')
)
filtro_piam20241final_beneficio = filtro_piam20241final_beneficio.rename(columns={
    'ESTADO CIVF': 'Estado de beneficio validado'
})
filtro_piam20241final_beneficio['Estado de beneficio validado'] = filtro_piam20241final_beneficio['Estado de beneficio validado'].replace(estado_beneficio)


dfl_piam20241_sqci_final['FLAG ETAPA PAGOS'].fillna('No ejecutados', inplace=True)
dfl_piam20241_sqci_final['ESTADO CIVF'].fillna('Extemporaneos', inplace=True)
filtro_piam20241final_beneficio_ejecucion = (
  dfl_piam20241_sqci_final
  .groupby(['ESTADO CIVF', 'FONDO VALIDADO', 'FLAG ETAPA PAGOS', 'Estado Actual'])
  .agg(Poblacion=('RECIBO', 'size'),
       Valor_aprobado = ('NETAAPL', 'sum'),
       Valor_Pagado=('Valor Pagado', 'sum'))
  .reset_index()
)
filtro_piam20241final_beneficio_ejecucion = filtro_piam20241final_beneficio_ejecucion.rename(columns={
    'ESTADO CIVF': 'Estado de beneficio validado'
})
filtro_piam20241final_beneficio_ejecucion['Estado de beneficio validado'] = filtro_piam20241final_beneficio_ejecucion['Estado de beneficio validado'].replace(estado_beneficio)




# Guarda los dataframe cruzados segun el insumo academico y financiero
output_path = "/content/AuditoriaConciliacionPiam20241.xlsx"
img_path = "/content/grafico_torta.png"
img_path1 = "/content/grafico_torta1.png"



with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:

  filtro_piam20241final_estadofacturacion.to_excel(writer, sheet_name='Generalidades', startrow=1, startcol=1, index=False)
  filtro_piam20241final_beneficio.to_excel(writer, sheet_name='Generalidades', startrow=1, startcol=4, index=False)
  filtro_piam20241final_beneficio_ejecucion.to_excel(writer, sheet_name='Generalidades', startrow=1, startcol=7, index=False)

  workbook  = writer.book
  worksheet = writer.sheets['Generalidades']
  formato = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
  worksheet.merge_range('B1', "CONCILIACION PIAM 2024-1 POR ESTADO DE RECAUDO",formato)
  worksheet.merge_range('E1', "CONCILIACION PIAM 2024-1 POR ESTADO DE BENEFICIO",formato)
  worksheet.merge_range('H1', "CONCILIACION PIAM 2024-1 POR ESTADO DE BENEFICIO Y EJECUCION",formato)

  # Insertar la imagen del gráfico en la hoja 'Graficos'
  worksheet_graficos = workbook.add_worksheet('Graficos')

  generar_grafico_torta(
    df=filtro_piam20241final_beneficio,
    columna_labels='Estado de beneficio validado',
    columna_valores='Poblacion',
    titulo='Distribución de la Población por Estado de Beneficio Validado',
    path_img=img_path
  )

  worksheet_graficos.insert_image('B2', img_path)

  generar_grafico_torta(
    df=filtro_piam20241final_estadofacturacion,
    columna_labels='Estado Actual',
    columna_valores='Poblacion',
    titulo='Distribución de la Población por Estado de Recaudo de Matricula',
    path_img=img_path1
  )

  worksheet_graficos.insert_image('B20', img_path1)

  df_final_20241 = condensar_estados_ejecucion(dfl_piam20241_sqci_final)
  df_final_20241.to_excel(writer, sheet_name='Piam20241Final', index=False)
  dfl_piam20241_ci_final.to_excel(writer, sheet_name='Piam20241CI', index=False)
  ajustar_ancho_columnas(writer=writer, dataframe=df_final_20241, sheet_name='Piam20241Final')
  ajustar_ancho_columnas(writer=writer, dataframe=dfl_piam20241_ci_final, sheet_name='Piam20241CI')


  if not dfl_piam20241_sqci_final.empty:
    df_no_validos = validar_tipo_documento(dfl_piam20241_sqci_final)
    if not df_no_validos.empty:
      #df_no_validos.to_excel(writer, sheet_name='RegistrosNoValidos', index=False)
      print('Se han guardado los registros no válidos')
    else:
      print('No se encontraron registros inconsistentes por tipo de identificacion')
  else:
    print('El DataFrame está vacío')


  if not dfl_piam20241_sqci_final.empty:

    # Valida las columnas obligatorias
    registrosVaciosTotal, resumenVacios = obtenerRegistrosVacios(dfl_piam20241_sqci_final, columnasValidacionObligatoriedad)
    if not registrosVaciosTotal.empty:
        #registrosVaciosTotal.to_excel(writer, sheet_name='RegistrosVacios', index=False)
        print('Se han guardado los registros vacíos')
    else:
        print('No se encontraron registros vacíos')

    # Valida la cantidad la congruencia en la cantidad de creditos del registro calificado de cada programa
    registrosConRCErrados = verificarInconsistenciasCreditos(dfl_piam20241_sqci_final, 'CRED_ACAD_PROGRAMA_RC')
    if not registrosConRCErrados.empty:
      #registrosConRCErrados.to_excel(writer, sheet_name='InconsistenciasRC', index=False)
      print('Se han guardado los registros con inconsistencias en los creditos del registro calificado')
    else:
      print('No se encontraron registros con inconsistencias en los creditos del registro calificado')

    # Valida la cantidad de creditos del RC de cada programa con respecto a los aporbados por cada estudainte
    registrosInconsistenciasRCAprobados = verificarInconsistenciasCreditosCantidad(dfl_piam20241_sqci_final, 'CRED_ACAD_PROGRAMA_RC', 'CREDIT_ACADEM_ACUM_SEM_ANTE')
    if not registrosInconsistenciasRCAprobados.empty:
      #registrosInconsistenciasRCAprobados.to_excel(writer, sheet_name='InconsistenciasRCAprobados', index=False)
      print('Se han guardado los registros con inconsistencias en los creditos del RC y los creditos aprobados')
    else:
      print('No se encontraron registros con inconsistencias en los creditos del RC y los creditos aprobados')


    # Ajusta el df final
    df_piamfinalajustado = ajustarCreditosAprobados(dfl_piam20241_sqci_final, 'CRED_ACAD_PROGRAMA_RC', 'CREDIT_ACADEM_ACUM_SEM_ANTE', 'SEMESTRES_RC')
    df_piamfinalajustado = ajustarRegistrosVacios(dfl_piam20241_sqci_final, columnasValidacionObligatoriedad)
    df_piamfinalajustado.to_excel(writer, sheet_name='Piam20241FinalAjustado', index=False)
    print('Se han guardado los registros con creditos ajustados')

    # Verificacion final de inconsistencias
    registrosInconsistenciasRCAprobadosFinal = verificarInconsistenciasCreditosCantidad(df_piamfinalajustado, 'CRED_ACAD_PROGRAMA_RC', 'CREDIT_ACADEM_ACUM_SEM_ANTE')
    if not registrosInconsistenciasRCAprobadosFinal.empty:
      print('Se han guardado los registros con inconsistencias en los creditos del RC y los creditos aprobados **')
    else:
      print('No se encontraron registros con inconsistencias en los creditos del RC y los creditos aprobados **')

    registrosVaciosTotal, resumenVacios = obtenerRegistrosVacios(df_piamfinalajustado, columnasValidacionObligatoriedad)
    if not registrosVaciosTotal.empty:
        print('Se han guardado los registros vacíos **')
    else:
        print('No se encontraron registros vacíos **')

    plantilla_matriculados, plantilla_caracterizacion = generarReportePlantillaMatriculados(df_piamfinalajustado)
    if plantilla_matriculados is not None:
      plantilla_matriculados.to_excel(writer, sheet_name='PlantillaMatriculados', index=False)
      print('Se ha generado la plantilla de matriculados')
    if plantilla_caracterizacion is not None:
      plantilla_caracterizacion.to_excel(writer, sheet_name='PlantillaCaracterizacion', index=False)
      print('Se ha generado la plantilla de caracterizacion')

print(f"Archivo guardado en {output_path}")

os.remove(img_path)
os.remove(img_path1)
