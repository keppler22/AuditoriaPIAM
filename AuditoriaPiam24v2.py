import os
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
    'ES_REINTEGRO_ESTD_ANTES_DE1998','AÑO_PRIMER_CURSO','SEMESTRE_PRIMER_CURSO','TELEFONO_CONTACTO','EMAIL_PERSONAL',
    'EMAIL_INSTITUCIONAL','CRED_ACAD_PROGRAMA_RC','SEMESTRES_RC','CREDIT_ACADEM_ACUM_SEM_ANTE','CREDITOSMATRICULADOS',
    'DERECHOS_MATRICULA','SEGURO_ESTUDIANTIL']

# Analisis estructural
def validar_tipo_documento(df):
    if 'ID_TIPO_DOCUMENTO' in df.columns:
        df_validacion_tipoId = df[~df['ID_TIPO_DOCUMENTO'].isin(valoresValidosTipoId)]
        return df_validacion_tipoId
    else:
        return pd.DataFrame()

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


# Función para ajustar el ancho de las columnas en un archivo Excel
def ajustar_ancho_columnas(writer, dataframe, sheet_name, startrow=0, startcol=0):
    worksheet = writer.sheets[sheet_name]  # Obtener la hoja de cálculo
    for i, col in enumerate(dataframe.columns):
        # Encuentra el ancho máximo entre el nombre de la columna y el contenido de la columna
        max_len = max(dataframe[col].astype(str).map(len).max(), len(col)) + 2
        # Ajustar el ancho de la columna en Excel
        worksheet.set_column(startcol + i, startcol + i, max_len)

def obtenerRegistrosVacios(df,columnas):
  registrosVaciosTotal = pd.DataFrame()
  resumenVacios = {}

  for columna in columnas:
    if columna not in df.columns:
      print(f"La columna '{columna}' no existe en el DataFrame.")
      continue

    registrosVacios = df[df[columna].isnull()]
    resumenVacios[columna] = len(registrosVacios)

    if registrosVacios.empty:
      print(f"No hay registros vacíos en la columna '{columna}'.")
    else:
      print(f"Hay {len(registrosVacios)} registros vacíos en la columna '{columna}'.")
      registrosVacios['BanderaRegistrosVacios'] = columna
      registrosVaciosTotal = pd.concat([registrosVaciosTotal, registrosVacios])

  return registrosVaciosTotal, resumenVacios


# Verifica la existencia del archivo en la ruta especifica
file_path = '/content/FinalMatriculaCero_2024-1.xlsx'

if not os.path.isfile(file_path):
    raise FileNotFoundError(f"{file_path} no encontrado.")
else:
    print(f"Archivo {file_path} encontrado.")

# Abre el archivo en modo binario para verificar problemas de acceso
try:
    with open(file_path, 'rb') as f:
        print(f"Archivo {file_path} abierto satisfactoriamente en modo binario.")
except OSError as e:
    print(f"Error al abrir el archivo {file_path}: {e}")


# Carga los DataFrames de trabajo
try:
    # Lectura de los insumos en un diccionario de dataframes
    dic_insumos = pd.read_excel(file_path, sheet_name=['PIAM20241', 'SQ010824', 'PIAM2024_1'], engine='openpyxl')

    # Limpia los nombres de columnas
    for df in dic_insumos.values():
        df.columns = df.columns.str.strip()

    piam20241, facturacion20241, PIAM20241CI = dic_insumos['PIAM20241'], dic_insumos['SQ010824'], dic_insumos['PIAM2024_1']

except Exception as e:
    print(f"Error al cargar los DataFrames: {e}")



# CONFRONTACION DE INSUMOS
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

PIAM20241CI['RECIBO'] = PIAM20241CI['RECIBO'].astype(str)
dfl_piam20241_sqci_final = pd.merge(
  dfl_piam20241_sq_final,
  PIAM20241CI[[
      'RECIBO',
      'ESTADO CIVF',
      'RESUL VAL2',
      'FLAG ETAPA PAGOS',
      'FONDO VALIDADO',
      'pagos 1',
      'Reintegro 1',
      'ESTADO FINANCIERO ICETEX',
      'GRADO PREVIO'
      ]],
  left_on='RECIBO',
  right_on='RECIBO',
  how='left'
)

dfl_piam20241_ci_final = pd.merge(
  PIAM20241CI,
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
  worksheet.merge_range('B1:C1', "CONCILIACION PIAM 2024-1 POR ESTADO DE RECAUDO",formato)
  worksheet.merge_range('E1:F1', "CONCILIACION PIAM 2024-1 POR ESTADO DE BENEFICIO",formato)
  worksheet.merge_range('H1:N1', "CONCILIACION PIAM 2024-1 POR ESTADO DE BENEFICIO Y EJECUCION",formato)

  worksheet.set_column('B:C', max(2, len("CONCILIACION PIAM 2024-1 POR ESTADO DE RECAUDO") + 1))
  worksheet.set_column('E:F', max(2, len("CONCILIACION PIAM 2024-1 POR ESTADO DE BENEFICIO") + 1))
  worksheet.set_column('H:N', max(2, len("CONCILIACION PIAM 2024-1 POR ESTADO DE BENEFICIO Y EJECUCION") + 1))

  # Insertar la imagen del gráfico en la hoja 'Graficos'
  worksheet_graficos = workbook.add_worksheet('Graficos')

  generar_grafico_torta(
    df=filtro_piam20241final_beneficio,
    columna_labels='Estado de beneficio validado',
    columna_valores='Poblacion',
    titulo='Distribución de la Población por Estado de Recaudo de Matricula',
    path_img=img_path
  )

  worksheet_graficos.insert_image('B2', img_path)

  generar_grafico_torta(
    df=filtro_piam20241final_estadofacturacion,
    columna_labels='Estado Actual',
    columna_valores='Poblacion',
    titulo='Distribución de la Población por Estado de Beneficio Validado',
    path_img=img_path1
  )

  worksheet_graficos.insert_image('B20', img_path1)

  dfl_piam20241_sqci_final.to_excel(writer, sheet_name='Piam20241Final', index=False)
  dfl_piam20241_ci_final.to_excel(writer, sheet_name='Piam20241CI', index=False)
  ajustar_ancho_columnas(writer, dfl_piam20241_sqci_final, 'Piam20241Final')
  ajustar_ancho_columnas(writer, dfl_piam20241_ci_final, 'Piam20241CI')


  if not dfl_piam20241_sqci_final.empty:
    df_no_validos = validar_tipo_documento(dfl_piam20241_sqci_final)
    if not df_no_validos.empty:
      df_no_validos.to_excel(writer, sheet_name='RegistrosNoValidos', index=False)
      print('Se han guardado los registros no válidos')
    else:
      print('No se encontraron registros por tipo Id no válidos por tipo de identificacion')
  else:
    print('El DataFrame está vacío')

  # Valida las columnas obligatorias
  registrosVaciosTotal, resumenVacios = obtenerRegistrosVacios(dfl_piam20241_sqci_final, columnasValidacionObligatoriedad)
  
  if not registrosVaciosTotal.empty:
      registrosVaciosTotal.to_excel(writer, sheet_name='RegistrosVacios', index=False)
      print('Se han guardado los registros vacíos')
  else:
      print('No se encontraron registros vacíos')

print(f"Archivo guardado en {output_path}")

os.remove(img_path)
os.remove(img_path1)
