import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment
from openpyxl.drawing.image import Image
from openpyxl.utils.dataframe import dataframe_to_rows
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet
import matplotlib.pyplot as plt
from reportlab.platypus import Spacer

# Leer el archivo de ventas en formato Excel
archivo_ventas = 'C:/Users/Usuario/OneDrive/Escritorio/AutoDataProcesor/ventas.xlsx'
df_ventas = pd.read_excel(archivo_ventas)

# Definir el período para el informe
periodo = '05/2023'

# Filtrar los datos de ventas por los productos A, B y C en el período especificado
productos = ['Producto A', 'Producto B', 'Producto C']
df_filtrado = df_ventas[df_ventas['Producto'].isin(productos)]

# Calcular la cantidad total vendida y el ingreso total para los productos filtrados
cantidad_total = df_filtrado.groupby('Producto')['Cantidad'].sum()
ingreso_total = df_filtrado['Cantidad'] * df_filtrado['Precio Unitario']
ingreso_total = ingreso_total.groupby(df_filtrado['Producto']).sum()

# Calcular otras estadísticas
margen_beneficio = ingreso_total / (df_filtrado['Cantidad'].sum() * df_filtrado['Precio Unitario'].mean())
rotacion_inventario = cantidad_total / df_filtrado['Cantidad'].sum()
rentabilidad_activo = ingreso_total / df_filtrado['Precio Unitario'].sum()
rentabilidad_patrimonio = ingreso_total / df_filtrado['Costo Unitario'].sum()
margen_bruto = ingreso_total / df_filtrado['Precio Unitario'].sum()
margen_neto = ingreso_total / df_filtrado['Costo Unitario'].sum()

# Generar informe en formato PDF
pdf = SimpleDocTemplate('informe_ventas.pdf', pagesize=letter)
styles = getSampleStyleSheet()

# Crear gráfico de comparación de productos
df_grafico = df_filtrado.groupby('Producto')['Cantidad'].sum().reset_index()

fig, ax = plt.subplots()
ax.bar(df_grafico['Producto'], df_grafico['Cantidad'])
ax.set_xlabel('Productos')
ax.set_ylabel('Cantidad Vendida')
ax.set_title('Comparación de Productos')

# Guardar el gráfico como imagen
plt.savefig('grafico_comparacion.png')

# Crear elementos del informe PDF
elements = []

# Agregar imagen al informe
image = RLImage('grafico_comparacion.png')
image.drawHeight = 400  # Ajusta la altura de la imagen según sea necesario
elements.append(image)

# Agregar texto al informe
informe_texto = []
for producto in productos:
    informe_texto.append(Paragraph(f"<b>Informe de ventas para el producto {producto} ({periodo}):</b>", styles['Heading1']))
    informe_texto.append(Spacer(1, 12))  # Agregar un espacio vertical entre párrafos
    informe_texto.append(Paragraph(f"<b>Cantidad total vendida:</b> {cantidad_total[producto]}", styles['Normal']))
    informe_texto.append(Paragraph(f"<b>Ingreso total:</b> {ingreso_total[producto]}", styles['Normal']))
    informe_texto.append(Paragraph(f"<b>Margen de beneficio:</b> {margen_beneficio[producto]}", styles['Normal']))
    informe_texto.append(Paragraph(f"<b>Rotación de inventario:</b> {rotacion_inventario[producto]}", styles['Normal']))
    informe_texto.append(Paragraph(f"<b>Rentabilidad del activo:</b> {rentabilidad_activo[producto]}", styles['Normal']))
    informe_texto.append(Paragraph(f"<b>Rentabilidad del patrimonio:</b> {rentabilidad_patrimonio[producto]}", styles['Normal']))
    informe_texto.append(Paragraph(f"<b>Margen bruto:</b> {margen_bruto[producto]}", styles['Normal']))
    informe_texto.append(Paragraph(f"<b>Margen neto:</b> {margen_neto[producto]}", styles['Normal']))
    informe_texto.append(Spacer(1, 12))  # Agregar un espacio vertical entre párrafos


# Agregar el texto al informe PDF
elements.extend(informe_texto)

# Generar el informe PDF
pdf.build(elements)

# Generar informe en formato Excel
wb = Workbook()
ws = wb.active
ws.title = 'Informe de Ventas'

# Añadir encabezados de columna
encabezados = list(df_ventas.columns)
ws.append(encabezados)

# Añadir filas de datos
for row in dataframe_to_rows(df_filtrado, index=False, header=False):
    ws.append(row)

# Ajustar alineación de celdas y ancho de columnas
for column in ws.columns:
    max_length = 0
    column_letter = column[0].column_letter
    for i, cell in enumerate(column, start=1):
        if cell.value is not None:
            cell.alignment = Alignment(wrap_text=True)
            if isinstance(cell.value, pd.Timestamp):
                cell.value = cell.value.strftime('%Y-%m-%d')  # Convertir a formato de fecha legible
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
    adjusted_width = (max_length + 2) * 1.2
    ws.column_dimensions[column_letter].width = adjusted_width

# Insertar imagen en el informe Excel
img = Image('grafico_comparacion.png')
img.anchor = 'E'
ws.add_image(img, 'E5')

# Ajustar el ancho de la columna para los datos debajo del gráfico
ws.column_dimensions['E'].width = 30

# Guardar el archivo Excel
wb.save('informe_ventas.xlsx')
