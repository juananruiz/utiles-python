import pandas as pd
from jinja2 import Environment, FileSystemLoader, TemplateNotFound
from weasyprint import HTML
import os
from datetime import datetime
import locale

# Establecer el locale a español
locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')  # Para sistemas Unix
# locale.setlocale(locale.LC_TIME, 'Spanish_Spain.1252')  # Para Windows

# Cargar el archivo Excel
df = pd.read_excel('datos-empleados.xlsx')

# Configurar Jinja2
env = Environment(loader=FileSystemLoader('templates'))

# Función para formatear fechas
def format_fecha(fecha):
    return fecha.strftime("%-d de %B de %Y")

# Crear carpeta "documentos" si no existe
os.makedirs('documentos', exist_ok=True)

# Crear contadores a cero
contador_empleados = 0
contador_documentos = 0

# Iterar sobre cada fila del DataFrame
for index, row in df.iterrows():
    nombre_empleado = row['nombre-empleado']
    modelos = [row['modelo-1'], row['modelo-2'], row['modelo-3'], row['modelo-4'], row['modelo-5']]
    
    # Filtrar modelos vacíos
    modelos = [modelo for modelo in modelos if not pd.isna(modelo)]  # Keep only non-empty models
    # Avisar si no hay modelos asignados
    if not modelos:
        print(f"Advertencia: El empleado '{nombre_empleado}' no tiene ningún modelo asignado.")
        continue  # Se asegura de que se continúe con el siguiente empleado
    
    contador_empleados += 1
    # Crear carpeta para el nombre dentro de "documentos"
    empleado_folder = os.path.join('documentos', nombre_empleado)
    os.makedirs(empleado_folder, exist_ok=True)
    
    # Renderizar y guardar cada modelo
    for modelo in modelos:
        if pd.isna(modelo):  # Verificar si el modelo específico está vacío
            continue  # No intentar cargar la plantilla si el modelo está vacío
        
        try:
            template = env.get_template(f'{modelo}.html')
            html_content = template.render(
                nombre_empleado=nombre_empleado,
                nif_empleado=row['nif-empleado'],
                puesto_empleado=row['puesto-empleado'],
                direccion_empleado=row['direccion-empleado'],
                fecha_inicio=format_fecha(row['fecha-inicio']),
                fecha_fin=format_fecha(row['fecha-fin']),
                duracion_contrato=row['duracion-contrato'],
                horas_semanales=row['horas-semanales'],
                salario_empleado=row['salario-empleado'],
            )
            
            # Convertir a PDF
            pdf_path = os.path.join(empleado_folder, f'{modelo}.pdf')
            HTML(string=html_content).write_pdf(pdf_path)
            contador_documentos += 1
        except TemplateNotFound:
            print(f"Advertencia: La plantilla '{modelo}.html' asignada al empleado '{nombre_empleado}' no se encontró en la carpeta 'templates'.")

print(f"{contador_documentos} documentos generados para {contador_empleados} empleados.")
