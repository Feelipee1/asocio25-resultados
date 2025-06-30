import pandas as pd
import os
from pathlib import Path

# Definir la ruta de la carpeta
folder_path = r"D:\asocio25-resultados\data"

# Listas para almacenar todos los DataFrames
all_assignments = []
all_meetings = []
all_slack = []

# Encontrar todos los archivos que coincidan con el patrón Instance*_900*.xlsx
file_pattern = "Instance*_900*.xlsx"
file_paths = list(Path(folder_path).glob(file_pattern))

if not file_paths:
    print(f"No se encontraron archivos que coincidan con {file_pattern} en {folder_path}")
else:
    print(f"Encontrados {len(file_paths)} archivos para procesar:")
    for file_path in file_paths:
        print(f"- {file_path.name}")

# Procesar cada archivo encontrado
for file_path in file_paths:
    try:
        # Extraer el número de instancia del nombre del archivo
        instance_num = file_path.stem.split('_')[0].replace('Instance', '')
        
        # Leer cada hoja y agregar una columna con el número de instancia
        assignments = pd.read_excel(file_path, sheet_name='Asignaciones')
        assignments['Instance'] = instance_num
        all_assignments.append(assignments)
        
        meetings = pd.read_excel(file_path, sheet_name='Reuniones')
        meetings['Instance'] = instance_num
        all_meetings.append(meetings)
        
        slack = pd.read_excel(file_path, sheet_name='Slack')
        slack['Instance'] = instance_num
        all_slack.append(slack)
        
    except Exception as e:
        print(f"Error procesando {file_path.name}: {str(e)}")
        continue

# Combinar todos los DataFrames si se encontraron archivos
if all_assignments:
    combined_assignments = pd.concat(all_assignments, ignore_index=True)
    combined_meetings = pd.concat(all_meetings, ignore_index=True)
    combined_slack = pd.concat(all_slack, ignore_index=True)
    
    # Guardar los datos combinados en un nuevo archivo Excel
    output_path = os.path.join(folder_path, "combined_results.xlsx")
    with pd.ExcelWriter(output_path) as writer:
        combined_assignments.to_excel(writer, sheet_name='Asignaciones_Combinadas', index=False)
        combined_meetings.to_excel(writer, sheet_name='Reuniones_Combinadas', index=False)
        combined_slack.to_excel(writer, sheet_name='Slack_Combinado', index=False)
    
    print(f"\nProcesamiento completado. Resultados guardados en: {output_path}")
    
    # Mostrar resumen estadístico
    print("\nResumen Estadístico:")
    print("\nAsignaciones por instancia:")
    print(combined_assignments['Instance'].value_counts().sort_index())
    
    print("\nReuniones por instancia:")
    print(combined_meetings['Instance'].value_counts().sort_index())
    
    print("\nUso de Slack por instancia:")
    print(combined_slack.groupby(['Instance', 'Slack usado']).size().unstack().fillna(0))
    
else:
    print("No se encontraron datos válidos para procesar.")


    # Procesar la hoja 'Reuniones' (página 2) de cada archivo y combinar los resultados
    all_meeting_days = []

    for file_path in file_paths:
        try:
            instance_num = file_path.stem.split('_')[0].replace('Instance', '')
            meeting_days = pd.read_excel(file_path, sheet_name='Reuniones', usecols=['Grupo', 'Dia de Reunion'])
            meeting_days['Instance'] = instance_num
            all_meeting_days.append(meeting_days)
        except Exception as e:
            print(f"Error procesando hoja 'Reuniones' en {file_path.name}: {str(e)}")
            continue

    if all_meeting_days:
        combined_meeting_days = pd.concat(all_meeting_days, ignore_index=True)
        output_meeting_days_path = os.path.join(folder_path, "combined_meeting_days.xlsx")
        combined_meeting_days.to_excel(output_meeting_days_path, index=False)
        print(f"\nCombinación de 'Dia de Reunion' completada. Resultados guardados en: {output_meeting_days_path}")
        print("\nResumen de reuniones por grupo e instancia:")
        print(combined_meeting_days.groupby(['Instance', 'Grupo'])['Dia de Reunion'].unique())
    else:
        print("No se encontraron datos válidos en la hoja 'Reuniones' para procesar.")