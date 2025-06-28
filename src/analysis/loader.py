import os
import pandas as pd

def cargar_resultados_multiples_xlsx(directorio="data"):
    registros = []
    
    for archivo in os.listdir(directorio):
        if archivo.endswith(".xlsx"):
            ruta = os.path.join(directorio, archivo)
            nombre = os.path.splitext(archivo)[0]
            
            try:
                xls = pd.read_excel(ruta, sheet_name=None)
                resumen = xls.get("Resumen")  # hoja opcional

                if resumen is not None:
                    fila = resumen.iloc[0].to_dict()
                    fila["Instancia"] = nombre
                    registros.append(fila)
                else:
                    # Cargar otras hojas manualmente y calcular KPIs si no hay hoja "Resumen"
                    asignaciones = xls["Asignaciones"]
                    slack = xls["Slack"]
                    empleados = asignaciones["Empleado"].nunique()
                    dias = len(asignaciones)
                    slack_total = slack["Slack usado"].sum()
                    
                    fila = {
                        "Instancia": nombre,
                        "Objetivo": None,  # si no está en hoja resumen
                        "Slack total": slack_total,
                        "Días asignados": dias,
                        "Empleados asignados": empleados
                    }
                    registros.append(fila)
            except Exception as e:
                print(f"⚠️ Error al procesar {archivo}: {e}")
    
    df = pd.DataFrame(registros)
    return df
