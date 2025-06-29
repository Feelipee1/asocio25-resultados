import os
import pandas as pd

def cargar_resultados_multiples_xlsx(directorio="data"):
    registros = []

    for archivo in os.listdir(directorio):
        if archivo.endswith(".xlsx"):
            ruta = os.path.join(directorio, archivo)
            try:
                # Cargar todas las hojas
                xls = pd.read_excel(ruta, sheet_name=None)

                # Procesar hojas esperadas
                asignaciones = xls.get("Asignaciones")
                slack_df = xls.get("Slack")

                if asignaciones is None or slack_df is None:
                    print(f"⚠️ Archivo {archivo} no contiene las hojas requeridas.")
                    continue

                # Calcular KPIs
                dias_asignados = len(asignaciones)
                empleados_asignados = asignaciones["Empleado"].nunique()
                slack_total = slack_df["Slack usado"].sum()

                registros.append({
                    "Instancia": os.path.splitext(archivo)[0],
                    "Slack total": slack_total,
                    "Días asignados": dias_asignados,
                    "Empleados asignados": empleados_asignados
                })

            except Exception as e:
                print(f"❌ Error procesando {archivo}: {e}")

    return pd.DataFrame(registros)

