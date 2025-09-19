import pandas as pd
import requests
from io import StringIO
from datetime import date, datetime
import matplotlib.pyplot as plt
from xlsxwriter import Workbook
import os

ruta_data = os.path.join(os.path.expanduser("~"), "Downloads") + '\\'
url = "https://www.datosabiertos.gob.pe/sites/default/files/Datos_abiertos_matriculas_2024_2_2025_1_0.csv"
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36"
}
response = requests.get(url, headers=headers)
df_uni = pd.read_csv(StringIO(response.content.decode("utf-8")))

df_uni["Periodo_ciclo"] = df_uni["ANIO"].astype(str) + "-" + df_uni["PERIODO"].astype(str)
df_uni["Descrip_ciclo"] = "Ciclo " + df_uni["CICLO_RELATIVO"].astype(str)

#EDAD ACTUAL DE MATRICULADOS
df_uni["EDAD_ACTUAL"] = date.today().year - df_uni["ANIO_NACIMIENTO"].astype(int)

#EDAD CUANDO SE MATRICULARON
df_uni["EDAD_MATRICULA"] = df_uni["ANIO"].astype(int) - df_uni["ANIO_NACIMIENTO"].astype(int)

def Grupo_edad_actual(row):
    if row["EDAD_ACTUAL"] <= 18:
        return f"{df_uni['EDAD_ACTUAL'].min()} - 18 años"
    elif row["EDAD_ACTUAL"] <= 24:
        return "19 - 24 años"
    elif row["EDAD_ACTUAL"] <= 30:
        return "25 - 30 años"
    elif row["EDAD_ACTUAL"] <= 40:
        return "31 - 40 años"
    elif row["EDAD_ACTUAL"] <= 50:
        return "41 - 50 años"
    else:
        return "+51 años"
    
df_uni["Grupo_edad_actual"] = df_uni.apply(Grupo_edad_actual, axis=1)

def Grupo_edad_matricula(row):
    if row["EDAD_MATRICULA"] <= 18:
        return f"{df_uni['EDAD_MATRICULA'].min()} - 18 años"
    elif row["EDAD_MATRICULA"] <= 24:
        return "19 - 24 años"
    elif row["EDAD_MATRICULA"] <= 30:
        return "25 - 30 años"
    elif row["EDAD_MATRICULA"] <= 40:
        return "31 - 40 años"
    elif row["EDAD_MATRICULA"] <= 50:
        return "41 - 50 años"
    else:
        return "+51 años"
    
df_uni["Grupo_edad_matricula"] = df_uni.apply(Grupo_edad_matricula, axis=1)
print(df_uni)

#print(df_uni["IDHASH"].duplicated().sum())

#Tablas dinamicas 
#Cantidad de alumnos matriculados por ciclo
alumnos_matriculados_ciclo = pd.pivot_table(df_uni, index="Periodo_ciclo", values="IDHASH", aggfunc="count")
#print("En los ciclos 2024-02 - 2025-01 se metricularon la siguiente cantidad de alumnos:")
print(alumnos_matriculados_ciclo)

#Matriculados por carreras 
matriculados_carreras = pd.pivot_table(df_uni, index="ESPECIALIDAD", columns="Periodo_ciclo", values="IDHASH", aggfunc="count")
matriculados_carreras['TOTAL'] = matriculados_carreras.sum(axis=1)  # Total por fila
#matriculados_carreras.loc['TOTAL'] = matriculados_carreras.sum(axis=0)  # Total por columna
matriculados_carreras = matriculados_carreras.sort_values(by="TOTAL", ascending=False)
print(matriculados_carreras)

#Matriculados por sexo

matriculados_sexo = pd.pivot_table(df_uni, index="SEXO", columns="Periodo_ciclo", values="IDHASH", aggfunc="count")
print(matriculados_sexo)

#Cant. matriculados por ciclo

matriculados_ciclo_ord = pd.pivot_table(df_uni, index="Descrip_ciclo", columns="Periodo_ciclo", values="IDHASH", aggfunc="count")
matriculados_ciclo_ord['TOTAL'] = matriculados_ciclo_ord.sum(axis=1) 
print(matriculados_ciclo_ord.sort_values(by='TOTAL', ascending=False))

#departamentos
departamento_matri = pd.pivot_table(df_uni, index="DOMICILIO_DEPA", columns="Periodo_ciclo", values="IDHASH", aggfunc="count")
departamento_matri['TOTAL'] = departamento_matri.sum(axis=1)
print(departamento_matri.sort_values(by='TOTAL', ascending=False))

#grupo de edad de los matriculados 2025-01
matriculados251 = df_uni[(df_uni["Periodo_ciclo"]=="2025-1")]
grupo_actual_matriculados = pd.pivot_table(matriculados251, index="Grupo_edad_actual", columns="Periodo_ciclo", values="IDHASH", aggfunc="count")
print(grupo_actual_matriculados.sort_values(by="2025-1", ascending=False))


with pd.ExcelWriter(ruta_data + 'REPORTE_MATRICULADOS_UNI.xlsx', engine='xlsxwriter') as writer:
    df_uni.to_excel(writer, sheet_name='DATA', index=False)
    alumnos_matriculados_ciclo.to_excel(writer, sheet_name='PERIODO')
    matriculados_carreras.to_excel(writer, sheet_name='CARRERA')
    matriculados_sexo.to_excel(writer, sheet_name='SEXO')
    matriculados_ciclo_ord.to_excel(writer, sheet_name='CICLO')
    departamento_matri.to_excel(writer, sheet_name='DEPARTAMENTO')
    grupo_actual_matriculados.to_excel(writer, sheet_name='Grupo edad')

print("Archivo generado")




    




