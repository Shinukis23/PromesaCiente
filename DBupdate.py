###Programa Funcional OK Mayo 23/2023
import os
import pandas as pd
from datetime import datetime
from oauth2client.service_account import ServiceAccountCredentials
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
import gspread
import sys
import warnings
from google.oauth2 import service_account
from googleapiclient.errors import HttpError
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
import gspread_dataframe as gd
from googleapiclient.http import MediaFileUpload
warnings.filterwarnings("ignore", category=DeprecationWarning)
warnings.simplefilter(action='ignore', category=FutureWarning)

# Get the current working directory
scopes = [
'https://www.googleapis.com/auth/spreadsheets',
'https://www.googleapis.com/auth/drive'
]
credentials = ServiceAccountCredentials.from_json_keyfile_name("monitor-eficiencia-3a13458926a2.json", scopes) #access the json key you downloaded earlier 
file = gspread.authorize(credentials)# authenticate the JSON key with gspread
directory = os.getcwd()

# Initialize an empty list to hold the dataframes from all .xlsx files
dfs = []

###################
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive

def reemplazar_archivo_en_drive(nombre_archivo, ruta_archivo):
    drive = build('drive', 'v3', credentials=credentials)
    
    # Buscar el archivo por su nombre
    #archivo_list = drive.files().list(q=f"name='{nombre_archivo}' and trashed=false").execute().get('files', [])
    #archivo_list = drive.files().list(q=f"name='{nombre_archivo}' and parents='117ZVbhyZ9xHtSdo7wkQ2Otww1boRC-Vc' and trashed=false").execute().get('files', [])
    carpeta_id = '117ZVbhyZ9xHtSdo7wkQ2Otww1boRC-Vc'
    archivo_list = drive.files().list(q=f"'{carpeta_id}' in parents and name='{nombre_archivo}' and trashed=false").execute().get('files', [])


    if archivo_list:
        archivo_drive = archivo_list[0]
        archivo_id = archivo_drive['id']
        
        # Actualizar el archivo con el nuevo contenido
        media_body = MediaFileUpload(ruta_archivo, resumable=True)
        archivo_actualizado = drive.files().update(fileId=archivo_id, media_body=media_body).execute()
        
        print(f"Archivo {nombre_archivo} actualizado en Google Drive")
    else:
        print(f"No se encontró el archivo {nombre_archivo} en Google Drive")

   

    
#############


# Iterate over all .xlsx files in the directory
for file_name in os.listdir(directory):
    if (file_name.startswith('JobsReport_')&file_name.endswith('_Logistica.xlsx')) or (file_name.startswith('ReporteProduccionDB')&file_name.endswith('resultado.xlsx')):
        file_path = os.path.join(directory, file_name)
        # Load the data from the current file into a dataframe
        print(file_path)
        data = pd.read_excel(file_path)
        # Append the dataframe to the dfs list
        indexDeleted = data[(data['Job Type'].str.upper().str.contains('CHECK'))|(data['Drop Location'].str.upper().str.contains('FOTOS'))].index
        data.drop(indexDeleted,inplace=True)
        dfs.append(data)


# Concatenate all dataframes in the dfs list into a single dataframe
concatenated_data = pd.concat(dfs, ignore_index=True)
concatenated_data.sort_values("Created", inplace=True)
concatenated_data.to_excel(r'ReporteProduccionDBsort2.xlsx', index=False)


concatenated_data.drop_duplicates(subset='Job #',keep='first', inplace=True)
concatenated_data.sort_values("Created",ascending=False, inplace=True)
#concatenated_data.to_excel(r'ReporteProduccionDBresultado.xlsx', index=False)
#ascending
#dl = pds.DataFrame()
#Job #   Job Status  Reason  Due
#Drop Location   R # Stock # Interchange Part Description Summary    Part Price  Created Pull Started    Pull Started By Pulled Finished Pulled Finish By    Ship Via    Inspector   Order Store #   Part Store #
#rutas= pd.read_excel(r'Rutas pendientes.xls')
#merged_data = pd.merge(rutas,concatenated_data[['Job #','Drop Location','R #','Stock #','Interchange','Part Description Summary',
#    'Part Price','Created','Ship Via','Order Store #','Part Store #','Due']],on=['Job #'],how="left")
# Write the concatenated data to a new .xlsx file
#merged_data["Delivery time"]= datetime.now() 
#indexDeleted = ds2[ds2['Job Status'] ==  'Pickup'].index
#ds2.drop(indexDeleted,inplace=True)
#indexDeleted = merged_data[merged_data['Drop Location'] == ' '].index  # dejando solo las 253 en copia de audit trial
#merged_data.drop(indexDeleted,inplace=True)
#merged_data=merged_data.dropna(subset=['Drop Location'])
output_file_path = os.path.join(directory, 'ReporteProduccionDBresultado.xlsx')
concatenated_data.to_excel(output_file_path, index=False)
#nombre_archivo = "ReporteProduccionDBresultado.xlsx"  # Nombre del archivo en Google Drive
#ruta_archivo = "ReporteProduccionDBresultado.xlsx"  # Ruta local del nuevo archivo Excel
#reemplazar_archivo_en_drive(nombre_archivo, ruta_archivo)
nombre_archivo = "ReporteProduccionDBresultado.xlsx"  # Nombre del archivo en Google Drive
ruta_archivo = "ReporteProduccionDBresultado.xlsx"  # Ruta local del nuevo archivo Excel
reemplazar_archivo_en_drive(nombre_archivo, ruta_archivo)