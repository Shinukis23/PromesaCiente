###Programa Funcional OK Mayo 23/2023
import os
import pandas as pd
from datetime import datetime

# Get the current working directory
directory = os.getcwd()

# Initialize an empty list to hold the dataframes from all .xlsx files
dfs = []

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
