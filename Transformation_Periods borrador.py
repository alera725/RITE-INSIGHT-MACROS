# -*- coding: utf-8 -*-
"""
Created on Thu Sep  2 11:34:01 2021

@author: alejandro.gutierrez
"""

#import os
import shutil
#os.chdir('C:\\Users\\alejandro.gutierrez\\OneDrive - Carlin Group - CA Fortune\\Documents\\ALEJANDRO RAMOS GTZ\\GIT\\RITE INSIGHT BOT') # relative path: scripts dir is under Lab

from openpyxl import Workbook
from openpyxl import load_workbook
import pandas as pd 
import numpy as np
from datetime import date, timedelta
import datetime

#Set paths 
import_path_periods = 'C:\\Users\\alejandro.gutierrez\\OneDrive - Carlin Group - CA Fortune\\Documents\\KROGER SELENIUM\\RITE INSIGHT\\PERIODS' 
import_path_ytd = 'C:\\Users\\alejandro.gutierrez\\OneDrive - Carlin Group - CA Fortune\\Documents\\KROGER SELENIUM\\RITE INSIGHT\\YTD' 
import_path_IML = 'C:\\Users\\alejandro.gutierrez\\OneDrive - Carlin Group - CA Fortune\\Documents\\KROGER SELENIUM\\RITE INSIGHT\\ITEM MASTER LISTING COMPETITIVE'
import_path_inventory = 'C:\\Users\\alejandro.gutierrez\\OneDrive - Carlin Group - CA Fortune\\Documents\\KROGER SELENIUM\\RITE INSIGHT\\INVENTORY'

export_path = 'C:\\Users\\alejandro.gutierrez\\OneDrive - Carlin Group - CA Fortune\\Documents\\ALEJANDRO RAMOS GTZ\\CARLIN\\RITE AID'

#Week number 
t1 = datetime.datetime.now()
week_number = str(t1.strftime("%U"))

# Creamos los nombres de los archivos que vamos a importar 
import_files = ['SBIC YTD 27DEC2020 TO LAST SATURDAY week %s.xlsx'%week_number, 
                'SBIC 1 WEEK week %s.xlsx'%week_number, 
                'SBIC 52 WEEKS week %s.xlsx'%week_number,
                'SBIC 26 WEEKS week %s.xlsx'%week_number,
                'SBIC 13 WEEKS week %s.xlsx'%week_number,
                'SBIC 4 WEEKS week %s.xlsx'%week_number,
                'ITEM MASTER LISTING 52 WEEKS week %s.xlsx'%week_number]

# !!! OJO: Si se agrega un nuevo cliente agregarlo aqui VER SI ES POSIBLE HACERLO HACER PRUEBA CON UNO EN ESCRITORIO
#import_inventory_files = ['KIND week %s.xlsx'%week_number, 
#                'KRAVE week %s.xlsx'%week_number, 
#                'EOS week %s.xlsx'%week_number,
#                'STERNO week %s.xlsx'%week_number,
#                'GOLDEN_EYE week %s.xlsx'%week_number,
#                'VIRMAX week %s.xlsx'%week_number,
#                'EVOLVE week %s.xlsx'%week_number, 
#                'SOYLENT week %s.xlsx'%week_number]

# Read files  
df = pd.read_excel(import_path_ytd + '\\' + import_files[0], sheet_name='Sales')

df_ytd = pd.read_excel(import_path_ytd + '\\' + import_files[0], sheet_name='Sales')
df_iml = pd.read_excel(import_path_IML + '\\' + import_files[6], sheet_name='Sales')
df = pd.read_excel(import_path_periods + '\\' + import_files[1], sheet_name='Sales')
df_52 = pd.read_excel(import_path_periods + '\\' + import_files[2], sheet_name='Sales')
df_26 = pd.read_excel(import_path_periods + '\\' + import_files[3], sheet_name='Sales')
df_13 = pd.read_excel(import_path_periods + '\\' + import_files[4], sheet_name='Sales')
df_4 = pd.read_excel(import_path_periods + '\\' + import_files[5], sheet_name='Sales')


#%% METEMOS EN EL MODULO 
# Seleccionar el rango de las fechas 
dates_between = df.iloc[2,10]

# Delete the first five rows using iloc selector
df = df.iloc[7:,]

# Delete 8 and 9 row
df = df.drop([8,9], axis=0) 

# Delete 1st column
df.drop(df.columns[[0]], axis = 1, inplace = True)

# Reacomodar columnas con saltos de linea
df = df.replace('\n',' ', regex=True)

# Set header 1st Column and take all rows except the 1sr and reset index
df = df.rename(columns=df.iloc[0])
df = df.iloc[1:,]
df = df.reset_index(drop=True)

# Agregamos columna con la fecha correspondiente al reporte y la colocamos como primera columna
df.insert(0, 'Dates', dates_between)

#not_interested_cols = [1,2,4,5,6,7,10,12,13,14,17,18,21,22,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54]
interested_cols = [0,3,8,9,11,15,16,19,20,23,24,55,56,57,58]

# Creamos un nuevo DataFrame con las columnas deseadas
df_1 = df.iloc[:,interested_cols]

# Rename multiple columns
df_1.columns.values[0] = "Dates"
df_1.columns.values[1] = "UPC"
df_1.columns.values[2] = "Status"
df_1.columns.values[3] = "Current_Yr_$"
df_1.columns.values[4] = "Prior_Yr_$"
df_1.columns.values[5] = "Current_Yr_U"
df_1.columns.values[6] = "Prior_Yr_U"
df_1.columns.values[7] = "Current_Yr_Promo$"
df_1.columns.values[8] = "Prior_Yr_Promo$"
df_1.columns.values[9] = "Current_Yr_PromoU"
df_1.columns.values[10] = "Prior_Yr_PromoU"
df_1.columns.values[11] = "Current_Yr_UniqSellStore"
df_1.columns.values[12] = "Prior_Yr_UniqSellStore"
df_1.columns.values[13] = "Current_Yr_POGStore"
df_1.columns.values[14] = "Prior_Yr_POGStore"

df_1.dtypes

#df_1['UPC'] = df_1['UPC'].astype('int64')
df_1['Current_Yr_$'] = df_1['Current_Yr_$'].astype('float64')
df_1['Prior_Yr_$'] = df_1['Prior_Yr_$'].astype('float64')
df_1['Current_Yr_U'] = df_1['Current_Yr_U'].astype('int64')
df_1['Prior_Yr_U'] = df_1['Prior_Yr_U'].astype('int64')
df_1['Current_Yr_Promo$'] = df_1['Current_Yr_Promo$'].astype('float64')
df_1['Prior_Yr_Promo$'] = df_1['Prior_Yr_Promo$'].astype('float64')
df_1['Current_Yr_PromoU'] = df_1['Current_Yr_PromoU'].astype('int64')
df_1['Prior_Yr_PromoU'] = df_1['Prior_Yr_PromoU'].astype('int64')
df_1['Current_Yr_UniqSellStore'] = df_1['Current_Yr_UniqSellStore'].astype('int64')
df_1['Prior_Yr_UniqSellStore'] = df_1['Prior_Yr_UniqSellStore'].astype('int64')
df_1['Current_Yr_POGStore'] = df_1['Current_Yr_POGStore'].astype('float64')
df_1['Prior_Yr_POGStore'] = df_1['Prior_Yr_POGStore'].astype('float64')

# -- Corroboramos data types con el archivo final de Magdiel 
df_2 = pd.read_excel(r'C:\Users\alejandro.gutierrez\OneDrive - Carlin Group - CA Fortune\Desktop\RA YTD segundo.xlsx', sheet_name='Sales')
#df2.dtypes


#%% Creamos otro modulo con los export names ??
names = ['RA YTD.xlsx', 'RA L04W.xlsx', 'RA L13W.xlsx', 'RA L26W.xlsx', 'RA L52W.xlsx', 'Weekly.xlsx', '%s'%dates_between]

df_1.iloc[0,0]

# Exportamos el archivo en excel file (Aqui entra la variable del nombre con if's)
df_1.to_excel(r'C:\Users\alejandro.gutierrez\OneDrive - Carlin Group - CA Fortune\Desktop\NUEVO YTD segundo.xlsx', sheet_name='Sales', index = False)



# -- Volvemos a leer el archivo exportado para corroborar
#df3 = pd.read_csv(r'C:\Users\alejandro.gutierrez\OneDrive - Carlin Group - CA Fortune\Desktop\NUEVO YTD segundo.csv')
#df3.dtypes


