# -*- coding: utf-8 -*-
"""
Created on Fri Sep  3 10:08:40 2021

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

#Week number 
t1 = datetime.datetime.now()
week_number = str(t1.strftime("%U"))

#import exprt paths
import_path_IML = 'C:\\Users\\alejandro.gutierrez\\OneDrive - Carlin Group - CA Fortune\\Documents\\KROGER SELENIUM\\RITE INSIGHT\\ITEM MASTER LISTING COMPETITIVE\\'
export_path = 'C:\\Users\\alejandro.gutierrez\\OneDrive - Carlin Group - CA Fortune\\Documents\\KROGER SELENIUM\\RITE INSIGHT\\FINAL WEEK FILES'


# Read files
df = pd.read_excel(r'C:\Users\alejandro.gutierrez\OneDrive - Carlin Group - CA Fortune\Documents\KROGER SELENIUM\RITE INSIGHT\ITEM MASTER LISTING COMPETITIVE\ITEM MASTER LISTING 52 WEEKS week 36.xlsx', skiprows= [0,1,2], sheet_name='ItemMaster')


#%% METEMOS EN EL MODULO 

# Delete the first five rows using iloc selector
#df = df.iloc[4:,]

# Reacomodar columnas con saltos de linea
df = df.replace('\n',' ', regex=True)

# Pasar el primer row como header del data frame
df = df.rename(columns=df.iloc[0])
df = df.iloc[1:,]
df = df.reset_index(drop=True)

# Renombrar columnas con espacios en blanco
df = df.rename(columns={'Category Description': 'Category_Description', 'Class Description': 'Class_Description', 'Sub Class Description': 'Sub_Class_Description', 'Item No': 'Item_No', 'Original Item No': 'Original_Item_No', 'Active Item No': 'Active_Item_No', 'Item Description': 'Item_Description', 'Sub Brand': 'Sub_Brand', 'Vendor Name': 'Vendor_Name'})
#df.dtypes

# Eliminamos duplicados en base a la columna UPC
df_1 = df.drop_duplicates(subset='UPC') #, inplace = True) #Fue antes de convertir las columnas a formato int64 y si funciono 

#df_1.groupby("Vendor#").count()

# Convertimos las columnas al tipo de dato deseado PARECE QUE NO SERA NECESARIO  YA QUE HAY MUCHOS QUE TIENEN 0s A LA IZQUIERDA Y ASI LOS DEJO MAGDIEL
#df_1['Item_No'] = df_1['Item_No'].astype('int64')
#df_1['Original_Item_No'] = df_1['Original_Item_No'].astype('int64')
#df_1['Active_Item_No'] = df_1['Active_Item_No'].astype('float64')
#df['UPC'] = df['UPC'].astype('int64')
#df_1['Vendor#'] = df_1['Vendor#'].astype('int64')


# -- Eliminamos duplicados en base a la columna UPC
#df_3drop = df.drop_duplicates(subset='UPC') #, inplace = True) #Si lo hacemos despues de convertir las columnas a int64 arroja mas duplicados
# -- localizar cuales rows fueron eliminados
#drop_after = df.loc[df.duplicated(subset=['UPC'], keep = False)]







# -- Corroboramos data types con el archivo final de Magdiel 
#df2 = pd.read_excel('C:\\Users\\alejandro.gutierrez\\OneDrive - Carlin Group - CA Fortune\\Desktop\\ITEM MASTER LISTING 52 WEEKS Segundo.xlsx', sheet_name='ItemMaster')
#df2.dtypes
#df2_drop = df2.drop_duplicates(subset='UPC') #, inplace = True)
#n_by_vendor = df2.groupby("Vendor#").count()


# Exportamos el archivo en excel file (Aqui entra la variable del nombre con if's)
df_1.to_excel(r'C:\Users\alejandro.gutierrez\OneDrive - Carlin Group - CA Fortune\Desktop\python Item master listing segundo.xlsx', sheet_name='ItemMaster', index = False)



# -- Volvemos a leer el archivo exportado para corroborar
#df3 = pd.read_csv(r'C:\Users\alejandro.gutierrez\OneDrive - Carlin Group - CA Fortune\Desktop\NUEVO YTD segundo.csv')
#df3.dtypes
