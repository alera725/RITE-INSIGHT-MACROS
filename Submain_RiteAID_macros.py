# -*- coding: utf-8 -*-
"""
Created on Thu Sep  2 22:22:28 2021

@author: alejandro.gutierrez
"""

class process_macros():
    
    def __init__(self):
        self.POS_REPORT_ID = '//*[@id="ext-comp-1006__ext-comp-1002"]'
        

    def macros(self, df):
        try:    
            
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
            return df_1
            
        except:
            print ("Something wrong please check what is happening with Macros Periods")
            
    def macros_IML(self,df):
        try:
            
            # Reacomodar columnas con saltos de linea
            df = df.replace('\n',' ', regex=True)
            
            # Pasar el primer row como header del data frame y resetear el indice
            df = df.rename(columns=df.iloc[0])
            df = df.iloc[1:,]
            df = df.reset_index(drop=True)
            
            # Renombrar columnas con espacios en blanco
            df = df.rename(columns={'Category Description': 'Category_Description', 'Class Description': 'Class_Description', 'Sub Class Description': 'Sub_Class_Description', 'Item No': 'Item_No', 'Original Item No': 'Original_Item_No', 'Active Item No': 'Active_Item_No', 'Item Description': 'Item_Description', 'Sub Brand': 'Sub_Brand', 'Vendor Name': 'Vendor_Name'})
            
            # Eliminamos duplicados en base a la columna UPC
            df_1 = df.drop_duplicates(subset='UPC')
            
            return df_1
            
        except:
            print ("Something wrong please check what is happening with Macros Item Master Listing")

            
    

            
            
            
            
 