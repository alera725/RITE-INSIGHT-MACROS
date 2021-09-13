# -*- coding: utf-8 -*-
"""
Created on Thu Sep  2 22:22:25 2021

@author: alejandro.gutierrez
"""

#Importar paqueterias
import os
os.chdir('C:\\Users\\alejandro.gutierrez\\OneDrive - Carlin Group - CA Fortune\Documents\\ALEJANDRO RAMOS GTZ\\GIT\\RITE INSIGHT MACROS') # relative path: scripts dir is under Lab

import unittest
import time 
import datetime
import pandas as pd
from datetime import date, timedelta, datetime

from Submain_RiteAID_macros import process_macros


class Macros_RITE_INSIGHT_DATA(unittest.TestCase):
    
    def setUp(self):
        
        self.PageProcess = process_macros()
        
        #paths to import each report 
        self.import_path_periods = 'C:\\Users\\alejandro.gutierrez\\OneDrive - Carlin Group - CA Fortune\\Documents\\KROGER SELENIUM\\RITE INSIGHT\\PERIODS\\' 
        self.import_path_ytd = 'C:\\Users\\alejandro.gutierrez\\OneDrive - Carlin Group - CA Fortune\\Documents\\KROGER SELENIUM\\RITE INSIGHT\\YTD\\' 
        self.import_path_IML = 'C:\\Users\\alejandro.gutierrez\\OneDrive - Carlin Group - CA Fortune\\Documents\\KROGER SELENIUM\\RITE INSIGHT\\ITEM MASTER LISTING COMPETITIVE\\'
        #self.import_path_inventory = 'C:\\Users\\alejandro.gutierrez\\OneDrive - Carlin Group - CA Fortune\\Documents\\KROGER SELENIUM\\RITE INSIGHT\\INVENTORY'

        #Export path CAMBIAR PARA HACER PRUEBAS
        #self.export_path = 'C:\\Users\\alejandro.gutierrez\\OneDrive - Carlin Group - CA Fortune\\Documents\\ALEJANDRO RAMOS GTZ\\CARLIN\\RITE AID'
        self.export_path = 'C:\\Users\\alejandro.gutierrez\\OneDrive - Carlin Group - CA Fortune\\Documents\\KROGER SELENIUM\\RITE INSIGHT\\FINAL WEEK FILES'

        #Week number 
        t1 = datetime.now()
        self.week_number = str(t1.strftime("%U"))
        
    #@unittest.skip('Not need now')
    def test_YTD(self):
        
        import_name = 'SBIC YTD 27DEC2020 TO LAST SATURDAY week %s.xlsx'%self.week_number
        export_name = 'RA YTD.xlsx'
        
        #Cambiar el nombre al archivo y aqui modificar cuando cambiemos de periodo anual
        #df = pd.read_excel(self.import_path_ytd + 'SBIC YTD 27DEC2020 TO LAST SATURDAY week %s.xlsx'%self.week_number, sheet_name='Sales')
        df = pd.read_excel(self.import_path_ytd + import_name, sheet_name='Sales') #DESBLOQUEAR EL COMENTARIO DE ARRIBA  REAL
        df_1 = self.PageProcess.macros(df)
        
        df_1.to_excel(self.export_path + '\\%s'%export_name, sheet_name='Sales', index = False)

        print("%s is READY!!"%export_name) 
        

    #@unittest.skip('Not need now')
    def test_52_weeks(self):
        
        import_name = 'SBIC 52 WEEKS week %s.xlsx'%self.week_number
        export_name = 'RA L52W.xlsx'
        
        #Import file and pass to macros module to transform 
        df = pd.read_excel(self.import_path_periods + import_name, sheet_name='Sales')
        df_1 = self.PageProcess.macros(df)
        
        #Export file 
        df_1.to_excel(self.export_path + '\\%s'%export_name, sheet_name='Sales', index = False)

        print("%s is READY!!"%export_name) 
        

    #@unittest.skip('Not need now')
    def test_26_weeks(self):
        
        import_name = 'SBIC 26 WEEKS week %s.xlsx'%self.week_number
        export_name = 'RA L26W.xlsx'
        
        #Import file and pass to macros module to transform 
        df = pd.read_excel(self.import_path_periods + import_name, sheet_name='Sales')
        df_1 = self.PageProcess.macros(df)
        
        #Export file 
        df_1.to_excel(self.export_path + '\\%s'%export_name, sheet_name='Sales', index = False)

        print("%s is READY!!"%export_name) 
        
    
    #@unittest.skip('Not need now')
    def test_13_weeks(self):
        
        import_name = 'SBIC 13 WEEKS week %s.xlsx'%self.week_number
        export_name = 'RA L13W.xlsx'
        
        #Import file and pass to macros module to transform 
        df = pd.read_excel(self.import_path_periods + import_name, sheet_name='Sales')
        df_1 = self.PageProcess.macros(df)
        
        #Export file 
        df_1.to_excel(self.export_path + '\\%s'%export_name, sheet_name='Sales', index = False)

        print("%s is READY!!"%export_name) 

        
    #@unittest.skip('Not need now')
    def test_04_weeks(self):
        
        import_name = 'SBIC 4 WEEKS week %s.xlsx'%self.week_number
        export_name = 'RA L04W.xlsx'
        
        #Import file and pass to macros module to transform 
        df = pd.read_excel(self.import_path_periods + import_name, sheet_name='Sales')
        df_1 = self.PageProcess.macros(df)
        
        #Export file 
        df_1.to_excel(self.export_path + '\\%s'%export_name, sheet_name='Sales', index = False)

        print("%s is READY!!"%export_name) 
        
 
    #@unittest.skip('Not need now')
    def test_01_week(self):
        
        import_name = 'SBIC 1 WEEK week %s.xlsx'%self.week_number
        export_name = 'Weekly.xlsx'
        
        #Import file and pass to macros module to transform 
        df = pd.read_excel(self.import_path_periods + import_name, sheet_name='Sales')
        df_1 = self.PageProcess.macros(df)
        
        #Export name 2 between dates 
        export_name_2 = str(df_1.iloc[0,0]+'.xlsx').replace('/','-')


        #Export file 
        df_1.to_excel(self.export_path + '\\%s'%export_name, sheet_name='Sales', index = False)
        df_1.to_excel(self.export_path + '\\%s'%export_name_2, sheet_name='Sales', index = False)

        print("%s is READY!!"%export_name) 
        
        
    #@unittest.skip('Not need now')
    def test_IML(self):
        
        import_name = 'ITEM MASTER LISTING 52 WEEKS week %s.xlsx'%self.week_number
        export_name = 'Item_Master_Listing.xlsx'
        
        #Import file and pass to macros module to transform 
        df = pd.read_excel(self.import_path_IML + import_name, skiprows= [0,1,2], sheet_name='ItemMaster')
        df_1 = self.PageProcess.macros_IML(df)
        
        #Export file 
        df_1.to_excel(self.export_path + '\\%s'%export_name, sheet_name='ItemMaster', index = False)
        
        
        print("%s is READY!!"%export_name) 
        
        
        
if __name__ == '__main__':
    unittest.main()