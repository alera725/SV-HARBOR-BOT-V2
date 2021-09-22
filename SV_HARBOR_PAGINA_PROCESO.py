# -*- coding: utf-8 -*-
"""
Created on Wed Aug 26 13:55:41 2020

@author: alejandro.gutierrez
"""

# PAGINA DEL PROCESO
#from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait # Guardar en las paginas
from selenium.webdriver.support import expected_conditions as EC #Guardar en las paginas import unittest
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.support.ui import Select

import time
import pandas as pd 
import numpy as np
from datetime import date, timedelta
import datetime


class process_page():
    def __init__(self,my_driver):
        self.driver = my_driver
        self.warehouse_store_shipments = (By.XPATH, '//*[@id="31"]')
        self.shipment_history = (By.XPATH, '//*[@id="pageMenu"]/dd[9]/div/a')
        self.buttonid = (By.ID, 'btnSignIn')
        
        self.clearall = '//*[contains(@id, "j_idt")]' #(By.ID, 'j_idt1056')
        
        self.favorites_button = (By.ID, 'facetMyFavSearch') 
        #table_id = 'myFavoriteSearches:j_idt123' xpath = '//*[@id="myFavoriteSearches:j_idt123"]' #//*[@id="myFavoriteSearches:j_idt123:4:j_idt125"]
        self.tablefavorite_searches = '//*[contains(@id, "myFavoriteSearches:")]' #El ID numero 5 es el bueno 
        #self.monthlydc_button = '//*[contains(@id, "Monthly DC")]' #(By.XPATH, '//*[@id="myFavoriteSearches:j_idt126:4:j_idt128"]') , "//*[@id='vendor:filterVendors']/tbody/tr/td/label[contains(text(), '%s')]
        #//*[@id="myFavoriteSearches:j_idt123"]
        #//*[@id="myFavoriteSearches:j_idt126"]
        
        self.vendor_button = (By.XPATH, '//*[@id="facetVendor"]/div[1]') 
        self.select_all_vendors = (By.ID,'vendor:j_idt948')
        
        self.clear_vendors = '//*[contains(@id, "vendor:j_idt")]' #AL FINAL USAMOS EL CONTAIN PARA BUSCAR EL MAS PARECIDO CON ESE TEXTO  Full xpath = '/html/body/div[1]/div[2]/div[2]/div[2]/div/form/div[2]/span[2]/span[4]/div/div[2]/span/div[1]/a'
        # '//*[@id="vendor:j_idt946"]' #'vendor:j_idt1101') #vendor:j_idt946  #//*[@id="vendor:j_idt714"] CLASS = 'facetsClearTxtLeft' 

        self.table_vendors_filter = (By.ID,'vendor:filterVendors')
        self.date_end_week_button = (By.XPATH,'//*[@id="facetWeekEnding"]/div[1]') #CLick en seccion End Week
        self.day_selected = (By.ID, 'weekEnding:filterWeekEnding:5') #Saturday end week 
        self.date_button = (By.XPATH, '//*[@id="facetDate"]/div[1]') #Click en seccion Date 
        self.date_range = '//*[@id="date:dateEntryType:1"]' # 'date:dateEntryType:1' ELEGIR DATE RANGE EN LUGAR DE SINGLE DAY  
        self.day_initial = 'date:rangeStartDate'
        self.day_end = 'date:rangeEndDate' # NEW 'date:rangeEndDate'
        
        self.pressok = (By.XPATH, '//*[@id="ui-datepicker-div"]/div[2]/button') #(By.CLASS_NAME, 'ui-datepicker-close ui-state-default ui-priority-primary ui-corner-all') #'/html/body/div[3]/div[2]/button')
        #'//*[@id="ui-datepicker-div"]/div[2]/button' //*[@id="ui-datepicker-div"]/div[2]/button

        self.add_dates_button = '//*[contains(@id, "date:j_idt")]' #AL FINAL USAMOS EL CONTAIN PARA BUSCAR EL MAS PARECIDO CON ESE TEXTO  '/html/body/div[1]/div[2]/div[2]/div[2]/div/form/div[2]/span[2]/span[8]/div/div[2]/span/div[2]/span[2]/div[2]/a' 
        #FUll xpath = '/html/body/div[1]/div[2]/div[2]/div[2]/div/form/div[2]/span[2]/span[8]/div/div[2]/span/div[2]/span[3]/div[2]/a'
        #'date:j_idt654' #'date:j_idt468' #date:j_idt195 class = 'short button2' #'//*[@id="date:j_idt189"]'
        
        self.itemcode = (By.XPATH, '//*[@id="facetItemCode"]/div[1]')
        self.ic_dropdown = (By.ID, 'itemCode:itemCodeType')
        self.ibox = (By.ID, 'itemCode:singleItemCd')
        self.add_ic_button = '//*[contains(@id, "itemCode:j_idt")]' #itemCode:j_idt229 es el 2
        
        self.download_button = (By.XPATH, '//*[@id="downloadmenu_label"]/a')
        self.store_level_detail = '//*[contains(@id, "j_idt")]' #(By.XPATH, '/html/body/div[1]/div[2]/div[2]/div[2]/div/form/div[1]/div[2]/div[1]/div[2]/div/div/div[3]') #'j_idt1317') 
        self.view_downloads = (By.ID, 'viewDownloads')
        self.Download_Hour = (By.ID, 'downloadDetailsTable:0:reqTS')
        self.Download_Table = (By.ID, 'downloadDetailsTable:tb')
        self.status_down = 'downloadDetailsTable:0:status' # full xpath = '/html/body/div[4]/div[2]/div/div[3]/div[1]/form/div[2]/div/table/tbody[1]/tr[1]/td[5]'
        self.refresh_button = 'refresh'
        self.file_name = (By.ID, 'downloadDetailsTable:0:docName')
        self.close_window_download = (By.XPATH ,'//*[@id="topLinks"]/a')
        
        self.week_ending = (By.ID ,'facetWeekEnding')
        self.week_ending_day = (By.ID ,'weekEnding:filterWeekEnding:5')
        
    def first_window(self):
        try:
            warehouse_store_button = WebDriverWait(self.driver,50).until(EC.element_to_be_clickable(self.warehouse_store_shipments))
            warehouse_store_button.click()
            
            #Cuidado a veces no es necesario dar click en el WholeSale ya esta clickado
            shipment_history_button = WebDriverWait(self.driver,50).until(EC.element_to_be_clickable(self.shipment_history))
            shipment_history_button.click()

        except TimeoutException:
            print ("Loading -first_window- took too much time!")
            
    def clear_all(self):
        try:
            
            #vendorbutt = WebDriverWait(self.driver,200).until(EC.element_to_be_clickable(self.vendor_button)) #Para hacer tiempo
            c_all = self.driver.find_elements_by_xpath(self.clearall)[3] #El 3 es el de clear all
            dynamic_id_clear_all = c_all.get_attribute("id")
            #print(dynamic_id_clear_all)
            dynamic_id_clear_all_by = (By.ID,dynamic_id_clear_all)
            clear_a = WebDriverWait(self.driver,50).until(EC.visibility_of_element_located(dynamic_id_clear_all_by)) 
            clear_a.click()
            time.sleep(2)
            
        except TimeoutException:
            print ("Loading -clear_all- took too much time!")
            
    def set_vendor(self,retailers,num):
        try:
            time.sleep(2)
            vendorbutt = WebDriverWait(self.driver,200).until(EC.element_to_be_clickable(self.vendor_button))
            vendorbutt.click()
            
        #    time.sleep(5)
            
            #MODIFICAR CON DYNAMIC ID
            #clear_filter_contain = self.driver.find_element_by_xpath(self.clear_vendors)
            #dynamic_id_clear_filter = clear_filter_contain.get_attribute("id")
            #print(dynamic_id_clear_filter)
            
        #    time.sleep(5)
            #vendorbutt = WebDriverWait(self.driver,200).until(EC.element_to_be_clickable(self.vendor_button))
           # vendorbutt.click()
           # time.sleep(5)
            
            #clear_filter = self.driver.find_element_by_id(dynamic_id_clear_filter)
            #clear_filter.click()
            
           # time.sleep(2)
           
            #Select all vendors
            #all_vendors = WebDriverWait(self.driver,20).until(EC.element_to_be_clickable(self.select_all_vendors))
            #all_vendors.click()
            
            #Select only the vendor we guess
            for i in retailers:
                retailer = i
                #print(retailer)
                table = self.driver.find_element_by_xpath("//*[@id='vendor:filterVendors']/tbody/tr/td/label[contains(text(), '%s')]" %retailer)
                table.click() #Segun eso aqui es donde fallo
                
            if(num==1):
                vendorbutt.click()
                time.sleep(8)  
            else:
                pass
            #//*[@id="vendor:filterVendors:1"]
            #/html/body/div[1]/div[2]/div[2]/div[2]/div/form/div[2]/span[2]/span[4]/div/div[2]/span/div[3]/table/tbody/tr[2]/td/input
            #table = webdriver.find_element_by_xpath("//*[@id='vendor:filterVendors']/tbody/tr/td/label[contains(text(),'%s')]" %retailer)
            #time.sleep(2)
            #retailer = retailers
            #table = self.driver.find_element_by_xpath("//*[@id='vendor:filterVendors']/tbody/tr/td/label[contains(text(), '%s')]" %retailer)
            #table.click()
            
            #select_retailer = Select(self.driver.find_element(*self.table_vendors_filter))
            #select_retailer.select_by_index(1)
           
           
            #Esto es para despues tal vez serviria para el if si son all vendors o uno en especial
           # if retailers!='all':
                #Select only the vendor we guess
           #     retailer = retailers
           #     select_retailer = Select(self.driver.find_element_by_id(self.first_date_filter))
           #     select_retailer.select_by_index(1)
                
           #     select_datasource = Select(self.driver.find_element(*self.select_dropdown))
           #     select_datasource.select_by_visible_text(newdata)
            
                
           # else:
                #Select all vendors
           #     all_vendors = WebDriverWait(self.driver,20).until(EC.element_to_be_clickable(self.select_all_vendors))
           #     all_vendors.click()


        except TimeoutException:
            print ("Loading -set_vendor- took too much time!")
    
    def set_MonthlyDC(self):
        try:
            
            fav_button = WebDriverWait(self.driver,50).until(EC.visibility_of_element_located(self.favorites_button))
            fav_button.click()
            
            #mdc_button = WebDriverWait(self.driver,50).until(EC.visibility_of_element_located(self.monthlydc_button))
            #mdc_button.click()
            
            fav_table = self.driver.find_elements_by_xpath(self.tablefavorite_searches)[4] #El ID 4 es el que me interesa
            dynamic_id_fav_table = fav_table.get_attribute("id")
            dynamic_id_fav_table_by = (By.ID, dynamic_id_fav_table)
            
            #Encontrar el td que contenga la palabra "Monthly DC" #//*[@id="myFavoriteSearches:j_idt123:4:j_idt125"]
            find_text = 'Monthly DC'
            #mdc = self.driver.find_element_by_xpath("//*[@id='%s']/tbody/tr/td[contains(text(), '%s')]" % (dynamic_id_fav_table, find_text))
            mdc_table =  WebDriverWait(self.driver,10).until(EC.visibility_of_element_located(dynamic_id_fav_table_by))
            mdc_table.find_element(By.LINK_TEXT,'%s' %find_text).click()
            #mdc_table.click()
            time.sleep(4)
            
           # mdc = self.driver.find_elements_by_xpath(self.monthlydc_button)[3] 
           # dynamic_id_clear_all = mdc.get_attribute("id")
            #print(dynamic_id_clear_all) 
           # dynamic_id_clear_all_by = (By.ID,dynamic_id_clear_all)
           # clear_a = WebDriverWait(self.driver,50).until(EC.visibility_of_element_located(dynamic_id_clear_all_by)) 
           # clear_a.click()
            
        except TimeoutException:
            print ("Loading -set_MotnhlyDC- took too much time!")   
            
            
    def set_week_ending(self):
        try:
            
            week_end_button = WebDriverWait(self.driver,50).until(EC.visibility_of_element_located(self.week_ending))
            week_end_button.click()
            
            week_end_table = WebDriverWait(self.driver,50).until(EC.visibility_of_element_located(self.week_ending_day))
            week_end_table.click()
            
            week_end_button.click()
            
            time.sleep(4)

            
        except TimeoutException:
            print ("Loading -set_week_ending- took too much time!")  
            
            
    def set_Halloween(self):
        try:
            
            fav_button = WebDriverWait(self.driver,50).until(EC.visibility_of_element_located(self.favorites_button))
            fav_button.click()
            
            fav_table = self.driver.find_elements_by_xpath(self.tablefavorite_searches)[4] #El ID 4 es el que me interesa
            dynamic_id_fav_table = fav_table.get_attribute("id")
            dynamic_id_fav_table_by = (By.ID, dynamic_id_fav_table)
            
            #Encontrar el td que contenga la palabra "Monthly DC" #//*[@id="myFavoriteSearches:j_idt123:4:j_idt125"]
            find_text = 'Haloween candy report1'
            mdc_table =  WebDriverWait(self.driver,10).until(EC.visibility_of_element_located(dynamic_id_fav_table_by))
            mdc_table.find_element(By.LINK_TEXT,'%s' %find_text).click()
            time.sleep(4)
            
        except TimeoutException:
            print ("Loading -set_Halloween- took too much time!") 
            
            
            
    def set_endweekday(self):
        try:
            date_end_button = WebDriverWait(self.driver,200).until(EC.visibility_of_element_located(self.date_end_week_button))
            date_end_button.click()
            
            time.sleep(2)
            WebDriverWait(self.driver,20).until(EC.visibility_of_element_located(self.day_selected)).click()
            

        except TimeoutException:
            print ("Loading -set_endweekday- took too much time!")


    def set_date_range(self,fecha_inicio,fecha_fin):
        try:
            time.sleep(5)
            
            datebutton = WebDriverWait(self.driver,20).until(EC.element_to_be_clickable(self.date_button))
            datebutton.click()
            
            time.sleep(15)
            
            rango = self.driver.find_element_by_xpath(self.date_range)
            rango.click()
            #select_multiple_range.click()
            
            time.sleep(15)
           
            dia_end = self.driver.find_element_by_id(self.day_end)
            dia_end.click()
            dia_end.send_keys(fecha_fin) #EL VALOR DE LA FECHA EN FORMATO MMDDYY EJ. 080120 (08 DE AGOSTO DEL 2020)
            
            #dia_end.submit()
            
            dia_start = self.driver.find_element_by_id(self.day_initial)
            dia_start.click()
            dia_start.send_keys(fecha_inicio) #EL VALOR DE LA FECHA EN FORMATO MMDDYY EJ. 080120 (08 DE AGOSTO DEL 2020)
            

            
            #Despues de esto otra vez el date y clickar add
            #datebutton = WebDriverWait(self.driver,20).until(EC.element_to_be_clickable(self.date_button))
            #datebutton.click()
            
            time.sleep(2)
            
            #shot_ok = WebDriverWait(self.driver,20).until(EC.visibility_of_element_located(self.pressok))
            #shot_ok.click()
            
            #MODIFICAR CON DYNAMIC ID
            
            add_button_contain = self.driver.find_elements_by_xpath(self.add_dates_button)[1] #VER QUE NUMERO DE ELEMENTO ES EL ID DE ADD DATES [4] EJ.
            dynamic_id_add_button = add_button_contain.get_attribute("id")
            #print(dynamic_id_add_button)
            add_button = self.driver.find_element_by_id(dynamic_id_add_button)
            add_button.click()
            
            time.sleep(2)
            

        except TimeoutException:
            print ("Loading -set_date_range- took too much time!")
            
            
    def set_Item_Code(self):
        try:
            
            # Item code click()
            drop_down_button = WebDriverWait(self.driver,50).until(EC.visibility_of_element_located(self.itemcode))
            drop_down_button.click()
            
            # Dropdown_click()
            select = Select(self.driver.find_element(*self.ic_dropdown)) 
            select.select_by_visible_text('Item UPC')            
            
            # Put item ID's
            items = '05543763118,05543763117,05543764101'
            items_box = WebDriverWait(self.driver,50).until(EC.visibility_of_element_located(self.ibox))
            items_box.send_keys(items)
            
            
            add_button_contain = self.driver.find_elements_by_xpath(self.add_ic_button)[2] #VER QUE NUMERO DE ELEMENTO ES EL ID DE ADD DATES [4] EJ.
            dynamic_id_add_button = add_button_contain.get_attribute("id")
            print('Id used: ' + dynamic_id_add_button)
            
            add_button_ic = self.driver.find_element_by_id(dynamic_id_add_button)
            add_button_ic.click()
            
            #time.sleep(50)
            
            #Encontrar el td que contenga la palabra "Monthly DC" #//*[@id="myFavoriteSearches:j_idt123:4:j_idt125"]
            #find_text = 'Haloween candy report1'
            #mdc_table =  WebDriverWait(self.driver,10).until(EC.visibility_of_element_located(dynamic_id_fav_table_by))
            #mdc_table.find_element(By.LINK_TEXT,'%s' %find_text).click()
            #time.sleep(4)
            
        except TimeoutException:
            print ("Loading -set_Item_Code- took too much time!") 
            


    def download_page(self):
        try:
            down_button = WebDriverWait(self.driver,20).until(EC.element_to_be_clickable(self.download_button))
            down_button.click()
            
            
            #Aqui falta acomodar el find elements 
            #sld_button = WebDriverWait(self.driver,20).until(EC.visibility_of_element_located(self.store_level_detail))
            #sld_button.click()
            
            sld = self.driver.find_elements_by_xpath(self.clearall)[2] #El 3 es el de clear all
            dynamic_id_sld = sld.get_attribute("id")
            #print(dynamic_id_sld)
            sld_button = self.driver.find_element_by_id(dynamic_id_sld)
            sld_button.click()
            
            time.sleep(1)
            view_downs = WebDriverWait(self.driver,20).until(EC.element_to_be_clickable(self.view_downloads))
            view_downs.click()
            
            #time.sleep(2)
            
            #CHECAR LO DEL ROW DE LA HORA Y EL NOMBRE PARA UBICAR EL ARCHIVO QUE NOSOTROS DESCARGAMOS
            #Aqui es donde debo de fijarme en el row de la descarga que yo hice, la hora y la descripcion que tenga: Store Level Detail
            #DOWNLOAD TYPE: VENDOR_SHIPMENT_HISTORY ID: 'downloadDetailsTable:0:type' XPATH: '//*[@id="downloadDetailsTable:0:type"]'
            #HORA 09/04/20 15:30:04 id = 'downloadDetailsTable:0:reqTS' xpath = '//*[@id="downloadDetailsTable:0:reqTS"]/div
            #DESCRIPCION Vendor Shipment Store Level Detail	 id = 'downloadDetailsTable:0:desc' xpath = '//*[@id="downloadDetailsTable:0:desc"]'  texto debe decir: 'Vendor Shipment Store Level Detail'
            
            time.sleep(2) #Usar los WebdriverWaits
            #DOWNLOAD_TYPE = self.driver.find_element_by_id('downloadDetailsTable:0:type').text
            HOUR = WebDriverWait(self.driver,20).until(EC.visibility_of_element_located(self.Download_Hour)) #webdriver.find_element_by_id('downloadDetailsTable:0:reqTS').text
            HOUR = HOUR.text
            print(HOUR)
            #DESC = self.driver.find_element_by_id('downloadDetailsTable:0:desc').text     
            
            #El status que debemos de estar checando es en donde este estos datos en el renglon y despues el click igual 
            
            #FORMA 1 TE REGRESA TODOS LOS ROWS DE LA COLUMNA INDICADA y da click en el row donde encontro un valor especifico
            table_id = WebDriverWait(self.driver,20).until(EC.visibility_of_element_located(self.Download_Table)) #self.driver.find_element(By.ID, 'downloadDetailsTable:tb')
            rows = table_id.find_elements(By.TAG_NAME, "tr") # get all of the rows in the table
            
            for row in rows:
                col1 = row.find_elements(By.TAG_NAME, "td")[1] #Revisamos la HORA
                print(col1)
                if(col1.text == HOUR):
                    col = row.find_elements(By.TAG_NAME, "td")[4] #Revisamos la columna donde tenemos que clickar
                    actual_status = col.text
                    print(actual_status)
                    break
                else:
                    continue            
            
            #actual_status =  self.driver.find_element_by_id(self.status_down).text
            
            while actual_status!='Ready':
                self.driver.find_element_by_id(self.refresh_button).click()
                time.sleep(8)
                table_id = WebDriverWait(self.driver,20).until(EC.visibility_of_element_located(self.Download_Table)) #self.driver.find_element(By.ID, 'downloadDetailsTable:tb')
                rows = table_id.find_elements(By.TAG_NAME, "tr") # get all of the rows in the table
                
                ##actual_status =  self.driver.find_element_by_id(self.status_down).text      
                for row in rows:
                    #WebDriverWait(self.driver,20).until(EC.visibility_of_element_located((By.TAG_NAME, "td")))
                    
                    
                    my_element_id = "td"
                    ignored_exceptions=(NoSuchElementException,StaleElementReferenceException)
                    WebDriverWait(self.driver,20,ignored_exceptions=ignored_exceptions)\
                                            .until(EC.presence_of_element_located((By.TAG_NAME, my_element_id)))
                                            
                                            
                    col1text = row.find_elements(By.TAG_NAME, "td")[1].text #Revisamos la HORA
                    #print(col1text)
                    #col1text = col1.text
                    if(col1text == HOUR):
                        #WebDriverWait(self.driver,20).until(EC.visibility_of_element_located((By.TAG_NAME, "td")))
                        actual_status = row.find_elements(By.TAG_NAME, "td")[4].text #Revisamos la columna donde tenemos que clickar
                        print(actual_status)
                        #actual_status = col.text
                        #Break Continue o que para que siga con las instrucciones tal vez meter aqui el if que sigue y hacer un else con continue
                        #if status es Failed cerrar todo y cancelar operacion
                        if (actual_status == 'Failed'):
                            break
                        else:
                            continue            
                    else:
                        continue
            
            #El texto ya esta en READY! o FAILED!
                    
                                       
                    
                #if status es Failed cerrar todo y cancelar operacion 
               # if (actual_status == 'Failed'):
               #     break
               # else:
               #     continue            
                  
            
                        
            #Falta lo de Actual Status**** checarlo en el row que es 
            
         #   actual_status =  self.driver.find_element_by_id(self.status_down).text
            
         #   while actual_status!='Ready':
         #       self.driver.find_element_by_id(self.refresh_button).click()
         #       time.sleep(1)
         #       actual_status =  self.driver.find_element_by_id(self.status_down).text
                #if status es Failed cerrar todo y cancelar operacion
         #       if (actual_status == 'Failed'):
         #           break
         #       else:
         #           continue
            
            if actual_status == 'Ready':
                #fname = WebDriverWait(self.driver,20).until(EC.visibility_of_element_located(self.file_name))
                #fname.click()
                #Esto va en el FNAME el IF despues del while 
                for row in rows:
                    col1 = row.find_elements(By.TAG_NAME, "td")[1] #Revisamos la HORA
                    if(col1.text == HOUR):
                        fname = row.find_elements(By.TAG_NAME, "td")[3] #Revisamos la columna donde tenemos que clickar
                        fname.click()
                        break
                    else:
                        continue
                
            else:
                print('Error: Download status "Failed" ')
            
            
        except TimeoutException:
            print ("Loading -download_page- took too much time!")
        
        

    def close_window(self):
        try:
            
            c_window = WebDriverWait(self.driver,10).until(EC.visibility_of_element_located(self.close_window_download))
            c_window.click()    
            time.sleep(1)
        except:
            print ("Loading -close_window- took too much time!")
            
            
            
    def macros_SV_Harbor(self, num_files, import_path, export_path, week_number):
        try:
            if (num_files==2):
                # Leer los archivos separados
                data_sv1 = pd.read_csv(import_path + '\\FIRST SUPER VALU MONTLHY DC week %s.csv'%week_number, skiprows=5) #leer week 53 'FIRST SUPER VALU MONTLHY DC week 53
                data_sv2 = pd.read_csv(import_path + '\\SECOND SUPER VALU MONTLHY DC week %s.csv'%week_number, skiprows=5) #leer week 1 'SECOND SUPER VALU MONTLHY DC week 1'
                
                #Si no funciona hacerlo manual sin el for 
                for i in range(2):
                    # Renombrar columnas con espacios en blanco
                    exec("data_sv%d = data_sv%d.rename(columns={'Invoice Week': 'Invoice_Week', 'Corp Code': 'Corp_Code', 'Order Type': 'Order_Type', 'Vendor#': 'Vendorno', 'Product Group': 'Product_Group', 'Class Group': 'Class_Group', 'Supporting DC Item': 'Supporting_DC_Item'})"%((i+1),(i+1)))
                    
                    # Cambiar el tipo de dato de dos columnas a numericas (no fue necesario pero ver VWPOUNDS A INT64 Y SVNETWT A INT64)
                    exec("data_sv%d['SVNetWt'] = data_sv%d['SVNetWt'].astype('int64')"%((i+1),(i+1)))
                    
                    # Es diferente el proceso ya que VWPounds era float64 pero si tenia valores decimales en la base no como SVNetWt
                    exec("data_sv%d['VWPounds'] = data_sv%d['VWPounds'].fillna(0).astype(np.int64, errors='ignore')"%((i+1),(i+1)))
                    
                    # Exportamos el archivo a la ubicacion deseada 
                    exec("data_sv%d.to_csv(export_path + '\\DataSV-%d.csv', index = False)"%((i+1),(i+1)))
                
            else:
                # Leer el archivo 
                data_sv = pd.read_csv(import_path + '\\SUPER VALU MONTLHY DC week %s.csv'%week_number, skiprows=5) #'SUPER VALU MONTLHY DC 25-Aug-2021 14Hr 44Min'
                
                # Renombrar columnas con espacios en blanco
                data_sv = data_sv.rename(columns={'Invoice Week': 'Invoice_Week', 'Corp Code': 'Corp_Code', 'Order Type': 'Order_Type', 'Vendor#': 'Vendorno', 'Product Group': 'Product_Group', 'Class Group': 'Class_Group', 'Supporting DC Item': 'Supporting_DC_Item'})
                
                # Cambiar el tipo de dato de dos columnas a numericas (no fue necesario pero ver VWPOUNDS A INT64 Y SVNETWT A INT64)
                data_sv['SVNetWt'] = data_sv['SVNetWt'].astype('int64')
                # Es diferente el proceso ya que VWPounds era float64 pero si tenia valores decimales en la base no como SVNetWt
                data_sv['VWPounds'] = data_sv['VWPounds'].fillna(0).astype(np.int64, errors='ignore')
                
                # Exportamos el archivo a la ubicacion deseada 
                data_sv.to_csv(export_path + '\\DataSV.csv', index = False)
                
        except:
            print ("An error occurred in macros_SV_Harbor")
        
            