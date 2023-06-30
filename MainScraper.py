# -*- coding: utf-8 -*-
"""
Created on Thu Jun 29 11:48:53 2023
Web Scraper to double check all addresses with Canada Post

For Dying with Dignity
@author: nigel
"""
import urllib, xml.dom.minidom
import pandas as pd
import xlsxwriter
import numpy as np
from random import randint

import time

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException




""" Pull CSV file- Real File - Nigel Cheriyan Address Cleanup.xlsx """

File_Name  = "Nigel Cheriyan Address Cleanup.xlsx"

Excel_Sheet = pd.read_excel(File_Name, sheet_name='Database cleanup Pulled 10 2023',na_values ='NaN').iloc[:,0:9]# pull data from file 


""" Create search input function """

def Search_Input(Row):
    Url_Row = Row[0:8].dropna()
    Unjoined_Search = Url_Row.to_string(header=False,index=False).split('\n')
    Joined_Search = ' '.join(Unjoined_Search)
    return Joined_Search

""" Pull up Website ~~~ CODE FROM CAN POST API ADDRESS COMPLETE"""
#Key,SearchTerm, LastId, SearchFor , Country, LanguagePreference, MaxSuggestions, MaxResults, Origin, Bias, Filters, GeoFence


Url = "https://www.canadapost-postescanada.ca/ac/support/api/addresscomplete-interactive-find/"

Driver = webdriver.Chrome()

Driver.get(Url)

Search_Bar_Position = '//*[@id="tryitnow"]/div/div/div/div/div/section[1]/div/table/tbody/tr[1]/td[2]/input'
Country_Position = '//*[@id="tryitnow"]/div/div/div/div/div/section[1]/div/table/tbody/tr[3]/td[2]/input'
Enter_Position  = '//*[@id="btnTest"]'
Address_Position = '//*[@id="pnlResults"]/table/tbody/tr/td[2]'
Description_Position = '//*[@id="pnlResults"]/table/tbody/tr/td[5]'
Second_Position = '//*[@id="pnlResults"]/table/tbody/tr[2]/td[2]'
"""function to get data""" 
Descriptions = []
def Get_Address(Joined_Search):
    Locate_Search = Driver.find_element(By.XPATH, Search_Bar_Position)  
    Locate_Search.clear()
    Locate_Search.send_keys(Joined_Search)
    Country_Search = Driver.find_element(By.XPATH,Country_Position)
    Country_Search.clear()
    Country_Search.send_keys(Country)
    Locate_Enter = Driver.find_element(By.XPATH, Enter_Position)
    Locate_Enter.click()
    WebDriverWait(Driver, 5).until(EC.presence_of_element_located((By.XPATH, Address_Position)))
    try:
        plastic = Driver.find_element(By.XPATH,Second_Position)
        return None
        
    except NoSuchElementException:                   
        Address_Output = Driver.find_element(By.XPATH, Address_Position)
        Address= [Address_Output.text]
        Description_Output = Driver.find_element(By.XPATH, Description_Position)
        
        Description = Description_Output.text
        Descriptions.append(Description)
        time.sleep(randint(1,5))
        if Description[-9:-1] == 'Addresse':
            return None
        else:
            if Country == 'United States':
                result = Address + Description.split(' ')
            else:    
                result = Address + Description.split(',')
            return result

    


"""Loop across all the clientel and check address"""
Header = Excel_Sheet.columns.ravel()
Header = np.append(Header, 'Successfull')
Fixed_Sheet_Data = pd.DataFrame([],columns = Header)
for Index, Row in Excel_Sheet.iterrows():
    Country = Row['PREFERRED ADDRESS LINE COUNTRY']
    results = Get_Address(Search_Input(Row))
    if results == None:
        row_unsuccessfull = Row
        row_unsuccessfull['Successfull'] =  'No'
        Fixed_Sheet_Data = Fixed_Sheet_Data.append(row_unsuccessfull)
        pass
    else:
        Address = results[0]
        City = results[1]
        Province = results[2]
        Postal_Code = results[3]
        Fixed_Sheet_Data.loc[len(Fixed_Sheet_Data)] = [Address,'','','','',City,Province,Postal_Code,Country,'Yes']
    
    
    

# if no postal code and no address , move on 


time.sleep(5) # Let the user actually see something!

Driver.quit()

Fixed_Sheet_Data.to_excel('Cleaned_Excel_Sheet_DWDC.xlsx')