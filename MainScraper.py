# -*- coding: utf-8 -*-
"""
Created on Thu Jun 29 11:48:53 2023
Web Scraper to double check all addresses with Canada Post

For Dying with Dignity
@author: nigel
"""
import pandas as pd
import numpy as np
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException




""" Pull CSV file- Real File - Test Sheet-NC.xlsx"""

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

Error_Position = '//*[@id="pnlError"]/table/tbody/tr/td[2]'
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
        Driver.find_element(By.XPATH,Error_Position)
    except NoSuchElementException:
        try:
            Driver.find_element(By.XPATH,Second_Position)
        except NoSuchElementException:
            Address_Output = Driver.find_element(By.XPATH, Address_Position)
            Address = Address_Output.text
            Description_Output = Driver.find_element(By.XPATH, Description_Position)
            Description = Description_Output.text
            Descriptions.append(Description)
            if Description[-9:-1] == 'Addresse':
                return None
        else:
            return None
    else:
        return None

    return Address, Description




""" Parse String Format for associated country"""

def Parse_String(Address, Description):
        if Country == 'Canada':
            Split_Description = Description.split(',')
            City = Split_Description[0]
            Province = Split_Description[1]
            Postal_Code = Split_Description[2]
            return [Address,'','','','',City,Province,Postal_Code,Country,'Yes']
        if Country == 'United States' or 'USA':
            Split_Description = Description.split(' ')
            City = Split_Description[0]
            Province = Split_Description[1]
            Postal_Code = str(Split_Description[2])
            return [Address,'','','','',City,Province,Postal_Code,Country,'Yes']

        if Country == 'Australia':
            Split_Description = Address.split(',')
            Address = Split_Description[0]
            Split_Description = Split_Description[1].split(' ')
            City = Split_Description[0]
            Province = Split_Description[1]
            Postal_Code = Split_Description[2]
            return [Address,'','','','',City,Province,Postal_Code,Country,'Yes']
        if Country in ['United Kingdom','England','UK','Belgium']:
            Split_Description = Description.split(',')
            City = Split_Description[0]
            Postal_Code = Split_Description[1]
            return [Address,'','','','',City,'',Postal_Code,Country,'Yes']
        else:
            Split_Description = Description.split(' ')
            City = Split_Description[0]
            Postal_Code = Split_Description[1]
            return [Address,'','','','',City,'',Postal_Code,Country,'Yes']



"""Loop across all the clientel and check address"""
Header = Excel_Sheet.columns.ravel()
Header = np.append(Header, 'Successfull')
Fixed_Sheet_Data = pd.DataFrame([],columns = Header)


for Index, Row in Excel_Sheet.iterrows():
    print(Index)
    Country = Row['PREFERRED ADDRESS LINE COUNTRY']
    Search = Search_Input(Row)
    if str(Country) == 'nan':
        Country = 'Canada'
    else:
        pass

    if Search == 'Series([], )':
        row_unsuccessfull = Row
        row_unsuccessfull['Successfull'] =  'No'
        Fixed_Sheet_Data.loc[len(Fixed_Sheet_Data)] = row_unsuccessfull
    else:
        results = Get_Address(Search)
        if results != None:
            try:
                Full_Parsed_String = Parse_String(results[0],results[1])
                Fixed_Sheet_Data.loc[len(Fixed_Sheet_Data)] = Full_Parsed_String
            except IndexError:
                print('There was an Index Error')
                pass
        else:
            row_unsuccessfull = Row
            row_unsuccessfull['Successfull'] =  'No'
            Fixed_Sheet_Data.loc[len(Fixed_Sheet_Data)] = row_unsuccessfull
            pass





# if no postal code and no address , move on


time.sleep(5) # Let the user actually see something!

Driver.quit()

Fixed_Sheet_Data.to_excel('Cleaned_Excel_Sheet_DWDC.xlsx')
