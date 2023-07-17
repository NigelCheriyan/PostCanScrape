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


from selenium.common.exceptions import TimeoutException



""" Pull CSV file"""
"Test Sheet-NC.xlsx"
File_Name  =  "Nigel Cheriyan Address Cleanup.xlsx"
Rest_Excel = pd.read_excel(File_Name, sheet_name='Database cleanup Pulled 10 2023',na_values ='NaN').iloc[:,9:]
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


"""function to write in country """
def Country_Search(Country):
    C_Search = Driver.find_element(By.XPATH,Country_Position)
    C_Search.clear()
    C_Search.send_keys(Country)

"""function to get data"""
Descriptions = []
def Get_Address(Joined_Search):
    Locate_Search = Driver.find_element(By.XPATH, Search_Bar_Position)
    Locate_Search.clear()
    Locate_Search.send_keys(Joined_Search)
    Locate_Enter = Driver.find_element(By.XPATH, Enter_Position)
    Locate_Enter.click()
    try:
        WebDriverWait(Driver, 5).until(EC.presence_of_element_located((By.XPATH, Address_Position)))
    except TimeoutException:
        return None
    else:
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
                    return Address, Description
            else:
                return None
        else:
            return None




""" Parse String Format for associated country"""
def Parse_String_CAN(Address,Description):
    Split_Description = Description.split(',')
    City = Split_Description[0]
    Province = Split_Description[1]
    Postal_Code = Split_Description[2]
    return [Address,'','','','',City,Province,Postal_Code,'Canada','Yes']




def Parse_String_Non_CAN(Address, Description):
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

""" function for what to do if info was NOT found on Canada Post website"""

def Unsuccessfull(Row):
    row_unsuccessfull = Row
    row_unsuccessfull['Successfull'] =  'No'
    return row_unsuccessfull

""" function for checking if data fits properly into array"""

def Index_Input_CAN(results):
        if results != None:
            try:
                return Parse_String_CAN(results[0],results[1])
            except IndexError:
                print('There was an Index Error')
                return Unsuccessfull(Row)
        else:
            return Unsuccessfull(Row)

def Index_Input_Non_CAN(results):
        if results != None:
            try:
                return Parse_String_Non_CAN(results[0],results[1])
            except IndexError:
                print('There was an Index Error')
                return Unsuccessfull(Row)
        else:
            return Unsuccessfull(Row)

"""Loop across all the clientel and check address"""
Header = Excel_Sheet.columns.ravel()
Header = np.append(Header, 'Successfull')
Fixed_Sheet_Data_CAN = pd.DataFrame([],columns = Header)
Fixed_Sheet_Data_Non_CAN = pd.DataFrame([],columns = Header)
Excel_Sheet_CAN = Excel_Sheet.loc[Excel_Sheet['PREFERRED ADDRESS LINE COUNTRY'] == 'Canada']
Excel_Sheet_Non_CAN = Excel_Sheet.loc[Excel_Sheet['PREFERRED ADDRESS LINE COUNTRY'] != 'Canada']



for Index, Row in Excel_Sheet_CAN.iterrows():
    print(Index)
    Search = Search_Input(Row)
    Fixed_Sheet_Data_CAN.loc[Index] = Unsuccessfull(Row)
    if Search != 'Series([], )':        
        results = Get_Address(Search)
        New_Address = Index_Input_CAN(results)
        if New_Address[-1]== 'No':
            Search = Row['PREFERRED ADDRESS LINE 1']
            print(Search)
            results = Get_Address(Search)
            New_Address = Index_Input_CAN(results)
            
            if New_Address[-1]== 'No':
                print('Moving on')
            else:
                if New_Address[7] == Row['PREFERRED ADDRESS POSTAL CODE']:
                    print(New_Address)
                    Fixed_Sheet_Data_CAN.loc[Index] = New_Address
                else: 
                    print( "Blah")
        else:
            Fixed_Sheet_Data_CAN.loc[Index] = New_Address

for Index, Row in Excel_Sheet_Non_CAN.iterrows():
    print(Index)
    Search = Search_Input(Row)
    Fixed_Sheet_Data_Non_CAN.loc[Index] = Unsuccessfull(Row)
    if Search != 'Series([], )':
        Country = Row['PREFERRED ADDRESS LINE COUNTRY']
        Country_Search(Country)
        results = Get_Address(Search)
        New_Address = Index_Input_Non_CAN(results)
        if New_Address[-1] == 'No':
            Search = Row['PREFERRED ADDRESS LINE 1']
            results = Get_Address(Search)
            New_Address = Index_Input_Non_CAN(results)

        else:   
            Fixed_Sheet_Data_Non_CAN.loc[Index] = New_Address



# if no postal code and no address , move on


time.sleep(5) # Let the user actually see something!

Driver.quit()

Fixed_Sheet_Data = pd.concat([Fixed_Sheet_Data_CAN,Fixed_Sheet_Data_Non_CAN], axis = 0)
Fixed_Sheet_Data.sort_index(inplace = True)
Full_Sheet = pd.concat([Fixed_Sheet_Data,Rest_Excel],axis = 1)
Full_Sheet.to_excel('Cleaned_Excel_Sheet_DWDC.xlsx', index = False)
