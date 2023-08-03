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

from selenium.common.exceptions import TimeoutException


from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive



st = time.time()
pts = time.process_time()
""" Pull CSV file"""

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
Error = Driver.find_elements(By.XPATH,Error_Position)
Second = Driver.find_elements(By.XPATH,Second_Position)
def Get_Address(Joined_Search):
    Locate_Search = Driver.find_element(By.XPATH, Search_Bar_Position)
    Locate_Search.clear()
    Locate_Search.send_keys(Joined_Search)
    Locate_Enter = Driver.find_element(By.XPATH, Enter_Position)
    Locate_Enter.click()
    WebDriverWait(Driver, 5).until(EC.element_to_be_clickable((By.XPATH,Enter_Position)))

    try:
        Second.text
        Error.text
    except AttributeError:
        Address_Output = Driver.find_element(By.XPATH, Address_Position)
        Address = Address_Output.text
        Description_Output = Driver.find_element(By.XPATH, Description_Position)
        Description = Description_Output.text
        Descriptions.append(Description)
        if Description[-9:-1] == 'Addresse':
            return None
        return Address, Description



""" Parse String Format for associated country"""
def Parse_String_CAN(Address,Description):
    Split_Description = Description.split(',')
    City = Split_Description[0]
    Province = Split_Description[1]
    Postal_Code = Split_Description[2]
    Success = [Address,'','','','',City,Province,Postal_Code,'Canada','Yes']
    return Success




def Parse_String_Non_CAN(Address, Description):
        if Country == 'United States' or 'USA':
            Split_Description = Description.split(' ')
            City = Split_Description[0]
            Province = Split_Description[1]
            Postal_Code = str(Split_Description[2])
            Success = [Address,'','','','',City,Province,Postal_Code,Country,'Yes']

        if Country == 'Australia':
            Split_Description = Address.split(',')
            Address = Split_Description[0]
            Split_Description = Split_Description[1].split(' ')
            City = Split_Description[0]
            Province = Split_Description[1]
            Postal_Code = Split_Description[2]
            Success =  [Address,'','','','',City,Province,Postal_Code,Country,'Yes']

        if Country in ['United Kingdom','England','UK','Belgium']:
            Split_Description = Description.split(',')
            City = Split_Description[0]
            Postal_Code = Split_Description[1]
            Success = [Address,'','','','',City,'',Postal_Code,Country,'Yes']
        else:
            Split_Description = Description.split(' ')
            City = Split_Description[0]
            Postal_Code = Split_Description[1]
            Success =  [Address,'','','','',City,'',Postal_Code,Country,'Yes']

        return Success
""" function for what to do if info was NOT found on Canada Post website"""

def Unsuccessful(Row):
    row_unsuccessful = Row
    row_unsuccessful['Successful'] =  'No'
    return row_unsuccessful

"""Highlighting Function"""

def rowStyle(row):
    if row['Successful'] == 'Yes':
        print('Highlight Green')
        return ['background-color: green'] * len(row)
    return [''] * len(row)


""" function for checking if data fits properly into array"""


def Index_Input_CAN(Results):
        if Results != None:
            try:
                return Parse_String_CAN(Results[0],Results[1])
            except IndexError:
                print('There was an Index Error')
                return Unsuccessful(Row)
        else:
            return Unsuccessful(Row)

def Index_Input_Non_CAN(Results):
        if Results != None:
            try:
                return Parse_String_Non_CAN(Results[0],Results[1])
            except IndexError:
                print('There was an Index Error')
                return Unsuccessful(Row)
        else:
            return Unsuccessful(Row)



"""Loop across all the clientel and check address"""
Header = Excel_Sheet.columns.ravel()
Header = np.append(Header, 'Successful')
Fixed_Sheet_Data_CAN = pd.DataFrame([],columns = Header)
Fixed_Sheet_Data_Non_CAN = pd.DataFrame([],columns = Header)
Excel_Sheet_CAN = Excel_Sheet.loc[Excel_Sheet['PREFERRED ADDRESS LINE COUNTRY'] == 'Canada']
Excel_Sheet_Non_CAN = Excel_Sheet.loc[Excel_Sheet['PREFERRED ADDRESS LINE COUNTRY'] != 'Canada']



for Index, Row in Excel_Sheet_CAN.iterrows():
    print(Index)
    Search = Search_Input(Row)
    Fixed_Sheet_Data_CAN.loc[Index] = Unsuccessful(Row)
    if Search != 'Series([], )':
        Results = Get_Address(Search)
        New_Address = Index_Input_CAN(Results)
        if New_Address[-1]== 'No':
            Search = Row['PREFERRED ADDRESS LINE 1']
            print(Search)
            Results = Get_Address(Search)
            New_Address = Index_Input_CAN(Results)
            if New_Address[-1] == 'No':
                print('Moving on')
            else:
                if New_Address[7] == Row['PREFERRED ADDRESS POSTAL CODE']:
                    print(New_Address)
                    Fixed_Sheet_Data_CAN.loc[Index] = New_Address
                else:
                    print( "Blah")
        else:
            Fixed_Sheet_Data_CAN.loc[Index] = New_Address


print('Starting on Non-Canada')
for Index, Row in Excel_Sheet_Non_CAN.iterrows():
    print(Index)
    Search = Search_Input(Row)
    Fixed_Sheet_Data_Non_CAN.loc[Index] = Unsuccessful(Row)
    if Search != 'Series([], )':
        Country = Row['PREFERRED ADDRESS LINE COUNTRY']
        Country_Search(Country)
        Results = Get_Address(Search)
        New_Address = Index_Input_Non_CAN(Results)
        if New_Address[-1] == 'No':
            Search = Row['PREFERRED ADDRESS LINE 1']
            Results = Get_Address(Search)
            New_Address = Index_Input_Non_CAN(Results)

        else:
            Fixed_Sheet_Data_Non_CAN.loc[Index] = New_Address



# if no postal code and no address , move on


Driver.quit()
et = time.time()
pte = time.process_time()
elapsed_time = st - et
elapsed_process_time = pts - pte
print("Time Passed:" ,elapsed_time/60, "Minutes")
print("Process time:", elapsed_process_time, "Seconds")
Fixed_Sheet_Data = pd.concat([Fixed_Sheet_Data_CAN,Fixed_Sheet_Data_Non_CAN], axis = 0)
Fixed_Sheet_Data.sort_index(inplace = True)
Full_Sheet = pd.concat([Fixed_Sheet_Data,Rest_Excel],axis = 1)
Full_Sheet.style.apply(rowStyle,axis = 1).to_excel('Cleaned_Excel_Sheet_DWDC_1.xlsx', index = False)
