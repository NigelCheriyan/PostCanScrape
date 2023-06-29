# -*- coding: utf-8 -*-
"""
Created on Thu Jun 29 11:48:53 2023
Web Scraper to double check all addresses with Canada Post

For Dying with Dignity
@author: nigel
"""
import urllib, xml.dom.minidom





""" Pull CSV file"""

file_name  = "Nigel Cheriyan Address Cleanup.xlsx"

dfs = pd.read_excel(file_name, sheet_name=None) # pull data from file 


""" Pull up Website ~~~ CODE FROM CAN POST API ADDRESS COMPLETE"""

def AddressComplete_Interactive_Find_v2_10(Key, SearchTerm, LastId, SearchFor, Country, LanguagePreference, MaxSuggestions, MaxResults, Origin, Bias, Filters, GeoFence):

      #Build the url
      requestUrl = "http://ws1.postescanada-canadapost.ca/AddressComplete/Interactive/Find/v2.10/xmla.ws?"
      requestUrl += "&" +  urllib.urlencode({"Key":Key})
      requestUrl += "&" +  urllib.urlencode({"SearchTerm":SearchTerm})
      requestUrl += "&" +  urllib.urlencode({"LastId":LastId})
      requestUrl += "&" +  urllib.urlencode({"SearchFor":SearchFor})
      requestUrl += "&" +  urllib.urlencode({"Country":Country})
      requestUrl += "&" +  urllib.urlencode({"LanguagePreference":LanguagePreference})
      requestUrl += "&" +  urllib.urlencode({"MaxSuggestions":MaxSuggestions})
      requestUrl += "&" +  urllib.urlencode({"MaxResults":MaxResults})
      requestUrl += "&" +  urllib.urlencode({"Origin":Origin})
      requestUrl += "&" +  urllib.urlencode({"Bias":Bias})
      requestUrl += "&" +  urllib.urlencode({"Filters":Filters})
      requestUrl += "&" +  urllib.urlencode({"GeoFence":GeoFence})

      #Get the data
      dataDoc = xml.dom.minidom.parseString(urllib.urlopen(requestUrl).read())

      #Get references to the schema and data
      schemaNodes = dataDoc.getElementsByTagName("Column")
      dataNotes = dataDoc.getElementsByTagName("Row")

      #Check for an error
      if len(schemaNodes) == 4 and schemaNodes[0].attributes["Name"].value == "Error":
         raise Exception, dataNotes[0].attributes["Description"].value

      #Work though the items in the response
      results = []
      for dataNode in dataNotes:
         rowData = dict()
         for schemaNode in schemaNodes:
              key = schemaNode.attributes["Name"].value
              value = dataNode.attributes[key].value
              rowData[key] = value
         results.append(rowData)

       return results

      #FYI: The output is an array of key value pairs, the keys being:
      #Id
      #Text
      #Highlight
      #Cursor
      #Description
      #Next


"""function to parse string data""" 





"""Loop across all the clientel and check address"""







# Pull postal code from csv



# if no postal code, pull address







# if no postal code and no address , move on 



