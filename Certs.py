from operator import index
import numpy as np3
import pandas as pd
import dataclasses
import fitz
from PyPDF2 import PdfFileWriter, PdfFileReader
from pandas.tseries.offsets import DateOffset
import pyodbc 
import pypyodbc
import win32com.client as win32
import os
import os.path
import concurrent.futures
from multiprocessing import freeze_support
from pathlib import Path

Work_Orders = []     #blank list of work Orders

n = int(input('Enter Number of Certs: '))  # request count of certs from user.

for i in range(0, n):
    WO = str(input('WO Number - '))
    Work_Orders.append(WO) # these blocks collect the certs from the user based on number of certs told

WO_Requested = pd.DataFrame(Work_Orders,columns= ['WoNumber'])

conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=STAUBCAD\SIGMANEST;'
                      'Database=SNDBase22;'
                      'Trusted_Connection=yes;')

SERVER_NAME = 'STAUBCAD\SIGMANEST'
DATABASE_NAME = 'SNDBase22'

sql_query = "SELECT WoNumber, SheetName, PartFileName  FROM [dbo].[PartArchive] "

part = pd.read_sql(sql_query, conn) # request SQL  Table from STAUBCAD

sql_query1 = "SELECT SheetName, HeatNumber, PrimeCode  FROM [dbo].[StockArchive] "

stock = pd.read_sql(sql_query1, conn) # request SQL  Table from STAUBCAD


#correct any length input to proper 7 digit + Lot number  

WO_Length = {8:'000',9:'00',10:'0',11:''} # sets thickness definitions for later pulling from to take 1 and two decimal point thicknesses up to 3 decimal points

WO_Requested['WO_Length'] = WO_Requested['WoNumber'].str.len() # collects the number of characters in the material thickness field 

WO_Requested['WoNumber']   = WO_Requested['WO_Length'].map(WO_Length) + WO_Requested['WoNumber']   # corrects the material thickness to 3 decimal points to match the PO Reciepts page


part = part[part['WoNumber'].isin(WO_Requested['WoNumber'])]          # Removes all un requested Work Orders from the parts list.

part_shortened = part.drop_duplicates(subset= ['WoNumber','SheetName'])

stock = stock[stock['SheetName'].isin(part['SheetName'])]          # removes all Sheets from the stock list that aernt required for the WO Numbers Requested.
stock_shortened= stock.drop_duplicates(subset= ['SheetName'])
#part['Company'] = part['PartFileName'].str[18:"\\"]

merged_inner = pd.merge(left=part_shortened, right=stock_shortened,how='left', left_on='SheetName', right_on='SheetName') # merges the two data frames of the database and the PO Recietps spreadsheet to matching PO_MTL fields.
merged_inner['PrimeCode'] = merged_inner['PrimeCode'].astype(str) + ".pdf"
merged_inner['PartFileName'] = merged_inner['PartFileName'].apply( lambda x : r'\\' + x.split('\\')[-2]  )
merged_inner['PartFileName'] = merged_inner['PartFileName'].str[2:]

filelist = []

for root, dirs, files in os.walk("G:\Materials Received"):
	for file in files:
        #append the file name to the list
		filelist.append(os.path.join(root,file))                 # pulls all files from the Materials Recieved folder on the G drive along with path


Materials_Recieved = pd.DataFrame( { 'filename':filelist })
Materials_Recieved['Docname'] = Materials_Recieved['filename'].apply( lambda x :  x.split('\\')[-1]  )   # extracts file name from the file list and gives it its own column to make it easier to search

df = pd.merge(left=merged_inner, right=Materials_Recieved,how='left', left_on='PrimeCode', right_on='Docname')
df['Job'] = df['PartFileName'] +" "+ df['WoNumber']

df.to_excel(r'C:\Users\GGehring\Documents\PO.xlsx',index = False)



for row in df:
    doc = df['filename']

    label = {}
    page_number = 0
    
   
    for page in doc:
    
        page_name = page.get_label()
        label[page_number] =  page_name 
        page_number += 1

print (label)


#### This marks the end of the cert modifications and the beginning of the email creation

Job_List_Unfiltered = list(df['Job'])
Job_List = []
for i in Job_List_Unfiltered:
    if i not in Job_List:
        Job_List.append(i)

Job = '<br>'.join(Job_List)                   # This section sets the text for the email based on the company name and work order number pulled from the Job Dataframe




def Emailer(text, subject, recipient):


    outlook = win32.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)
    mail.To = recipient
    mail.Subject = subject
    mail.HtmlBody = text
    ###

   # attachment1 = os.getcwd() +"\\file.ini"

    #mail.Attachments.Add(attachment1)

    ###
    mail.Display(False)

MailSubject= "Certs"
MailInput="""
<p>
Please find attached the certs for: </p>"""+ Job + """

                                         <br><br><br>

Thanks,                                  <br>
George Gehring
                                         <br>
                                         <br>
Staub Manufacturing Solutions            <br>
2501 Thunderhawk Ct.                     <br>
Dayton, OH 45414                         <br>
e-mail: ggehring@staubmfg.com            <br>
Website: www.staubmfg.com                <br>
Phone: 937-890-4486                      <br>
Fax: 937-890-4487                        <br>
We are an ISO 9001:2015 certified company.
</p>
"""
MailAdress=""

Emailer(MailInput, MailSubject, MailAdress ) #that open a new outlook mail even outlook closed.













