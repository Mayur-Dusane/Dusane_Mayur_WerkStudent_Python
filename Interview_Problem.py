
#importing required libraries 

from tabula import read_pdf
import pandas as pd
import numpy as np
import os
import jpype

#extracting the sample files path
try:
    exe_file_path = os.path.dirname(os.path.abspath('Interview_Problem.exe')) #to find current exe file path
    invoice1_path = os.path.join(exe_file_path, "sample_invoice_1.pdf") #to find invoice file 1 path
    invoice2_path = os.path.join(exe_file_path, "sample_invoice_2.pdf") #to find invoice file 2 path
except Exception as e:
    print(f"Error determining file paths: {e}")
    

#extracting tables from first invoice

try:
    Tables_1 = read_pdf(invoice1_path, pages = "all", multiple_tables=True, lattice=True) #Reading tables from invoice 1
    Tables_2 = read_pdf(invoice2_path, pages = "all", area = [55, 55, 1000, 1000], multiple_tables=True, lattice=True) #Reading tables from invoice 2

except Exception as e:
    print(f"Error extracting tables: {e}")

Tables_1_filter = pd.DataFrame()

#finding rows from tables_1 which has 1st required value
try:
    for i in Tables_1:
        for j in i.columns:
            if (i[j].replace({pd.NA: None}).astype(str).str.contains('Gross Amount incl. VAT', case=False).any().any()):
                Tables_1_filter = i[i[j] == 'Gross Amount incl. VAT']
            else:
                pass

    Sample_value_1 = Tables_1_filter.dropna(axis=1).values
    Sample_value_1 = Sample_value_1[0][1]
    Sample_value_1

    import re

    match = re.search(r'\d+', Sample_value_1) #Extracting numerical value from sample value 1.

    if match:
        Value_1 = int(match.group())
        print(f"First value from Sample invoice 1 is : {Value_1}")
        
except Exception as e:
    print(f"Error determining value 1: {e}")
       
#finding rows from tables_2 which has 2nd required value
try:
    Tables_2_filter = pd.DataFrame()
    for i in Tables_2:
        for j in i.columns:
            if (i[j].replace({pd.NA: None}).astype(str).str.contains('Total', case=False).any().any()):
                Tables_2_filter = i[i[j] == 'Total']
            else:
                pass

    Sample_value_2 = Tables_2_filter.dropna(axis=1).values
    Sample_value_2 = Sample_value_2[0][1]
    Sample_value_2

    match = re.search(r'\d+', Sample_value_2) #Extracting numerical value from sample value 2.

    if match:
        Value_2 = int(match.group())
        print(f"First value from Sample invoice 2 is : {Value_2}")

except Exception as e:
    print(f"Error determining value 2: {e}")
    

#Scrape the date from sample invoice 1 pdf


try:
    Tables_1_date = pd.DataFrame
    for i in Tables_1:
        for j in i.columns:
            if j == 'Date':
                Tables_1_date = i[j].replace({pd.NA: None})
            else:
                pass

    sample_date1 = Tables_1_date.values

except Exception as e:
    print(f"Error determining date 1. {e}")
    
try:
    from datetime import datetime
    import locale
    locale.setlocale(locale.LC_TIME, "deu_deu")
    Date_1 = datetime.strptime(sample_date1[0], "%d. %B %Y")
    Date_1 = Date_1.strftime("%Y-%m-%d")
    Date_1

except Exception as e:
    print(e)
    
#Scrape the date from sample invoice 2 pdf

try:
    Tables_2_1 = read_pdf(invoice2_path, pages = "all", area = [55, 55, 1000, 1000], multiple_tables=True, stream=True)
    Filter_date = Tables_2_1[0].iloc[:,0].str.split(":", expand=True).replace({pd.NA: None})
    Sample_date2 = Filter_date[Filter_date[0] == "Invoice date"][1].values
    Date_2 = datetime.strptime(Sample_date2[0], " %b %d, %Y")
    Date_2 = Date_2.strftime("%Y-%m-%d")
    Date_2

except Exception as e:
    print(f"Error determining date 2: {e}")
    
#Creating Excel sheets


try:
    sheet_1 = pd.DataFrame({"File Name": ['sample_invoice_1', 'sample_invoice_2'], "Date" : [Date_1, Date_2], "Value" : [Value_1, Value_2] }) #Creating Excel sheet 1
    sheet_1.to_excel('sheet_1.xlsx')
    
    sheet_1.pivot_table( values='Value', index= 'Date', 
                        columns='File Name', aggfunc='sum', 
                        fill_value=0).to_excel('sheet_2.xlsx') #Creating Excel sheet 2

except Exception as e:
    print(f"Error creating excel sheets: {e}")

#Adding all data into csv file

try:
    blank_df = pd.DataFrame()
    blank_df.to_csv('Invoice_data.csv', mode='w', index=False, sep=';')
    for i in range(len(Tables_1)):
        Tables_1[i]['Tables_invoice_1'] = f'Table_{i + 1}'
        Tables_1[i].to_csv('Invoice_data.csv', mode='a', index=False, sep=';')
    for i in range(len(Tables_2)):
        Tables_2[i]['Tables_invoice_2'] = f'Table_{i + 1}'
        Tables_2[i].to_csv('Invoice_data.csv', mode='a', index=False, sep=';')

except Exception as e:
    print(f"Error creating csv file: {e}")
    
print("All required Excel sheets and CSV are created")


