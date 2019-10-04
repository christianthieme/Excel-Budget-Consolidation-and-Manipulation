# -*- coding: utf-8 -*-
"""
Created on Mon Jun 10 13:42:29 2019

@author: cthieme
"""

import win32com.client


##############all_templates is a list that needs to contain the full_path to the files you want to update##########################
all_templates = [r"\\collab.micron.com@SSL\DavWWWRoot\corp\Legal\AcctFinRpts\Shared Documents\Budget and Expenses\Python Test\107300.xlsx",
                 r"\\collab.micron.com@SSL\DavWWWRoot\corp\Legal\AcctFinRpts\Shared Documents\Budget and Expenses\Python Test\107302.xlsx",
                 r"\\collab.micron.com@SSL\DavWWWRoot\corp\Legal\AcctFinRpts\Shared Documents\Budget and Expenses\Python Test\107306.xlsx"] 

##################### write the name of the Macro you want to run ##################################################################

macro_name = 'workbook_rollover'

####################################################################################################################################

excel = win32com.client.Dispatch("Excel.Application")
#file path to your personal macro workbook
macro = excel.Workbooks.Open(r"C:\Users\cthieme\AppData\Roaming\Microsoft\Excel\XLSTART\PERSONAL.XLSB")

#The below code loops through each file in the list above and calls the macro that you have recorded and named. 
for i, template in enumerate(all_templates):
         
    wb = excel.Workbooks.Open(template)
    wb.Worksheets[1]
    try:
        print("File: " + str(i + 1) + " is in process................................")
        wb.Application.Run("PERSONAL.XLSB!" + macro_name)
        wb.Close(True)
        print("File " + str(i + 1) +": " + template + " completed successfully")
    except:
        print("Something went wrong with file " + str(i + 1) +": " + template + " in the 'all_templates' list of files. Please fix the issue and try again. To see the error, run script again and click the 'Debug' option in the Microsoft Visual Basic window that appears. This will take you to the code behind your macro that you can then debug.")
        wb.Close(True)
        macro.Close(True)
        break
try:
    macro.Close(True)     
except: 
    None
excel.Quit()
del excel   
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    


