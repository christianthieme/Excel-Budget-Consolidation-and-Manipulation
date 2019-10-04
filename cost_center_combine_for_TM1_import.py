# -*- coding: utf-8 -*-
"""
Created on Mon Jun 17 11:17:04 2019

@author: cthieme
"""
######################################### PURPOSE OF THIS CODE ##################################################################

# This Code is written to loop through all of the cost center templates and 1. Aggregate the data into a loadable TM1 format and 
# 2. To put in a format that can be easily digested by Power BI. The end result of this script is 2 excel files. One for TM1 and 
# one for Power BI

#################################################################################################################################

#Necessary imports
import numpy as np
import pandas as pd
from datetime import datetime

#all_template is a list that contains the FULL-FILE PATH to all the templates you want to loop through to create your ending aggregated files
all_templates = [r'C:FilePath\Templates\107300.xlsx',
                r'C:\FilePath\Templates\107302.xlsx', 
                r'C:\FilePath\Templates\107306.xlsx']

#Function to GROUPBY GL & Description Column
def grouped_by(template_loc):
    cost_center = pd.read_excel(template_loc, sheet_name = "Budget", index_col=0)
    cost_center.reset_index(inplace = True)
    new_cols = [str(x) for x in cost_center.columns]
    cost_center.columns = new_cols
    cost_center.replace(r'^\s*$', np.nan, regex=True, inplace = True)
    #if you need are adding or removing column names, change the index below two lines of code
    date_cols = cost_center.columns[8:]
    text_cols = cost_center.columns[1:8]
    cost_center.update(cost_center[date_cols].fillna(0))
    cost_center.update(cost_center[text_cols].fillna(''))
    #the line below (.groupby) is what we are grouping by. 
    grouped_df = cost_center.groupby([ 'Cost Center','GL & Description'], as_index = False)[date_cols].sum()
    #Below 2 lines splits the GL Code from the description for the TM1 Import file
    GL_code = grouped_df['GL & Description'].str.split(' ', n=1, expand = True)
    grouped_df.insert(loc=1, column='GL Code', value=GL_code[0])
    #drop the old GL & Description column since we only need GL Code 
    grouped_df.drop('GL & Description', axis =1, inplace = True)
    
    text_list = ['Cost Center', 'GL Des']
    date_list = []    
    #This will format the dates in the same format as required for TM1 upload
    for date in date_cols:
        new_date = datetime.strptime(date,'%Y-%m-%d %H:%M:%S')
        right_date = new_date.strftime('%b-%y')
        str_date = str(right_date)
        date_list.append(str_date)
        
    column_name_list = text_list + date_list
    
    grouped_df.columns = column_name_list
    
    return grouped_df

#This function creates the output for the data to be ingested by Power BI
def detailed_data(template_loc):
    detail_cost_center = pd.read_excel(template_loc, sheet_name = "Budget", index_col=0)
    detail_cost_center.reset_index(inplace = True)
    new_cols = [str(x) for x in detail_cost_center.columns]
    detail_cost_center.columns = new_cols
    detail_cost_center.replace(r'^\s*$', np.nan, regex=True, inplace = True)
    #The line below drops rows where everything is NULL
    detail_cost_center.dropna(inplace = True, how = 'all', axis = 0)
     #if you need are adding or removing column names, change index in the below two lines of code
    date_cols = detail_cost_center.columns[8:]
    text_cols = detail_cost_center.columns[1:8]
    #fills date_column Nulls with 0's
    detail_cost_center.update(detail_cost_center[date_cols].fillna(0))
    #fills text_column Nulls with spaces
    detail_cost_center.update(detail_cost_center[text_cols].fillna(''))
    detail_cost_center.set_index(['Cost Center', 'GL & Description', 'Cost Element - Description', 'GL Helper', 'Vendor', 'PO', 'Team/ Function', 
             'Initiative/ Project'], inplace= True)
    #Pivots data to necessary format
    stacked = detail_cost_center.stack()
    frame = pd.DataFrame(stacked)
    frame.reset_index(inplace = True)
    frame.columns = ['Cost Center', 'GL & Description', 'Cost Element - Description', 'GL Helper', 'Vendor', 'PO', 'Team/ Function', 
                'Initiative/ Project', 'Date', 'Value']
    return frame


detail_data_to_append = []    
grouped_data_to_append = []

# For loop to loop through my functions above
for template in all_templates:
    grouped_cost_center = grouped_by(template)
    detail_cost_center =detailed_data(template)
    #print(next_cost_center)
    grouped_data_to_append.append(grouped_cost_center)
    detail_data_to_append.append(detail_cost_center)
    
grouped_all_cost_centers = pd.concat(grouped_data_to_append)
detail_all_cost_centers = pd.concat(detail_data_to_append)

#The below filepaths are where you want the ending files to go for the grouped data and the detailed data, respectively
grouped_all_cost_centers.to_excel(r"C:FilePath\TM1 Publish Data\grouped_test.xlsx", index = False)
detail_all_cost_centers.to_excel(r"C:FilePath\detail_test.xlsx", index = False)
    
    
