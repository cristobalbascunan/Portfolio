#!/usr/bin/env python
# coding: utf-8

# # Automation of output of client files

print('Importing packages and defining function...')
import os
import pandas as pd
import numpy as np
import errno
import shutil
import xlsxwriter
import time
import pyodbc
from datetime import date, timedelta, datetime
import xlwings as xw
import sys


def create_connection():
    'This function stablishes a connexion with the database.'
    # Information necessary to create the engine.
    server = 'server' 
    database = 'database' 
    username = 'username' 
    password = 'password'
    # Parse different elements to connect to the database.
    params = urllib.parse.quote_plus("""DRIVER={SQL Server Native Client 11.0};
                                        SERVER=server;
                                        DATABASE=database;
                                        UID=username;
                                        PWD=password""")
    
# Definition of a new function to convert the day (today) to the format used on the output file and stick it to the whole name plus its extension.

def excel_file_name_output(name, date_format, date_first, file_format, reportdate):
    """
    Return an string with the output name depending on the format of the parameters.
    This will always refer as if the file has been done with the 'today' date.
    :param name: Name of the file without date.
    :param date_format: Date format style. i.e: dd-mm-yy.
    :param date_first: Condition to see if the date goes first. Yes or No.
    :param file_format: Extention of the file.
    :param reportdate: In which reportdate are we in order to not use the 'today' but the publication in which we are.
    :returns: output name of the file
    """
    ##Using today as the date.
    date_today = date.today()
    ##Depending on the format of the date written in the "Format" column, this will treat each case.
    if date_format == "yymmdd":
        date_converted = date_today.strftime("%y%m%d")
    elif date_format == "yyyy-mm-dd":
        date_converted = date_today.strftime("%Y-%m-%d")
    elif date_format == "ddmmyy":
        date_converted = date_today.strftime("%d%m%y")
    elif date_format == "yyyymmdd":
        date_converted = date_today.strftime("%Y%m%d")
    elif date_format == "yyyy.mm.dd":
        date_converted = date_today.strftime("%Y.%m.%d")
    elif date_format == "Month yyyy":
        date_converted = date_today.strftime("%B %Y")
    elif date_format == "Current Report Date":
        date_converted = reportdate.strftime("%B %Y")
    elif date_format == "yyyy":
        date_converted = reportdate.strftime("%Y")
    else:
        date_converted = ''
        
    ##Conditional if the date should be the first word in the name or the last    
    if date_first =="Yes":
        file_name = date_converted + name + file_format
    else:
        file_name = name + date_converted + file_format
        
    return file_name


def change_date_format(df, column, date_format):

    ##Depending on the format of the date written in the "Format" column, this will treat each case.
    if date_format == "mmddyy":
        df[column] = pd.to_datetime(df[column], format="%m/%d/%Y")
    elif date_format == "ddmmyy":
        df[column] = pd.to_datetime(df[column], format='%d/%m/%Y')
    elif date_format == "mm/dd/yy":
        df[column] = pd.to_datetime(df[column], format='%Y/%m/%d')
    return df[column]


def change_date_format_shifted_columns(df, columns_to_shift, date_format):
    ##Depending on the format of the date written in the "Format" column, this will treat each case.

    if date_format == "mmddyy":
        df[columns_to_shift] = pd.to_datetime(df[columns_to_shift], format="%m/%d/%Y")
    elif date_format == "ddmmyy":
        df[columns_to_shift] = pd.to_datetime(df[columns_to_shift], format='%d/%m/%Y')

    return df[columns_to_shift]


def save_csv_with_date_format(df, date_format):
    ##Depending on the format of the date written in the "Format" column, this will treat each case.

    if date_format == "mm/dd/yy":
        file.to_csv(output_name, index=False, encoding='utf-8-sig', date_format="%m/%d/%y")    
    else:
        file.to_csv(output_name, index=False, encoding='utf-8-sig')    


def run_excel_macro(file_path):
    """
    Execute an Excel macro
    :param file_path: path to the Excel file holding the macro
    :return: None
    """
    try:
        xl_app = xw.App(visible=False, add_book=False)
        wb = xl_app.books.open(file_path)

        run_macro = wb.app.macro('Module1.Macro' + x['Client'])
        run_macro()

        wb.save()
        wb.close()

        xl_app.quit()

    except Exception as ex:
        template = "An exception of type {0} occurred. Arguments:\n{1!r}"
        message = template.format(type(ex).__name__, ex.args)
        print(message)

print('Libraries imported.')

# ## 0. Preparing data
os.chdir(r'C:\...\...\...\....\Destination folder')
print(os.getcwd()) ## This will show us where are we (we need to be in the folder where the Excel template file is)


# Connection with SQL so we can run the queries.
cnxn = create_connection()
##Reading the template made in Excel
client_excel =pd.read_excel('... .xlsx')

# ### 1. Queries with macros
# We will only take the clients that are `OK` and are included in this week.
excel_file = client_excel.loc[(client_excel["..."] == week_number) & (client_excel["..."] == 'OK')].reset_index(drop=True)

##IF THERE IS ANY PROBLEM GENERATING FILES, UNCOMMENT THIS LINE AND CHANGE THE NUMBER THAT WILL SET THE START. 
## FILES ALREADY PRINTED + 1
#excel_file = excel_file[4:]

# This loop will start the creation of output
##Loop to do the inside for each client 'OK'
for i in range(len(excel_file)):
    x = excel_file.iloc[i]
    ## Using the function defined before, creating the name for the client file.
    output_name = excel_file_name_output(x["..."], x["..."], x["..."], x["..."], reportdate)
    print("Creating the file " + output_name + "...")
    query_path = x["Query"]
    # Open and read the file as a single buffer
    fd = open(query_path, 'r')
    sqlFile = fd.read()
    fd.close()
    query = 'SET NOCOUNT ON; '+ sqlFile
    file = pd.read_sql_query(query, cnxn)
    #Adding date format to the columns (for those that are .xlsx) 
    if x["File format"] == ".xlsx":
        if str(x["..."]) != 'nan':
            date_columns = list(x["..."].split(", "))
            for column in date_columns:
                file[column] = change_date_format(file, column, x["..."]) 

        if x["Months shifted"] == 'Yes':            
            date_columns_shifted = list(x["..."].split(", "))
            for shifts in date_columns_shifted:
                file[shifts] = change_date_format_shifted_columns(file, shifts, x["..."])
    #Fill empty spaces that could appear with the SQL 'NULL' value
    file = file.fillna("NULL")
    ##First of all is seleted the Path for each file. This will be changing each iteration
    os.chdir(x["Path"])  
    ##Exception for those clients that need the file inside a new folder (named as i.e. 2106).Creation of a folder if it doesn't exist
    if x["In folder?"] == 'Yes':
        f = reportdate.strftime("%y%m")    
        try:
            os.mkdir(f)
        except OSError as e:
            if e.errno != errno.EEXIST:
                raise
    ## This will change the eventual path to inside the folder
        path_inside = str(x["Path"])+str(f)+'\\' 
        print(path_inside)      
    ## Depending on the output, this will create different kind of files.
    ### CSV files
    if x["File format"] == ".csv":
        save_csv_with_date_format(output_name, x["..."])

    ### XLSX files
    elif x["File format"] == ".xlsx":
        dirpath = r'C:\...\...\Macros\\'
        os.chdir(dirpath)
        name_while_open = 'Macro' + x["..."]+'_tmp'+'.xlsx'
        with pd.ExcelWriter(name_while_open, engine='xlsxwriter') as writer:
            # Write the data from the file variable to the Excel file in a sheet specified by the "Tab name"
            file.to_excel(writer, sheet_name=x["Tab name"], index=False)
            # Get the workbook object to call it in next lines
            workbook = writer.book
            filename_macro = 'Macro' + x["..."]+ '.xlsm'
            # Assign the filename with extension .xlsm
            workbook.filename = filename_macro
            # add VBA project to the .xlsm file that already contains the data 
            file_vba = 'vbaProject'+x["..."]+'.bin'
            workbook.add_vba_project(file_vba)
            #save the excel file
            writer.save()
            # Run client custom excel macro on the saved file to apply the 
            run_excel_macro(filename_macro)
            
        old_file_name = 'Macro' + x["..."]+ '.xlsx'
        #Moving file to its destination path
        os.rename(old_file_name, output_name)            
        # Check if the file should be moved to a folder or a specific path
        if x["..."] == 'Yes':
            # If the file should be moved to a folder, use the shutil.move() function to move it to the specified path_inside folder
            try:
                shutil.move(output_name, path_inside)
            except:
                # If the file already exists in the destination folder, print a message and overwrite the existing file
                print("It already exist in the destintation folder!! Overwriting existing file")
                os.remove(path_inside+str(output_name))
                shutil.move(output_name, path_inside)
        else:
            # If the file should be moved to a specific path, use the shutil.move() function to move it to the specified path
            try:
                shutil.move(output_name, x["Path"])
            except:
                # If the file already exists in the destination folder, print a message and overwrite the existing file
                print("It already exist in the destintation folder!! Overwriting existing file")
                os.remove(x["Path"]+str(output_name))
                shutil.move(output_name, x["Path"])
        # Check if the filename_macro and name_while_open files exist
        try:
            if os.path.exists(filename_macro):
                # If the filename_macro file exists, remove it
                os.remove(filename_macro)
            if os.path.exists(name_while_open):
                # If the name_while_open file exists, remove it
                os.remove(name_while_open)
            else:
                # If the files do not exist, print a message
                print("The file does not exist")
        except:
            # In case of an exception, do nothing
            pass
# TXT files
    elif x["File format"] == ".txt":
        file.to_csv(output_name, index=None, sep='\t')
    print("done!")
print('All files created!')


# ### 2. Masterfiles with macros
excel_file = client_excel.loc[(client_excel["Week"] == week_number) & (client_excel["..."] == 'UPDATE')].reset_index(drop=True)
##Loop to do the inside for each client 'UPDATE'
for i in range(len(excel_file)):
    x = excel_file.iloc[i]
    ## Using the function defined before, creating the name for the client file.
    output_name = excel_file_name_output(x["..."], x["..."], x["..."], x["..."], reportdate)
    print("Creating the file " + output_name + "...")
    dirpath = r'C:\...\...\Macros'
    os.chdir(dirpath)    
    if x["..."] == 'Yes':
        f = reportdate.strftime("%Y")    
        try:
            os.mkdir(f)
        except OSError as e:
            if e.errno != errno.EEXIST:
                raise
    ## This will change the eventual path to inside the folder
        path_inside = str(x["Path"])+str(f)+'\\' 
        print(path_inside)
    #ATTACH MACRO AND CREATE .XLSM FILE# 
    filename_macro = 'Macro' + x["..."]+ '.xlsm'
    run_excel_macro(filename_macro)
    filename_excel = 'Macro' + x["..."]+ '.xlsx'    
    os.rename(filename_excel, output_name)
    if x["..."] == 'Yes':
        try:
            os.remove(path_inside+str(output_name))
            print("It already exist in the destintation folder!! Overwriting existing file")
        except:
            pass
        shutil.move(output_name, path_inside)
    else:
        try:
            os.remove(x["Path"]+str(output_name))
            print("It already exist in the destintation folder!! Overwriting existing file")
        except:
            pass        
        shutil.move(output_name, x["Path"])    
    print("done!")
print('All Master-depending files created!')

# ### 3. Masterfiles with auto-output
excel_file = client_excel.loc[(client_excel["..."] == week_number) & (client_excel["..."] == 'DIRECTLY')].reset_index(drop=True)
##Loop to do the inside for each client 'DIRECTLY'

for i in range(len(excel_file)):
    x = excel_file.iloc[i]
    ## Using the function defined before, creating the name for the client file.
    print("Creating the file " + x["Client"] + "...")
    dirpath = r'C:\...\...\Macros'
    os.chdir(dirpath)    
    #ATTACH MACRO AND CREATE .XLSM FILE# 
    filename_macro = 'Macro' + x["Client"]+ '.xlsm'
    run_excel_macro(filename_macro)    
    print("done!")
print('All files created!')