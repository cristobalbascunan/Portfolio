#!/usr/bin/env python
# coding: utf-8

# # CONVERT DATABANK FORECASTS LATAM

# #### Libraries used

import sys
import os
import pandas as pd
import numpy as np
from pandas.tseries.offsets import MonthEnd
import pyodbc
import fnmatch
import sys
import errno
import re
from datetime import datetime

#Region
Region = 'Latam'
path = r"C:\...\Regions\\" + Region

os.chdir(path)
os.getcwd()


# ##### SQL Credentials
#Conection with the database
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

########################################################
cnxn = create_connection()

#ONLY APPLICABLE FOR DATABANK
sourcedata = 185
########################################################

#######################################
path_files = r"C:\...\DATABANK_Python\\"

for filename in os.listdir(path_files):
    match =  re.match("^\d.*"+Region+" Annual.xlsx", filename, re.IGNORECASE)
    if match:
        excel_name_ann = path_files+ filename
    match =  re.match("^\d.*"+Region+" Quarterly.xlsx", filename, re.IGNORECASE)
    if match:
        excel_name_qt = path_files+ filename
        
for filename in os.listdir('.'):
    if fnmatch.fnmatch(filename, 'Indicators*Ann*by*.xlsx'):
        file_indicators_ann = filename
    elif fnmatch.fnmatch(filename, 'Indicators*Qt*by*.xlsx'):
        file_indicators_qt = filename
#######################################

print("Using the following files:", excel_name_ann, excel_name_qt, "(Be sure you are uploading the correct day, otherwise check the download status from the webpage)",sep="\n")


# ## 0. Defined Functions

# #### Filenames and function

# #### TRANFORMATIONS FOR YEARLY DATA

def transform_excel_ann(file_excel, file_indicators):
    print("Processing yearly data...")
    excel_file =pd.ExcelFile(file_excel)
    full_df = pd.DataFrame()
    df_ind = pd.read_excel(file_indicators, sheet_name=file_indicators[:-5])
    df_ind = df_ind[["Entity","Entity_DATABANK", "FE_Ticker","DATABANK_Ticker", "Calculations"]]
    for sheet in df_ind["Entity_DATABANK"].unique():
        df = excel_file.parse(sheet_name=sheet)
        row_headers_n = df[df[df.columns[0]] == 'Series'].index[0]
        df = df[row_headers_n:].reset_index(drop=True)
        header_row = 0
        a = df.loc[header_row].astype(str)
        df.columns = [s.replace(".0", "") for s in a]
        df = df.drop(header_row).reset_index(drop=True)
        df["Entity_name"]= sheet
        df = df.drop(columns=["Currency", "Units", "Source", "Definition", "Note", "Series"])
        df.replace("n.a.", np.nan, inplace=True)
        df.replace("–", np.nan, inplace=True)
        # Calculated values for general (TRADEUSD and positive IMPUSD):    
        row_exp_imp = pd.concat([df[df["Code"]=="MEXP"],df[df["Code"]=="MIMP"]]).reset_index(drop=True)
        df = df[df.Code != 'MIMP']
        col_list = list(row_exp_imp.filter(regex='^[0-9]').columns)
        row_exp_imp[col_list] = abs(row_exp_imp[col_list])
        row_exp_imp.loc[0, col_list] = row_exp_imp[col_list].iloc[0] - row_exp_imp[col_list].iloc[1]
        row_exp_imp.loc[0,'Code'] = "TRADEUSD"
        df = pd.concat([df,row_exp_imp[0:2]], ignore_index=True)
        df = pd.melt(df, id_vars=["Code", "Entity_name","Published"])
        df=df.dropna(subset=['value'])
        full_df=pd.concat([full_df,df])
        
    full_df = df_ind.merge(full_df, left_on=['Entity_DATABANK', 'DATABANK_Ticker'], right_on=['Entity_name', 'Code'])
    #Values already calculated (ADD HERE IF THERE IS A NEW CALCULATION) 
        #Calculating /1000
    full_df.loc[full_df["Calculations"] == "/1000", "value"] = full_df.loc[full_df["Calculations"] == "/1000", "value"]/1000
        #Turning negatives to absolute values
    full_df.loc[full_df["Calculations"] == "ABS", "value"] = abs(full_df.loc[full_df["Calculations"] == "ABS", "value"])
        #1 / CURRENCYY
    full_df.loc[full_df["Calculations"] == "1/ENDR", "value"] = 1/(full_df.loc[full_df["Calculations"] == "1/ENDR", "value"])

    full_df = full_df.drop(columns=["Calculations", "Entity_DATABANK", "DATABANK_Ticker", "Code", "Entity_name"])
    full_df = full_df.rename(columns={'variable':"Period", "value":"value", "Published":"Last forecasts"})
    full_df = full_df.dropna(subset=['value', 'FE_Ticker'])
    
    full_df['Period'] = pd.to_datetime(full_df['Period'])
    full_df['Last forecasts'] = pd.to_datetime(full_df['Last forecasts'], format='%d-%m-%Y')
    full_df['Frequency'] = 1
    
    return full_df


# #### TRANFORMATIONS FOR QUARTERLY DATA

def transform_excel_Q(file_excel, file_indicators):
    print("Processing quarterly data...")
    excel_file =pd.ExcelFile(file_excel)
    full_df = pd.DataFrame()
    df_ind = pd.read_excel(file_indicators, sheet_name=file_indicators[:-5])
    df_ind = df_ind[["Entity","Entity_DATABANK", "FE_Ticker","DATABANK_Ticker", "Calculations"]]
    for sheet in df_ind["Entity_DATABANK"].unique():
        df = excel_file.parse(sheet_name=sheet) 
        row_headers_n = df[df[df.columns[0]] == 'Series'].index[0]
        df = df[row_headers_n:].reset_index(drop=True)
        header_row = 0
        a = df.loc[header_row].astype(str)
        df.columns = [s.replace("'", "20") for s in a]
        df = df.drop(header_row).reset_index(drop=True)
        df["Entity_name"]= sheet
        df = df.drop(columns=["Currency", "Units", "Source", "Definition", "Note", "Series"])
        df.replace("n.a.", np.nan, inplace=True)
        df.replace("–", np.nan, inplace=True)
        df = pd.melt(df, id_vars=["Code", "Entity_name", "Published"])
        df = df.rename(columns={'variable':"Period", "value":"value", " Last forecasts":"Last forecasts"," Code":"Code", "Published":"Last forecasts"})
        full_df=pd.concat([full_df,df])

    #Mapping indicators for each country        

    full_df = df_ind.merge(full_df, left_on=['Entity_DATABANK', 'DATABANK_Ticker'], right_on=['Entity_name', 'Code'])

    full_df.loc[full_df["Calculations"] == "1/ENDR", "value"] = 1/(full_df.loc[full_df["Calculations"] == "1/ENDR", "value"])
   
    full_df = full_df.drop(columns=["Calculations", "Entity_DATABANK", "DATABANK_Ticker", "Code", "Entity_name"])
    full_df=full_df.dropna(subset=['value', 'FE_Ticker'])
    full_df['Last forecasts'] = pd.to_datetime(full_df['Last forecasts'], format='%d-%m-%Y')
    full_df['Frequency'] = 2    
    
    new_period = full_df["Period"].str.split(" ", n = 1, expand = True)
    new_period = new_period[1].apply(lambda x: x.replace("'", "")) + new_period[0] 

    full_df['Period'] = new_period
    full_df['Period'] = full_df['Period'].astype(str)
    full_df['Period'] = pd.to_datetime(['-'.join(x.split()[::-1]) for x in full_df['Period']] )
    full_df['Period'] = full_df['Period'] + MonthEnd(3) 

    return full_df     


# ##### To check the output of the file as an Excel file

def create_excel_in_template(table_processed, file_indicators, excel_name):
    df_ind = pd.read_excel(file_indicators, sheet_name=file_indicators[:-5])
    df_ind = df_ind[["Entity","FE_Ticker", "FE_Definition","DATABANK_Ticker", "DATABANK_Long_Definition"]]
    df_excel = table_processed.copy()
    df_excel["Period"] = df_excel["Period"].astype(str) 
    df_excel = df_excel.pivot_table(values=['value'] , index=['Entity','FE_Ticker'], columns='Period', aggfunc='first')
    df_excel.columns = df_excel.columns.get_level_values(1)
    df_excel = df_excel.reset_index()
    df_ind = df_ind.merge(df_excel, how='right', on=['Entity', 'FE_Ticker'])
    name = re.findall("\\\\*DATABANK .*[.]xlsx", excel_name)[0]
    path = r"C:\...\DATABANK\\"
    df_ind.to_excel(path+str(name), index = False)


# ##### Function definition to swap the names for ID numbers:

def change_codes(df, cnxn, sourcedata):
    df_mappingcountries = pd.read_sql_query("SELECT EntityName, NumericCode FROM databasetable",cnxn)
    df_mappingindicator = pd.read_sql_query("SELECT ID, ShortName from databasetable",cnxn)
    df = df.merge(df_mappingcountries, left_on="Entity", right_on="EntityName")
    df = df.merge(df_mappingindicator, left_on="FE_Ticker", right_on="ShortName")
    df = df.rename(columns={'ID':"IndicatorID", "NumericCode":"EntityID"})
    df = df[['IndicatorID','EntityID','Last forecasts','Period','value','Frequency']]
    
    return df


# ##### Function definition to add extra columns:

def get_report_date(df, region):
    query = 'SELECT * FROM databasetable'
    reportdate = pd.read_sql_query(query, cnxn)
    reportdate = reportdate.loc[reportdate['EntityAggregate'] == region]['Reportdate'].iloc[0]
    print("The current report date is:",reportdate)    
    
    return reportdate


# ##### Function definition to add extra columns:

def extra_columns(df):
    df['RowState'] = 'I'
    df['FinalUpload'] = 0
    df['FileType'] = 'Panelist'
    df['TypeData'] = 1
    df['SourceData'] = sourcedata
    df['StatusData'] = 1
    df['ReportDate'] = reportdate
    df['Comment'] = 'NULL'
    df['Private'] = 0
    df['UserName'] = 'Python'
    df = df.rename(columns = {'value': 'Value'})
    new_period = df['Period']
    new_period = new_period.apply(lambda dt: dt.replace(day=1))
    df['Period'] = new_period
    df = df.rename(columns = {'Last forecasts':'UpdateDate'})
    df['Value'] = df['Value'].astype(float)
    df['Private'] = df['Private'].astype(bool)
    df['FinalUpload'] = df['FinalUpload'].astype(bool)
    df['UploadTime'] = datetime.now()    
    return df 


# ## 1.  Applying transformations

df_ann = transform_excel_ann(excel_name_ann, file_indicators_ann)
create_excel_in_template(df_ann, file_indicators_ann, excel_name_ann)

df_qt = transform_excel_Q(excel_name_qt, file_indicators_qt)
create_excel_in_template(df_qt, file_indicators_qt, excel_name_qt)


df_final = pd.concat([df_ann,df_qt])
df_final = change_codes(df_final, cnxn, sourcedata)
reportdate = get_report_date(df_final, region=3550)
df_final = extra_columns(df_final)

print("Done! Wait for upload")


# ## 2. To SQL
#get the conncetion with our Database
from sqlalchemy import create_engine
import urllib
params = urllib.parse.quote_plus("DRIVER={SQL Server Native Client 11.0};SERVER= server;DATABASE=data;UID=username;PWD=password")

engine = create_engine("mssql+pyodbc:///?odbc_connect=%s" % params)
print("Uploading", len(df_final), "datapoints to tmp_python_DATABANK...")

df_final.to_sql('tmp_python_DATABANK',con=engine,if_exists='replace',index=False)

def execute_query(query):
    '''This function executes the query passed as argument.
    
    Inputs:
    query (str): Query passed.'''
    conn = create_connection()
    conn = conn.execution_options(autocommit=True)
    # Execute the query.
    conn.execute(query)


query = """drop table if exists #dtemp
use Data

declare @sd as int set @sd= 

select * into #dtemp from tmp_python_DATABANK where sourcedata=@sd

merge tmp t using #dtemp 
on  (t.EntityID=#dtemp.EntityID and t.indicatorid=#dtemp.indicatorid and t.period= #dtemp.period and t.Frequency= #dtemp.Frequency and t.ReportDate= #dtemp.ReportDate and t.Sourcedata= @sd and t.typedata= #dtemp.typedata)
when matched then update
set
t.EntityID=#dtemp.EntityID, t.UpdateDate=#dtemp.UpdateDate, t.value= #dtemp.Value, t.Comment= #dtemp.Comment, t.username=#dtemp.username, t.rowstate=#dtemp.rowstate, t.UploadTime=#dtemp.UploadTime
when not matched then insert
(entityid, period, frequency, indicatorid, value, typedata, sourcedata, statusdata, reportdate, updatedate, comment, private, username, rowstate, filetype, finalupload, UploadTime)
VALUES (#dtemp.EntityID, #dtemp.Period, #dtemp.Frequency, #dtemp.IndicatorID, #dtemp.Value, #dtemp.Typedata, @sd, #dtemp.StatusData, #dtemp.Reportdate, #dtemp.Updatedate, #dtemp.Comment, #dtemp.Private, #dtemp.UserName, #dtemp.RowState, #dtemp.FileType, #dtemp.FinalUpload, #dtemp.UploadTime)
;
drop table #dtemp"""


create_connection()
print("Merging that data to tmp...")


if df_final.duplicated(subset=["IndicatorID", "EntityID", "Period"]).any() == True:
    print("ERROR!, There are some duplicated rows, so data won't be uploaded :( )")
else:
    execute_query(query)
    print("Done!")