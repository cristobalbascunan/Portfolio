#!/usr/bin/env python
# coding: utf-8

# # OneHub: move PDFs to OneHub

import requests
import sys
import json
import pandas as pd
import os
import numpy as np
import pyodbc
from datetime import date, timedelta, datetime
import re

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
    
#Selecting path and changing it
path = 'C:\...\...\OneHub'
os.chdir(path)

# ### 0. Getting the authorization token and connection
#Credentials to use to make the Oauth2 in order to get the token
username = "username_Onehub"
password = "password_Onehub"
client_id = 'client_id'
client_secret = 'client_secret'
#Part of the request needs to create the request
url = "https://ws-api.onehub.com/oauth/token"
data = "grant_type=password&client_id="+client_id+"&client_secret="+client_secret+"&username="+username+"&password="+password+""
headers = { 'accept': "application/json", 'content-type': "application/x-www-form-urlencoded" }
#The response will come as JSON in the POST request
response = requests.request("POST", url, data=data, headers=headers).json()
#Extracting the acces token from the response, so the header will contain the Autorizathion that will allow to interact with the API.
header = {'Authorization': 'Bearer ' + response['access_token']}

# #### 1. Getting values from Excel template
#Get the publication in which we are
try:
    ##Quert in SQL to download the WeeklyPublication Table
    query="""query1"""
    ##Reading the query and converting to a dataframe, and dropping the 'Period' column
    week_publi = pd.read_sql_query(query, cnxn).drop(columns="Period")
    week_publi = week_publi[(week_publi["..."] <= 3550)]
    ##Since the table has the value in wich the week ends, by substracting 6 days, we will have the day the publication week starts
    week_publi['start'] = week_publi['RealDate'] - timedelta(days=6)
    ##Creating a column with the day of today
    week_publi['today'] = datetime.today()
    week_publi['today'] = pd.to_datetime(week_publi['today']).dt.date
    week_publi['today'] = pd.to_datetime(week_publi['today'])
    ##We will get the number of the publication week, when the day of today is between the start and end of a week publication
    week_number = week_publi.loc[week_publi['today'].between(week_publi['start'], week_publi['RealDate']), ['...']].values.item()
except:
    week_number = input('There has been an error, please introduce manually the week number: ')
    week_number = int(week_number)
    print(f'Selected the week number {week_number}')
    
#Getting the reportdate in which that reportdate is ¡   
query = 'query2'
reportdate = pd.read_sql_query(query, cnxn)
reportdate = reportdate.loc[reportdate['...'] == week_number]['...'].iloc[0]
#Date to delete in onehub (minus 1 year) 
reportdate_1yearago = reportdate - pd.DateOffset(years=1)
print('Printing clients from region ' + str(week_number) + ' with Report Date ' + str(reportdate))

#Read OneHub specifications that are in a custom excel
df_files = pd.read_excel('... .xlsx')
#Selecting just the clients of the selected region/week
df_files = df_files.loc[(df_files["..."] == week_number)].reset_index(drop=True)

#Year formats
YY = reportdate.strftime("%Y")
yymm =  reportdate.strftime("%y%m")
#Empty lists
temp_path =[]
errors_in=[]
#Get every PDFs current publication names
for i in range(len(df_files)):
    df = df_files.iloc[i]
    #Path of the PDFs
    path = str(df['Path']).replace("YY",YY).replace("yymm", yymm)+'\\'
    #Latin america exceptions
    if df['EntityName'] in ['...']:
        #There are some exceptions in Latin America and special names
        try: 
            dict_latam = {'...':"..."}
            for filename in os.listdir(path):
                regex = re.compile(str('.*... '+df.replace(dict_latam)['...']+'-.*'))
                if regex.match(filename):
                    excel_name =  path+filename
            temp_path.append(excel_name)
            del excel_name
        except:
            errors_in.append(df['...'])
            print('error in', df['...'], 'you have to upload it manually.')        
    #For the rest of the regions
    else:
        try: 
            for filename in os.listdir(path):
                regex = re.compile(str('.*Forecast .*'+str(df['...'])+' -.*'))
                if regex.match(filename):
                    excel_name =  path+filename
            temp_path.append(excel_name)
            del excel_name
        except:
            errors_in.append(df['...'])
            print('error in', df['...'], 'you have to upload it manually.')

#Those that gave an error (non existant PDFs) and delete it of the list to upload
df_files = df_files[~df_files['...'].isin(errors_in)]
df_files["..."] = temp_path
#Change format of the 'file to remove'
YY_1ago = reportdate_1yearago.strftime("%Y")
#Getting the exact name of a list of file to remove
df_files['...'] = df_files['...'].str.replace(YY,YY_1ago)

#Loop to upload to Onehub one by one
for i in range(len(df_files)):
    try:
        #Select name of the PDF
        x = df_files.iloc[i]
        #ID in Onehub URL
        id_folder = x['id_Link']
        #Post string for the request        
        folder_onehub = 'https://ws-api.onehub.com'+id_folder+'/files'
        file_path = x['Path']
        #Files part of the request
        files = {'files': open(file_path, 'rb')}
        #POST request
        requests.post(folder_onehub, files = files, headers=header)
        print(x['EntityName'], "uploaded")
    except: 
        x = df_files.iloc[i]
        #If not possible, print a message
        print(x['EntityName'], "doesn't exist")
        
## This part is to get the -1 year report and delete it, so we keep just 12 PDFs in onehub.        
    try:    
        #Post string for the request        
        folder_onehub = 'https://ws-api.onehub.com'+id_folder+'/files'
        #Request to get a list of items where the file was uploaded
        df_contents = requests.get(folder_onehub, headers=header).json()
        #get content of the folder
        df_contents = pd.json_normalize(df_contents, record_path =['items'])
        file_to_remove = x['File_to_remove'].split('\\')[-1]
        df_folder = pd.DataFrame(data={'id': df_contents['file.id'], 'filename':df_contents['file.filename']})
        #Look for the id to remove
        id_to_remove = df_folder.loc[df_folder["filename"] == file_to_remove, "id"]
        path_to_remove = 'https://ws-api.onehub.com/files/'+str(id_to_remove.iloc[0])
        #DELETE request
        requests.delete(path_to_remove, headers=header)
    except:
        pass                 
print("You can close!")