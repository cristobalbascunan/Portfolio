#!/usr/bin/env python
# coding: utf-8

# # Automation of CLIENT

import os
import pandas as pd
import numpy as np
import errno
import pyodbc
from datetime import date, timedelta, datetime
import time
import pywinauto
from pywinauto.application import Application
from pywinauto.keyboard import send_keys
from pywinauto import keyboard as kb

print('Importing packages and defining function...')
#When it's impossible to install pyodbc
# pip install --upgrade pip
# pip install pyodbc
# pip install openpyxl

#Selecting path and changing it
os.chdir(r'C:\...\...\Specifications for clients')

# Connection with SQL so we can run the queries.
########################################################
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
    
cnxn = create_connection()
cursor = cnxn.cursor()
########################################################

####THIS CODE WAS USED IN ORDER TO EXTRACT THE INFORMATION FROM CLIENT FILES
# #### 0.0 Code used to get every country-indicator-... from whitin the CLIENT files (no need to use it again)

# # import required module
# from pathlib import Path
# import glob 

# regions = {'...':2051, '...':2051, "...":2052, '...':2053, '...':2053, '...':3550}
# df_content = pd.DataFrame()
# for region, code in regions.items():
#     # assign directory
#     directory = r"C:\...\\" + region+"\...\CLIENT\\"
#     # iterate over files in that directory
#     files = glob.glob(directory+'[a-zA-Z]*_client.xlsm')    
#     for file in files:
#         df = pd.read_excel (file, sheet_name='Data')
#         df_ann=df[['...', '...']]
#         df_ann=df_ann.dropna()
#         df_ann['...'] = 'Annual'
#         df_ann.columns = ["...", "...", "..."]
#         try:
#             df_q=df[['Quarterly Indicators', '....1']]
#             df_q=df_q.dropna()
#             df_q['...'] = 'Quarterly'
#             df_q.columns = ["Indicator", "...", "..."]
#             df_ann = pd.concat([df_ann, df_q], ignore_index=True)
#         except:
#             pass
#         df_ann["..."] = df[".."].iloc[0]
#         df_ann["..."] = str(code)
#         df_content = pd.concat([df_content, df_ann], ignore_index=True)
#         df_content.to_excel('list.xlsx')

# #### 0.1 Getting week number in order to just generate the publication files
try:
    ##Quert in SQL to download the WeeklyPublication Table
    query="""SELECT * FROM databasetable"""
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
    week_number = int(input('There has been an error, please introduce manually the week number: '))
    print(f'Selected the week number {week_number}')
# #### 0.2 Getting report name for the file's content
query = 'SELECT *  FROM databasetable'
reportdate = pd.read_sql_query(query, cnxn)
reportdate = reportdate.loc[reportdate['...'] == week_number]['...'].iloc[0]
###IF ITS NEEDED TO SELECT WEEK AND REPORTDATE MANUALLY
#reportdate = pd.to_datetime('2023-04-01')
#week_number= 2052 
print('Printing clients from region ' + str(week_number) + ' with Report Date ' + str(reportdate))
# #### 0.3 Max and min year for the file
date_today = datetime.today().strftime("%y%m%d")
min_year = pd.to_datetime(reportdate.year - 2, format="%Y")
max_year = pd.to_datetime(reportdate.year + 1, format="%Y")
min_qt = pd.to_datetime(reportdate.year, format="%Y")
max_qt = pd.to_datetime(reportdate.year + 1, format="%Y")
# ### 1. Reading the excel template
##Reading the template made in Excel
excel_file =pd.read_excel('Client country-indicator.xlsx')
excel_file = excel_file.loc[(excel_file["Region"] == week_number)].reset_index(drop=True)
countries = list(excel_file["..."].unique())
name_countries = list(excel_file[".."].unique())
#Getting from the database the right data, this comes from datatable
query = """SELECT EntityName, Period, Frequency, TypeName, value, ShortName
FROM databasetable 
WHERE Frequency <=2 AND ReportDate='"""+str(reportdate)+"""' 
AND period BETWEEN '"""+str(min_year)+"""' AND '"""+str(max_year)+"""' 
ORDER BY EntityName, TypeName, Frequency"""
df_f = pd.read_sql_query(query, cnxn)
df_f = df_f[df_f["EntityName"].isin(countries)]
df_f.loc[df_f['...'] == 1, ['...']] = "Annual"
df_f.loc[df_f['...'] == 2, ['...']] = "Quarterly"
#And from HData
query = """SELECT EntityName, Period, FrequencyName AS Frequency, ShortName, value
FROM databasetable 
WHERE FrequencyName IN ('Annual','Quarterly') AND period BETWEEN '"""+str(min_year)+"""' AND '"""+str(max_year)+"""' """
df_h = pd.read_sql_query(query, cnxn)
df_h = df_h[df_h["..."].isin(countries)]
df_h['...'] = "Consensus"
#Concatenate both tables
df_full = pd.concat([df_f, df_h], ignore_index=True)

#Change path directory
os.chdir(r'C:\...\CLIENT')
print("Creating files...")
#Create file if it doesnt exist (to save all the files)
try:
    os.mkdir(date_today)
except OSError as e:
    if e.errno != errno.EEXIST:
        raise
#Loop to create each one of the files        
for country in countries:
    #Table of country specifications 
    excel_file_country = excel_file[excel_file["..."] == country]
    name_country = excel_file_country["..."].iloc[0]
    indicators = list(excel_file_country['...'].unique())
    frequencies = list(excel_file_country['...'].unique())
    #Create files for different frequencies (Annual - Qt)
    for frequency in frequencies:
        if frequency == 'Annual':
            df_ann= df_full[df_full["..."] == 'Annual']
            df_ann= df_ann[df_ann["..."] == country]
            df_ann = df_ann[df_ann["..."].isin(indicators)]
            df_ann['...'] = df_ann['...'].dt.strftime("%Y")
            #Add GDPCAPUD minimm and maximum
            try:
                gdp_cap = df_ann[df_ann['...']=='GDPCAPUSD'][0:1]
                gdp_cap['...'] = 0
                gdp_cap['...'] = 'Minimum'
                gdp_cap_min = gdp_cap.copy()
                gdp_cap['...'] = 'Maximum'
                gdp_cap_max = gdp_cap.copy()
                df_ann = pd.concat([df_ann, gdp_cap_min, gdp_cap_max], ignore_index=True)
            except:
                pass
            df_ann = df_ann.merge(excel_file_country, on=['...', '...','...'])
            CLIENT_Code = df_ann['Code'].iloc[0]
            CLIENT_Region= df_ann['Region_code'].iloc[0]
            CLIENT_Country= df_ann['Country'].iloc[0]
            CLIENT_Frequency = str(df_ann['Frequency'].iloc[0]) + " Data" 
            df_ann = df_ann[['...',"...", "...",'...', '...']]
            df_pivot = pd.pivot_table(df_ann, values='value', index=['...', '...', 'index'],
                        columns=['Period'])
            df_pivot.columns = df_pivot.columns.get_level_values(0)
            df_pivot = df_pivot.reset_index()
            df_pivot = df_pivot.rename_axis(None, axis=1)
            df_pivot = df_pivot.set_index('TypeName')
            df_pivot = df_pivot.iloc[pd.Categorical(df_pivot.index,["Consensus","Minimum", "Maximum"]).argsort()].reset_index()
            df_pivot = df_pivot.sort_values(by=['index']).reset_index()
            df_pivot = df_pivot.sort_values(by=['index', "level_0"], ascending=True).drop(columns=['level_0', 'index'])
            try:
                df_pivot.loc[(df_pivot['...'] == "...")&(df_pivot['TypeName'] != "Consensus"),"2023"] = np.nan
            except:
                pass
            #Convert NaN to '-'
            df_pivot = df_pivot.replace(np.nan, '-')
            df_pivot = df_pivot[['...', 'TypeName']+ list(df_pivot.columns)[2:]]
            df_pivot = df_pivot.rename(columns={'...':CLIENT_Frequency, "TypeName":CLIENT_Code})
            df_pivot[CLIENT_Region] =""
            #Export CSV
            df_pivot.to_csv(date_today+"\\"+str(name_country) +'_CLIENT_'+date_today+".csv", index=False)
    
        if frequency == 'Quarterly':
            try: 
                df_qt = df_full[df_full["Frequency"] == 'Quarterly']
                df_qt= df_qt[df_qt["..."] == country]
                df_qt = df_qt[df_qt["..."].isin(indicators)]
                df_qt = df_qt[(df_qt["Period"] > min_qt) & (df_qt["Period"] < max_qt)]
                df_qt['Period'] = df_qt['Period'].dt.to_period('Q').dt.strftime('Q%q %y')
                df_qt= df_qt[df_qt["..."] == 'Consensus']
                df_qt = df_qt.merge(excel_file_country, on=['...', '...','...'])
                CLIENT_Code = df_qt['Code'].iloc[0]
                CLIENT_Region= df_qt['...'].iloc[0]
                CLIENT_Country= df_qt['...'].iloc[0]
                CLIENT_Frequency = str(df_qt['Frequency'].iloc[0]) + " Data" 
                df_qt =df_qt[['...',"...", "Period",'value', 'index']]
                df_pivot = pd.pivot_table(df_qt, values='value', index=['...', 'TypeName', 'index'],
                            columns=['Period'])
                df_pivot.columns = df_pivot.columns.get_level_values(0)
                df_pivot = df_pivot.reset_index()
                df_pivot = df_pivot.rename_axis(None, axis=1)
                df_pivot = df_pivot.set_index('TypeName')
                df_pivot = df_pivot.iloc[pd.Categorical(df_pivot.index,["...","Minimum", "Maximum"]).argsort()].reset_index()
                df_pivot = df_pivot.sort_values(by=['index']).reset_index()
                df_pivot = df_pivot.sort_values(by=['index', "level_0"], ascending=True).drop(columns=['level_0', 'index'])

                df_pivot = df_pivot.replace(np.nan, '-')
                df_pivot = df_pivot[['...', 'TypeName']+ list(df_pivot.columns)[2:]]
                df_pivot = df_pivot.rename(columns={'...':CLIENT_Frequency, "TypeName":CLIENT_Code})
                df_pivot[CLIENT_Region] =""
                df_pivot.to_csv(date_today+"\\"+str(name_country) +'_CLIENT_Quarterly_'+date_today+".csv", index=False)
            except:
                pass
print("Done!")


# 2. Open WinSCP and apply the correspondant process
app = Application(backend="win32").start(cmd_line=r"C:\Program Files (x86)\WinSCP\WinSCP.exe")
time.sleep(0.5)
#Anti first window of updates
wnd = app.top_window()
first_window = str(wnd.texts()[0])
if 'Update' in first_window:
    win0 = app["TMessageForm"]
    win0.set_focus()
    win0["CloseButton"].click()
#Login window 
app.Login.set_focus()
win=app.TLoginDialog
#Try to login, otherwise insert the credentials
try:
    win.Login.click()
except:
    win["Edit"].set_text(u"21")    
    win["Edit4"].set_text("ftp2.CLIENT.de")
    win["Edit3"].set_text("user")
    win["Save"].click()
    win0_5 = app["TSaveSessionDialog"]
    win0_5.set_focus()
    win0_5["OKButton"].click()
    win.Login.click()
time.sleep(0.25)
win1 = app["TAuthenticateForm"]
#Inserting passwords
win1["Edit"].set_text("password")
win1.OK.click()
win1=app.TScpCommanderForm
win1.set_focus()
time.sleep(1)

#Open slection window
send_keys("^O")
app[u'Open directory'].set_focus()
win2 = app.TOpenDirectoryDialog
time.sleep(1)
#Type source of files
src_folder = r"C:\...\CLIENT"+'\\'+date_today
win2["Edit"].set_text(src_folder)
time.sleep(2)
win2.OK.click()

#If it creates an error, there is a loop that has to type outside the current folder and then come back
try: 
    app.Error.OK.click()
    time.sleep(1)
    win1=app.TScpCommanderForm
    win1.set_focus()
    time.sleep(1)
    send_keys("^O")
    app[u'Open directory'].set_focus()
    win2 = app.TOpenDirectoryDialog
    time.sleep(1)
    #Select root folder
    src_folder = r"C:\...\CLIENT"
    win2["Edit"].set_text(src_folder)
    time.sleep(1)
    win2.OK.click()  
    time.sleep(1)
    win1=app.TScpCommanderForm
    win1.set_focus()
    time.sleep(1)
    send_keys("^O")
    app[u'Open directory'].set_focus()
    win2 = app.TOpenDirectoryDialog
    time.sleep(1)
    #Select proper folder
    src_folder = r"C:\...\CLIENT"+'\\'+date_today
    win2["Edit"].set_text(src_folder)
    time.sleep(1)
    win2.OK.click()   
except:
     pass
#Upload files
time.sleep(1)
send_keys("^A")
time.sleep(1)
send_keys("{F5}")
time.sleep(1)
send_keys("{ENTER}")
time.sleep(2)
send_keys("%F")
#Trying to close
for x in range(0, 6):  # try 4 times
    try:
        app.Information.Yes.click()
        str_error = None
    except Exception as e:
        str_error = str(e)
        pass

    if str_error:
        time.sleep(3)  # wait for 3 seconds
    else:
        break

