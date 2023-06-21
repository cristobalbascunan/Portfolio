#!/usr/bin/env python
# coding: utf-8

import pandas as pd
import numpy as np
import pyodbc
import re
import win32com.client
import pandas
from pandas.tseries.offsets import MonthEnd
import datetime
from datetime import datetime
from macrobond_api_constants import SeriesFrequency as f
from macrobond_api_constants import SeriesToHigherFrequencyMethod as h #Search for its 
from macrobond_api_constants import SeriesToLowerFrequencyMethod as l
from macrobond_api_constants import SeriesMissingValueMethod as m
import sys


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


# ### 0. Functions to be used
#Function: Get Macrobond Series to Pandas
def toPandasSeries(series):
    pdates = pd.to_datetime([d.strftime('%Y-%m-%d') for d in series.DatesAtStartOfPeriod])
    return pd.Series(series.values, index=pdates)

#Fucntion: Get one series from Macrobond as a pandas time series
def getOneSeries(db, name):
    return toPandasSeries(db.FetchOneSeries(name))

#Fucntion: Get several series from Macrobond as pandas time series
def getSeries(db, names):
    series = db.FetchSeries(names)
    return [toPandasSeries(s) for s in series]

#Function: Get several series from Macrobond as a pandas data frame with a common calendar
def getDataframe(db, unifiedSeriesRequest):
    series = db.FetchSeries(unifiedSeriesRequest)
    return pd.DataFrame({s.Name: toPandasSeries(s) for s in series})

def change_codes(df, cnxn):
    df_mappingcountries = pd.read_sql_query("SELECT EntityName, NumericCode FROM databasetable",cnxn)
    entitynames = df_mappingcountries['EntityName']
    numericcodes = df_mappingcountries['NumericCode']            
    dict_ch_entitycode = dict(zip(entitynames, numericcodes))
    df['Entity'] = df['Entity'].map(dict_ch_entitycode).astype('int32')
    df = df.rename(columns={'Entity':'EntityID'})
    
    df_mappingindicator = pd.read_sql_query("select * from databasetable",cnxn)
    indicnames = df_mappingindicator['ShortName']
    indicID = df_mappingindicator['ID']
    dict_ch_indicatorcode = dict(zip(indicnames, indicID))
    df['ShortName'] = df['ShortName'].map(dict_ch_indicatorcode).astype(int)
    df = df.rename(columns={'ShortName':'IndicatorID'})    
    return df

def extra_columns(df):
    df['UpdateDate'] = updatedate
    df['FileType'] = 'Panelist'
    df['TypeData'] = 1
    df['SourceData'] = sourcedata
    df['Comment'] = 'NULL'
    df['UserName'] = 'Python_WEO'
    df = df.rename(columns = {'value': 'Value'})
    new_period = df['Period']
    new_period = new_period.apply(lambda dt: dt.replace(day=1))
    df['Period'] = new_period
    df['Value'] = df['Value'].astype(float)
    df['Private'] = df['Private'].astype(bool)
    df['UploadTime'] = datetime.now()
    return df 

def add_reportdate(df):
    query="""SELECT e.entityname as Entity, reportdate as ReportDate 
            FROM databasetable e
            INNER JOIN databasetable1 c ON c.column=ea.column
            WHERE  e.column in (SELECT columns 
                                    FROM databasetable 
                                    WHERE column IN (2051,2052,2053,3550,5005))
    """
    reportdate_df = pd.read_sql_query(query, cnxn)
    df = df.merge(reportdate_df, on="Entity")
    return df


# ##### Connection with the database
cnxn = create_connection()
#source data
sourcedata = 251 
# ### 1. Getting data from Macrobond API
#Macrobond API Connection
c = win32com.client.Dispatch("Macrobond.Connection")
d = c.Database
m = d.CreateEmptyMetadata()
r = d.CreateUnifiedSeriesRequest()
#Read mapped indicators Excel
mb_weo_concept = pd.read_excel (r'C:\...\MB WEO Concept.xlsx')
#Remove indicators that are removed and shown as '-'
mb_weo_concept = mb_weo_concept[mb_weo_concept.Concept != "-"]
#Read mapped countries Excel
entities = pd.read_excel (r'C:\...\Macrobond_entities.xlsx')
#Extract series of entities
entities_all = entities['Code']
#Extract series of indicators (in MB Concept)
concepts = mb_weo_concept['Concept']
Entity = []
tickers = []
#Iterate and 'query' each indicator 
for x in concepts:
    query = d.CreateSearchQuery()
    query.SetEntityTypeFilter("TimeSeries")
    #Filter by indicator
    query.AddAttributeValueFilter("RegionKey", x)
    #Filter and select all countries for that indicator
    query.AddAttributeValueFilter("Region", entities_all) 
    #Save result
    result = d.Search(query).Entities
    #Append to empty lists the countries and indicatos get
    for s in result:
        r.AddSeries(s.Name)
        m = s.Metadata
        Country = m.GetFirstValue("Region")
        Entity.append(Country)
        tickers.append(s.Name)
s = d.FetchOneSeries(tickers[0])
#Extract metadata
releaseName = s.Metadata.GetFirstValue("Release")
if (releaseName is not None):
    rel = d.FetchOneEntity(releaseName)
    nextReleaseTime = rel.Metadata.GetFirstValue("LastReleaseEventTime")
#Get from release date the update date
updatedate = datetime.strptime(str(nextReleaseTime)[0:10], '%Y-%m-%d')
print("The data was updated on", updatedate)

#Empty dataframe where entity and tickers will be append
df = pd.DataFrame()
df['Entity'] = Entity
df['MB Code'] = tickers
#Rename columns in Excel entities
entities.rename(columns={'Code': 'Entity', 'Description': 'Full_Name'}, inplace=True)
#Merge info downlaoded from MB and merge it with excel 
EntityTicker = df.merge(entities,how='left')
#Request in order to get the data (x= ticker in Macrobond)
for x in EntityTicker['MB Code']:
    r.Addseries(x)
    #Frequency annual
    r.Frequency=f.ANNUAL
    #Starting date
    r.startDate ='2021-01-01'
data = getDataframe(d, r)
data = data.reset_index()
#Melting data (from wide to long)
data_pivoted = pd.melt(data, id_vars=data.columns[0], value_vars=data.columns[1:])
#Rename columns
data_pivoted.rename(columns={'index': 'Period', 'variable':'MB Code'}, inplace=True)
#Merging data with information in each ticker of Macrobond and ours.
data_merged = data_pivoted.merge(EntityTicker, how ='inner').drop_duplicates().dropna().reset_index(drop=True)
code_filtered = []
#Since the tickers brought from Macrobond have a different composition (they add the name of the country in the middle)
#in order to clasify them by indicator, it is necessary to use regex to extract that info. For local currency (USLDCUEOP)
#they use the country name so thats why is treated differently
for i in range(len(data_merged['MB Code'])):
    if 'weo' in str(data_merged['MB Code'][i]):
        x = re.sub('(^[a-z]..)|((?<=weo)_[a-z]*?(?=_))','',data_merged['MB Code'][i])
        code_filtered.append(x)
    else: 
        #USDLCUEOP
        x = re.sub('.$','',data_merged['MB Code'][i])
        code_filtered.append(x)
data_merged['Code_filtered'] = code_filtered
#Merging with the table that contains the conversion of indicators to ours.
data_merged = data_merged.merge(mb_weo_concept, left_on='Code_filtered', right_on='Concept', how='left').drop_duplicates().reset_index(drop=True)
#Those that have the same name are equal to USDLCUEOP
data_merged['ShortName'] = np.where(data_merged['Code_filtered'] == data_merged['Entity'],'USDLCUEOP',data_merged['ShortName'])
#Dropping columns
data_merged = data_merged.drop(columns=['MB Code', 'Entity', 'Full_Name', 'Code_filtered', 'Concept'])
data_merged.rename(columns={'Focus name': 'Entity', 'value':'Value'}, inplace=True)
#Replacing '-' with nans
data_merged = data_merged.replace('-', np.nan).dropna()
#Dropping nans
data_merged = data_merged.dropna(subset=['ShortName'])

# ### 2. Transformations in data
complete_df = data_merged.copy()
#Convert column to datetime
complete_df['Period'] = pd.to_datetime([x for x in complete_df['Period']])
#adding column 'frequency'
complete_df['Frequency'] = 1
##Adding some /10000... calculations
complete_df.loc[complete_df["ShortName"] == "GDPUSD", "Value"] = complete_df.loc[complete_df["ShortName"] == "GDPUSD", "Value"]/1000000000
complete_df.loc[complete_df["ShortName"] == "GDPEUR", "Value"] = complete_df.loc[complete_df["ShortName"] == "GDPEUR", "Value"]/1000000000
omplete_df.loc[complete_df["ShortName"] == "POP", "Value"] = complete_df.loc[complete_df["ShortName"] == "POP", "Value"]/1000000
##Drop those countries that contaisn GDPEUR but shouldn't (avoiding useless data).
list_countries_gdpeur = ['Austria', 'Belgium', 'Cyprus', 'Estonia', 'Finland', 'France', 'Greece', 'Germany', 'Ireland', 'Italy', 'Latvia', 'Lithuania', 'Luxembourg', 'Malta', 'Portugal', 'Slovakia', 'Slovenia', 'Spain', 'Netherlands', 'Kosovo', 'Montenegro']
complete_df['Euro'] = 'No'
#Those that have GDPEUR switched to 'Yes'
complete_df.loc[(complete_df["Entity"].isin(list_countries_gdpeur))&(complete_df["ShortName"] == "GDPEUR"),"Euro"] = 'Yes'
#Getting the index of those that shouldn't have GDPEUR and dropping them
complete_df = complete_df.drop(complete_df[(complete_df['ShortName'] == 'GDPEUR') & (complete_df['Euro'] == 'No')].index).drop(columns='Euro')
##GDPEUR conversions for those that are in Trillions instead of Billions 
complete_df.loc[(complete_df["ShortName"] == "GDPEUR") & (complete_df["Entity"] == "Norway"), "Value"] = complete_df.loc[(complete_df["ShortName"] == "GDPEUR") & (complete_df["Entity"] == "Norway"), "Value"]/10
complete_df.loc[(complete_df["ShortName"] == "GDPEUR") & (complete_df["Entity"] == "Norway"), "Value"] = complete_df.loc[(complete_df["ShortName"] == "GDPEUR") & (complete_df["Entity"] == "Norway"), "Value"]/10

#Add report date depending on the region of the country
df_final = add_reportdate(complete_df)
#Change names to IDs
df_final = change_codes(df_final, cnxn)
#Add extra columns to final df
df_final = extra_columns(df_final)
#Changing years to adapt Ethiopia to Fiscal Year
df_final.loc[(df_final["EntityID"] == 231) & (df_final["IndicatorID"]== 3) & (df_final["Frequency"]==1), "Period"] = (df_final.loc[(df_final["EntityID"] == 231) & (df_final["IndicatorID"]== 3) & (df_final["Frequency"]==1), "Period"] - MonthEnd(12)).replace("30", "31").map(lambda x: x.replace(day=1))
# ### 3. Uploading data to the database
# get the conncetion with our Database
from sqlalchemy import create_engine
import urllib
params = urllib.parse.quote_plus("DRIVER={SQL Server Native Client 11.0};SERVER=server;DATABASE=database;UID=user;PWD=password")
engine = create_engine("mssql+pyodbc:///?odbc_connect=%s" % params)
df_final.to_sql('tmp_table',con=engine,if_exists='replace',index=False)


def execute_query(query):
    '''This function executes the query passed as argument.    
    Inputs:
    query (str): Query passed.'''
    conn = create_connection()
    conn = conn.execution_options(autocommit=True)
    # Execute the query.
    conn.execute(query)

query = """drop table if exists #dtemp
USE database
DECLARE @sd as int set @sd=251
SELECT * into #dtemp from databasetable where sourcedata=@sd

MERGE databasetable1 t using #dtemp 
on  (t.EntityID=#dtemp.EntityID and t.indicatorid=#dtemp.indicatorid and t.period= #dtemp.period and t.Frequency= #dtemp.Frequency and t.ReportDate= #dtemp.ReportDate and t.Sourcedata= @sd and t.typedata= #dtemp.typedata)
when matched then update
set
t.EntityID=#dtemp.EntityID, t.UpdateDate=#dtemp.UpdateDate, t.value= #dtemp.Value, t.Comment= #dtemp.Comment, t.username=#dtemp.username, t.rowstate=#dtemp.rowstate
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
    print("Done! You can close the window! :)")