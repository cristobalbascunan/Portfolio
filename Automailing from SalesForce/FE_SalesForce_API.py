#!/usr/bin/env python
# coding: utf-8

# # SalesForce API Connection

import requests
import sys
import json
import pandas as pd
import os
import numpy as np
import re
from simple_salesforce import Salesforce
import time
from base64 import a85decode


def get_pass(path:str):
    '''
    This opens a flat file (.txt), if we have enconded the information in ASCII85
    -Parameters:
    :path: Path of the file
    -Returns: content in JSON format
    '''
    with open(path) as f:
        contents = f.read()
    contents = str(a85decode(contents))[2:-1]
    res = json.loads(contents)
    
    return res


def get_contacts_sf(groupby: str, drop_old_innactive=True):
    """
    -Parameters:
    :groupby: Use Region-Country or Region. Choose between grouping by Countries and Regions suscribed (and get their aggregates) or just by countries. 
    :drop_old_innactive: If True it will filter to Actives and Passive panelist only. Otherwise it would bring all records.
    -Returns: df
    """
    #We use Joan credentials (as he has a developer account) and then we use the salesforce library/package to log in.
    auth = get_pass("C:\...\SF_auth.txt")
    instance = auth['instance']
    username = auth['username']
    password = auth['password']
    token = auth['token']   
    sf = Salesforce(username = username,
                    password = password,
                    security_token = token,
                    instance_url = instance)    
    #Getting the first df of contacts
    query = """ SOQL query1"""
    df_contact = sf.query_all(query)
    df_contact = pd.json_normalize(df_contact, record_path =['...'])
    df_contact = df_contact[[...]]
    #Gettting the regions for each one of the panelists is subscripted (this means, which regions do we send them)
    query = """SOQL query2"""
    df_region = sf.query_all(query)
    df_region = pd.json_normalize(df_region, record_path =['...'])
    df_region = df_region[["..."]].rename(columns={"...":"..."})  
    #Getting the countries that each of the panelists are sending info.
    query = """SOQL query3"""
    df_countries = sf.query_all(query)
    df_countries = pd.json_normalize(df_countries, record_path =['...'])
    df_countries = df_countries[["..."]].rename(columns={"...":"..."})
    #Getting the ID of previous results equivalence (Countries and Regions)
    query = """SOQL query4"""
    df_map = sf.query(query)
    df_map = pd.json_normalize(df_map, record_path =['...'])
    df_region_map = df_map[["..."]].rename(columns={"...":"..."})
    df_country_map = df_map[["..."]].rename(columns={"...":"..."})
    #Merging all these tables into the main one.
    ## Countries and Regions IDs
    df = df_contact.merge(df_countries, on=["..."], how='left')
    df = df.merge(df_region, on=["..."], how='left')
    df = df.merge(df_region_map, on=["..."], how='left')
    df = df.merge(df_country_map, on=["..."], how='left')
    #Drop columns that won't have more use
    df = df.drop(columns=['...'])
    if drop_old_innactive==True:
        df = df[df['...'].isin(['...'])]
    if groupby=='...':
        #Grouping the different Countries/regions into nested lists
        df = df.groupby(['...']).agg({'...': lambda x: x.unique().tolist(), '...': lambda x: x.unique().tolist()}).reset_index()
    elif groupby=='...':
        #Grouping the different Countries/regions into nested lists
        df = df.groupby(['...']).agg({'...': lambda x: x.unique().tolist()}).reset_index()

    #Rename of some columns 
    df = df.rename(columns={"...": "..."})
    
    return df