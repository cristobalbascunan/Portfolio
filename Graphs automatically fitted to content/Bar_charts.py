#!/usr/bin/env python
# coding: utf-8


import sys
import numpy as np
import matplotlib as mpl
import matplotlib.pyplot as plt 
import pandas as pd
from pandas.api.types import CategoricalDtype
from math import ceil, floor
import time
from matplotlib.ticker import FormatStrFormatter
pd.options.mode.chained_assignment = None  # default='warn'
from datetime import datetime, timedelta
import matplotlib.patches as mpatches
import matplotlib.path as mpath



"""
This module contains the function to check if a legend is overlaping 
with the series being ploted.
"""

def check_if_serie_overlaps(legend: plt.Axes, ax1: plt.Axes = None, ax2: plt.Axes = None) -> bool:
    """
    Return if the legend overlaps with the series being ploted.
    Currently the implementation only takes into account if the serie overlaps or not.
    It doesn't matter if it overlaps with columns or lines. TODO explore if it can be implemented.
    I tried and ended in an infinite loop.
    :param legend:
    :param ax1: Left vertical axis of the plot.
    :param ax2: Right vertical axis of the plot.
    :returns: True if overlaps, False if don't.
    """
    # Get the space filled by the legend.
    legend_bbox = legend.get_window_extent() # Això és display space
    if (ax1 is None) and (ax2 is None):
        for line in plt.gca().lines:
            # Get the path the lines are passing in Display space.
            points_coordinates = plt.gca().transData.transform(line.get_xydata())
            path = mpath.Path(points_coordinates)
            # Check if both cross at some point.
            if path.intersects_bbox(legend_bbox):
                return True
    elif (ax1 is not None) and (ax2 is None):
        for line in ax1.lines:
            # Get the path the lines are passing in Display space.
            points_coordinates = plt.gca().transData.transform(line.get_xydata())
            path = mpath.Path(points_coordinates)
            # Check if both cross at some point.
            if path.intersects_bbox(legend_bbox):
                return True
        # Check the display space of all the columns.
        rectangles = [child for child in ax1.get_children() if isinstance(child, mpatches.Rectangle)][:-1]
        for rectangle in rectangles:
            if (rectangle.get_window_extent().overlaps(legend_bbox))&(str(rectangle.get_window_extent())!= 'Bbox(x0=inf, y0=inf, x1=-inf, y1=-inf)'):
                return True
    else:
        to_return = []
        rectangles = [child for child in ax1.get_children() if isinstance(child, mpatches.Rectangle)][:-1]
        for rectangle in rectangles:
            if rectangle.get_window_extent().overlaps(legend_bbox):
                to_return.append((True, 'columns'))
                break
        for line in ax2.lines:    
            points_coordinates = plt.gca().transData.transform(line.get_xydata())
            path = mpath.Path(points_coordinates)
            if path.intersects_bbox(legend_bbox):
                to_return.append((True, 'line'))
        if len(to_return) > 0:
            return to_return


def round_to_multiple(number, multiple, where='up'):
    '''
    This will round a number to its a higher (ceil) integer that's a multiple of the number decided. In case of negative
    numbers it is used the floor value (in absolute terms this would be the upper value). This will be used 
    
    :param number: number to round 
    :param multiple: multiple decided

    :return: value
    '''
    if where == 'up':
        #If number is higher or equal to 0, use ceil
        if number >= 0:
            return multiple * ceil(number / multiple)
        #If number is higher or equal to 0, use floor
        if number < 0:
            return multiple * floor(number / multiple)
    elif where =='down':
        #If number is higher or equal to 0, use ceil
        if number >= 0:
            return multiple * floor(number / multiple)
        #If number is higher or equal to 0, use floor
        if number < 0:
            return multiple * ceil(number / multiple)       


def get_rounded_limits(df_column):
    '''
    Gets the minimum and miximum value from a list (dataframe column). How to get this value and the necessary conversion
    depends if the value is positive or negative. When it's positive, the minimum value takes the value closest to the left
    value (in negative values would mean a ceil, in positives would be a floor) and the maximum the oposite.
    
    :param df_column: dataframe column

    :return: returns a list of two values [min_value, max_value]
    '''
    #Conversions in max_value
    if ~pd.isna(df_column.max().max()):
        max_value = df_column.max().max() * 1.05
    elif ~pd.isna(df_column.max()):
        max_value = df_column.max() * 1.05
    else:
        print('error')
    if max_value < -10:
        max_value_rounded = -floor(abs(max_value))
    elif max_value < 0:
        max_value_rounded = -floor(abs(max_value)*2)/2
    elif max_value == 0:
        max_value_rounded = 0
    elif max_value < 10:
        max_value_rounded = ceil(abs(max_value)*2)/2
    else:
        max_value_rounded = ceil(abs(max_value))
    #Conversions in min_value
    min_value = df_column.min().min() * 1.05
    if min_value < -10:
        min_value_rounded = -ceil(abs(min_value))
    elif min_value < 0:
        min_value_rounded = -ceil(abs(min_value)*2)/2
    elif min_value == 0:
        min_value_rounded = 0
    elif max_value < 10:
        min_value_rounded = floor(abs(min_value)*2)/2
    else:
        min_value_rounded = floor(abs(min_value))


    return [min_value_rounded, max_value_rounded]       


def adjusting_limits(interval, solo=False, ticks=6):
    '''
    This is a complex way to get the most fitted limits for the axis in order to have proper values for correct
    visualizations. This is done after getting the rounded limits.
    
    :param interval: interval containing a maximum and minimum to adjust. In combination with 'solo', it can contain just 1 value.
    :param solo: False as default. Turn to True when adding just one limit to be changed.

    :return: returns a list of two values [min_value, max_value]. If 'solo'=True it will return just the value to change.
    '''

    #This is only used for CFCharts_H. When a value is X times higher it need to be redimensioned in order to have proper
    #limits.
    #If solo is turned to 0 it will take both parts of the interval in order to check if both need to start from 0.
    if solo == False:
        interval_min=interval[0]
        interval_max=interval[1]
        #So if both of them turns to be higher or lower than zero. It would change the graph origin to 0 (right when negative, left when positive)
        if (interval_min>0) & (interval_max > 0):
            interval_min=0       
        if (interval_min<0) & (interval_max < 0):
            interval_max=0
        listed_interval = [interval_min, interval_max]
    #Otherwise, it will take the interval (just one value) just as it is
    else:
        listed_interval = [interval]
        
    #Starting and empty list where next generated values will be stored     
    limits_rounded=[]
    #This loop will use the previous function to get the rounded multiple depending on the own value.
    for x in listed_interval:
        #The value will be used in absolute terms. If it's negative it will be stored the minus sign in order to add it later.
        is_negative = x < 0
        x=abs(x)
        #If it's higher than 10 it will store how much times it can be divided by 10 until reaching >=10, in order to treat them as lower values, so a 47.343 would follow the same logic than a 47.
        zeros = 0
        while x >= 100:
            x = x/10
            zeros += 1
        #Depending on the number, to make them more stetical, there are some default multiples to each range.
        if ticks == 6:
            if x > 90:
                n = round_to_multiple(x, 20, where='up')
            elif x > 75:
                n = round_to_multiple(x, 100, where='up')
            elif x > 60:
                n = round_to_multiple(x, 25, where='up')
            elif x > 35:
                n = round_to_multiple(x, 10, where='up')
            elif x > 12.5:
                n = round_to_multiple(x, 5, where='up')
            elif x > 6:
                n = round_to_multiple(x, 2.5, where='up')
            elif x > 5:
                n = round_to_multiple(x, 1, where='up')
            elif x > 4:
                n = round_to_multiple(x, 2.5, where='up')
            elif x > 0: 
                n = round_to_multiple(x, 0.5, where='up')
            elif x == 0:
                n = 0
        else:
            if x > 90:
                n = round_to_multiple(x, 20, where='up')
            elif x > 35:
                n = round_to_multiple(x, 20, where='up')
            elif x > 12.5:
                n = round_to_multiple(x, 4, where='up')
            elif x > 4:
                n = round_to_multiple(x, 2, where='up')
            elif x > 0: 
                n = round_to_multiple(x, 0.4, where='up')
            elif x == 0:
                n = 0            
        #Adding back the zeros    
        n = n*int("1"+"0"*zeros)
        #Adding back the sign if it was negative    
        if is_negative == True:
            n = -n
        limits_rounded.append(n)
    #Depending in the 'solo' feature, it will return both limits or just the one to be changed.   
    if solo ==False:
        return limits_rounded[0], limits_rounded[1]
    else:
        return limits_rounded[0]


def rescale_intervals(min_limit, max_limit, ticks):
    """
    """
    zeros = 0
    while max_limit >= 100:
        max_limit = max_limit/10
        zeros += 1
    min_limit = min_limit/int("1"+"0"*zeros)
    min_limit = min_limit * 0.95
    if ticks == 6:
        diff = abs(max_limit) - abs(min_limit)
        list_multiples = [25, 5, 2.5, 1, 0.5]
        n = 0
        if diff <= 1:
            while int(diff*100) % int(list_multiples[n]*100) != 0:
                n += 1
                min_limit = round_to_multiple(min_limit, 0.5, where='down')
                max_limit = round_to_multiple(max_limit, 0.5, where='up')
        elif diff <= 4.5:
            while int(diff*100) % int(list_multiples[n]*100) != 0:
                n += 1
                min_limit = round_to_multiple(min_limit, 1, where='down')
                max_limit = round_to_multiple(max_limit, 1, where='up')
                diff = max_limit - min_limit
        elif diff <= 15:
            while int(diff*100) % int(list_multiples[n]*100) != 0:
                n += 1
                min_limit = round_to_multiple(min_limit,2.5, where='down')
                max_limit = round_to_multiple(max_limit, 2.5, where='up')
                diff = max_limit - min_limit
        elif diff <= 30:
            while int(diff*100) % int(list_multiples[n]*100) != 0:
                n += 1
                min_limit = round_to_multiple(min_limit, 5, where='down')
                max_limit = round_to_multiple(max_limit, 5, where='up')
                diff = max_limit - min_limit

        elif diff <= 100:
            while int(diff*100) % int(list_multiples[n]*100) != 0:
                n += 1
                min_limit = round_to_multiple(min_limit, 25, where='down')
                max_limit = round_to_multiple(max_limit, 25, where='up')
                diff = max_limit - min_limit

        max_limit = max_limit*int("1"+"0"*zeros)      
        min_limit = min_limit*int("1"+"0"*zeros)
    else:
        diff = abs(max_limit) - abs(min_limit)
        list_multiples = [20, 10, 4, 2, 1, 0.5]
        n = 0
        if diff <= 1:
            while int(diff*100) % int(list_multiples[n]*100) != 0:
                n += 1
                min_limit = round_to_multiple(min_limit, 0.4, where='down')
                max_limit = round_to_multiple(max_limit, 0.4, where='up')
                diff=max_limit - min_limit
        elif diff <= 2.5:
            while int(diff*100) % int(list_multiples[n]*100) != 0:
                n += 1
                min_limit = round_to_multiple(min_limit, 2, where='down')
                max_limit = round_to_multiple(max_limit, 2, where='up')
                diff = max_limit - min_limit

        elif diff <= 20:
            while int(diff*100) % int(list_multiples[n]*100) != 0:
                n += 1
                min_limit = round_to_multiple(min_limit,2, where='down')
                max_limit = round_to_multiple(max_limit, 2, where='up')
                diff = max_limit - min_limit

        elif diff <= 100:
            while int(diff*100) % int(list_multiples[n]*100) != 0:
                n += 1
                min_limit = round_to_multiple(min_limit, 10, where='down')
                max_limit = round_to_multiple(max_limit, 10, where='up')
                diff = max_limit - min_limit

        max_limit = max_limit*int("1"+"0"*zeros)      
        min_limit = min_limit*int("1"+"0"*zeros)        
    return min_limit, max_limit



def adjust_zero_intervals(min_limit, max_limit, ticks):
    '''
    This functions is created in order to create limits in the axis that matches the 0 line. This way we can create a good
    and functional graph that has negative and positive values and both cross the 0 tick.
    
    :param min_limit: minimum for the axis, already adjusted previously. 
    :param max_limit: maximum for the axis, already adjusted previously. 
    :param ticks: Number of gridlines that the graph will contain

    :return: returns a new maximum and minimum that divided by the number of tickers, will have the 0 in it.
    '''
    ticks = int(ticks)
    #It will be used a loop in which, until 0 appears in the tickers, will keep working. This will be divided if the min is higher than the max or the opposite.
    if 0 in np.linspace(min_limit, max_limit, num=ticks):
        return min_limit, max_limit, ticks
    elif 0 in np.linspace(min_limit, max_limit, num=ticks-1):
        ticks = ticks -1 
        return min_limit, max_limit, ticks
    
    while 0 not in np.linspace(min_limit, max_limit, num=ticks):
        ##The logic between this method is to try to add to the smaller side values, so this way we avoid creating blanks spaces
        #If the minimum value is lowe or equal it will follow this logic:
        if (abs(min_limit) <= 12.5) & (abs(max_limit) <= 12.5):
            value_to_sum = 0.5
            min_limit = floor(min_limit*2)/2
            max_limit = ceil(max_limit*2)/2            
        else:
            min_limit = floor(min_limit)
            max_limit = ceil(max_limit)
            value_to_sum = 1
        
        if abs(min_limit) <= max_limit:
            new_max_limit=max_limit
            while abs(new_max_limit) <= max_limit*1.3:
                new_max_limit= round(new_max_limit + value_to_sum,1)
                if 0 in np.linspace(min_limit, new_max_limit, num=ticks):
                    return min_limit, new_max_limit, ticks
            for i in np.linspace(0, 1, num=11):    
                new_max_limit=max_limit
                new_min_limit=min_limit
                while abs(new_min_limit) <= abs(max_limit)*i: 
                    new_min_limit= round(new_min_limit - value_to_sum,1)
                    if abs(new_min_limit) == new_max_limit:
                        ticks = 5
                        return new_min_limit, new_max_limit, ticks
                    if 0 in np.linspace(new_min_limit, max_limit, num=ticks):
                        return new_min_limit, max_limit, ticks
                    new_max_limit = max_limit
                    while abs(new_max_limit) <= max_limit*1.3:
                        if 0 in np.linspace(new_min_limit, new_max_limit, num=ticks):
                            return new_min_limit, new_max_limit, ticks                    
                        new_max_limit= round(new_max_limit + value_to_sum,1)
                        if 0 in np.linspace(new_min_limit, new_max_limit, num=ticks):
                            return new_min_limit, new_max_limit, ticks
            min_limit=round(min_limit - value_to_sum,1)                    
        #This will happen exactly in the opposite way when the minimum is higher than the maximum
        elif abs(min_limit) > max_limit: 
            new_min_limit=min_limit
            while abs(new_min_limit) <= abs(min_limit)*1.3:
                new_min_limit= round(new_min_limit - value_to_sum,1)
                if 0 in np.linspace(new_min_limit, max_limit, num=ticks):
                    return new_min_limit, max_limit, ticks
            for i in np.linspace(0, 1, num=11):    
                new_max_limit=max_limit
                new_min_limit=min_limit
                while abs(new_max_limit) <= abs(min_limit)*i: 
                    new_max_limit= round(new_max_limit + value_to_sum,1)
                    if abs(new_min_limit) == new_max_limit:
                        ticks = 5
                        return new_min_limit, new_max_limit, ticks
                    if 0 in np.linspace(min_limit, new_max_limit, num=ticks):
                        return min_limit, new_max_limit, ticks
                    new_min_limit = min_limit
                    while abs(new_min_limit) <= abs(min_limit)*1.3:
                        if 0 in np.linspace(new_min_limit, new_max_limit, num=ticks):
                            return new_min_limit, new_max_limit, ticks                    
                        new_min_limit= round(new_min_limit - value_to_sum,1)
                        if 0 in np.linspace(new_min_limit, new_max_limit, num=ticks):
                            return new_min_limit, new_max_limit, ticks    
            max_limit=round(max_limit + value_to_sum,1)                
    #If there is a 0 in the first loop and there's need of no conversion it will end here returning the values
    return min_limit, max_limit, ticks


def change_names_for_graph(df, column_name):
    '''
    This will contain the names that will be used in the graphs. Since we have to rename this after taking from SQL,
    and avoiding to use mapping in Excel, here everything will be contained.
    
    :param df: dataframe 
    :param column_name: column from that dataframe where to replace.

    :return: returns a Pandas Series to add to a dataframe.
    '''
    #Dictionary of Country SQL : country graph
    entities_dict =  {'...'}
    #If it's being used in the Column 'Countries' it will be used in the list within the series.
    if column_name == 'Countries':
        for entity0, entity1 in entities_dict.items():
            df[column_name] = df[column_name].apply(lambda x: x.replace(entity0, entity1))
    #Otherwise, it will use directly into the series.
    else:
        df = df.replace({column_name: entities_dict})
    
    return df


def load_presets_for_regions(sheet_name:str, region:list, code_of_region:'str'):
    #Path where the masterfile is contained
    path_excel = r'C:\...\Regional_graphs_master.xlsx'
    #Read excel
    df_master = pd.read_excel(path_excel, sheet_name=sheet_name)
    #Use only those sleceted before.
    if region==None:
        pass
    else:
        if len(code_of_region)==1:
            df_master = df_master[df_master['...'].isin(region)].reset_index(drop=True)
        else: 
            df_master = df_master[df_master['...'].isin(region)].reset_index(drop=True)
    #This will prepared the list of countries to be used inside the SQL query IN clause.
    countries = df_master['...'].str.split(", ", expand=True).stack().str.strip("[]'").unique().tolist()
    countries = str(countries).replace("[","").replace("]","")+ ", '...'"
    #This will prepared the list of indicators to be used inside the SQL query IN clause.
    indicators = str(list(df_master["..."].dropna().unique()) + list(df_master["..."].dropna().unique()) + list(df_master["..."].dropna().unique()))
    indicators = str(indicators).replace("[","").replace("]","")
    
    return df_master, countries, indicators


def current_report_date(cnxn):
    ##Quert in SQL to download the WeeklyPublication Table
    query="""query1"""
    ##Reading the query and converting to a dataframe, and dropping the 'Period' column
    week_publi = pd.read_sql_query(query, cnxn).drop(columns="Period")
    week_publi = week_publi[(week_publi["..."] <= 3550)]
    ##Since the table has the value in wich the week ends, by substracting 6 days, we will have the day the publication week starts
    week_publi['start'] = week_publi['RealDate'] - timedelta(days=6)
    ##Creating a column with the day of today
    week_publi['today'] = datetime.today().strftime('%Y-%m-%d')
    week_number = week_publi.loc[week_publi['today'].between(week_publi['start'], week_publi['RealDate']), ['...']].values.item()
    query = 'query2'
    reportdate = pd.read_sql_query(query, cnxn)
    reportdate = reportdate.loc[reportdate['...'] == week_number]['...'].iloc[0].strftime('%Y-%m-%d')
    return reportdate


def get_df(df_master, countries, indicators, cnxn):
    publi_report_date = "'"+str(current_report_date(cnxn))+"'"
    #Getting historical values
    query_annual_h = f"""
    SELECT distinct ...
    FROM ...

    INNER JOIN ... e ON e....=t....
    INNER JOIN ... s ON s....=t....
    INNER JOIN .... i ON i....=t....

    WHERE ((... = '1970-01-01' AND ...=1) or (... = {publi_report_date} AND ...=2)) AND 
        ... = 1 AND 
        YEAR(period) BETWEEN year(DATEADD(YEAR, -3, getdate())) AND year(DATEADD(YEAR, 1, getdate())) AND
        ... IN ({countries}) AND
        ... IN ({indicators})                   

    ORDER BY ..., ..., ...
    """ 
    df_h = pd.read_sql_query(query_annual_h, cnxn)

    #Getting forecasts values FROM tmp_merged
    query_annual_f = f"""
    SELECT distinct ...
    FROM ... t
    INNER JOIN ... e ON e....=t....
    INNER JOIN ... s ON s....=t....
    INNER JOIN ... i ON i....=t....

    WHERE 
        YEAR(period) BETWEEN year(DATEADD(YEAR, -1, getdate())) AND year(DATEADD(YEAR, 2, getdate()))  AND
        ... =1 AND
        ... = 2 AND
        ReportDate = {publi_report_date} AND
        ... IN ({countries}) AND
        ... IN ({indicators})                   

    ORDER BY ..., ..., ... DESC
    """
    df_f = pd.read_sql_query(query_annual_f, cnxn)
    max_report_date = max(df_f['...'])

    euro_q_tmp = f"""
    SELECT distinct ..., ..., ..., ..., ..., ...
    FROM ... t
    INNER JOIN ... e ON e....=t....
    INNER JOIN ... s ON s....=t....
    INNER JOIN ... i ON i....=t....
    WHERE 
        YEAR(period) BETWEEN year(DATEADD(YEAR, -2, getdate())) AND  year(DATEADD(YEAR, 1, getdate())) AND
        ... =1 AND
        ... = 2 AND
        ... = {publi_report_date} AND
        ... IN ('Euro Area') AND
        ... IN ({indicators})                   
    ORDER BY ..., ..., ... DESC, ...
    """
    euro_tmp = pd.read_sql_query(euro_q_tmp, cnxn)

    #Getting forecasts values from regional aggregates
    query_annual_agg = f"""
    SELECT distinct ...
    FROM ...
    WHERE 
        YEAR(period) BETWEEN year(DATEADD(YEAR, -3, getdate())) AND year(DATEADD(YEAR, 2, getdate()))  AND
        ... =1 AND
        ... = {publi_report_date} AND
        ... IN ({countries}) AND
        ... IN ({indicators})                   
    ORDER BY ..., ..., ... DESC
    """
    df_agg = pd.read_sql_query(query_annual_agg, cnxn)
    
    query_annual_rateten = f"""
    SELECT distinct ...
    FROM ...
    WHERE 
        YEAR(period) BETWEEN year(DATEADD(YEAR, -2, getdate())) AND year(DATEADD(YEAR, 2, getdate()))  AND
        ... =1 AND
        ... = {publi_report_date} AND
         ...='rateten' and ..='eurozone'
    ORDER BY ..., ..., ... DESC
    """
    df_rateten = pd.read_sql_query(query_annual_rateten, cnxn)
    df_rateten['...'] = 'Euro Area'
    
    #Concat all the dataframes
    df = pd.concat([euro_tmp, df_agg, df_h, df_f,  df_rateten]).reset_index(drop=True)
    #This will drop all rows repeated just differencing them by reportdate and update date.
    df = df.drop_duplicates(['...'], keep='first').drop(columns=['...']).dropna(subset=['Value']).sort_values(by=["..."]).reset_index(drop=True)
    #MANUAL CALCULATIONS
    #Greece RATEPOL == Euro Area
    rates_cyprus = df[(df['...'] == 'Euro Area') & (df['...']=='RATEPOL')]
    rates_cyprus['...'] = 'Cyprus'
    df = pd.concat([df, rates_cyprus]).reset_index(drop=True)
    #Take only 2 last years as history if its november or december
    if int(max_report_date.month) >= 11:
        minyear = pd.to_datetime(max_report_date.year-1, format='%Y') 
        maxyear = pd.to_datetime(max_report_date.year+2, format='%Y') 
        df = df[(df['Period']>=minyear) & (df['Period']<=maxyear)]
    else:
        minyear = pd.to_datetime(max_report_date.year-2, format='%Y') 
        maxyear = pd.to_datetime(max_report_date.year+1, format='%Y') 
        df = df[(df['Period']>=minyear) & (df['Period']<=maxyear)]    
    #Convert date to YY format
    df['Period'] = df['Period'].dt.strftime('%Y')
    df = df.replace('...', '...')
    # Changing names for the master file
    df_master = change_names_for_graph(df_master, '...')
    # Changing names for the data table
    df = change_names_for_graph(df, '...')
    df = df[~((df['...'] == 'Norway') & (df['...'] =='CURUSDGDP'))].reset_index(drop = True)
    df = df[~((df['...'] == 'Sweden') & (df['...'] =='CURUSDGDP'))].reset_index(drop = True)
    df = df.replace([np.inf, -np.inf], np.nan)
    return df_master, df


def settings_for_plots():
    #Font type
    mpl.rcParams['font.family'] = 'sans-serif'
    mpl.rcParams['font.sans-serif'] = 'Arial'
    mpl.rcParams['font.weight'] = 400
    mpl.rcParams['font.stretch'] = 'normal'
    #Font size
    mpl.rcParams['font.size'] = 19
    #Box container colour and width 
    mpl.rc('axes', axisbelow=True, linewidth=1)
    ##Legend
    #Legend squares style
    mpl.rcParams['legend.handlelength'] = 1
    mpl.rcParams['legend.handleheight'] = 1.125    
    mpl.rcParams['legend.fontsize'] = 16
    # the vertical space between the legend entries
    mpl.rcParams["legend.labelspacing"]=  0.6 # 
    # the space between the legend line and legend text
    mpl.rcParams["legend.handletextpad"]= 0.6  # 
    # the border between the axes and legend box
    mpl.rcParams["legend.borderaxespad"]= 1.4  # 
    # column separation between labels
    mpl.rcParams["legend.columnspacing"]= 1.2 # 
    # legend box size
    mpl.rcParams["legend.borderpad"]= 1 
    #Legend edge color
    mpl.rcParams["legend.edgecolor"]= 'white' 


def get_axis_limits(df_i, percentage_up_to_avoid_overlapping, ticks, no_reescale):
    limits = get_rounded_limits(df_i) 

    if limits[0] <= 0 and limits[1] <= 0:
        limits[0] = limits[0] * percentage_up_to_avoid_overlapping
    else:
        limits[1] = limits[1] * percentage_up_to_avoid_overlapping

    #This will round to a multiple making it more stetical
    min_limit, max_limit = adjusting_limits(limits, solo=False, ticks=ticks)
    #After having a proper limits, if there are a graph that contains negative and positive values to make it pass in the 0 tick.
    min_limit, max_limit, ticks = adjust_zero_intervals(min_limit, max_limit, ticks)
    #This will create the ticks in the range of the limits
    y_tickers_col = np.linspace(min_limit, max_limit, num=ticks)
    ### Rescale
    rescale=False
    if (df_i.min().min() < 0) & (df_i.max().max() < 0):
        if df_i.max().max() < y_tickers_col[-2]:
            rescale = True
    elif (df_i.min().min() > 0) & (df_i.max().max() > 0):
        if df_i.min().min() > y_tickers_col[1]:
            rescale = True
    try:
        if no_reescale == True:
            rescale=False
    except:
        pass
    if rescale == True:
        if (df_i.min().min() <= 0) & (df_i.max().max() < 0):
            max_limit, min_limit= rescale_intervals(df_i.max().max()*0.975, df_i.min().min()*1.025, ticks)
            y_tickers_col = np.linspace(min_limit, max_limit, num=ticks)  
        elif (df_i.min().min() > 0) & (df_i.max().max() > 0):
            min_limit, max_limit = rescale_intervals(df_i.min().min()*0.975, df_i.max().max()*1.025, ticks)
            y_tickers_col = np.linspace(min_limit, max_limit, num=ticks)
    ##This checks if there are empty spaces in the graphs and retires them.
    #If value max is minor than the 2nd last axis it will remove the major tick.
    if df_i.max().max() > 0:
        while (df_i.max().max() < y_tickers_col[-2]) & (len( y_tickers_col) >5):
            y_tickers_col = y_tickers_col[:-1]
            ticks -=1
    if df_i.min().min() < 0:
        while (abs(df_i.min().min()) < abs(y_tickers_col[1])) & (len( y_tickers_col) >5):
            y_tickers_col = y_tickers_col[1:]         
            ticks -=1

    if (rescale == True) &((df_i.min().min() <= 0) & (df_i.max().max() < 0)) & (abs(df_i.max().max()*0.975) > abs(y_tickers_col[-2])) & (ticks == 6):
        y_tickers_col = y_tickers_col[:-1]
        ticks=5
    elif (rescale == True) &((df_i.min().min() > 0) & (df_i.max().max() > 0)) & (df_i.min().min()*0.975 > y_tickers_col[1]) & (ticks == 6):
        y_tickers_col = y_tickers_col[1:]
        ticks=5 

        
    return y_tickers_col, rescale, ticks, max_limit


def generate_the_plots(df, df_master):
    mpl.use('SVG')
    #Loop in every row of the df_master, so every graph turned 'On' will be processed
    count = 0
    #Colours codes
    gray_background = '#D9D9D9'
    for i in range(len(df_master)):
        count+=1
        overlapping = True
        percentage_up_to_avoid_overlapping = 1
        ticks = 6
        no_reescale = False
        ##Prepare data for graph
        df_master_i = df_master.iloc[i]
        #Extract countries from rows
        countries = df_master_i['...'].split(", ")
        #Select indicators and countries for that chart, creating a temporal df_i (This is made like this in order to keep
        #an index that separates the first usable from the others, so if it has one )
        #Extract indicators
        indicator1 = df_master_i['...']
        indicator2 = df_master_i['...']
        indicator3 = df_master_i['...']
        #Search for that indicator
        df_i_1 = df[df['...'].isin([indicator1]) & df['...'].isin(countries)]
        df_i_2 = df[df['...'].isin([indicator2]) & df['...'].isin(countries)]
        df_i_3 = df[df['...'].isin([indicator3]) & df['...'].isin(countries)]
        df_i = pd.concat([df_i_1, df_i_2, df_i_3]).reset_index(drop=True)
        df_i = df_i.drop_duplicates(['...','...'], keep='first')
        #extraction of names in list (this will keep the order)
        columns_sort = [x for x in countries if x in df_i['...'].unique()]
        cat_order = CategoricalDtype(columns_sort, ordered=True)
        #Conversion of the column into categorical (using previous features as type)
        df_i['...'] = df_i['...'].astype(cat_order)
        #Pivot table
        df_i = pd.pivot_table(df_i, values='Value', index=['...'], columns=['...'])
        #Drop countries that don't have forecasts
        df_i =df_i.dropna(subset=df_i.columns.to_list()[2:], how='all')  
        ##Prepare the limits for axis
        #Tickers
        #This will round to an integer the max and min value
        while overlapping == True:
            print('overlap')
            y_tickers_col, rescale, ticks, max_limit = get_axis_limits(df_i, percentage_up_to_avoid_overlapping, ticks, no_reescale)
            ###Starting the plot 
            fig, ax = plt.subplots(figsize = (13.46, 8.296))
            #Determine width of bars
            number_of_years= len(list(df_i.columns))+0.8 #+0.8 means a ticker before and after this gives a perfect fit
            barWidth = 1/number_of_years
            #Set position of bar on X axis
            r1 = np.arange(len(df_i.iloc[:,0]))
            r2 = [x + barWidth for x in r1]
            r3 = [x + barWidth for x in r2]
            r4 = [x + barWidth for x in r3]
            # Make the bar plot. This is done in a try-except way because for some countries they may not be the whole amount of years, and it would raise an error.
            try:
                plt.bar(r1, df_i.iloc[:,0], color='#CCDFE5', width=barWidth, label=df_i.iloc[:,0].name)
                plt.bar(r2, df_i.iloc[:,1], color='#729FA2', width=barWidth, label=df_i.iloc[:,1].name)
                plt.bar(r3, df_i.iloc[:,2], color='#3B7394', width=barWidth, label=df_i.iloc[:,2].name)
                plt.bar(r4, df_i.iloc[:,3], color='#164B75', width=barWidth, label=df_i.iloc[:,3].name)
            except:
                pass
            #This will change the dot to a comma for thousands values
            if max_limit >= 1000: 
                y_tickers_col = [int(i) for i in y_tickers_col]        
                ax.get_yaxis().set_major_formatter(mpl.ticker.FuncFormatter(lambda x, p: format(x, ',')))
            #Add xticks on the middle of the group bars (So the name is centered)
            plt.xticks([r + barWidth*1.5 for r in range(len(df_i.iloc[:,0]))], df_i.index)
            #Set the numeric values limits of the graph
            plt.ylim([y_tickers_col[0], y_tickers_col[-1]])
            #Set the tickers in values    
            ax.set_yticks(y_tickers_col)
            #Margins betweenbox and axis?
            plt.tick_params(axis='x', which='major')
            #Erase the ticks in the bottom side
            plt.tick_params(bottom = False)
            #Colours and position of the gridlines.    
            plt.grid(axis='y', color = gray_background, linewidth=1)
            #Remove little lines in ticks
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            #Setting one decimal for X axis
            if df_master_i["Type"] == '%':
                ax.yaxis.set_major_formatter(FormatStrFormatter('%.1f'))
            #Setting for invisible legend box
            bars, labels = ax.get_legend_handles_labels()
            legend = plt.legend(bars, labels, ncol=4, markerscale=0.75)
            plt.draw()
            plt.gcf().canvas.draw()
            if check_if_serie_overlaps(legend, ax) == True:
                if ticks == 6:
                    ticks = 5 
                elif ticks == 5 :
                    ticks = 6
                    percentage_up_to_avoid_overlapping +=0.2
                no_reescale = True
            else:
                overlapping = False
        #Path where to save the graphs
        path_to_save = 'C:\...\\' + df_master_i["..."] + "\\" + df_master_i["..."] + '.svg'
        #Save graph in an tight form (this will keep the dimension)
        plt.savefig(path_to_save, bbox_inches="tight")     
        plt.close(fig)
        plt.close('all')
        print('printed bar graph!')
    return count


def export_images(region, code_of_region):    
    start_time = time.time()
    cnxn = create_connection()
    df_master, countries, indicators = load_presets_for_regions(sheet_name = 'CFS_Bar',  region=region, code_of_region=code_of_region)
    df_master, df = get_df(df_master, countries=countries, indicators=indicators, cnxn=cnxn)
    settings_for_plots()
    count = generate_the_plots(df, df_master)    
    end_time = time.time()
    print("Printed", str(count), "charts in",int(end_time - start_time),"seconds, this means an average of", "{0:.2f}".format((end_time - start_time)/count), "seconds per chart")

