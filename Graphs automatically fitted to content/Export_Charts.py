#!/usr/bin/env python
# coding: utf-8



import pandas as pd
import sys
sys.path.append(r'C:\...\Regional_graphs')
import Generate_Arrow_text_graphs
import Generate_CF_charts
import Generate_CFS_Quarterly_charts
import Generate_Bar_charts
import Generate_H_Charts
import Generate_M_charts
import time
import os
import logging

# System call
os.system("")
# Class of different styles
class style():
    BLACK = '\033[30m'
    RED = '\033[31m'
    GREEN = '\033[32m'
    YELLOW = '\033[33m'
    BLUE = '\033[34m'
    MAGENTA = '\033[35m'
    CYAN = '\033[36m'
    WHITE = '\033[37m'
    UNDERLINE = '\033[4m'
    RESET = '\033[0m'


def get_query_from_file(path:str):
    # Open and read the file as a single buffer
    fd = open(path, 'r')
    sqlFile = fd.read()
    fd.close()
    query = sqlFile
    return query


dict_regions ={
    '0': None,
    '1': ['...', '...'],
        '1.1': ['...', '...'],
            '1.1.1': ['...'],
            '1.1.2': ['...'],
        '1.2': ['...'],
    '2': ['...'],
        '2.1': ['...'],
        '2.2': ['...'],    
        '2.3': ['...'],
    '3': ['...', '...'],
        '3.1': ['...', '...', '...'],
            '3.1.1': ['...'],
            '3.1.2': ['...'],        
            '3.1.3': ['...'],
        '3.2': ['...'],
    '4': ['...', '...'],
        '4.1': ['...'],
        '4.2': ['...']
    }
printable_dict_regions = str(dict_regions).replace('],', '\n').replace("'",'').replace("[",'').replace("]",'').replace("{",'').replace("}",'').replace('None', 'All regions')
print('Select the region to print:\n#############################################\n', printable_dict_regions,'\n#############################################')
code_of_region_to_print =  str(input())


while code_of_region_to_print not in dict_regions.keys():
    code_of_region_to_print = str(input(style.RED +'Insert a valid code number of a region/subregion please:'))
print(style.CYAN +'The code will print the region(s):', str(dict_regions[code_of_region_to_print]).replace("'",'').replace("[",'').replace("]",'')+style.RESET)
region = dict_regions[code_of_region_to_print]
# logging.info('The code will print the region(s):', str(dict_regions[code_of_region_to_print]).replace("'",'').replace("[",'').replace("]",''))

dict_graphs ={
    '0': 'All graphs',
    '1': 'Bar Charts',
    '2': 'H Charts',
    '3': 'Quarterly CFS Charts',
    '4': 'CF Charts and Arrow text charts',
    '5': 'M Charts and  Arrow text charts', 
    }
printable_dict_graphs = str(dict_graphs).replace(",", '\n').replace("'",'').replace("{",'').replace("}",'').replace('None', 'All graphs')
print('Select the graph(s) to print:\n------------------------------------------\n', printable_dict_graphs,'\n------------------------------------------')
code_of_graph_to_print =  str(input())
while code_of_graph_to_print not in dict_graphs.keys():
    code_of_graph_to_print = str(input(style.RED +'Insert a valid code number to get graphs please:'))
print(style.CYAN +'The code will print:', str(dict_graphs[code_of_graph_to_print]).replace("'",'').replace("[",'').replace("]",'')+style.RESET)


start_time = time.time()
weights_query = get_query_from_file(r"C:\...\(...).sql")
print("Recalculating weights for regions...")
FE_database.execute_query(weights_query)
regional_agg_query = get_query_from_file(r"C:\...\(...).sql")
print("Recalculating regional aggregates...")
FE_database.execute_query(regional_agg_query)
end_time = time.time()
print("The calculations has taken %s seconds." % int(end_time - start_time))


start_time = time.time()
if code_of_graph_to_print == '0':
    Generate_Bar_charts.export_images(region, code_of_region_to_print)
    Generate_H_Charts.export_images(region, code_of_region_to_print)
    Generate_CFS_Quarterly_charts.export_images(region, code_of_region_to_print)
    Generate_CF_charts.export_images(region, code_of_region_to_print)
    Generate_M_charts.export_images(region, code_of_region_to_print)
    Generate_Arrow_text_graphs.export_images(region, code_of_region_to_print)
elif code_of_graph_to_print == '1':
    Generate_Bar_charts.export_images(region, code_of_region_to_print)
elif code_of_graph_to_print == '2':
    Generate_H_Charts.export_images(region, code_of_region_to_print)
elif code_of_graph_to_print == '3':
    Generate_CFS_Quarterly_charts.export_images(region, code_of_region_to_print)
elif code_of_graph_to_print == '4':
    Generate_CF_charts.export_images(region, code_of_region_to_print)
    Generate_Arrow_text_graphs.export_images(region, code_of_region_to_print)
elif code_of_graph_to_print == '5':
    Generate_M_charts.export_images(region, code_of_region_to_print)
    Generate_Arrow_text_graphs.export_images(region, code_of_region_to_print)
    
end_time = time.time()
print(style.GREEN + "It has taken %s seconds to finish. Now you can close!" % int(end_time - start_time)+style.RESET)

