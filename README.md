# Portfolio
For any question you can contact me on my LikedIn profile. https://www.linkedin.com/in/cristobal-bl/

## Project samples

**API POST, GET, DELETE connection**: this project involves connecting to OneHub through the API using OAuth2 to obtain the authentication token. The token is then incorporated into an existing process that utilizes POST requests to upload files, GET requests to check if an older file exists, and DELETE requests to remove files.
-    *Libraries used: requests, sys, json, pandas, os, numpy, pyodbc, datetime, re.*

**Automailing from SalesForce**: This project involves establishing an API connection with the SalesForce platform to extract data in JSON format. The extracted data is then used as recipients for an email sending process in Python. This allows for sending customized emails without relying on a third-party company system. 
-    *Libraries used: requests, sys, json, pandas, os, numpy, re, simply_salesforce, time, base64, smtplib, pathlib, email.*

**ETL complete through python**: This project focuses on automatically searching for the latest version of an target Excel file composed multiple tabs. The information is extracted and transformed as necessary to connect with the database and upload/add only the values that have changed compared to the existing database.
-    *Libraries used: sys, os, pandas, numpy, pyodbc, fnmatch, sys, errno, re, datetime.*

**Generate custom files from DB**: each client asked for different formats to be delivered the data: .txt, .csv or .xlsx. The first two only depend on SQL queries, but to customize the output in Excel I used xlsxwriter and xlwings in order to apply exact formats through VBA code resulting in a custom process totally automatized. Additionally, for delivering files through Windows applications that lack an API, the project utilizes pywinauto to control the interface. This includes opening a program, logging in, uploading the file, and then closing it.
-    *Libraries used: os, pandas, numpy, errno, shutil, xlsxwriter, time, pyodbc, datetime, xlwings, sys, pywinauto.*

**Graphs automatically fitted to content**: This project automates the process of creating graphs using matplotlib. The goal is to ensure that the graphs are perfectly fitted to display the correct economic scales for comparison. This involves defining a set of rules that can be customized using Python. The rules ensure that the graphs always display the same number of ticks, cross 0 if there are positive-negative values, use integer values, and leave enough space to avoid overlapping the graph's limits. This enables the generation of hundreds of perfect graphs within seconds.
-    *Libraries used: sys, numpy, matplotlib, pandas, math, time, datetime*


