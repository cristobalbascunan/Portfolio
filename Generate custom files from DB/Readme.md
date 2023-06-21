# Files clients

Here there are the code needed to run the scripts in the different processes for each week.

## Generate Client Files for regions or commodities.
This code produces different ouputs (format .txt, .xlsx or .csv) depending on the 'current publication region' and the reportdate.
There is a 'master' file were all the specifications are already written and where new clients can be added. `C:\...\Excel_with_specifications.xlsx`
The process is to read the queries in `C:\...\Queries` and then create a file.


## Client to WinSCP
The code includes 3 separate things:

0. Code in order to extract the information from old Excel templates used to generate the files. Now this is commented because is not necessary but not deleted just in case is 
needed in this process or in other. The extracted data is now stored on `C:\...\Excel_with_specifications.xlsx`
and if there's some need to change something, that's the place.

1. Create the output in .csv format.

2. Execution of `pywinauto` that will control WinSCP in order to send instructions to upload the previous generated files and when finished, close it.
