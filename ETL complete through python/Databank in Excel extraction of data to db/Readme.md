# Client
Client provides their data in a .xlsx format. Each of the countries are separated in different sheets that should be processed one by one. To download it the source is in 
(...). 

Each region has two excels (Annual and Qt) where all the mapping is done . Here it's easy to add, modify or remove indicators `C:\***\folder`

The columns to modify in order to get the desired output are:

| Entity | Entity_Client | FE_Ticker | Client_Ticker | Calculation |
| :------------------ |:-----------------| :-----------------| :-----------------| :-----------------|
| Name in DB | Name in Client <br/>sheet | Name in DB | Column code <br/>in Excel | Specific calculation that <br/> is developed in the code|

There are different *bats* to upload the different regions, and one that run the whole process for all regions. 
