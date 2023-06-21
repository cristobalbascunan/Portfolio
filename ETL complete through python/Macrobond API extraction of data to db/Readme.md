## Macrobond API gathering of data.py
Usee the Macrobond API in order to have the information contained there as a priority above what is in the webpage.

This API isn't used with a token but with the current connection made in Windows. Therefore, this connection will only work in windows1 user. Through some functions,
its possible to retrieve the series from MB and convert them into DataFrames. 


- Those series and countries to be collected are described in two 'mapping' excels in `C:\...\MB source`:
    * `Mapping1.xlsx` contains the short code of the countries, the transcription as they will be shown from MB and how they are called in our DB.
    * `Mapping2.xlsx` contains the indicator code and their name in the DB.

Then the code proceeds as in others databanks, cleaning of data, proper calculations, adding or reportdates and columns and uploads.

