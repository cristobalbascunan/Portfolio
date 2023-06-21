# OneHub

We use the API in order to interact with OneHub. It's necessary some credentials that come from the developer account settings to get the access token (Oauth2) that will allow 
us the autorization to proceed: username, password, client id and client secret.


After having this token, we will use some of the request methods: **POST**, **GET** and **DELETE**.


## Upload Client files & PDFs (and Excel) files

First of all, the code searches the files created in the previous process (in case of the PDFs, when their are ready by the production unit. 
After getting a match of which files/PDFs between the names and current date, it proceeds to upload them to the proper place (ID url in Onehub)

In the case of PDF's reports, it deletes the file name that contains the previous year (i.e.:  if the current PDF uploaded is *file - May 2023.pdf*, the one that will be deleted is *file - May 2022.pdf*) 
