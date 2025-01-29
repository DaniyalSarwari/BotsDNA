
### Server Setup Challenge

The requirement is to create a server by selecting the required option get from excel file and send detail over email.

This challenge is solved in three steps:
- Bot navigate to the BotsDNA Server Creation challenge website.
- Download required excel template file over HTTP (If downloaded file exist use that one)
- Loop over each record in excel file to extract data and based on that data perform selection over website to create server.
- Once the server is created. Get the required detail (**IP**, **Username**, **Password**) and send Email to the recipient mention in excel record.

HTML Log file also placed in **logs** folder to check the successful execution of the BOT.







