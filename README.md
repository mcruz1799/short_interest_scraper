# Short Interest & Financial Data - Web Scraping Project

I was tasked by my employer to create a web scraping application that gathers data for the performance of two stocks in the comapany's portfolio. The application had to do the following: 
1) Scrape short interest data from Fintel.io
2) Scrape financial data from Yahoo finance
3) Send three separate automated emails daily containing the scraped data (one email for one stock, one for the other, one for both)
4) Add the scraped data to an Excel sheet that is stored in Dropbox

Therefore to complete this project, I had to develop 3 core features.

1) Web Scraping - Scrape the data from Fintel & Yahoo Finance
2) Automated Emails - Send the data to the specified recipients
3) Spreadsheet Updates - Update the Excel sheet stored in Dropbox 

I used various Python libraries to develop this application. 

For the web scraping portion of the project, I used Selenium to create an automated web browser, I used bs4 (Beautiful Soup) to parse the HTML, and finally pandas to clean and manipulate the scraped data. 

For the email portion of the project, I used the SMTPLib, SSL, and Email.MIME libraries to send secure automated emails through Gmail.

For the spreadsheet portion of the project, I used the Dropbox SDK to access and save the file, then pandas and openpyxl to manipulate the data within the file. 

I then used Windows Task Scheduler to run this script daily. 

See "main.py" for the code. 

