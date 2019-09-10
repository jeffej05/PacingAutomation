## Revenue-at-Risk Report Automation


This was my first big automation project with python. The Revenue-at-Risk Report tracks all of IBM Watson Advertising's current campaign performances in terms of delivery and percentage pacing towards our overall booked revenue goal. I've broken down the features of this report below and any thoughts on how it can be improved for the future. Please note that all confidential company and personal information has been taken out of the script.

1.) Data Import- We are currently manually downloading the raw data CSV file from Qlik Sense but for the future, the plan is to either set up a direct connection to the database using pyodbc or utilize Selenium to automatically log into Qlik Sense to retrieve the csv data file for us. This way, we will have an automated report pulling process for this report.

2.) Data Manipulation- The program will automatically clean up the datasets used for the report.

3.) Calculations- The report tracks week-over-week improvements so the program will automatically import the previous report from our report directory and calculate week-over-week changes in revenue and delivery risk. 

4.) Excel Export- The weekly report features pivot tables and raw data sets that is sent as an excel file to account managers and senior leadership. The program will automatically create these excel files for us. In the future I would love to include more graphs to visualize under-delivery via matplotlib or seaborn. 

5.) Word Export- The weekly report features an email that lists all the current accounts that show under-delivery risk as well as updated notes sent by account executive on what is being done to mitigate this risk and if things have improved. The program will automatically write out this email into a word doc so it can be reviewed by the report owner just in case the owner wants to add any additional comments.

6.) Google Sheets Export- The program will push the "late creative" pivot in Google Sheets via gspread so Account Executives can add in their weekly notes and comments regarding the status of their accounts.

6.) Email Automation- The program will log into the users' email account, create an email subject and body, and attach the excel and word documents that were previously created, and will send it to all relevant recipients.

6.) Cron Job- The report is sent out every Monday and Thursday at around 11am. A cron job was set up in the terminal to mimic this execution schedule.
