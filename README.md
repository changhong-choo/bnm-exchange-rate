# bnm-exchange-rate
To extract exchange rates from [Bank Negara Malaysia](https://www.bnm.gov.my/) and send the content by email.

# Description
This project is aim to automate the process of extracting monthly closing exchange rates and send to receipients by email. This project covers SGD, THB, PHP, GBP, USD, AUD, and IDR.  
The overall process is as follow:  
1. API call to extract exchange rates from [Bank Negara Malaysia OpenAPI](https://apikijangportal.bnm.gov.my/) and save in excel file.
2. Excel Power Query transform the data into the required format.
3. Send the content to receipients by email.

# Steps to complete before run the program
1. Input the necessary value for the following variables in the python file
- file_path -> path to save the exchange rate data
- excel_path -> file_path + excel_name + excel_type (xlsx)
- image_path -> path + image_name + image_type (png)
- sheet_name -> excel sheet name
- copy_range -> excel data range to be copy
- sender -> sender's email
- sender_pw -> sender's email password
- receivers -> one or more receivers email
