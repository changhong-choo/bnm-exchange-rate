
# 1. API call to extract exchange rate

def extract_exchange_rate():
    
    import requests
    import pandas as pd
    
    headers = {'Accept' : 'application/vnd.BNM.API.v1+json'}
    parameters = {'session' : '1700',
              'quote' : 'rm'}
    
    dfs = []
    
    global year 
    year = input('Enter year: ')
    global month 
    month = input('Enter month: ')    
     
    urls = [f'https://api.bnm.gov.my/public/exchange-rate/SGD/year/{year}/month/{month}',
            f'https://api.bnm.gov.my/public/exchange-rate/THB/year/{year}/month/{month}',
            f'https://api.bnm.gov.my/public/exchange-rate/PHP/year/{year}/month/{month}',
            f'https://api.bnm.gov.my/public/exchange-rate/GBP/year/{year}/month/{month}',
            f'https://api.bnm.gov.my/public/exchange-rate/USD/year/{year}/month/{month}',
            f'https://api.bnm.gov.my/public/exchange-rate/AUD/year/{year}/month/{month}',
            f'https://api.bnm.gov.my/public/exchange-rate/IDR/year/{year}/month/{month}']

    for url in urls:
        
        response = requests.get(url, headers = headers, params = parameters)
        # print(response.url) 
        
        data = response.json()
        data2 = data['data']['rate'][-1]
        data2['currency_code'] = data['data']['currency_code']
        dfs.append(data2)
        
    df = pd.DataFrame(dfs)
    
    file_path = ...
    
    df.to_csv(f'{file_path}\\bnm-exchange-rate-{year}{month}.csv', index=False)
    
   
    
# ---------------------------------------------------------------------

# 2. Refresh data in excel working file

excel_path = ...


def excel_refresh_data():
    import win32com.client
    import os
    import time
    import pythoncom
    
    
    # Refresh data in Excel
    xl = win32com.client.Dispatch("Excel.Application",pythoncom.CoInitialize())
    # xl = win32com.client.DispatchEx("Excel.Application")
    wb = xl.workbooks.open(excel_path)
    xl.Visible = True
    wb.RefreshAll()
    time.sleep(3)
    wb.Close(True)
    xl.Quit()
    
    os.system("taskkill /f /im excel.exe")

# ------------------------------------------------------------------

# 3. Export excel data as image & Send the content by gmail
image_path = ...
sheet_name = ...
copy_range = ...

sender = ...
receivers = [..., ...]
sender_pw = ...


def send_email():

    from email.mime.text import MIMEText
    from email.mime.image import MIMEImage
    from email.mime.multipart import MIMEMultipart
    import excel2img
    
    # Transform exchange rate data to image
    excel2img.export_img(excel_path, image_path, sheet_name, copy_range)
    
    msg = MIMEMultipart()
    msg['Subject'] = f'Forex rate {year}-{month}'
    msg['From'] = sender
    msg['To'] = ','.join(receivers)
    
    msgAlternative = MIMEMultipart('alternative')
    msg.attach(msgAlternative)
    
    # We reference the image in the IMG SRC attribute by the ID we give it below
    msgText = MIMEText('<p>Dear all, <br><br>Below forex rate for %s-%s.</p><br><img src="cid:image1"><br>' % (year, month), 'html')
    msgAlternative.attach(msgText)
    
    # This example assumes the image is in the current directory
    fp = open(image_path, 'rb')
    msgImage = MIMEImage(fp.read())
    fp.close()
    
    # Define the image's ID as referenced above
    msgImage.add_header('Content-ID', '<image1>')
    msg.attach(msgImage)
    
    # Send the email
    import smtplib
    s = smtplib.SMTP_SSL(host = 'smtp.gmail.com', port = 465)
    s.login(user = sender, password = sender_pw)
    s.sendmail(sender, receivers, msg.as_string())
    s.quit()


# Call function to execute
extract_exchange_rate()
excel_refresh_data()
send_email()











