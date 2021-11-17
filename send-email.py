import smtplib
import os
import datetime as dt
import pandas as pd
import numpy as np
import logging
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Creating a log file #
logging.basicConfig(filename="send-email.log", level=logging.INFO, filemode='a', format='%(asctime)s %(levelname)s %(message)s', datefmt='%d-%b-%Y %H:%M:%S')
# with open('send-email.log', 'w'):
#     pass

# Extracting data from the Excel files #
shops = pd.read_excel('Shops.xlsx')
perf = pd.read_excel('WeeklyPerformance.xlsx')
perf = perf.replace(np.nan, '', regex=True) # Convert empty cells from Nan to empty string

#Taking login details from Environment variables
myEmail = os.environ.get('GMAIL_USER')
myEmailPass = os.environ.get('GMAIL_PASS')

# Time variables #
todayDate = dt.datetime.today()
toDate = todayDate - dt.timedelta(days=2)
fromDate = todayDate - dt.timedelta(days=9)
toDate= toDate.strftime("%d, %b %Y")
fromDate= fromDate.strftime("%d, %b %Y")
time = str(dt.datetime.now().strftime("%d-%b-%Y %H:%M "))

# First log of the run #
logging.info("")
logging.info("Process Initiated")

for x in shops.index:
    chainName = shops['Chain'][x]
    shopName= shops['Shop'][x]
    shopRecipients = shops['Recipients'][x]

    # Format common email part #
    msg = MIMEMultipart('alternative')
    msg['Subject'] = chainName +" - Weekly online sales performance"
    msg['From'] = myEmail
    msg['To'] = shopRecipients

    html1 = """\
            <!DOCTYPE html>
            <head>
                <style>
                    body{{
                        font-weight: normal;
                    }}
                    h2{{
                        font-family: Verdana;
                        font-weight: bold;
                        font-size: 18px;
                    }}
                    p{{
                        line-height: 18px;
                        font-family: Verdana;
                        font-size: 14px;
                    }}
                    table {{
                        margin-top: 20px;
                        margin-bottom: 20px;
                        width: 100%;
                        border-collapse: collapse;
                        font-size: 12px;
                        text-align: center;
                    }}
                    th {{
                        background-color: #FF80FF;
                        color: white;
                    }}
                    td, th {{
                        font-family: Verdana;
                        padding: 8px;
                        border: 1px solid gray;
                    }}
                </style>
            </head>
            <html>
                <body>
                    <h2>Dear Team,</h2>
                    <p>
                        Please find below the weekly online sales summary report for the period {FromDate} - {ToDate}.<br>
                    </p>
                    <table>
                        <tr>
                            <th style="text-align: left;">Shop Name</th>
                            <th>Orders Performed</th>
                            <th>Online Sales</th>
                            <th>Less Bank Variable</th>
                            <th>Add: Adjustments</th>
                            <th>Add: InstaPoints</th>
                            <th>Less: Commission</th>
                            <th>Final Amount Payable</th>
                        </tr>
                        <tr style="font-weight: bold; background-color: #FFFFCC">
                            <td colspan="2"></td>
                            <td>AED</td>
                            <td>AED</td>
                            <td>AED</td>
                            <td>AED</td>
                            <td>AED</td>
                            <td>AED</td>
                        </tr>
    """.format(FromDate=fromDate, ToDate=toDate)
    html = html1

    print(shopName)
    shopName = shopName.split(", ")

    for y in shopName:
        logging.info("Selected mainshop: "+y)

        for z in perf.index:
            perfShopName = perf['Shop Name'][z]
            logging.info("Looking for shop: "+perfShopName)

            if y == perfShopName:
                logging.info("Found shop")

                # Collecting data from Performance sheet #
                perfOrders = perf['Orders Performed'][z]
                perfOnline = perf['Online Sale'][z]
                perfBank = perf['Less Bank Variable'][z]
                perfAdjust = perf['Add Adjustments'][z]
                perfIP = perf['Add InstaPoints'][z]
                perfComm = perf['Less Commission'][z]
                perfAmount = perf['Final Amount'][z]

                if(perfAmount==0):
                    print('Zero')
                    payTerm="No transfer needed."
                elif(perfAmount>500):
                    print('Above 500')
                    payTerm="Please also note that the funds have been transfered to your account."
                elif(perfAmount<500):
                    print('Below 500')
                    payTerm="Please also note that the transfer will be made once amount exceeds 500DHS."

                # Adding table data for each shop#
                html2 = """\
                            <tr>
                                <td style="text-align: left;">{ShopName}</td>
                                <td>{Orders}</td>
                                <td>{Online}</td>
                                <td>{Bank}</td>
                                <td>{Ad}</td>
                                <td>{IP}</td>
                                <td>{Comm}</td>
                                <td style="font-weight: bold;">{Amount}</td>
                            </tr>
                """.format(ShopName =y, Orders=perfOrders, Online=perfOnline, Bank=perfBank, Ad=perfAdjust, IP=perfIP, Comm=perfComm ,Amount=perfAmount)
                
                html += html2
                sendEmail = True
                break
            else:
                logging.warning("Shop name not found")
       
    html3 = """\
                    </table>
                    <p>
                        {PaymentTerm}
                        <br><br>Feel free to let us know if you have any queries.
                    </p>
                </body>
            </html>
    """.format(PaymentTerm=payTerm)

    # logging.info(msg)
    #logging.info(time+": Sending email to: "+chainName)
    logging.info("Sending email to: "+chainName)

    # Attaching files #
    # files = [shopName+'.pdf']

    # for file in files:
    #     with open(file, 'rb') as f:
    #         file_data = f.read()
    #         file_name = f.name
    #     msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)

    # Combining html parts of email #
    html += html3
    part1=MIMEText(html, 'html')
    msg.attach(part1)

    # Sending Email using SSL connection #
    try:
        if(sendEmail==True):
            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                smtp.login(myEmail, myEmailPass)
                smtp.send_message(msg)

            logging.info("Email sent to: "+chainName)
    except:
        logging.error("Email aborted as no shop data found for: "+chainName)
