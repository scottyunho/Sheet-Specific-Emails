import sys
sys.path.append('/home/scottho/.local/Library/Frameworks/Python.framework/Versions/3.7/bin/python3/site-packages/')
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import gspread
from oauth2client.service_account import ServiceAccountCredentials


scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('/Users/scottho/Desktop/WFN_Medium_Automation/Test-70f387e79130.json', scope)
client = gspread.authorize(creds)

mainSpreadsheet = client.open('Medium Form Responses')
sheet1 = mainSpreadsheet.get_worksheet(0)

fromEmail = "email@gmail.com"
fromEmailPass = "password"



    
## SHEET 2 WORK
sheet2 = mainSpreadsheet.get_worksheet(1)
empty2 = (sheet2.acell('A2').value == "")
if (empty2 == False):
    
    toEmail = 'email@gmail.com' ##delegating sheet 3 work to ...

    sheet2_firstName = sheet3.col_values(2)
    sheet2_lastName = sheet3.col_values(3)
    sheet2_email = sheet3.col_values(6)

    del sheet2_firstName[0]
    del sheet2_lastName[0]
    del sheet2_email[0]

    printLines = ("Set up channels of communciation with the following people and determine meeting dates. "+"\n"+"\n")
    for i in range(0, len(sheet2.get_all_values())-1):
        printLines += ('Name: '+str(sheet2_firstName[i])+" "+str(sheet2_lastName[i])+", "+str(sheet2_email[i])+"\n")

    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login(fromEmail, fromEmailPass)

    msg = MIMEMultipart()
    msg['From'] = fromEmail
    msg['To'] = toEmail
    msg['Subject'] = 'Publishing, New Submission'
    message = (printLines)
    msg.attach(MIMEText(message))

    server.send_message(msg)
    server.quit()





## SHEET 3 WORK
sheet3 = mainSpreadsheet.get_worksheet(2)
empty3 = (sheet3.acell('A2').value == "")
if (empty3 == False):

    toEmail = 'email@gmail.com' ##delegating sheet 3 work to ...

    sheet3_firstName = sheet3.col_values(2)
    sheet3_lastName = sheet3.col_values(3)
    sheet3_email = sheet3.col_values(6)
    sheet3_ideas = sheet3.col_values(8)

    del sheet3_firstName[0]
    del sheet3_lastName[0]
    del sheet3_email[0]
    del sheet3_ideas[0]

    printLines = ("Set up channels of communciation with the following people and decide whether or not to proceed with article ideas. "+"\n"+"\n")
    for i in range(0, len(sheet3.get_all_values())-1):
        printLines += ('Name: '+str(sheet3_firstName[i])+" "+str(sheet3_lastName[i])+", "+str(sheet3_email[i])+", "+str(sheet3_ideas[i])+"\n")

    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login(fromEmail, fromEmailPass)

    msg = MIMEMultipart()
    msg['From'] = fromEmail
    msg['To'] = toEmail
    msg['Subject'] = 'Publishing, New Submission'
    message = (printLines)
    msg.attach(MIMEText(message))

    server.send_message(msg)
    server.quit()





## SHEET 4 WORK
sheet4 = mainSpreadsheet.get_worksheet(3)
empty4 = (sheet4.acell('A2').value == "")
if (empty4 == False):

    toEmail = 'email@gmail.com' ##delegating sheet 4 work to ...
    sheet4 = mainSpreadsheet.get_worksheet(3)

    sheet4_firstName = sheet4.col_values(1)
    sheet4_lastName = sheet4.col_values(2)
    sheet4_email = sheet4.col_values(6)
    sheet4_article = sheet4.col_values(9)

    del sheet4_firstName[0]
    del sheet4_lastName[0]
    del sheet4_email[0]
    del sheet4_article[0]

    printLines = ("Determine whether to proceed with article. If approved, revise then manually post on Medium blog. "+"\n"+"\n")
    for i in range(0, len(sheet4.get_all_values())-1):
        printLines += ('Name: '+str(sheet4_firstName[i])+" "+str(sheet4_lastName[i])+", "+str(sheet4_email[i])+", "+str(sheet4_article[i])+"\n")

    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login(fromEmail, fromEmailPass)

    msg = MIMEMultipart()
    msg['From'] = fromEmail
    msg['To'] = toEmail
    msg['Subject'] = 'Publishing, New Submission'
    message = (printLines)
    msg.attach(MIMEText(message))

    server.send_message(msg)
    server.quit()



## RESET MAIN SPREADSHEET
sheet1.resize(rows=2)
sheet1.resize(rows=30)
sheet1.delete_row(2)


