import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication
from os.path import join, dirname, abspath
from datetime import datetime
import xlrd
def send_mail():
    sender_email = "danghuyhoang18@gmail.com"
    loc = join(dirname(dirname(abspath(__file__))), 'send_payroll_over_email' ,'database.xlsx')
    wb=xlrd.open_workbook(loc)
    sheet_names = wb.sheet_names()
    sheet=wb.sheet_by_index(0)
    for i in range(1,sheet.nrows):
        text = MIMEMultipart('alternative')
        text.attach(MIMEText("<html><head></head><body>Hi "+sheet.cell_value(i,1)+",<br>On behalf of CoAsia SEMI Vietnam, we would like to send you the payslip of your salary in " + datetime.strptime(sheet.cell_value(i,39),'%d/%m/%Y').strftime('%m-%Y') +  ", please kindly check attached file herewith for more details. <br>Should you have any queries, don't hesitate to let us know. <br>Thank you for your efforts during the month. <br>Best Regards,</body></html>", "html", _charset="utf-8"))
        msg = MIMEMultipart('mixed')
        msg.attach(text)
        msg['Subject'] = 'Payslip for period from '+ sheet.cell_value(i,38)+ ' to ' + sheet.cell_value(i,39)
        msg['From'] = sender_email
        msg['To'] = sheet.cell_value(i,2)
        pdf = MIMEApplication(open("./payslip_files/"+datetime.strptime(sheet.cell_value(i,39),'%d/%m/%Y').strftime('%Y%m')+" - CoAsia SEMI Vietnam - Payslip - "+sheet.cell_value(i,0)+".pdf", 'rb').read())
        pdf.add_header('Content-Disposition', 'attachment', filename= datetime.strptime(sheet.cell_value(i,39),'%d/%m/%Y').strftime('%Y%m')+" - CoAsia SEMI Vietnam - Payslip - "+sheet.cell_value(i,0)+".pdf")
        msg.attach(pdf) 

        try:
            with smtplib.SMTP('smtp.gmail.com', 587) as smtpObj:
                smtpObj.ehlo()
                smtpObj.starttls()
                smtpObj.login("danghuyhoang18@gmail.com", "cuzztiiarsglzfda")
                smtpObj.sendmail(sender_email, sheet.cell_value(i,2), msg.as_string())
                print("Sent to "+sheet.cell_value(i,2)+" successfully")
        except Exception as e:
            print(e)

    print ("OK!")

if __name__ == "__main__":
    exec(open('pdf_generator.py').read())
    send_mail()