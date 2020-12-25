import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication
from os.path import join, dirname, abspath
import xlrd
def send_mail():
    sender_email = "danghuyhoang18@gmail.com"
    loc = join(dirname(dirname(abspath(__file__))), 'send_payroll_over_email' ,'database.xlsx')
    wb=xlrd.open_workbook(loc)
    sheet_names = wb.sheet_names()
    sheet=wb.sheet_by_index(0)
    for i in range(1,sheet.nrows):
        text = MIMEMultipart('alternative')
        text.attach(MIMEText("<html><head></head><body>Hi "+sheet.cell_value(i,1)+",<br>I would like to send your payslip.<br> Thank you for your contribution this month.</body></html>", "html", _charset="utf-8"))
        msg = MIMEMultipart('mixed')
        msg.attach(text)
        msg['Subject'] = 'Luong nhan vien '+ sheet.cell_value(i,1)
        msg['From'] = sender_email
        msg['To'] = "danghuyhoang18@gmail.com"
        pdf = MIMEApplication(open("./payslip_files/payslip00"+str(i)+".pdf", 'rb').read())
        pdf.add_header('Content-Disposition', 'attachment', filename= "payslip00"+str(i)+".pdf")
        msg.attach(pdf) 

        try:
            with smtplib.SMTP('smtp.gmail.com', 587) as smtpObj:
                smtpObj.ehlo()
                smtpObj.starttls()
                smtpObj.login("danghuyhoang18@gmail.com", "cuzztiiarsglzfda")
                smtpObj.sendmail(sender_email, 'danghuyhoang18@gmail.com', msg.as_string())
                print("Sent to "+sheet.cell_value(i,3)+" successfully")
        except Exception as e:
            print(e)

    print ("OK!")

if __name__ == "__main__":
    exec(open('pdf_generator.py').read())
    send_mail()