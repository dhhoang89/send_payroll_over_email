from xhtml2pdf import pisa
from os.path import join, dirname, abspath 
import jinja2
import xlrd
templateLoader = jinja2.FileSystemLoader(searchpath="./")
templateEnv = jinja2.Environment(loader=templateLoader)
TEMPLATE_FILE = "payslip_template.html"
template = templateEnv.get_template(TEMPLATE_FILE)

# This data can come from database query
loc = join(dirname(dirname(abspath(__file__))), 'send_payroll_over_email' ,'database.xlsx')
wb=xlrd.open_workbook(loc)
sheet_names = wb.sheet_names()
sheet=wb.sheet_by_index(0)
for i in range(1,sheet.nrows):
    body = {
        "data":{
            "Name": sheet.cell_value(i,1),
            "Salary": sheet.cell_value(i,2),
        }
    }

    # This renders template with dynamic data 
    sourceHtml = template.render(json_data=body["data"]) 
    outputFilename = "payslip"+str(sheet.cell_value(i,0))+".pdf"
    # Utility function
    def convertHtmlToPdf(sourceHtml, outputFilename):
        # open output file for writing (truncated binary)
        resultFile = open("./payslip_files/"+outputFilename, "w+b")

        # convert HTML to PDF
        pisaStatus = pisa.CreatePDF(
                src=sourceHtml,            # the HTML to convert
                dest=resultFile)           # file handle to receive result

        # close output file
        resultFile.close()

        # return True on success and False on errors
        print(pisaStatus.err, type(pisaStatus.err))
        return pisaStatus.err

    if __name__ == "__main__":
        pisa.showLogging()
        convertHtmlToPdf(sourceHtml, outputFilename)