from xhtml2pdf import pisa
from os.path import join, dirname, abspath 
import jinja2
import xlrd
import locale
from datetime import datetime
templateLoader = jinja2.FileSystemLoader(searchpath="./")
templateEnv = jinja2.Environment(loader=templateLoader)
TEMPLATE_FILE = "payslip_template.html"
template = templateEnv.get_template(TEMPLATE_FILE)

# This data can come from database query
loc = join(dirname(dirname(abspath(__file__))), 'send_payroll_over_email' ,'database.xlsx')
wb=xlrd.open_workbook(loc)
sheet_names = wb.sheet_names()
sheet=wb.sheet_by_index(0)
locale.setlocale( locale.LC_ALL, '' )
for i in range(1,sheet.nrows):
    body = {
        "data":{
            "Name": sheet.cell_value(i,1),
            "Code_fm": sheet.cell_value(i,0),
            "days_contract": sheet.cell_value(i,3),
            "days_probation": sheet.cell_value(i,4),
            "Auto_Working_day": sheet.cell_value(i,5),
            "Working_day": sheet.cell_value(i,6),
            "Full_paid_leave": sheet.cell_value(i,7),
            "Special_leave": sheet.cell_value(i,8),
            "Unpaid_leave": sheet.cell_value(i,9),
            "Sick_leave": sheet.cell_value(i,10),
            "Maternity_leave": sheet.cell_value(i,11),
            "Gross_Salary_level_fm": f"{int(sheet.cell_value(i,12)):,}",
            "Probation_salary_fm": f"{int(sheet.cell_value(i,13)):,}",
            "Actual_sal_contract": f"{int(sheet.cell_value(i,14)):,}",
            "Actual_pro_sal": f"{int(sheet.cell_value(i,15)):,}",
            "Gross_salary_fm": f"{int(sheet.cell_value(i,16)):,}",
            "Total_additional_fm": sheet.cell_value(i,17),
            "Total_OT_hour": sheet.cell_value(i,18),
            "Total_OT_fm": f"{int(sheet.cell_value(i,19)):,}",
            "Lunch_allowance_fm": f"{int(sheet.cell_value(i,20)):,}",
            "Relotation_support__package_fm": f"{int(sheet.cell_value(i,21)):,}",
            "Referral_bonus_fm": f"{int(sheet.cell_value(i,22)):,}",
            "Remedy_fm": f"{int(sheet.cell_value(i,23)):,}",
            "Total_deduction_fm": f"{int(sheet.cell_value(i,24)):,}",
            "SI_by_employee": f"{int(sheet.cell_value(i,25)):,}",
            "HI_by_employee": f"{int(sheet.cell_value(i,26)):,}",
            "UI_by_employee": f"{int(sheet.cell_value(i,27)):,}",
            "Total_SI_by_employee_fm": f"{int(sheet.cell_value(i,28)):,}",
            "Late_time": sheet.cell_value(i,29),
            "Late_time_fm": f"{int(sheet.cell_value(i,30)):,}",
            "Early_time": sheet.cell_value(i,31),
            "Early_time_fm": f"{int(sheet.cell_value(i,32)):,}",
            "Dependent_Qty": sheet.cell_value(i,33),
            "Taxable_income_fm": f"{int(sheet.cell_value(i,34)):,}",
            "Income_tax_fm": f"{int(sheet.cell_value(i,35)):,}",
            "Other__deduction_fm": f"{int(sheet.cell_value(i,36)):,}",
            "net_taken_home_fm": f"{int(sheet.cell_value(i,37)):,}",
            "from_date": sheet.cell_value(i,38),
            "to_date": sheet.cell_value(i,39)
        }
    }

    # This renders template with dynamic data 
    sourceHtml = template.render(json_data=body["data"]) 
    outputFilename = datetime.strptime(sheet.cell_value(i,39),'%d/%m/%Y').strftime('%Y%m')+" - CoAsia SEMI Vietnam - Payslip - "+sheet.cell_value(i,0)+".pdf"
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
        print("Generated payslip for",sheet.cell_value(i,1))
        return pisaStatus.err

    if __name__ == "__main__":
        pisa.showLogging()
        convertHtmlToPdf(sourceHtml, outputFilename)