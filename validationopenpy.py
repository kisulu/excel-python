import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation

wb = openpyxl.load_workbook("Book1.xlsx")

sheet = wb.active

valid_options='"Not started, In Progress, Completed"'

rule = DataValidation(type="list", formula1=valid_options, allow_blank=True)

rule.error = 'Your entry is not valid.'
rule.errorTitle = 'Invalid Entry'


rule.promptTitle = 'Select Option'
rule.prompt = 'please selet from the list.'


sheet.add_data_validation(rule)

rule.add('C2:C100')

wb.save("validation.xlsx")


