from openpyxl import load_workbook
wb = load_workbook("Invoice.xlsx")

sheet = wb.active


name = sheet['B19']#B19

trade = sheet['D19']#D19

monday = sheet['E19']#E19
tuesday = sheet['F19']#F19
wednesday = sheet['G19']#G19
thursday = sheet['H19'] #H19
friday = sheet['I19'] #I19
saturday = sheet['J19'] #J19

hourly = sheet['P19'] #P19

invoicenumber = sheet['G11'] #G11

address = sheet['G12'] #G12

date = sheet['G13'] #G13


wb.save("Invoice1.xlsx")
