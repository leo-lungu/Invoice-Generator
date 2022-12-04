from openpyxl import load_workbook
from flask import Flask, request, render_template
import datetime

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

ordinary = sheet['L19']

totalhours = sheet['O19']

totalpay = sheet['Q19']

hourly = sheet['P19'] #P19

invoicenumber = sheet['G11'] #G11

address = sheet['G12'] #G12

date = sheet['G13'] #G13

app = Flask(__name__)  
print("aaaa")
@app.route('/invoice', methods = ['POST', 'GET'])
def invoice():
    if request.method == 'POST':
        name.value = request.form.get("name")

        trade.value = request.form.get("trade")

        mondayForm = float(request.form.get("monday"))
        tuesdayForm = float(request.form.get("tuesday"))
        wednesdayForm = float(request.form.get("wednesday"))
        thursdayForm = float(request.form.get("thursday"))
        fridayForm = float(request.form.get("friday"))
        saturdayForm = float(request.form.get("saturday"))

        monday.value = mondayForm
        tuesday.value = tuesdayForm
        wednesday.value = wednesdayForm
        thursday.value = thursdayForm
        friday.value = fridayForm
        saturday.value = saturdayForm

        hourlyForm = float(request.form.get("hourlyrate"))

        hourly.value = hourlyForm
        invoicenumber.value = request.form.get("invoicenumber")

        address.value = request.form.get("address")
        
        totalhoursForm = mondayForm + tuesdayForm + wednesdayForm + thursdayForm + fridayForm + saturdayForm
        totalhours.value = totalhoursForm
        ordinary.value = totalhoursForm
        totalpay.value = totalhoursForm * hourlyForm
        print(request.form.get("date"))
        updatedDate = datetime.datetime.strptime(request.form.get("date"), "%Y-%m-%d").strftime("%d/%m/%Y")
        date.value = datetime.updatedDate
        wb.save("Invoice1.xlsx")
        exit(0)

    return render_template('invoice.html')
 
if __name__ == '__main__':
    app.debug = True
    app.run()
