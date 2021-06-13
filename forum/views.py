from django.shortcuts import render,redirect
from django.http import HttpResponse
from .models import *
from .forms import *
import xlsxwriter
import PyPDF2
import re

def pdf_to_text(request):
    mobile_list=[]
    email_list=[]
    line=""
    file_path=""
    if request.method=='POST':
        form=DocumentForm(request.POST, request.FILES)
        if form.is_valid():
            f=request.FILES['document']
            file_path='/media/documents'+request.FILES['document'].name
            pdf=PyPDF2.PdfFileReader(f)
            for page_num in range(pdf.numPages):
                line=line+pdf.getPage(page_num).extractText()
            email = re.findall(r'[\w\.-]+@[\w\.-]+', line)
            mobile_number=re.findall(r'\+?\d[\d -]{8,12}\d',line)
            for e in email:
                email_list.append(e)
            for m in mobile_number:
                mobile_list.append(m)
    form=DocumentForm()
    context={
        'form':form,
        'phone_number':mobile_list,
        'emails':email_list,
        'file_path':file_path,
    }
    return render(request,'forum/home.html',context)

def excel_sheet(request):
    if(request.method=="GET"):
        response = HttpResponse(content_type='application/ms-excel')
        response['Content-Disposition'] = 'attachment; filename="ThePythonDjango.xls"'
        wb = xlwt.Workbook(encoding='utf-8')
        ws = wb.add_sheet("sheet1")
        row_num = 0
        font_style = xlwt.XFStyle()
        font_style.font.bold = True
        columns = ['Column 1', 'Column 2', 'Column 3', 'Column 4', ]
        for col_num in range(len(columns)):
            ws.write(row_num, col_num, columns[col_num], font_style)
        font_style = xlwt.XFStyle()
        data = get_data()
        for my_row in data:
            row_num = row_num + 1
            ws.write(row_num, 0, my_row.name, font_style)
            ws.write(row_num, 1, my_row.start_date_time, font_style)
            ws.write(row_num, 2, my_row.end_date_time, font_style)
            ws.write(row_num, 3, my_row.notes, font_style)  
        wb.save(response)
        return response