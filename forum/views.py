from django.shortcuts import render,redirect
from django.http import HttpResponse
from .models import *
from .forms import *
import xlsxwriter
import xlwt
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
            

            response = HttpResponse(content_type='application/ms-excel')
            response['Content-Disposition'] = 'attachment; filename="detail.xls"'
            wb = xlwt.Workbook(encoding='utf-8')
            ws = wb.add_sheet('Users')
            
            row_num = 0
            font_style = xlwt.XFStyle()
            font_style.font.bold = True
            columns = ['email', 'contact number']
            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            font_style = xlwt.XFStyle()

            for email in email_list:
                row_num=row_num+1
                ws.write(row_num,0,email,font_style)
            row_num=0
            for contact in mobile_list:
                row_num=row_num+1
                ws.write(row_num,1,contact,font_style)
            wb.save(response)
            return response
            


    form=DocumentForm()
    context={
        'form':form,
        'phone_number':mobile_list,
        'emails':email_list,
        'file_path':file_path,
    }
    return render(request,'forum/home.html',context)


def export_users_xls(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="detail.xls"'
    wb = xlwt.Workbook(encoding='utf-8')
    ws = wb.add_sheet('Users')
    # Sheet header, first row
    row_num = 0
    font_style = xlwt.XFStyle()
    font_style.font.bold = True
    columns = ['email', 'contact number']
    for col_num in range(len(columns)):
        ws.write(row_num, col_num, columns[col_num], font_style)

    # Sheet body, remaining rows
    font_style = xlwt.XFStyle()

    rows = Document.objects.all().values_list('description', 'uploaded_at')
    for row in rows:
        row_num += 1
        for col_num in range(len(row)):
            ws.write(row_num, col_num, row[col_num], font_style)

    wb.save(response)
    return response