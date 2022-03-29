from cgi import print_environ
from io import StringIO
from re import template
from django.shortcuts import render
from django.http import HttpResponse, JsonResponse
from django.views.generic import TemplateView
from httplib2 import Response
from rest_framework.views import APIView
from django.contrib.auth.models import User
import xlsxwriter
import csv
# from pyexcel_xls import get_data as xls_get
# from pyexcel_xlsx import get_data as xlsx_get
# from django.utils.datastructures import MultiValueDictKeyError

# Create your views here.

class Index(TemplateView):
    template_name = 'main.html'

    def get_context_data(self, *args, **kwargs):
        context = super(Index, self).get_context_data(*args, **kwargs)
        context['message'] = User.objects.all()
        return context


# class ParseExcel(APIView):
#     def post(self, request, format=None):
#         try:
#             excel_file = request.FILES['files']
#         except MultiValueDictKeyError:
#             return Response('Failed to upload')
#         if (str(excel_file).split('.')[-1] == "xls"):
#             data = xls_get(excel_file, column_limit=4)
#         elif (str(excel_file).split('.')[-1] == "xlsx"):
#             data = xlsx_get(excel_file, column_limit=4)
#         else:
#             return Response("Failed to xlsx")

import gspread
from oauth2client.service_account import ServiceAccountCredentials




def export(request):
    upload(request)
    return JsonResponse({'message':'success'}, status=200)




def upload(request):
    # print(request.user.email)
    filename = request.GET.get('filename','New_sheet')
    scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
            "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]

    credentials = ServiceAccountCredentials.from_json_keyfile_name('django-auth-345510-628361dbc7e5.json', scope)
    client = gspread.authorize(credentials)
    # client = gspread.oauth(
    #     credentials_filename='django-auth-345510-628361dbc7e5.json',
    #     authorized_user_filename='client_secret.json',
    #     scopes = scope
    # )
    print(client)
    spreadsheet = client.create(filename)
    spreadsheet.share(request.user.email,perm_type='user', role='owner')
    for spread in client.openall():
        print(f' {spread.title}, {spread.url}, {spread.id}')
        # client.insert_permission(file_id=spread.id, value='nileshpayghan7@gmail.com', perm_type='anyone', role='owner', notify=True, with_link=True)

    with open('data.csv', 'r') as file_obj:
        # content = file_obj.read()
        # print(type(content))
        # print(len(content))
        response = HttpResponse (content_type='text/csv')
        writer = csv.writer(response)
        writer.writerows(User.objects.values_list('id','username','date_joined'))
        # print(response.content)
        # print(type(response.content))
        client.import_csv(spreadsheet.id, data=response.content)
    return HttpResponse('Success')







def writetoexcel(request):
    # output = StringIO.StringIO()
    workbook = xlsxwriter.Workbook('Reporte3a4.xlsx', 
                                   {'constant_memory': True})
    worksheet = workbook.add_worksheet()

    bold = workbook.add_format({'bold': True})

    worksheet.write('A1', 'username', bold)
    worksheet.write('B1', 'email', bold)

    Users = User.objects.all()

    row = 1
    col = 0

    for data in Users:
        worksheet.write(row, col, data.username)
        worksheet.write(row, col + 1, data.email)
        worksheet.write(row, col + 2, data.first_name)

    workbook.close()
    # xlsx_Data = output.getvalue()
    return  worksheet