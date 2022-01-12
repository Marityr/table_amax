import os
import requests

import openpyxl
from openpyxl import Workbook, load_workbook

from dotenv import load_dotenv

load_dotenv()


class Exelfile:
    """Экспорт данных ексель из файла"""

    def read_file_brand() -> list:
        """считывание данных из файла"""

        sheet_list = load_workbook('export.xlsx')
        brand = sheet_list["brand"]

        brand_list = list()
        for cell in brand['A']:
            brand_list.append(cell.value)

        return brand_list

    def read_file_sheet():
        sheet_list = load_workbook('export.xlsx')
        sheet = sheet_list["sheet"]

        #tmp = sheet.cell(row=rowit, column=3).value
        number_list = list()
        item = 0
        for cell in sheet['E']:
            number_list.append(cell.value)
            item += 1
            if item >= 1000:
                break
        return number_list

class API_set:
    """Взаимодействие с API сервиса"""

    def connect_api(number_list):
        """Подключение по API и получение json"""

        HOST_API = os.getenv('HOST_API')
        USER_API = os.getenv('USER_API')
        PASSWORD_API = os.getenv('PASSWORD_API')

        listen = list()
        items = 0
        for i in number_list:
            conect_url = f"https://{HOST_API}/search/brands/?userlogin={USER_API}&userpsw={PASSWORD_API}&number={i}"
            response = requests.get(conect_url)
            datajson = response.json()
            listen.append(datajson)
            items += 1
            print(items)
        return listen



def max_row() -> int:
    """максимальное количество строк с данными"""

    wb = openpyxl.load_workbook('export.xlsx')
    sheet = wb['sheet']
    nb_row = sheet.max_row
    return int(nb_row)

listen = Exelfile.read_file_sheet()
listen.pop(0)
#print(listen)
tempel = API_set.connect_api(listen)
print(tempel)
# print(len(API_set.connect_api()))

# for i in range(2, 10):
#     print(len(API_set.connect_api(i)))
