import access
import datetime
import os
import openpyxl
from win32com.client import Dispatch


class Excel:
    def __init__(self):
        # заносим в переменную перечень используемых цветов заливки
        self.color = {'green': r'#CCFFCC', 'red': r'#FF9999', 'white': None}
        # заносим в переменные путь к каталогу права на папки которого нужно получить и путь к файлу excel
        path = get_all_path()
        self.file_path = path['file']
        self.dir_path = path['dir']
        if os.path.exists(self.file_path):
            self.wb = openpyxl.open(self.file_path)
            self.ws_now = self.wb.create_sheet(self.ws_now_name(), 0)  # insert at first position
        else:
            self.wb = openpyxl.Workbook()
            self.ws_now = self.wb.active
            self.ws_now.title = self.ws_now_name()

    def ws_now_name(self):
        """создает название листа на основании текущей даты (если такой лист уже есть, добавляет время)"""
        ws_name1 = datetime.datetime.today().strftime('%d-%m-%Y')
        ws_name2 = datetime.datetime.today().strftime('%d-%m-%Y %H-%M-%S')
        return ws_name1 if ws_name1 not in self.wb.sheetnames else ws_name2

    def wb_save(self):
        self.wb.save(self.file_path)

    def write(self):
        self.ws_now['A1'].value = 'data'
        self.ws_now['C9'] = 'hello world'
        for cell in tuple(self.ws_now.values):
            print(cell)




def get_all_path():
    """читает пути для работы из файла"""
    if os.path.exists(f'{os.getcwd()}\\config.ini'):
        all_path = {}
        with open(f'{os.getcwd()}\\config.ini', 'r', encoding='cp1251') as file:
            for id, string in enumerate(file):
                if id == 1:
                    all_path['dir'] = string.strip().lower()
                elif id == 3:
                    all_path['file'] = string.strip().lower()
        return all_path
    else:
        return {'file': f'{os.getcwd()}\\access_list.xlsx', 'dir': f'{os.getcwd()}'}


def auto_size_column(book_path, sheet_name):
    """выбирает оптимальную ширину столбцов листа"""
    # from win32com.client import Dispatch
    excel = Dispatch('Excel.Application')
    wb = excel.Workbooks.Open(book_path)
    excel.Worksheets(sheet_name).Activate()
    excel.ActiveSheet.Columns.AutoFit()
    wb.Save()
    wb.Close()
    excel.Quit()


if __name__ == '__main__':
    # wb = openpyxl.open(r'\\fs\SHARE\Documents\OTDEL-IT\access_list.xlsx')
    # print(wb.sheetnames)
    xxl = Excel()
    xxl.write()
    xxl.wb_save()
