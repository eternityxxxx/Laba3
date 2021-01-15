import win32com.client
import os


class ExcelApplication(object):
    def __init__(self):
        self.excel = win32com.client.Dispatch('Excel.Application')
        self.workbooks = self.open_workbook()
        self.sheet = self.activate_sheet()

    def show(self):
        self.excel.Visible = True

    def hide(self):
        self.excel.Visible = False

    def open_workbook(self):
        wb = self.excel.WorkBooks.Open(os.path.abspath('Lab3.1.xlsm'))
        return wb

    def activate_sheet(self):
        sh = self.workbooks.ActiveSheet

        return sh

    def get_functions(self, ran1, ran2):
        """
        Получение значений из ячеек в заданном диапозоне

        :param ran1: начало диапозона
        :param ran2: конец диапозона
        :return: список значений
        """
        values = [r[0].value for r in self.sheet.Range(f'{ran1}:{ran2}')]

        return values

    def get_xvalue(self):
        """
        Получение значения X из F3
        :return:
        """
        xvalue = self.sheet.Cells(3, 6).value

        return xvalue

    def set(self, row, col, val):
        """
        Установка значния в опредленную ячейку

        :param row: строка
        :param col: столбец
        :param val: значение
        """
        self.sheet.Cells(row, col).value = val

    def set_data_list(self, y):
        """
        Установка списка значений в столбец

        :param y: список значений
        """
        i = 1
        for value in y:
            self.sheet.Cells(i, 8).value = value
            i += 1

    def quit(self):
        """
            Завершаем работу с Excel приложением
        """

        self.excel.Quit()
