import math
import excel
import xlsxwriter


def print_graph(_x, _y):
    """
    Функция вывода графика

    :param _x: список значений X
    :param _y: список значений Y
    """
    workbook = xlsxwriter.Workbook('graph.xlsx')
    sheet = workbook.add_worksheet()

    sheet.write_column('A1', _x)
    sheet.write_column('B1', _y)

    chart = workbook.add_chart({'type': 'bar'})

    chart.add_series(
        {
            'categories': '=Sheet1!$A$1:$A$40',
            'values': '=Sheet1!$B$1:$B$40',
        }
    )

    sheet.insert_chart('E2', chart)
    workbook.close()


def make_data(func=None):
    """
    Функция создает данные для графика

    :param func: математическая функция, выбранная пользователем
    :return: списки значений
    """
    x_list = []
    y_list = []

    if func is None:
        i = -10
        while i <= 10:
            x_list.append(i)
            y_list.append((i ** 2) + (5 * i) - 4)
            i += 0.5

        return x_list, y_list

    i = -10
    while i <= 10:
        x_list.append(i)
        y_list.append(func(i))
        i += 0.5

    return x_list, y_list


def user_choise(app):
    """
    Функция выбора математической функции для построения графика

    :param app: excel приложение
    :return: пользовательский выбор
    """

    print('Доступные функции: ')
    for i, val in enumerate(app.get_functions('A1', 'A4')):
        print(f'{i + 1}.) {val}')

    return input('Ваш выбор: ')


if __name__ == '__main__':
    # Создаем экземпляр Excel приложения
    excel_app = excel.ExcelApplication()
    # Флаг Visible ставим True
    excel_app.show()

    # Данные
    x = []
    y = []

    # Значение X из ячейки F3
    x_val = excel_app.get_xvalue()
    # Это будет значение для ячейки Y в F6
    y_val = None

    # Выбор функции
    ch = user_choise(excel_app)
    if ch == '1':
        x, y = make_data()
        y_val = x_val ** 2 + 5 * x_val - 4
        excel_app.set(2, 6, 1)
        excel_app.set(5, 6, 'y = x^2 + 5x - 4')
    elif ch == '2':
        x, y = make_data(math.sin)
        y_val = math.sin(x_val)
        excel_app.set(2, 6, 2)
        excel_app.set(5, 6, 'sin(x)')
    elif ch == '3':
        x, y = make_data(math.cos)
        y_val = math.cos(x_val)
        excel_app.set(2, 6, 3)
        excel_app.set(5, 6, 'cos(x)')
    elif ch == '4':
        x, y = make_data(math.tan)
        y_val = math.tan(x_val)
        excel_app.set(2, 6, 4)
        excel_app.set(5, 6, 'tan(x)')
    else:
        raise Exception('Ошибка ввода')

    # Занесем полученное значени Y в ячейку F6
    excel_app.set(6, 6, y_val)

    # Заносим данные результатов математических функций в таблицу
    excel_app.set_data_list(y)

    # Выводим график
    print_graph(x, y)

    excel_app.quit()
