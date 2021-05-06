# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook
import PySimpleGUI as sg
import openpyxl
import datetime
import xlsxwriter


def simple_read_commands():
    wb = load_workbook(filename='test_book.xlsx')
    sheet_ranges = wb['range names']
    print(sheet_ranges['D18'].value)


def simple_write_commands():
    wb = Workbook()
    dest_filename = 'test_book.xlsx'
    ws1 = wb.active
    ws1.title = "range names"

    for row in range(1, 40):
        ws1.append(range(600))

    ws1.merge_cells('A2:D3')
    ws1.unmerge_cells('A2:D3')
    ws1.insert_rows(7)

    ws2 = wb.create_sheet(title="PI")
    ws2['F5'] = 3.14

    ws3 = wb.create_sheet(title="Data")
    for row in range(10, 20):
        for col in range(27, 54):
            _ = ws3.cell(column=col, row=row, value="{0}".format(get_column_letter(col)))

    print(ws3['AA10'].value)

    ws4 = wb.create_sheet(title="Fold")
    ws4.column_dimensions.group('A', 'D', hidden=True)
    ws4.row_dimensions.group(1, 10, hidden=True)

    wb.save(filename=dest_filename)


def transfer(source, destination):
    row_start = int(input("row_start:"))
    row_end = int(input("row_end:"))
    column_start = int(input("column_start:"))
    column_end = int(input("column_end:"))
    print("rows", row_start, "-", row_end, "columns", column_start, "-", column_end)

    wb1 = load_workbook(filename=source)
    wb2 = load_workbook(filename=destination)
    ws1 = wb1.active
    ws2 = wb2.active
    for row in range(row_start, row_end + 1):
        for col in range(column_start, column_end + 1):
            c = ws1.cell(row=row, column=col)
            ws2.cell(row=row, column=col).value = c.value
    wb2.save(destination)
    return


def copy(source, target):
    wb = Workbook
    source = wb.active
    target = wb.copy_worksheet(source)

    return


if __name__ == '__main__':
    sg.theme('DarkAmber')
    layout = [
        [sg.Text('Source file:'), sg.Input(), sg.FileBrowse()],
        [sg.Text('Destination file:'), sg.Input(), sg.FileBrowse()],
        [sg.Button('Ok'), sg.Button('Cancel')]]

    window = sg.Window('Window Title', layout, )
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == 'Cancel':  # if user closes window or clicks cancel
            break
        print('source file ', values[0])
        print('Destination file', values[1])

    window.close()
    simple_write_commands()
    simple_read_commands()
    transfer(values[0], values[1])
# See PyCharm help at https://www.jetbrains.com/help/pycharm/
