# This program demonstrates the usage of a module called xlsxwriter which can:
#   Only write to a -NEW- xlsx document.
# Theoretically, this means that I may be able to import data from Google sheets to a new Excel file.
import xlsxwriter as xw

for num in range(0, 4):
    my_book = xw.Workbook('Demo' + str(num) + '.xlsx')
    wsHello = my_book.add_worksheet(name='HelloSheet')

    wsHello.write('A1', 'Hello World')
    wsHello.write(1,0,'Goodbye') # (row, column, data)
    for int in range(2,10):
        wsHello.write(int, int, 'This int: ' + str(num))
    my_book.close() # saves the file


