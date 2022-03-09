import xlwings as xw
#将二维数组写入Excel
def outPutData(xlsx_path,sheet_name,begin_cell,data_2d_array):
    app=xw.app(visible=False, add_book=False)
    try:
       wb=app.books.open(xlsx_path)
       sheet=wb.sheets[sheet_name]
       sheet.range(begin_cell).value= data_2d_array
    except ZeroDivisionError as e:
        print('except:', e)
    finally:
        wb.save()
        wb.close()
        app.quit()