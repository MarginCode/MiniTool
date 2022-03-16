from ast import Return
from typing import List
import xlwings as xw
import win32

#替换三维列表中的空list项,将三维列表转换为二维（输入三维列表中的第三维直接删除）
def noneListToNA(data3darray):
    for i in range(len(data3darray)):
        for j in range(len(data3darray[i])):
            if type(data3darray[i][j])==List: data3darray[i][j]=="NA"
    return data3darray


#将二维列表写入Excel
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

#读取excel表格中所出现的PDF和png文件路径，然后进行插入操作
def inputAttachment(xlsx_path,sheetname,lanenum):
    xlsx=win32.gencache.EnsureDis
