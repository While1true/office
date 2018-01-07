#encoding:gbk
from tkMessageBox import showwarning
'''
https://www.cnblogs.com/qqnnhhbb/p/3601844.html
'''
import win32com.client as handler
from Tkinter import Tk
from time import sleep
warn = lambda : showwarning("EXCEL",'Exit?')

def makeExcel():
    work=handler.gencache.EnsureDispatch('Excel.Application')
    book=work.Workbooks.Add()
    work.DisplayAlerts = False
    sh=book.ActiveSheet
    work.Visible=True
    sleep(1)
    sh.Cells(1,1).Value='来自py 自动生成'
    sleep(1)
    for i in range(2,30):
        print(type(i),i)
        cells = sh.Cells(i, 1)
        cells.Value='Line %d' % i
        sleep(0.2)
    sh.Cells(i+2,1).Value='from py automatic '
    # warn()
    file=r'C:\Users\ck\Desktop\123.xlsx'
    book.SaveAs(file)
    work.Application.Quit()
if __name__=='__main__':
    Tk().withdraw()
    makeExcel()

