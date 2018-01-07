#encoding:gbk
import win32com.client as handler
from  time import  sleep
from Tkinter import Tk



def email():
    Outlook = handler.DispatchEx('Outlook.Application')
    mail = Outlook.CreateItem(handler.constants.olMailItem)
    Outlook.Visible=True
    recp=mail.Recipients.Add('893825883@qq.com')
    sub=mail.Subject='来自 py'
    boay=['Line %s'% i for i in range(1,20)]
    boay.insert(0,'%s\r\n',sub)
    boay.append('\r\n 来自 朋友、py')
    mail.Send()

    ns=Outlook.GetNamespace('MAPI')
    folder = ns.GetDefaultFolder(handler.constants.olFolderOutbox)
    folder.Display()
    folder.Items.Item(1).Display()
    Outlook.Quit()
    sleep(0.5)

    # file = r'C:\Users\ck\Desktop\123.doc'
    # doc.SaveAs(file)
    # word.Application.Quit()
if __name__=='__main__':
    Tk().withdraw()
    email()
