import win32com.client as win32
import os
olApp = win32.Dispatch("Outlook.Application")
olNS = olApp.GetNamespace("MAPI")

def get_file_content(file_name):
    with open(file_name, 'rb') as f:
        return f.read()


def get_file_by_lines(file_name):
    with open(file_name, 'r') as f:
        return f.readlines()



for lines in get_file_by_lines("list_destinataire.txt"):
    mailItem = olApp.CreateItem(0)
    mailItem.To = lines
    mailItem.Subject = "Test"
    mailItem.BodyFormat = 2
    mailItem.HTMLBody = get_file_content("mailformat.html").decode("utf-8")
    for filename in os.listdir("./File_to_send"):
        mailItem.Attachments.Add(os.path.join(os.getcwd() + "\File_to_send" ,filename))
    mailItem.Display()
    mailItem.Send()
