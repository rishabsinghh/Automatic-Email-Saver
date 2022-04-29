import win32com.client
import re
from win32com.client import Dispatch
import os
import tkinter as tk
from tkinter import *
from datetime import date
import shutil
root=tk.Tk()
root.geometry("300x200")

global parent_dir
parent_dir=os.getcwd()
global directory
directory =str(date.today())
global path2
path2 = os.path.join(parent_dir, directory)
def main():
    def error():
        screen1=Toplevel(root)
        screen1.geometry("150x90")
        screen1.title("Warning!")
        Label(screen1,text="All Fields Required",fg="red").pack()
    def file_exists():
        screen2=Toplevel(root)
        screen2.geometry("150x90")
        screen2.title("Warning!")
        Label(screen2,text="File Already Exists",fg="red").pack()
    def invalid_val():
        screen2=Toplevel(root)
        screen2.geometry("150x90")
        screen2.title("Warning!")
        Label(screen2,text="Invalid Values",fg="red").pack()
        
    def get_emails(email,date_from,date_to,folder_name,folder):
        out_app = win32com.client.gencache.EnsureDispatch("Outlook.Application")
        out_namespace = out_app.GetNamespace("MAPI")
        sFilter = "[ReceivedTime]>= '" + date_from + "' AND [ReceivedTime]<= '" + date_to + "'"
        out_iter_folder = out_namespace.Folders[email].Folders[folder]
        filteredEmails = out_iter_folder.Items.Restrict(sFilter)
        item_count = filteredEmails.Count
        for i in range(item_count,0,-1):
            x=filteredEmails[i]
            name = str(x.Subject)
            #to eliminate any special charecters in the name
            name = re.sub('[^A-Za-z0-9]+', '', name)+'.msg'
            #to save in the current working directory
            x.SaveAs(folder_name+'//'+name)
    if os.path.exists(path2)==False:
        os.mkdir(path2)
    email_name=Emailvalue.get()
    date_from=date_startvalue.get()
    date_to=date_endvalue.get()
    folder=Foldervalue.get()
    if email_name=="" or date_from=="" or date_to=="" or folder=="":
        error()
        
    curr_dirr=email_name
    parentdir=path2
    email_path = os.path.join(parentdir, curr_dirr)
    folder_path= os.path.join(email_path, folder)
    if os.path.exists(email_path):
        if os.path.exists(folder_path):
            file_exists()
        else:
            os.mkdir(folder_path)
            try:
                get_emails(email_name,date_from,date_to,folder_path,folder)
            except:
                shutil.rmtree(email_path)
                invalid_val()
    else:
        os.mkdir(email_path)
        os.mkdir(folder_path)
        try:
            get_emails(email_name,date_from,date_to,folder_path,folder)
        except:
            shutil.rmtree(email_path)
            invalid_val()


Label(root,text="Automatic Email Saver",font="comicansms 13 bold",pady=15).grid(row=0,column=3)
Email = Label(root, text="Email")
date_start = Label(root, text="Start Date")
date_end = Label(root, text="Ending Date")
Folder=Label(root,text="Folder")
Email.grid(row=1, column=2)
date_start.grid(row=2, column=2)
date_end.grid(row=3, column=2)
Folder.grid(row=4,column=2)
Emailvalue = StringVar()
date_startvalue = StringVar()
date_endvalue = StringVar()
Foldervalue=StringVar()
Emailentry = Entry(root, textvariable=Emailvalue)
date_toentry = Entry(root, textvariable=date_startvalue)
date_endentry = Entry(root, textvariable=date_endvalue)
Folder_entry = Entry(root, textvariable=Foldervalue)
Emailentry.grid(row=1, column=3)
date_toentry.grid(row=2, column=3)
date_endentry.grid(row=3, column=3)
Folder_entry.grid(row=4, column=3)
Button(text="Save", command=main).grid(row=5, column=3)
#Label(root,text="All Files Successfuly Saved")
root.resizable(False,False)
root.mainloop()

