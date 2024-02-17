from tkinter import *
from tkinter import ttk
import tkinter.messagebox as tmsg
import os
import openpyxl


def Enter():

    accept = termsandcondtions.get()

    if accept=="accepted":
        firstname = firstnameentry.get()
        lastnames = lastnameentry.get()
        if firstname and lastnames:
            title = title_combobox.get()
            age = age_spinbox.get()
            nationality = nationality_combobox.get()

            # courses information
            registration = registrationchecks.get()
            completecourse = completecoursespinbox.get()
            semester = semesterspinbox.get()
            termsandcondtion=termsandcondtions.get()

            # terms and conditions

            print("First name:", firstname, "Last name:", lastnames)
            print("Title:", title, "Age:", age, "Nationality:", nationality)
            print("--------------------------------------------------------------------")
            print("registration status:",registration)
            print("Complete course:", completecourse, "Semester:", semester)


            filepath=r"C:\Users\panta\Desktop\entry\Book2.xlsx"
            if not os.path.exists(filepath):
                workbook = openpyxl.Workbook()
                sheet = workbook.active
                heading = ['First name', 'Last name', 'Title', 'Age', 'Nationality', 'Courses', 'Semester',
                           'Registration status']
                sheet.append(heading)
                workbook.save(filepath)
            workbook=openpyxl.load_workbook(filepath)
            sheet=workbook.active
            sheet.append([firstname, lastnames, title, age, nationality, completecourse, semester, registration])
            workbook.save(filepath)
        else:
            tmsg.showwarning(title="error",message="plz fill the first and last name")
    else:
        tmsg.showwarning(title='error',message="you have not accpeted the terms and conditions")


root = Tk()
root.geometry("600x500")
root.title('student form')

frame = Frame(root)
frame.pack()

userinfo = LabelFrame(frame, text='User Information',font="arial 12 bold")
userinfo.grid(row=0, column=0, padx=20, pady=20)
Label(userinfo, text="First name").grid(row=1, column=0)
Label(userinfo, text="Last name").grid(row=1, column=1)

firstnames = StringVar()
lastnames = StringVar()
title = StringVar()

firstnameentry = Entry(userinfo, textvariable=firstnames)
firstnameentry.grid(row=2, column=0)

lastnameentry = Entry(userinfo, textvariable=lastnames)
lastnameentry.grid(row=2, column=1)

Label(userinfo, text="Title").grid(row=1, column=3)
title_combobox = ttk.Combobox(userinfo, textvariable=title, values=["", "Mr.", "Ms.", "Dr."])
title_combobox.grid(row=2, column=3)

Label(userinfo, text="Age").grid(row=3, column=0)
age_spinbox = Spinbox(userinfo, from_=18, to=110)
age_spinbox.grid(row=4, column=0)

Label(userinfo, text="Nationality").grid(row=3, column=1)
nationality_combobox = ttk.Combobox(userinfo, values=["Indian", "American", "Chinese","British","Brazilian"])
nationality_combobox.grid(row=4, column=1)

for widget in userinfo.winfo_children():
    widget.grid_configure(padx=10, pady=5)

frame2 = Frame(root)
frame2.pack()

secondframe = LabelFrame(frame2)
secondframe.grid(row=1, column=0, sticky='news')
registrationchecks = StringVar(value="Not registered")
Label(secondframe, text="Registration status").grid(row=0, column=0)
registrationcheck = Checkbutton(secondframe, text="currently registered", variable=registrationchecks,onvalue="registerd",offvalue="not registered")
registrationcheck.grid(row=1, column=0)

Label(secondframe, text="Complete course").grid(row=0, column=1)
completecoursespinbox = Spinbox(secondframe, from_=0, to=8)
completecoursespinbox.grid(row=1, column=1)

Label(secondframe, text="Semester").grid(row=0, column=2)
semesterspinbox = Spinbox(secondframe, from_=0, to=6)
semesterspinbox.grid(row=1, column=2)

for widget in secondframe.winfo_children():
    widget.grid_configure(padx=14, pady=5)

frame3 = Frame(root)
frame3.pack()

thirdframe = LabelFrame(frame3, text="Terms & condtions")
thirdframe.grid(row=0, column=0, sticky='news', padx=20, pady=20)

termsandcondtions = StringVar(value="not accepted")

termsandcondtionscheck = Checkbutton(thirdframe, text="I accept terms and condtions", variable=termsandcondtions,onvalue="accepted",offvalue="not accepted")
termsandcondtionscheck.grid(row=-0, column=0)

frame4 = Frame(root)
frame4.pack()

Button(frame4, text='Enter data', command=Enter).grid(row=3, column=0, sticky='news', padx=20, pady=10)

root.mainloop()

