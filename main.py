import tkinter.font
import speak
from tkinter import *
from tkinter import messagebox
from voice_assistant_ENG import main_eng
from voice_assistant_RUS import main_rus
import pyodbc

login = ''
password = ''
lang = ''


def language():
    if lang == 'english':
        main_eng.main_eng_va(name=login)
        pass
    elif lang == 'russian':
        main_rus.main_rus_va(name=login)
    else:
        speak.speak_eng("I know only Russian and English. Which will you choose?")
        language()


def change_lang(var, window3):
    print(type(var))
    global lang
    window3.destroy()
    if var == "0":
        lang = "english"
    elif var == "1":
        lang = "russian"


def global_var(txt=None, pswrd=None, window2=None, type=None):
    conn = pyodbc.connect('Driver={SQL Server};'
                          'Server=USER-PC\SQLEXP;'
                          'Database=Voice_assistant;'
                          'Trusted_Connection=yes;')
    cursor = conn.cursor()
    try:
        cursor.execute('''
            CREATE TABLE dbo.CREDENTIALS
            (
                LOGIN varchar(300) unique,
                PASSWORD varchar(300)
            )
        ''')
        cursor.execute('''
            CREATE TABLE dbo.HISTORY
            (
                NAME varchar(300),
                COMMAND_N varchar(300)
            )
        ''')
    except pyodbc.ProgrammingError:
        pass
    global login
    global password
    try:
        if type == "up":
            login = txt.get()
            password = pswrd.get()
            cursor.execute(f"INSERT INTO dbo.CREDENTIALS values('{login}', '{password}')")
            cursor.execute(f"INSERT INTO dbo.HISTORY values('{login}', 'Created a new user')")
            cursor.commit()
        else:
            cursor.execute('SELECT Login, Password FROM Voice_assistant.dbo.Credentials')
            for lo, p in cursor:
                if lo == txt.get() and p == pswrd.get():
                    login = txt.get()
                    password = pswrd.get()
                    cursor.execute(f"INSERT INTO dbo.HISTORY values('{login}', 'Authorization')")
                    cursor.commit()
                    break
                else:
                    raise pyodbc.DataError

        messagebox.showinfo("Status", "Success")
        window3 = Tk()
        label4 = Label(window3, text=f"{login} Please choose the language for voice assistant", font=my_font)
        label4.pack()
        v = StringVar(window3, "0")
        values = {"English": "0",
                  "Русский": "1"}
        for (text, value) in values.items():
            Radiobutton(window3, text=text, variable=v,
                        value=value, font=my_font).pack(side=TOP, ipady=2)
        button7 = Button(window3, text="OK", command=lambda: [change_lang(var=v.get(), window3=window3)], font=my_font)
        button7.pack()
        window2.destroy()
        window3.mainloop()
    except pyodbc.IntegrityError:
        messagebox.showerror("Status", "FAILD: Login is existed. Please choose another one")
    except pyodbc.DataError:
        messagebox.showerror("Status", "FAILD: You have typed a wrong credentials")


def clicked(type):
    window.destroy()
    window2 = Tk()
    label1 = Label(window2, text='Please input your login and password', font=my_font)
    label1.pack()
    label2 = Label(window2, text='Login: ', font=my_font)
    label2.pack()
    txt = Entry(window2, font=my_font)
    txt.pack()
    label3 = Label(window2, text='Password: ', font=my_font)
    label3.pack()
    pswrd = Entry(window2, show="*", width=15, font=my_font)
    pswrd.pack()
    button3 = Button(window2, text='OK', command=lambda: [global_var(txt, pswrd, window2, type)], font=my_font)
    button3.pack()
    window2.mainloop()


if __name__ == "__main__":
    window = Tk()
    my_font = tkinter.font.Font(family='Helvetica')
    window.geometry('400x200')
    window.title('Introduction')
    label = Label(window, text='Hello user! Please log in or log up below', font=my_font)
    label.place(x=50, y=50)
    button1 = Button(window, text='Log In', font=my_font, command=lambda: [clicked("in")])
    button1.place(x=100, y=80)
    button2 = Button(window, text='Sign Up', font=my_font, command=lambda: [clicked("up")])
    button2.place(x=210, y=80)
    window.mainloop()
    print("here")
    print(login + password)
    print(lang)
    language()







