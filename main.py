from tkinter import filedialog
from tkinter import *
import os
#mport convertor
import mail
import platform

os_type = platform.system()


filepath=''
pdfpath='test.pdf'
login_flag=False
username=''
password=''


def trigLogin():
    def gmailLogin():
        global username
        global password
        global login_flag
        username = str(e.get())
        password = str(p.get())
        login_flag=True
        loginWin.destroy()

    loginWin = Toplevel(root)
    Label(loginWin, text="Gmail ID:").grid(row=1,column=0)
    e = Entry(loginWin)
    e.grid(row=1,column=1)
    Label(loginWin, text="Password:").grid(row=3, column=0)
    p = Entry(loginWin, show="*")
    p.grid(row=3,column=1)


    b = Button(loginWin, text="Login", command=gmailLogin)
    b.grid(row=4, column=1)


root = Tk()
root.title("PDF Convertor ")
root.geometry('600x400+0+0')

menubar = Menu(root)

menubar.add_command(label = "Login", command = trigLogin)
menubar.add_command(label = "About", command = trigLogin)
menubar.add_command(label = "Exit", command = exit)

root.config(menu=menubar)


lblmin = Label(font = ('arial',15), text = "PDF Convertor")
lblmin.grid(row=1, column=0)

def trigSendEmail():

    def trigemail():
        sendEmail()

    if login_flag == False:
        trigLogin()

    elif login_flag == True:
        emailWin = Toplevel(root)
        #emailWin.geometry('500x200+12+12')
        Label(emailWin, text="Username: ").grid(row=1,column=0)
        Label(emailWin, text=username).grid(row=1,column=1)

        Label(emailWin, text="To: ").grid(row=3, column=0)
        to = Entry(emailWin)
        to.grid(row=3, column=1)

        Label(emailWin, text="Subject: ").grid(row=4, column=0)
        subject = Entry(emailWin)
        subject.grid(row=4,column=1)


        send_btn = Button(emailWin, text="Send", command = trigemail)
        send_btn.grid(row=7,column=3)

    def sendEmail():
        global username
        global password

        to_ = str(to.get())
        subject_ = str(subject.get())
        print("username :" + username)
        print("password: "+password)
        print("to :" + to_)
        print("subject:" + subject_)

        mail.sendEmail(username, password,to_, subject_, pdfpath)
        emailWin.destroy()


def trigSavePdf():
    outputName = ''

    def ok():
        global outputName
        outputName = str(e.get())
        print("value is " + outputName)
        top.destroy()

    #getting file name to be saved for pdf
    top = Toplevel(root)
    Label(top, text="File Name:").grid(row=0,column=0)

    e = Entry(top)
    e.grid(row=1,column=0)

    b = Button(top, text="Save", command=ok)
    b.grid(row=2, column=0)



    ##outputName mhnje output pdf file ch nav kay asel te
    ##extension mhnje file chi extension kay load kelellya
    global filepath   ## he ji file load keliye tyacha path
    global pdfpath    ## ha jith file store karaichiye ,, basically same dir
    extension=filepath.split('.')[1]
    if extension == 'docx':
            # pdfpath = convertor.wordpdf(filepath,pdfpath,outputName)
            pass #he pass nntr kadhun takaich

    elif extention == 'txt':
            # pdfpath = convertor.texttopdf(filepath,pdfpath,outputName)
            pass

    elif extension == 'xlsx':
            # pdfpath = convertor.exceltopdf(filepath,pdfpath,outputName)
            pass

    #error else case add karaichiye
    #pass kadhun tak nntr



def temp():
    global filepath
    global fileSize
    print(int(fileSize)/1000)
    print(filepath)



def trigGetFile():
    global filepath
    global fileSize
    filepath = getFile()
    fileSize = os.path.getsize(filepath)
    global name

    a = Label(root, text="File Path: ")
    name = Label(root,text=filepath, font="Times")
    a.grid(row=6,column=1)
    name.grid(row=6,column=2)
    #a.pack(side="left")
    #name.pack(side="left")

    b = Label(root,text="File Size: ")
    temp = str(fileSize/1000) + "Kb"
    size = Label(root, text=temp)
    b.grid(row=7,column=1)
    size.grid(row=7, column=2)
    #size.pack(side = "bottom")
    #b.pack(side="bottom")
    #temp()

    save_pdf_btn = Button(root, text="Save as PDF", command=trigSavePdf)
    save_pdf_btn.grid(row=10, column=0)


    send_email_btn = Button(root, text="Send file through Email", command=trigSendEmail)
    send_email_btn.grid(row=15,column=0)
    return filepath

def getFile():
    global filepath
    if os_type == "Linux":
        filepath =  filedialog.askopenfilename(
                                initialdir = "/home/shardulind/Documents",
                                title = "Select file",
                                filetypes = (("Document","*.docx"),
                                            ("spreadsheet","*.xlxs"),
                                            ("text","*.txt")))
        return filepath
    elif os_type == "Windows":
        filepath =  filedialog.askopenfilename(
                                initialdir = "C:\\Users\\",
                                title = "Select file",
                                filetypes = (("Document","*.docx"),
                                            ("spreadsheet","*.xlxs"),
                                            ("text","*.txt")))
        return filepath





load_file_btn = Button(root, text="Load File", command=trigGetFile).grid(row=2, column=0)



root.mainloop()
