""" pdfconvertor main driving program with gui
    Rev. 1
"""

#importing all necessary libraries
from tkinter import filedialog
from tkinter import *
import os
import convertor    #importing convertor module (convertor.py )
import mail         #importing mail module (mail.py)
import platform


#getting information about operating system in os_type variable
os_type = platform.system()


#############################################################
#global variables declaration

#initially set to False, once file gets converted to pdf becomes True
isFileConverted = False

#Output name of file, will be taken from user 
outputName = ''

#the file which is to be converted its path will be stored in it
filepath=''

#the default directory where the converted pdf file will be stored.
# hardcoded, make sure to change it as per users local directories.
# will be made automatic in further updates
pdfpath='C:/Users/lenovo/Desktop/'

#if user is logged in into gmail or not?
login_flag=False

#gmail username and passwords
username=''
password=''
###############################################################






## function responsible to start user login routine
def trigLogin():
    
    # function which logs in into the gmail acccount
    def gmailLogin():
        
        ## global login related variables are accessed
        global username
        global password
        global login_flag
        username = str(e.get())
        password = str(p.get())
        login_flag=True
        loginWin.destroy()
    
    
    # small dialog window for login to  take username and pasword
    loginWin = Toplevel(root)
   
    Label(loginWin, text="Gmail ID:").grid(row=1,column=0)
    #e is entry field to get username 
    e = Entry(loginWin)
    e.grid(row=1,column=1)
    Label(loginWin, text="Password:").grid(row=3, column=0)
    
    #p is password field to get password 
    p = Entry(loginWin, show="*")
    p.grid(row=3,column=1)

    #login buttton
    b = Button(loginWin, text="Login", command=gmailLogin)
    b.grid(row=4, column=1)


# object of tkinter 
root = Tk()
root.title("PDF Convertor ")
# size of gui window
root.geometry('600x400+0+0')


#setting up menu bar 
menubar = Menu(root)
#things present in menu bar and its associated functions
menubar.add_command(label = "Login", command = trigLogin)
menubar.add_command(label = "About", command = trigLogin)
menubar.add_command(label = "Exit", command = exit)

root.config(menu=menubar)




#Label name in GUI , whatever written in text =""
lblmin = Label(font = ('arial',15), text = "PDF Convertor")
lblmin.grid(row=1, column=0)        ## allignment of the text label in window 


# triggering function to start sendEmail function
def trigSendEmail():
    
    ## checks if FIle is converetd to pdf or not.,
    #if not then first it gets converted to pdf. and then program is ready to send it through gmail
    global isFileConverted
    if isFileConverted == False:
        trigSavePdf()
        isFileConverted = True
    
    #after successfull conversion email sending is triggered
    def trigemail():
        sendEmail()

    ## is user logged in to Gmail account of his own or not is checked?
    #if not ,, then he is first logged in into gmail account 
    if login_flag == False:
        trigLogin()

     # if he is logged in,, emailing routine beginss
    elif login_flag == True:
        
        #small toplevel window is opened
        emailWin = Toplevel(root)
        
        #displays users logged in gmail username
        Label(emailWin, text="Username: ").grid(row=1,column=0)
        Label(emailWin, text=username).grid(row=1,column=1)

        #To:
        Label(emailWin, text="To: ").grid(row=3, column=0)
        # entry field to take email address of whom the mail is to be sent. stored in "to" object
        to = Entry(emailWin)
        to.grid(row=3, column=1)

        #subject:
        Label(emailWin, text="Subject: ").grid(row=4, column=0)
        #entry field to take email subject through textfiled.
        subject = Entry(emailWin)
        subject.grid(row=4,column=1)

        #send button.. on pressing email is sent
        send_btn = Button(emailWin, text="Send", command = trigemail)
        send_btn.grid(row=7,column=3)

        
    ### function resposible to send eemail.
    def sendEmail():
        
        ### username and password of users gmail is accessed
        global username
        global password
        
        #to is set to receivers email ids.
        to_ = str(to.get())
        #subject of email
        subject_ = str(subject.get())
        print("username :" + username)
        print("password: "+password)
        print("to :" + to_)
        print("subject:" + subject_)

        #mail is composed and mail modules sendEmail funciton is called.
        #parameters sent are usrname, password, to , subject and conveted pdf files path.
        mail.sendEmail(username, password,to_, subject_, pdfpath)
        emailWin.destroy()


        
## funciton resposinble to start conversion of file to pdf
def trigSavePdf():
    

    def ok():
        global outputName
        outputName = str(e.get())
        print("value is " + outputName)
        convert()
        top.destroy()

    #getting file name to be saved for pdf
    top = Toplevel(root)
    Label(top, text="File Name:").grid(row=0,column=0)

    e = Entry(top)
    e.grid(row=1,column=0)

    b = Button(top, text="Save", command=ok)
    b.grid(row=2, column=0)


    def convert():
        ##outputName mhnje output pdf file ch nav kay asel te
        ##extension mhnje file chi extension kay load kelellya
        global filepath   ## he ji file load keliye tyacha path
        global pdfpath    ## ha jith file store karaichiye ,, basically same dir
        extension=filepath.split('.')[1]

        ## checks the extension of file to be converted
        
        #if file is docx, 
        if extension == 'docx':
             pdfpath = convertor.wordtopdf(filepath,pdfpath,outputName)
             print(pdfpath)         

        #if file is .txt
        elif extension == 'txt':
             pdfpath = convertor.texttopdf(filepath,pdfpath,outputName)
                 
        #if file is an excel file
        elif extension == 'xlsx':
             pdfpath = convertor.exceltopdf(filepath,pdfpath,outputName)
                
        #after successful convefrsion of file,,, isFileCOnverted flag is set to True
        isFileConverted = True
        


def temp():
    global filepath
    global fileSize
    print(int(fileSize)/1000)
    print(filepath)



    
## called when load file button gets pressed    
def trigGetFile():
    
    ##accessing filepath and filesize as global variables
    global filepath
    global fileSize
    
    #getting file 
    filepath = getFile()
    fileSize = os.path.getsize(filepath)
    global name

    
 ####displaying selected file attributes ############

    #displaying the selected file path in front window.
    a = Label(root, text="File Path: ")
    name = Label(root,text=filepath, font="Times")
    a.grid(row=6,column=1)
    name.grid(row=6,column=2)

    
    # displaying selected file size    
    b = Label(root,text="File Size: ")
    temp = str(fileSize/1000) + "Kb"
    size = Label(root, text=temp)
    b.grid(row=7,column=1)
    size.grid(row=7, column=2)
########################################################
    

    #save button to CONVERT the selected file into PDF and save locally 
    save_pdf_btn = Button(root, text="Save as PDF", command=trigSavePdf)
    save_pdf_btn.grid(row=10, column=0)

    
    #send email button to send email of conveted pdf file to your receipts 
    send_email_btn = Button(root, text="Send file through Email", command=trigSendEmail)
    send_email_btn.grid(row=15,column=0)
    return filepath


#function for small window to load File from file selector which is to be converted
def getFile():
    #accessing filepath global variable
    global filepath
    
    #checking the operating sysem
    if os_type == "Linux":
        filepath =  filedialog.askopenfilename(
                                initialdir = "/home/shardulind/Documents",
                                title = "Select file",
                                filetypes = (("Document","*.docx"),
                                            ("spreadsheet","*.xlsx"),
                                            ("text","*.txt")))
        return filepath
    
    #for windows
    elif os_type == "Windows":
        #file dialog box which can select .docs, .xlxs, .txt files
        filepath =  filedialog.askopenfilename(
                                initialdir = "C:\\Users\\",
                                title = "Select file",
                                filetypes = (("Document","*.docx"),
                                            ("spreadsheet","*.xlsx"),
                                            ("text","*.txt")))
        return filepath




### On clicking this button a file selector dialoag window will open , from where file which is to be converted gets selected.
load_file_btn = Button(root, text="Load File", command=trigGetFile).grid(row=2, column=0)


#essential to keep the gui window in loop
root.mainloop()
