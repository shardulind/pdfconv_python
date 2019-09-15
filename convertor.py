import sys
import os
import comtypes.client
from win32com import client
import win32api
import docx
def wordtopdf():
    wdFormatPDF = 17
    in_file = os.path.abspath('C:\\Users\\Adhiksha\\Desktop\\csv files\\story.docx')
    out_file = os.path.abspath('C:\\Users\\Adhiksha\\Desktop\\csv files\\out.pdf')
    word = comtypes.client.CreateObject('Word.Application')
    doc = word.Documents.Open(in_file)
    doc.SaveAs(out_file, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()

def wordpdf():
    wdFormatPDF = 17
    in_file = os.path.abspath('C:\\Users\\Adhiksha\\Desktop\\csv files\\document.docx')
    out_file = os.path.abspath('C:\\Users\\Adhiksha\\Desktop\\csv files\\out.pdf')
    word = comtypes.client.CreateObject('Word.Application')
    doc = word.Documents.Open(in_file)
    doc.SaveAs(out_file, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()




def exceltopdf():
    excel = client.DispatchEx("Excel.Application")
    excel.Visible = 0

    wb = excel.Workbooks.Open('C:\\Users\\Adhiksha\\Desktop\\csv files\\inner.xlsx')
    ws = wb.Worksheets[0]

    try:
        wb.SaveAs('C:\\Users\\Adhiksha\\Desktop\\csv files\\file.pdf', FileFormat=57)
    except Exception (e):
        print ("Failed to convert")
        print (e)
    finally:
        wb.Close()
        excel.Quit()

def texttopdf():
    f=open(r'C:\Users\Adhiksha\Desktop\csv files\codefordfa.txt','r')
    text = f.read()
    docu=docx.Document()
    para=docu.add_paragraph(text)
    docu.save('C:\\Users\\Adhiksha\\Desktop\\csv files\\document.docx')
    wordpdf()
    f.close()




texttopdf()
