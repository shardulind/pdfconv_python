
import win32com.client
import sys
import os
import io
import win32api
import docx
def wordtopdf(filepath,pdfpath,outputName):
    print("inside wordtopdf")
    
    wdFormatPDF = 17
    in_file = os.path.abspath(filepath)
    out = str(pdfpath) + outputName + '.pdf'
    print(out)
    out_file = os.path.abspath(out)
    word = win32com.client.Dispatch('Word.Application')
    doc = word.Documents.Open(in_file)
    doc.SaveAs(out_file, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()

    return out

def wordpdf(filepath, pdfpath, outputName):
    wdFormatPDF = 17
    out = str(pdfpath) + outputName + '.pdf'
    in_file = os.path.abspath(filepath)
    out_file = os.path.abspath(out)
    word = win32com.client.Dispatch('Word.Application')
    doc = word.Documents.Open(in_file)
    doc.SaveAs(out_file, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()
    return out



def exceltopdf(filepath, pdfpath, outputName):
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = 0

    wb = excel.Workbooks.Open(filepath)
    ws = wb.Worksheets[0]

    out = str(pdfpath) + outputName + '.pdf'
    try:
        wb.SaveAs(out, FileFormat=57)
    except Exception (e):
        print ("Failed to convert")
        print (e)
    finally:
        wb.Close()
        excel.Quit()

    return out

def texttopdf(filepath, pdfpath, outputName):
    #f=open(r'C:\Users\Adhiksha\Desktop\csv files\codefordfa.txt','r')
    f=open(filepath,'r')

    out = str(pdfpath) + outputName + '.pdf'
    text = f.read()
    docu=docx.Document()
    para=docu.add_paragraph(text)
    docu.save(out)
    wordpdf(filepath, pdfpath, outputName)
    f.close()
    return out



