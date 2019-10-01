""" convertor module as functions to do specific task called in main driver program
    Rev. 1
"""



import win32com.client
import sys
import os
import io
import win32api
import docx

# wordtopdf converts word (.docx) file to pdf file.
# params : filepath : directory (string) where the file is stored which is about to converted
#           pdfpath : output directory (use in string format) where the converted pdf file is to be stored
#           outputname: filename with which the pdf file is to be saved. (dont attach extension)
# returns: out : output path of converted pdf file.

def wordtopdf(filepath,pdfpath,outputName):
    # debuggin purpose 
    #    print("inside wordtopdf")
    
    #scaling 
    wdFormatPDF = 17
    
    #varaible in which to filepath is loaded
    in_file = os.path.abspath(filepath)
    
    #output path of coverted file
    out = str(pdfpath) + outputName + '.pdf'
       
    #debugging purpose
    #print(out)
    
    out_file = os.path.abspath(out)
    
    #calling the win32com librarys dispatcher to start word application
    word = win32com.client.Dispatch('Word.Application')
    
    #opening a word document file in opened word application
    doc = word.Documents.Open(in_file)
    
    #saving the opened file in pdf format at out_file locaiton
    doc.SaveAs(out_file, FileFormat=wdFormatPDF)
    
    #closing the file
    doc.Close()
    
    #closing the word.Application
    word.Quit()

    return out


#similar to above, redundancy in code. ignore if needed anywhere
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


# exceltopdf 
# params : filepath : directory (string) where the file is stored which is about to converted
#           pdfpath : output directory (use in string format) where the converted pdf file is to be stored
#           outputname: filename with which the pdf file is to be saved. (dont attach extension)
#returns : output file location

def exceltopdf(filepath, pdfpath, outputName):

    #dispatching the excel application
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = 0

    #starting or opening a new excel workbook with the input file
    wb = excel.Workbooks.Open(filepath)
    ws = wb.Worksheets[0]  # selecting starting location of workbook

    #saving as the input file as pdf
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



# texttopdf
#params : filepath : directory (string) where the file is stored which is about to converted
#           pdfpath : output directory (use in string format) where the converted pdf file is to be stored
#           outputname: filename with which the pdf file is to be saved. (dont attach extension)
#returns : output file location

def texttopdf(filepath, pdfpath, outputName):

    #file handling inbuilt function 
    f=open(filepath,'r')

    # setting the output url
    out = str(pdfpath) + outputName + '.pdf'
    # reading the txt file in text object
    text = f.read()
    
    # setting a docx object and saving the textfile text into word format
    docu=docx.Document()
    para=docu.add_paragraph(text)
    docu.save(out)
    
    #sending the saved format in wordpdf convertor for conversion purpose
    wordpdf(filepath, pdfpath, outputName)
    
    #closing the file
    f.close()
    return out



