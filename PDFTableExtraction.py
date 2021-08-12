from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfinterp import resolve1
import pdfplumber
import pandas as pd


def pdfExtraction(pdfFilePath,folderPath):

    #pdfFilePath = "C:\\Users\\SPECTRE\\Downloads\\June 2021.pdf"
    pdfFilePath = pdfFilePath
    file = open(pdfFilePath, 'rb')
    parser = PDFParser(file)
    document = PDFDocument(parser)

    # This will give you the count of pages
    pdfPageCount = resolve1(document.catalog['Pages'])['Count']
    #print(pdfPageCount)
    i = 0

    while i < pdfPageCount:

        pdf = pdfplumber.open(pdfFilePath)
        page = pdf.pages[i]
        tables = page.find_tables()
        tableCount = len(tables)
        if tableCount == 0:
            print("No Table")
        else:
            print(tableCount)
            j=0
            with pd.ExcelWriter(folderPath + str(i) + ".xlsx") as writer:
                while j < tableCount:

                    DSP = tables[j].extract(x_tolerance=5)
                    dfDSP = pd.DataFrame(DSP[1:], columns=DSP[0])
                    dfDSP.to_excel(writer, sheet_name='Sheet_name_'+str(j))
                     #dfDSP.to_excel(folderPath+str(i)+".xlsx",sheet_name=str(j),na_rep="")
                    j += 1

        i += 1


if __name__ == '__main__':
    pdfExtraction("C:\\Users\\SPECTRE\\Downloads\\29-July-2021.pdf","C:\\Users\\SPECTRE\\Documents\\akaBot\\ITSA\\")



