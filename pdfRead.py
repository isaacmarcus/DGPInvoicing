# import packages
import PyPDF2
import re
import os
import tkinter as tk
from tkinter.filedialog import askdirectory

root = tk.Tk()
root.withdraw()

# Call the tkinter library to ask user to choose invoice and do folders
invFolderPath = askdirectory(title='Select Invoices Folder') + "/"  # ask user to select invoices folder
doFolderPath = askdirectory(title='Select Delivery Orders Folder') + "/"  # ask user to select invoices folder

# make a directory for invoices missing DO as well as completed merged invoices
try:
    os.mkdir(invFolderPath + "Merged Invoices")
    os.mkdir(invFolderPath + "Invoices Missing DO")
except Exception as e:
    print(e)

outputFolder = invFolderPath + "Merged Invoices/"
missingDOFolder = invFolderPath + "Invoices Missing DO/"

# define key terms
re_param = "A€[\d]{6}"
#re_param2 = "A[\d]{6}"  # for use in the future if we need to check more variations of DO No.

# loop through files in specified folder path
for filename in os.listdir(invFolderPath):
    if filename.endswith(".pdf"):
        # open the current pdf file
        currentInvoiceFile = PyPDF2.PdfFileReader(invFolderPath + filename)

        # get number of pages for looping through later
        NumPages = currentInvoiceFile.getNumPages()

        # Loop through pages of the current pdf to find delivery order number
        doNumber = ""
        doMatch = None
        for i in range(0, NumPages):
            # Get current page and store in object
            PageObj = currentInvoiceFile.getPage(i)
            # Extract text from current page
            Text = PageObj.extractText()
            # Search for a match of the do number
            doMatch = re.search(re_param, Text)
            if doMatch:
                # if match is found, store in doNumber object
                doNumber = doMatch.string[doMatch.start():doMatch.end()].replace("€", "")

                # Loop through do folder to check for matches
                for curDOFile in os.listdir(doFolderPath):
                    # If match is found, start merging that file
                    if curDOFile.endswith(".pdf") and doNumber in curDOFile:
                        pdfMergeObj = PyPDF2.PdfFileMerger()
                        pdfMergeObj.append(currentInvoiceFile)
                        pdfMergeObj.append(PyPDF2.PdfFileReader(doFolderPath + curDOFile))
                        # Create new file in designated output folder
                        pdfMergeObj.write(outputFolder + filename.replace(".pdf", "") + "_" + doNumber + ".pdf")
                        # close the merge object after done merging and writing
                        pdfMergeObj.close()
            else:
                continue

        # Handle case where no Delivery order was found in the pdf
        if doMatch is None:
            print(filename)
            missingDOMerger = PyPDF2.PdfFileMerger()
            missingDOMerger.append(currentInvoiceFile)
            missingDOMerger.write(missingDOFolder + filename)
            missingDOMerger.close()
        else:
            continue

    else:  # ending for IF CONTAINS .PDF IN NAME ELSE
        continue
