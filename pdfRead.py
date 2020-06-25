# import packages
import PyPDF2
import re
import os
import tkinter as tk
import copypdfto
from tkinter import messagebox
from tkinter.filedialog import askdirectory



root = tk.Tk(className="DG Merge Invoices/Delivery Orders")
root.withdraw()
Height = 450
Width = 500

# Main canvas widget for our window
# mainCanvas = tk.Canvas(root, height=Height, width=Width)
# mainCanvas.pack()
#
# mainFrame = tk.LabelFrame(root)
# mainFrame.place(relx=0.025, rely=0.025, relwidth=0.95, relheight=0.95)
#
# invPathFrame = tk.Frame(mainFrame, bg="#e0e0e0", bd=5)
# invPathFrame.place(relx=0.025, rely=0.025, relheight=0.1, relwidth=0.95)
#
# invPathLabelFrame = tk.Frame(invPathFrame, bg="#cccccc", bd=5)
# invPathLabelFrame.place(relx=0, rely=0.025, relheight=0.9, relwidth=0.65)
#
# invPathButton = tk.Button(invPathFrame)

# titleLabel = tk.Label(root)


messagebox.showinfo(message="Please select Invoice Folder, followed by Delivery Orders Folder")

# Call the tkinter library to ask user to choose invoice and do folders
invFolderPath = askdirectory(title='Select Invoices Folder') + "/"  # ask user to select invoices folder
doFolderPath = askdirectory(title='Select Delivery Orders Folder') + "/"  # ask user to select invoices folder

# make a directory for invoices missing DO as well as completed merged invoices
try:
    os.mkdir(invFolderPath + "Merged Invoices Incomplete")
    os.mkdir(invFolderPath + "Merged Invoices")
    os.mkdir(invFolderPath + "Invoice no Do Number")
    os.mkdir(invFolderPath + "Invoice no Matching DO")
except Exception as e:
    print("Directory already exists, no need to create again..")

incMergedFolder = invFolderPath + "Merged Invoices Incomplete/"
outputFolder = invFolderPath + "Merged Invoices/"
invNoDOFolder = invFolderPath + "Invoice no Do Number/"
invNoMatchFolder = invFolderPath + "Invoice no Matching DO/"

# define key terms
re_param = "A€[\d]{6}"
# re_param2 = "A[\d]{6}"  # for use in the future if we need to check more variations of DO No.

# loop through files in specified folder path
for filename in os.listdir(invFolderPath):
    if filename.endswith(".pdf"):
        # open the current pdf file
        currentInvoiceFile = PyPDF2.PdfFileReader(invFolderPath + filename)

        # get number of pages for looping through later
        NumPages = currentInvoiceFile.getNumPages()

        # Loop through pages of the current pdf to find delivery order number
        doNumber = ""
        doMatchStatus = "unfound"
        for i in range(0, NumPages):
            # Get current page and store in object
            PageObj = currentInvoiceFile.getPage(i)
            # Extract text from current page
            Text = PageObj.extractText()
            # Search for a match of the do number
            doMatch = re.findall(re_param, Text)
            doMatchSize = len(doMatch)

            # If any DO number numbers were found, proceed
            if doMatchSize > 0:
                doMatchCounter = 0
                doOutputName = ""
                pdfOutMerge = PyPDF2.PdfFileMerger()
                pdfOutMerge.append(currentInvoiceFile)
                # Parse through list of DO numbers found
                for curDONumber in doMatch:
                    tempDoNumber = curDONumber.replace("€", "")  # syntax
                    # loop through all do files in do folder to find match for current do number
                    for curDOFile in os.listdir(doFolderPath):
                        if curDOFile.endswith(".pdf") and tempDoNumber in curDOFile:
                            # append the current do file if a match is found
                            pdfOutMerge.append(PyPDF2.PdfFileReader(doFolderPath + curDOFile))
                            doOutputName += tempDoNumber
                            doMatchCounter += 1
                # case where all matches were found
                if doMatchCounter == doMatchSize:
                    doMatchStatus = "allfound"
                    pdfOutMerge.write(outputFolder + filename.replace(".pdf", "") + "_" + str(doMatchSize) + "DO" + "_" + doOutputName + ".pdf")
                    pdfOutMerge.close()
                # case where some matches were found
                elif doMatchSize > doMatchCounter > 0:
                    doMatchStatus = "somefound"
                    pdfOutMerge.write(incMergedFolder + filename.replace(".pdf", "") + "_" + str(doMatchSize) + "DO" + "_" + doOutputName + ".pdf")
                    pdfOutMerge.close()
                # case where no matches were found
                else:
                    doMatchStatus = "nonefound"
                    copypdfto.copy_file_to(currentInvoiceFile, invNoMatchFolder, filename)
            else:
                continue

        # handle case where no do number found in invoice pdf
        if doMatchStatus == "unfound":
            copypdfto.copy_file_to(currentInvoiceFile, invNoDOFolder, filename)
        else:
            continue

    else:  # ending for IF CONTAINS .PDF IN NAME ELSE
        continue

messagebox.showinfo(title="DO Invoice Merging", message="Merging Complete")




