import copypdfto
import PyPDF2
import os
import re

def parse_invoice_folder(invFolderPath, doFolderPath, incMergedFolder, mergedFolder, invNoDOFolder, invNoMatchFolder, re_param):
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
                        pdfOutMerge.write(
                            mergedFolder + filename.replace(".pdf", "") + "_" + str(doMatchSize) + "DO" + "_"
                            + doOutputName + ".pdf")
                        pdfOutMerge.close()
                    # case where some matches were found
                    elif doMatchSize > doMatchCounter > 0:
                        doMatchStatus = "somefound"
                        pdfOutMerge.write(
                            incMergedFolder + filename.replace(".pdf", "") + "_" + str(doMatchSize) + "DO" +
                            "_" + doOutputName + ".pdf")
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

