import copypdfto
import PyPDF2
import os
import re


def parse_invoice_folder(invFolderPath, doFolderPath, incMergedFolder, mergedFolder, invNoDOFolder, invNoMatchFolder, re_param):
    # Create Variables for returning details of files processed
    unFoundCount = 0
    allFoundCount = 0
    someFoundCount = 0
    noneFoundCount = 0

    # loop through files in specified folder path
    for filename in os.listdir(invFolderPath):
        # start parsing the file only if its a pdf file
        if filename.endswith(".pdf"):
            # open the current pdf file
            currentInvoiceFile = PyPDF2.PdfFileReader(invFolderPath + filename)

            # get number of pages for looping through later
            NumPages = currentInvoiceFile.getNumPages()

            # Loop through pages of the current pdf to find delivery order number
            doNumberList = []
            for i in range(0, NumPages):
                # Get current page and store in object
                PageObj = currentInvoiceFile.getPage(i)
                # Extract text from current page
                Text = PageObj.extractText()
                # Search for a match of the do number
                doMatch = re.findall(re_param, Text)
                doNumberList += doMatch

            doMatchSize = len(doNumberList)
            # If any DO number numbers were found, proceed
            if doMatchSize > 0:
                doMatchCounter = 0
                doOutputName = ""
                pdfOutMerge = PyPDF2.PdfFileMerger()
                pdfOutMerge.append(currentInvoiceFile)
                # Parse through list of DO numbers found
                for curDONumber in doNumberList:
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
                    allFoundCount += 1
                    pdfOutMerge.write(
                        mergedFolder + filename.replace(".pdf", "") + "_" + str(doMatchSize) + "DO" + "_"
                        + doOutputName + ".pdf")
                    pdfOutMerge.close()
                # case where some matches were found
                elif doMatchSize > doMatchCounter > 0:
                    someFoundCount += 1
                    pdfOutMerge.write(
                        incMergedFolder + filename.replace(".pdf", "") + "_" + str(doMatchSize) + "DO" +
                        "_" + doOutputName + ".pdf")
                    pdfOutMerge.close()
                # case where no matches were found
                else:
                    noneFoundCount += 1
                    copypdfto.copy_file_to(currentInvoiceFile, invNoMatchFolder, filename)
            # handle case where no do number found in invoice pdf
            else:
                unFoundCount += 1
                copypdfto.copy_file_to(currentInvoiceFile, invNoDOFolder, filename)

        else:  # ending for IF CONTAINS .PDF IN NAME ELSE
            continue

    # return the count of all files processed
    return [allFoundCount, someFoundCount, noneFoundCount, unFoundCount]
