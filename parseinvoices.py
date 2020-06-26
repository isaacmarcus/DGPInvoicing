import copypdfto
import PyPDF2
import os
import re

# Regex patterns for DO Numbers to be found
# 4€May€2020\n7077852
re_param2 = "[\d]{1,2}€[A-Za-z]{3}€[\d]{4}\\n[\d]{7}\\n"
# A€123456
re_param = "A€[\d]{6}"


def parse_invoice_folder(invFolderPath, doFolderPath, incMergedFolder, mergedFolder, invNoDOFolder, invNoMatchFolder):
    # Create Variables for returning details of files processed
    unFoundCount = 0
    allFoundCount = 0
    someFoundCount = 0
    noneFoundCount = 0

    # check how many files are to be processed
    invFileCount = 0
    invFileIndex = 0
    for filename in os.listdir(invFolderPath):
        if filename.endswith(".pdf"):
            invFileCount += 1

    print("Now processing all " + str(invFileCount) + " invoices found in folder selected... \n")
    # loop through files in specified folder path
    for filename in os.listdir(invFolderPath):
        # start parsing the file only if its a pdf file
        if filename.endswith(".pdf"):
            # terminal info for process
            invFileIndex += 1
            print("Now processing invoice " + str(invFileIndex) + "/" + str(invFileCount) + "...")

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
                # Search for all matches of do numbers without "A"
                doMatch = re.findall(re_param, Text)
                # Search for all matches of do numbers with "A"
                dateMatch = re.findall(re_param2, Text)
                dateMatch = [newDO[-8:-1] for newDO in dateMatch]  # only take the do number
                # append all found do numbers to master doNumberList
                doNumberList += doMatch
                doNumberList += dateMatch

            doMatchSize = len(doNumberList)  # Store size of do list
            print(str(doMatchSize) + " DO Numbers found in file")
            # If any DO number numbers were found, proceed
            if doMatchSize > 0:
                print("Now finding for matches in DO folder...")
                firstDONumber = doNumberList[0].replace("€", "")
                doMatchCounter = 0
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
                            doMatchCounter += 1

                # case where all matches were found
                if doMatchCounter == doMatchSize:
                    allFoundCount += 1
                    pdfOutMerge.write(
                        mergedFolder + filename.replace(".pdf", "") + "_" + str(doMatchSize) + "DO" + "_"
                        + firstDONumber + ".pdf")
                    pdfOutMerge.close()
                # case where some matches were found
                elif doMatchSize > doMatchCounter > 0:
                    someFoundCount += 1
                    pdfOutMerge.write(
                        incMergedFolder + filename.replace(".pdf", "") + "_" + str(doMatchSize) + "DO" +
                        "_" + firstDONumber + ".pdf")
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
