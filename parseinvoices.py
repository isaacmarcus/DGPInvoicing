import copypdfto
import PyPDF2
import os
import re
import pandas as pd
from tkinter import messagebox


# Regex patterns for DO Numbers to be found
# 4€May€2020\n7077852
re_param2 = "[\d]{1,2}€[A-Za-z]{3}€[\d]{4}\\n[\d]{7}\\n"
# A€123456
re_param = "A€[\d]{6}"


class ParseInvoices:
    def __init__(self):
        self.unFoundCount = 0
        self.allFoundCount = 0
        self.someFoundCount = 0
        self.noneFoundCount = 0
        self.masterDoList = []

    def parse_invoice_folder(self, invFolderPath, doFolderPath, incMergedFolder, mergedFolder, invNoDOFolder, invNoMatchFolder, root, progressBar):
        # reset Variables for returning details of files processed
        self.unFoundCount = 0
        self.allFoundCount = 0
        self.someFoundCount = 0
        self.noneFoundCount = 0
        # master do list to check against external excel later on
        self.masterDoList = []

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
                    # print(Text)
                    # Search for all matches of do numbers with "A"
                    doMatch = re.findall(re_param, Text)
                    # Search for all matches of do numbers without "A"
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
                    pdfOutMerge.append(currentInvoiceFile)  # append the invoice first as the first page of pdf
                    # Parse through list of DO numbers found
                    for curDONumber in doNumberList:
                        tempDoNumber = curDONumber.replace("€", "")  # syntax
                        # loop through all do files in do folder to find match for current do number
                        for curDOFile in os.listdir(doFolderPath):
                            if curDOFile.endswith(".pdf") and tempDoNumber in curDOFile:
                                print("match found for " + str(tempDoNumber))
                                # append the current do file if a match is found
                                pdfOutMerge.append(PyPDF2.PdfFileReader(doFolderPath + curDOFile))
                                doMatchCounter += 1

                    # case where all matches were found
                    if doMatchCounter == doMatchSize:
                        self.allFoundCount += 1
                        pdfOutMerge.write(
                            mergedFolder + filename.replace(".pdf", "") + "_" + str(doMatchSize) + "DO" + "_"
                            + firstDONumber + ".pdf")
                        pdfOutMerge.close()
                        self.masterDoList += doNumberList  # append do numbers to master list if all matches found
                    # case where some matches were found
                    elif doMatchSize > doMatchCounter > 0:
                        self.someFoundCount += 1
                        pdfOutMerge.write(
                            incMergedFolder + filename.replace(".pdf", "") + "_" + str(doMatchSize) + "DO" +
                            "_" + firstDONumber + ".pdf")
                        pdfOutMerge.close()
                    # case where no matches were found
                    else:
                        self.noneFoundCount += 1
                        copypdfto.copy_file_to(currentInvoiceFile, invNoMatchFolder, filename)
                # handle case where no do number found in invoice pdf
                else:
                    self.unFoundCount += 1
                    copypdfto.copy_file_to(currentInvoiceFile, invNoDOFolder, filename)
                # Update progress bar
                progressBar['value'] = invFileIndex / invFileCount * 100
                root.update_idletasks()  # call to update ui after changing progress bar value
            else:  # ending for IF CONTAINS .PDF IN NAME ELSE
                continue

        # return the count of all files processed
        # self.checkForMissing()
        return [self.allFoundCount, self.someFoundCount, self.noneFoundCount, self.unFoundCount]

    # function to check succesfully merged DOs against external Done List
    def checkForMissing(self, invPath, exDoPath):
        if self.masterDoList:
            for i in range(len(self.masterDoList)):
                self.masterDoList[i] = self.masterDoList[i].replace("€", "")
            masterDF = pd.DataFrame(self.masterDoList, columns=["do_number"])
            # call method to write list of Do numbers to an excel sheet that can be used later
            self.writeToExcel(masterDF, invPath, 'mergedDOList_' + str(self.allFoundCount) + "DO.xlsx", ["do_number"])

            try:
                # read external do file
                exDoDf = pd.read_excel(exDoPath, sheet_name="Done", header=0, dtype=str)
                exDoList = []
                # reformat all the DOs so they have no space included
                for do in exDoDf["D.O. #"].iteritems():
                    exDoList.append(do[1].replace(" ", ""))  # append to exDoList
                    exDoDf.replace(do[1], do[1].replace(" ", ""), inplace=True)  # replace dataframe with do without space
                # rename column of DO number to sth more accessible
                exDoDf.rename(columns={"D.O. #": "do_number"}, inplace=True)
                exDoDf.drop(columns=["DO Status", "Invoice number"], inplace=True)  # drop columns we dont need

                # compare the two lists and create a new list with the ones missing
                missingDoList = list(set(exDoList).symmetric_difference(set(self.masterDoList)))
                excludeMatchDf = exDoDf[~exDoDf["do_number"].isin(self.masterDoList)]  # create a new df without the ones matching
                self.writeToExcel(excludeMatchDf, invPath, 'missingDOList_' + str(excludeMatchDf['do_number'].count()) + "DO.xlsx", excludeMatchDf.columns)

                messagebox.showinfo(title="Cross-check complete",
                                    message="Delivery Orders cross-checked: " + str(len(exDoList)) +
                                    "\nDelivery Orders merged: " + str(len(self.masterDoList)) +
                                    "\nDelivery Orders missing: " + str(excludeMatchDf['do_number'].count()))
            except Exception as e:
                messagebox.showerror(title="Error",
                                    message="Error cross-checking: \n" + str(e) + "\n\nPlease check if you have selected the correct file, and if the sheet / column names are labelled correctly")
        else:
            messagebox.showinfo(title="DOs have not been merged",
                                message="The DO merge has not been complete yet.")

    def writeToExcel(self, listToExport, invPath, fileName, headerList):
        df = pd.DataFrame(listToExport)
        writer = pd.ExcelWriter(invPath + "Merged Invoices/" + fileName, engine='xlsxwriter')
        df.to_excel(writer, index=False, header=headerList)
        writer.save()



