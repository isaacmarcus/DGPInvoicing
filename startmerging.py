# import packages
import os
from tkinter import messagebox

# local packages
import parseinvoices


class MergeHandler:
    def __init__(self):
        self.invFolderPath = ""
        self.doFolderPath = ""
        self.invoiceParser = parseinvoices.ParseInvoices()
        self.masterDict = {}

    # Function to process invoices in folder
    def start_merging(self, invFolderSelected, doFolderSelected, root, progressBar):
        self.invFolderPath = str(invFolderSelected)
        self.doFolderPath = str(doFolderSelected)

        if self.invFolderPath == "select invoices folder..." or self.invFolderPath == "/":
            print("No invoice folder selected...")
            return messagebox.showinfo(title="", message="Please select an invoice folder...")
        elif self.doFolderPath == "select delivery orders folder..." or self.invFolderPath == "/":
            print("No DO folder selected...")
            return messagebox.showinfo(title="", message="Please select a delivery orders folder...")
        else:
            print("Invoice folder path: " + self.invFolderPath)
            print("DO folder path: " + self.doFolderPath + "\n")

        # Names for output folders to be defined here;
        mergedName = "Merged Invoices"
        mergedNameIMG = mergedName + "/imgversion"
        incMergedName = "Merged Invoices Incomplete"
        invNoDoName = "Invoice no Do Number"
        invNoMatchName = "Invoice no Matching DO"
        folderList = [mergedName, mergedNameIMG, incMergedName, invNoDoName, invNoMatchName]

        # make a directory for invoices missing DO as well as completed merged invoices
        try:
            print("Creating output folders...")
            for folder in folderList:
                if not os.path.exists(self.invFolderPath + folder):
                    os.mkdir(self.invFolderPath + folder)
        except Exception as e:
            print("Error creating folders: " + e + "\n")

        # Create absolute folder path name variables
        incMergedFolder = self.invFolderPath + incMergedName + "/"
        mergedFolder = self.invFolderPath + mergedName + "/"
        invNoDOFolder = self.invFolderPath + invNoDoName + "/"
        invNoMatchFolder = self.invFolderPath + invNoMatchName + "/"

        # Function to parse invoices
        parsedInfo = self.invoiceParser.parse_invoice_folder(self.invFolderPath, self.doFolderPath, incMergedFolder, mergedFolder,
                                                        invNoDOFolder, invNoMatchFolder, root, progressBar)

        messagebox.showinfo(title="DO Invoice Merging Complete", message="Invoices merged with no errors: " + str(parsedInfo[0])
                            + "\nInvoices merged with some missing DOs: " + str(parsedInfo[1]) + "\n"
                            + "Invoices with no matching DOs: " + str(parsedInfo[2]) + "\n"
                            + "Invoices with no DO Numbers found: " + str(parsedInfo[3]))

        # take master dict from lower down the chain
        self.masterDict = self.invoiceParser.masterDict

    def checkExternalDo(self, invPath, exDoPath):
        if exDoPath == "" or exDoPath == "select external delivery order list...":
            messagebox.showinfo(title="Please select an excel file",
                                message="No excel file containing DOs have been selected yet")
        else:
            self.invoiceParser.checkForMissing(invPath, exDoPath)

    # TODO finish method to send out emails using dictionary, call method from mailtest.py
    def sendEmails(self):
        print("HI")



