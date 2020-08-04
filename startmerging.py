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
        incMergedName = "Merged Invoices Incomplete"
        invNoDoName = "Invoice no Do Number"
        invNoMatchName = "Invoice no Matching DO"

        # make a directory for invoices missing DO as well as completed merged invoices
        try:
            # TODO fix bug where if one of the folders has already been created, it stops it entirely, or does not create remaining folders
            print("Creating output folders...")
            os.mkdir(self.invFolderPath + incMergedName)
            os.mkdir(self.invFolderPath + mergedName)
            os.mkdir(self.invFolderPath + mergedName + "/imgversion")
            os.mkdir(self.invFolderPath + invNoDoName)
            os.mkdir(self.invFolderPath + invNoMatchName)
        except Exception as e:
            print("Output directory already exists, program will not create them again" + "\n")

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



