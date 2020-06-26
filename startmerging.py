# import packages
import os
import tkinter as tk
from tkinter import messagebox
from tkinter.filedialog import askdirectory

# local packages
import parseinvoices


# Function to process invoices in folder
def start_merging(invFolderSelected, doFolderSelected):
    invFolderPath = str(invFolderSelected)
    doFolderPath = str(doFolderSelected)

    if invFolderPath == "select invoices folder...":
        print("No invoice folder selected...")
        return messagebox.showinfo(title="", message="Please select an invoice folder...")
    elif doFolderPath == "select delivery orders folder...":
        print("No DO folder selected...")
        return messagebox.showinfo(title="", message="Please select a delivery orders folder...")
    else:
        print("Invoice folder path: " + invFolderPath)
        print("DO folder path: " + doFolderPath + "\n")

    # Names for output folders to be defined here;
    mergedName = "Merged Invoices"
    incMergedName = "Merged Invoices Incomplete"
    invNoDoName = "Invoice no Do Number"
    invNoMatchName = "Invoice no Matching DO"

    # make a directory for invoices missing DO as well as completed merged invoices
    try:
        print("Creating output folders...")
        os.mkdir(invFolderPath + incMergedName)
        os.mkdir(invFolderPath + mergedName)
        os.mkdir(invFolderPath + invNoDoName)
        os.mkdir(invFolderPath + invNoMatchName)
    except Exception as e:
        print("Output directory already exists, program will not create them again" + "\n")

    # Create absolute folder path name variables
    incMergedFolder = invFolderPath + incMergedName + "/"
    mergedFolder = invFolderPath + mergedName + "/"
    invNoDOFolder = invFolderPath + invNoDoName + "/"
    invNoMatchFolder = invFolderPath + invNoMatchName + "/"

    # Function to parse invoices
    parsedInfo = parseinvoices.parse_invoice_folder(invFolderPath, doFolderPath, incMergedFolder, mergedFolder,
                                                    invNoDOFolder, invNoMatchFolder)

    messagebox.showinfo(title="DO Invoice Merging Complete", message="Invoices merged with no errors: " + str(parsedInfo[0])
                        + "\nInvoices merged with some missing DOs: " + str(parsedInfo[1]) + "\n"
                        + "Invoices with no matching DOs: " + str(parsedInfo[2]) + "\n"
                        + "Invoices with no DO Numbers found: " + str(parsedInfo[3]))



