# import packages
import os
import tkinter as tk
from tkinter import messagebox
from tkinter.filedialog import askdirectory

# local packages
import parseinvoices

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

# Names for output folders to be defined here;
mergedName = "Merged Invoices"
incMergedName = "Merged Invoices Incomplete"
invNoDoName = "Invoice no Do Number"
invNoMatchName = "Invoice no Matching DO"

# make a directory for invoices missing DO as well as completed merged invoices
try:
    os.mkdir(invFolderPath + incMergedName)
    os.mkdir(invFolderPath + mergedName)
    os.mkdir(invFolderPath + invNoDoName)
    os.mkdir(invFolderPath + invNoMatchName)
except Exception as e:
    print("Directory already exists, no need to create again..")

# Create absolute folder path name variables
incMergedFolder = invFolderPath + incMergedName + "/"
mergedFolder = invFolderPath + mergedName + "/"
invNoDOFolder = invFolderPath + invNoDoName + "/"
invNoMatchFolder = invFolderPath + invNoMatchName + "/"

# define key terms
re_param = "A€[\d]{6}"
# re_param2 = "A[\d]{6}"  # TODO Create key terms for DO numbers without "A" starter

# Function to parse invoices
parseinvoices.parse_invoice_folder(invFolderPath, doFolderPath, incMergedFolder, mergedFolder, invNoDOFolder, invNoMatchFolder, re_param)

messagebox.showinfo(title="DO Invoice Merging", message="Merging Complete")




