import tkinter as tk
import hoverbutton
import startmerging
from tkinter import *
from tkinter.ttk import *
import tkinter.ttk as ttk
from tkinter.filedialog import askdirectory

appVersion = 1.1
print("Merge Invoice Tool " + str(appVersion) + "\n")


def get_folder(folderPath, title):
    pathSelected = askdirectory(title=title) + "/"  # ask user to select invoices folder
    folderPath.set(pathSelected)


# colours
mainBGColour = "#ffffff"
pathSelectedColour = "#f2f2f2"
buttonColour = "#e0e0e0"
hoveringButtonColour = "#c9e1ff"
progressBarColour = "#7db7ff"

# measurements
titleLabelY = 0.025
titleLabelHeight = 0.09
bodyLabelY = titleLabelY + titleLabelHeight
bodyLabelHeight = 0.15
notesLabelFrameY = bodyLabelY + bodyLabelHeight + 0.0075
notesLabelFrameHeight = 0.4
notesLabelHeight = 0.9
startButtonY = 0.875
doPathHolderY = startButtonY - 0.095
invPathHolderY = doPathHolderY - 0.09
progressBarY = 0.875

root = tk.Tk(className="DG Merge Invoices/Delivery Orders")
root.iconbitmap(r"C:\Users\DGP-Yoga1\PycharmProjects\DGPInvoicing\dg_merge.ico")

# Styling for progress bar
style = ttk.Style()
style.theme_use('clam')
style.configure("blue.Horizontal.TProgressbar", troughcolor=buttonColour, background=progressBarColour, foreground=progressBarColour,
                darkcolor=progressBarColour, lightcolor=progressBarColour, bordercolor=buttonColour, bd=0)

invFolderPath = StringVar()
invFolderPath.set("select invoices folder...")
doFolderPath = StringVar()
doFolderPath.set("select delivery orders folder...")

Height = 400
Width = 500

# Main canvas widget for our window
mainCanvas = tk.Canvas(root, height=Height, width=Width, bg=mainBGColour)
mainCanvas.pack()

# Main Frame holding all widgets
mainFrame = tk.Frame(root, bd=0, bg=mainBGColour)
mainFrame.place(relx=0, rely=0, relwidth=1, relheight=1)

# Title Label
titleLabel = tk.Label(mainFrame, bg=mainBGColour, text="Instructions", anchor="nw", font=20)
titleLabel.place(relx=0.025, rely=titleLabelY, relheight=titleLabelHeight, relwidth=1)

# Body Label
bodyLabel = tk.Label(mainFrame, bg=mainBGColour, text="1. Select folder containing invoices"
                                                      "\n2. Select folder containing delivery orders"
                                                      "\n3. Click 'Start' to begin merging"
                     , anchor="nw", justify="left")
bodyLabel.configure(font=("Helvetica", 12))
bodyLabel.place(relx=0.025, rely=bodyLabelY, relheight=bodyLabelHeight, relwidth=1)

# notes Label
notesLabelFrame = tk.LabelFrame(mainFrame, bg=mainBGColour, text=" output notes ", bd=1.5)
notesLabelFrame.place(relx=0.035, rely=notesLabelFrameY, relheight=notesLabelFrameHeight, relwidth=0.93)
notesLabel = tk.Label(notesLabelFrame, bg=mainBGColour, text= "4 folders will be created in the invoice folder selected;"
                                                        "\n'Invoice no Do Number' : contains invoices with no do number found"
                                                        "\n'Invoice no Matching DO' : contains invoices with no matching do file"
                                                        "\n'Merged Invoices' : contains successfully merged invoices"
                                                        "\n'Merged Invoices Incomplete' : contains invoices with some do files found",
                      anchor="nw", justify="left")
notesLabel.configure(font=("Helvetica", 9))
notesLabel.place(relx=0.01, rely=0, relheight=notesLabelHeight, relwidth=0.99)

# ------------------------------------------------
# Invoice frame holding the path and button for it
invPathFrame = tk.Frame(mainFrame, bg=mainBGColour, bd=5)
invPathFrame.place(relx=0.025, rely=invPathHolderY, relheight=0.0975, relwidth=0.95)
# invoice path name box
invPathLabelFrame = tk.LabelFrame(invPathFrame, bg=pathSelectedColour, bd=0.5)
invPathLabelFrame.place(relx=0, rely=0.1, relheight=0.9, relwidth=0.75)
# invoice label holding path name
invPathLabel = tk.Label(invPathLabelFrame, textvariable=invFolderPath, bg=pathSelectedColour, font=("Helvetica", 8))
invPathLabel.config(fg="#666666")
invPathLabel.place(relx=0.025, rely=0.1)
# invoice folder button for browsing
invPathButton = hoverbutton.HoverButton(invPathFrame, text="Browse", relief="flat", bd=1.25, bg=buttonColour,
                                        activebackground=hoveringButtonColour, command=lambda: get_folder(invFolderPath, "Select invoices folder"))
invPathButton.place(relx=0.775, rely=0.1, relheight=0.9, relwidth=0.225)

# --------------------------------------------------------
# DELIVERY ORDER frame holding the path and button for it
doPathFrame = tk.Frame(mainFrame, bg=mainBGColour, bd=5)
doPathFrame.place(relx=0.025, rely=doPathHolderY, relheight=0.0975, relwidth=0.95)
# DO path name box
doPathLabelFrame = tk.LabelFrame(doPathFrame, bg=pathSelectedColour, bd=0.5)
doPathLabelFrame.place(relx=0, rely=0.1, relheight=0.9, relwidth=0.75)
# DO label holding path name
doPathLabel = tk.Label(doPathLabelFrame, textvariable=doFolderPath, bg=pathSelectedColour, fg="#666666", font=("Helvetica", 8))
doPathLabel.place(relx=0.025, rely=0.1)
# DO folder button for browsing
doPathButton = hoverbutton.HoverButton(doPathFrame, text="Browse", bd=1.25, relief="flat", bg=buttonColour,
                                       activebackground=hoveringButtonColour, command=lambda: get_folder(doFolderPath, "Select delivery orders folder"))
doPathButton.place(relx=0.775, rely=0.1, relheight=0.9, relwidth=0.225)

# Start button Frame to start running processing
startButtonFrame = tk.Frame(mainFrame, bg=mainBGColour, bd=5)
startButtonFrame.place(relx=0.025, rely=startButtonY, relheight=0.0975, relwidth=0.95)
# Progress bar
progressBar = Progressbar(startButtonFrame, orient=HORIZONTAL, style="blue.Horizontal.TProgressbar", length=100, mode="determinate")
progressBar.place(relx=0, rely=0.1, relheight=0.9, relwidth=0.75)
# Start button Widget
startButton = hoverbutton.HoverButton(startButtonFrame, text="Start", bd=1.25, relief="flat", bg=buttonColour,
                                      activebackground=hoveringButtonColour, command=lambda: startmerging.start_merging(invFolderPath.get(), doFolderPath.get(), root, progressBar))
startButton.place(relx=0.775, rely=0.1, relheight=0.9, relwidth=0.225)


root.mainloop()
