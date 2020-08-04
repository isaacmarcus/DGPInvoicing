import resourcepath as rp
import tkinter as tk
import hoverbutton
import startmerging
import tkinter.ttk as ttk
from tkinter.filedialog import askdirectory, askopenfilename

appVersion = 1.9
print("Merge Invoice Tool " + str(appVersion) + "\n")

# colours
mainBGColour = "#ffffff"
pathSelectedColour = "#f2f2f2"
buttonColour = "#e0e0e0"
hoveringButtonColour = "#c9e1ff"
progressBarColour = "#7db7ff"

# default constants
holderHeight = 40
relHolderHeight = 0.075

# measurements
Height = 550
Width = 550

titleLabelY = 0.025
titleLabelHeight = 0.09
bodyLabelY = titleLabelY + titleLabelHeight
bodyLabelHeight = 0.15
notesLabelFrameY = bodyLabelY + bodyLabelHeight + 0.0075
notesLabelFrameHeight = 0.375
notesLabelHeight = 0.95
exDoPathHolderY = 0.875
exDoPathHolderAbsY = Height - 45
startButtonY = exDoPathHolderY - relHolderHeight
startButtonAbsY = exDoPathHolderAbsY - 43
progressBarY = startButtonY
doPathHolderY = startButtonY - relHolderHeight
doPathHolderAbsY = startButtonAbsY - 43
invPathHolderY = doPathHolderY - relHolderHeight
invPathHolderAbsY = doPathHolderAbsY - 43


class Window(tk.Frame):
    def __init__(self, master=None):
        tk.Frame.__init__(self, master)
        self.master = master

        menu = tk.Menu(self.master)
        # remove tear off in menu
        menu.option_add('*tearOff', False)
        self.master.config(menu=menu)

        # set up variables for email/pass/port entries
        self.emailString = tk.StringVar()
        self.passString = tk.StringVar()
        self.subjectString = tk.StringVar()
        self.subjectString.set("Type subject here...")
        self.bodyString = tk.StringVar()
        self.bodyString.set("Type body here...")

        # create file drop down menu
        self.fileMenu = tk.Menu(menu)
        self.fileMenu.add_command(label="Email Login", command=self.createEmailWindow)
        self.fileMenu.add_command(label="Exit", command=self.exitProgram)
        # cascade file drop down to main menu line
        menu.add_cascade(label="File", menu=self.fileMenu)

    def exitProgram(self):
        exit()

    def createEmailWindow(self):
        self.emailWindow = tk.Toplevel(self.master, width=Width, height=Height)
        self.emailWindow.wm_title("Configure Email")
        self.emailWindow.resizable(False, False)
        self.emailWindowTitle = tk.Label(self.emailWindow, text="Configure Emailer", anchor='w', font=20)
        self.emailWindowTitle.grid(row=0, column=0, padx=(7, 10), pady=(10, 10), sticky="w")
        self.subjectText = tk.Entry(self.emailWindow, bd=1, width=50, textvariable=self.subjectString, font="Calibri 10")
        self.subjectText.grid(row=1, column=0, padx=(10, 10), pady=(0, 10), sticky="w")
        self.bodyText = tk.Text(self.emailWindow, height=15, width=50, font="Calibri 10")
        self.bodyText.grid(row=2, column=0, padx=(10, 10), pady=(0, 10), sticky="w")
        self.emailSaveButton = hoverbutton.HoverButton(self.emailWindow, text="Save", relief="flat", bd=1.25,
                                                       bg=buttonColour,
                                                       activebackground=hoveringButtonColour,
                                                       command=lambda: self.updateVariables(),
                                                       width=10)
        self.emailSaveButton.grid(row=3, column=0, padx=(10, 10), pady=(0, 10), sticky='w')
        self.emailSendButton = hoverbutton.HoverButton(self.emailWindow, text="Send", relief="flat", bd=1.25,
                                                       bg=buttonColour,
                                                       activebackground=hoveringButtonColour,
                                                       command=lambda: self.updateVariables(),
                                                       width=10)
        self.emailSendButton.grid(row=3, column=0, padx=(10, 10), pady=(0, 10), sticky='e')
        # TODO finish up UI for email login

    def updateVariables(self):
        self.subjectString.set(self.subjectText.get())
        self.bodyString.set(self.bodyText.get("1.0", 'end-1c'))
        # print(self.bodyString.get())


class MainGui:
    def __init__(self):
        self.root = tk.Tk(className="DG Merge Invoices/Delivery Orders")
        self.root.resizable(False, False)
        try:
            self.root.iconphoto(True, tk.PhotoImage(file=rp.resource_path("dg_merge_png.png")))
        except Exception as ex:
            print(ex)
        self.mergeObj = startmerging.MergeHandler()
        # Setup UI
        self.setup_gui()

        # self.root.mainloop()

    def setup_gui(self):

        # TODO fully implement email login using smtplib

        # create menu bar at the top of window
        self.menuWindow = Window(self.root)

        # Styling for progress bar
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.style.configure("blue.Horizontal.TProgressbar", troughcolor=buttonColour, background=progressBarColour,
                             foreground=progressBarColour,
                             darkcolor=progressBarColour, lightcolor=progressBarColour, bordercolor=buttonColour, bd=0)

        # set up string variables for inv,do and external do path
        self.invFolderPath = tk.StringVar()
        self.invFolderPath.set("select invoices folder...")
        self.doFolderPath = tk.StringVar()
        self.doFolderPath.set("select delivery orders folder...")
        self.exDoFilePath = tk.StringVar()
        self.exDoFilePath.set("select external delivery order list...")

        # Main canvas widget for our window
        self.mainCanvas = tk.Canvas(self.root, height=Height, width=Width, bg=mainBGColour)
        self.mainCanvas.pack()

        # Main Frame holding all widgets
        self.mainFrame = tk.Frame(self.root, bd=0, bg=mainBGColour)
        self.mainFrame.place(relx=0, rely=0, relwidth=1, relheight=1)

        # Title Label
        self.titleLabel = tk.Label(self.mainFrame, bg=mainBGColour, text="Instructions", anchor="nw", font=20)
        self.titleLabel.place(relx=0.025, rely=titleLabelY, relheight=titleLabelHeight, relwidth=1)

        # Body Label
        self.bodyLabel = tk.Label(self.mainFrame, bg=mainBGColour, text="1. Select folder containing invoices"
                                                                        "\n2. Select folder containing delivery orders"
                                                                        "\n3. Click 'Start' to begin merging"
                                  , anchor="nw", justify="left")
        self.bodyLabel.configure(font=("Helvetica", 12))
        self.bodyLabel.place(relx=0.025, rely=bodyLabelY, relheight=bodyLabelHeight, relwidth=1)

        # notes Label
        self.notesLabelFrame = tk.LabelFrame(self.mainFrame, bg=mainBGColour, text=" output notes ", bd=1.5)
        self.notesLabelFrame.place(relx=0.035, rely=notesLabelFrameY, relheight=notesLabelFrameHeight, relwidth=0.93)
        self.notesLabel = tk.Label(self.notesLabelFrame, bg=mainBGColour,
                                   text="4 folders will be created in the invoice folder selected;"
                                        "\n'Invoice no Do Number' : contains invoices with no do number found"
                                        "\n'Invoice no Matching DO' : contains invoices with no matching do file"
                                        "\n'Merged Invoices' : contains successfully merged invoices"
                                        "\n'Merged Invoices Incomplete' : contains invoices with some do files found"
                                        "\n\nCross-checker will output two excel files upon completion:"
                                        "\n1. mergedDOList - containing DO numbers successfully merged"
                                        "\n2. missingDOList - containing DO numbers that are missing",
                                   anchor="nw", justify="left")
        self.notesLabel.configure(font=("Helvetica", 9))
        self.notesLabel.place(relx=0.01, rely=0, relheight=notesLabelHeight, relwidth=0.99)

        # ------------------------------------------------
        # Invoice frame holding the path and button for it
        self.invPathFrame = tk.Frame(self.mainFrame, bg=mainBGColour, bd=5)
        self.invPathFrame.place(relx=0.025, y=invPathHolderAbsY, relheight=relHolderHeight, relwidth=0.95)
        # invoice path name box
        self.invPathLabelFrame = tk.LabelFrame(self.invPathFrame, bg=pathSelectedColour, bd=0.5)
        self.invPathLabelFrame.place(relx=0, rely=0.1, relheight=0.9, relwidth=0.75)
        # invoice label holding path name
        self.invPathLabel = tk.Label(self.invPathLabelFrame, textvariable=self.invFolderPath, bg=pathSelectedColour,
                                     font=("Helvetica", 8))
        self.invPathLabel.config(fg="#666666")
        self.invPathLabel.place(relx=0.025, rely=0.1)
        # invoice folder button for browsing
        self.invPathButton = hoverbutton.HoverButton(self.invPathFrame, text="Browse", relief="flat", bd=1.25,
                                                     bg=buttonColour,
                                                     activebackground=hoveringButtonColour,
                                                     command=lambda: self.get_folder(self.invFolderPath,
                                                                                     "Select invoices folder"))
        self.invPathButton.place(relx=0.775, rely=0.1, relheight=0.9, relwidth=0.225)

        # --------------------------------------------------------
        # DELIVERY ORDER frame holding the path and button for it
        self.doPathFrame = tk.Frame(self.mainFrame, bg=mainBGColour, bd=5)
        self.doPathFrame.place(relx=0.025, y=doPathHolderAbsY, relheight=relHolderHeight, relwidth=0.95)
        # DO path name box
        self.doPathLabelFrame = tk.LabelFrame(self.doPathFrame, bg=pathSelectedColour, bd=0.5)
        self.doPathLabelFrame.place(relx=0, rely=0.1, relheight=0.9, relwidth=0.75)
        # DO label holding path name
        self.doPathLabel = tk.Label(self.doPathLabelFrame, textvariable=self.doFolderPath, bg=pathSelectedColour,
                                    fg="#666666",
                                    font=("Helvetica", 8))
        self.doPathLabel.place(relx=0.025, rely=0.1)
        # DO folder button for browsing
        self.doPathButton = hoverbutton.HoverButton(self.doPathFrame, text="Browse", bd=1.25, relief="flat",
                                                    bg=buttonColour,
                                                    activebackground=hoveringButtonColour,
                                                    command=lambda: self.get_folder(self.doFolderPath,
                                                                                    "Select delivery orders folder"))
        self.doPathButton.place(relx=0.775, rely=0.1, relheight=0.9, relwidth=0.225)

        # Start button Frame to start running processing
        self.startButtonFrame = tk.Frame(self.mainFrame, bg=mainBGColour, bd=5)
        self.startButtonFrame.place(relx=0.025, y=startButtonAbsY, relheight=relHolderHeight, relwidth=0.95)
        # Progress bar
        self.progressBar = ttk.Progressbar(self.startButtonFrame, orient='horizontal',
                                           style="blue.Horizontal.TProgressbar",
                                           length=100, mode="determinate")
        self.progressBar.place(relx=0, rely=0.1, relheight=0.9, relwidth=0.75)
        # Start button Widget
        self.startButton = hoverbutton.HoverButton(self.startButtonFrame, text="Start", bd=1.25, relief="flat",
                                                   bg=buttonColour,
                                                   activebackground=hoveringButtonColour,
                                                   command=lambda: self.mergeObj.start_merging(self.invFolderPath.get(),
                                                                                               self.doFolderPath.get(),
                                                                                               self.root,
                                                                                               self.progressBar))
        self.startButton.place(relx=0.775, rely=0.1, relheight=0.9, relwidth=0.225)

        # --------------------------------------------------------
        # EXTERNAL DO EXCEL frame holding the path and button for it
        self.exDoPathFrame = tk.Frame(self.mainFrame, bg=mainBGColour, bd=5)
        self.exDoPathFrame.place(relx=0.025, y=exDoPathHolderAbsY, relwidth=0.95, relheight=relHolderHeight)
        # exDO path name box
        self.exDoPathLabelFrame = tk.LabelFrame(self.exDoPathFrame, bg=pathSelectedColour, bd=0.5)
        self.exDoPathLabelFrame.place(relx=0, rely=0.1, relheight=0.9, relwidth=0.5)
        # exDO label holding path name
        self.exDoPathLabel = tk.Label(self.exDoPathLabelFrame, textvariable=self.exDoFilePath, bg=pathSelectedColour,
                                      fg="#666666",
                                      font=("Helvetica", 8))
        self.exDoPathLabel.place(relx=0.025, rely=0.1)
        # exDO file button for browsing
        self.exDoPathButton = hoverbutton.HoverButton(self.exDoPathFrame, text="Browse", bd=1.25, relief="flat",
                                                      bg=buttonColour,
                                                      activebackground=hoveringButtonColour,
                                                      command=lambda: self.get_file(self.exDoFilePath,
                                                                                    "Select external delivery order file..."))
        self.exDoPathButton.place(relx=0.525, rely=0.1, relheight=0.9, relwidth=0.225)
        # exDO button to run check
        self.exDoPathCheckButton = hoverbutton.HoverButton(self.exDoPathFrame, text="Check", bd=1.25, relief="flat",
                                                           bg=buttonColour,
                                                           activebackground=hoveringButtonColour,
                                                           command=lambda: self.mergeObj.checkExternalDo(
                                                               self.invFolderPath.get(), self.exDoFilePath.get()))
        self.exDoPathCheckButton.place(relx=0.775, rely=0.1, relheight=0.9, relwidth=0.225)

    def get_folder(self, folderPath, title):
        pathSelected = askdirectory(title=title) + "/"  # ask user to select invoices folder
        folderPath.set(pathSelected)

    def get_file(self, filePath, title):
        pathSelected = askopenfilename(title=title, filetypes=[
            ("Excel Files", "*.xlsm *.xls *.xlsx *.xlsb")])  # ask user to select invoices folder
        filePath.set(pathSelected)


if __name__ == "__main__":
    ui = MainGui()
    ui.root.mainloop()

# pyinstaller.exe -w --onefile --clean --icon dg_merge.ico -n MergeInvoiceTool1.5 maingui.py
