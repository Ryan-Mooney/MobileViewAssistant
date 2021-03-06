#Mobileview Assistant - Asset Finder and Organizer
#Developed by Ryan Mooney, BMET II, at Mercy Medical Center - Cedar Rapids, IA

#Requires IEDriverSoftware and ChromeDriver
    #found here: https://stackoverflow.com/questions/24925095/selenium-python-internet-explorer
    #and here: http://chromedriver.chromium.org/
    #Use 32 bit one
    #Drivers must also be added to PATH to function properly
#To compile this yourself, install pyinstaller and run this in the current directory:
#pyinstaller --paths C:\Windows\WinSxS\x86_microsoft-windows-m..namespace-downlevel_31bf3856ad364e35_10.0.17134.1_none_50c6cb8431e7428f --hidden-import tkinter MVAssistantPopupDialog.py


import os, time, sys
#from PIL import Image, ImageTk
from tkinter import *
from tkinter import filedialog
from tkinter.ttk import Frame, Label, Style, Button, Entry
from assetFileLoader import *
from MobileViewAssetFinder import *
from dbManagement import *
from resultSheetExporter import *
from resultsEmailer import *


#Uncheck this when compiling into standalone app with pyinstaller
#os.chdir("../..")
class MainDialog(Frame):

    def __init__(self, root):
        super().__init__()
        self.assetFile=('./DefaultAssetList.xlsx')
        self.root=root
        self.initUI()
        
    def initUI(self):
        #Task menu
        menubar = Menu(self.master)
        self.master.config(menu=menubar)
        
        fileMenu = Menu(menubar)
        fileMenu.add_command(label="Load Asset File", command=self.onOpen)
        fileMenu.add_command(label="Close", command=sys.exit)
        menubar.add_cascade(label="File", menu=fileMenu)
        
        #Note: Each time a .pack is called, it creates a new frame
        self.master.title("MobileView Assistant")

        self.root.rowconfigure(1, pad=10)
        self.root.rowconfigure(2, pad=10)
        self.root.rowconfigure(3, pad=10)
        self.root.rowconfigure(4, pad=10)
        self.root.rowconfigure(5, pad=10)
        self.root.rowconfigure(6, pad=10)
        self.root.rowconfigure(7, pad=10)
        self.root.rowconfigure(8, pad=10)
        self.root.rowconfigure(9, pad=10)
        self.root.rowconfigure(10, pad=10)
        self.root.rowconfigure(11, pad=10)

        self.root.columnconfigure(1, pad=5)
        self.root.columnconfigure(2, pad=5)
        self.root.columnconfigure(3, pad=5)
        self.root.columnconfigure(4, pad=5)

        #Username area
        lbl1 = Label(self.root, text="MV Username:", width=15)
        lbl1.grid(row=1, column=1)           
       
        MVUsername = Entry(self.root)
        MVUsername.grid(row=1, column=2)

        lbl9 = Label(self.root, text="RSQ Username:", width=15)
        lbl9.grid(row=1, column=3)           
       
        RSQUsername = Entry(self.root)
        RSQUsername.grid(row=1, column=4)

        #Password area
        lbl2 = Label(self.root, text="MV Password:", width=15)
        lbl2.grid(row=2, column=1)        

        MVPassword = Entry(self.root, show="*")
        MVPassword.grid(row=2, column=2)

        lbl10 = Label(self.root, text="RSQ Password:", width=15)
        lbl10.grid(row=2, column=3)        

        RSQPassword = Entry(self.root, show="*")
        RSQPassword.grid(row=2, column=4)

        #File area        
        lbl3 = Label(self.root, text="Asset File:", width=10)
        lbl3.grid(row=3, column=1)
        assetButton = Button(self.root, text="Browse", command=self.onOpen)
        assetButton.grid(row=3, column=2, sticky="w")
        self.lbl4 = Label(self.root, text=os.path.basename(self.assetFile))
        self.lbl4.grid(row=4, column=2, columnspan=2, sticky="w")

        #Trial Type Selector
        lbl7 = Label(self.root, text="Trial Type:", width=10)
        lbl7.grid(row=5, column=1)
        options=["All Assets", "PM Month: January", "PM Month: February", "PM Month: March"
                 , "PM Month: April", "PM Month: May", "PM Month: June", "PM Month: July"
                 , "PM Month: August", "PM Month: September", "PM Month: October",
                 "PM Month: November", "PM Month: December"]
        trial_type=StringVar()
        trial_type.set(options[0])
        trial_selector=OptionMenu(self.root, trial_type, *options)
        trial_selector.grid(row=5, column=2, columnspan=2, sticky="WE")

        #Admin Access Checker
        self.admin_access=IntVar(value=1)
        Checkbutton(self.root, text="I have access to Mobile View Admin.", variable=self.admin_access).grid(row=6, column=1, columnspan=2, sticky=W)

        #Active PMs only Checker
        self.cross_checker=IntVar(value=1)
        Checkbutton(self.root, text="Print only the active PMs.", variable=self.cross_checker).grid(row=7, column=1, columnspan=2, sticky=W)

        #Run a test only Checker
        self.test_case=IntVar(value=0)
        Checkbutton(self.root, text="Run a test trial only.", variable=self.test_case).grid(row=6, column=3, columnspan=2, sticky=W)

        #Email Results Checker
        self.email_results=IntVar(value=0)
        Checkbutton(self.root, text="Email results to ResultsMailingList.txt.", variable=self.email_results).grid(row=7, column=3, columnspan=2, sticky=W)
        
        #Status Indicator
        lbl5 = Label(self.root, text="Status:", width=10)
        lbl5.grid(row=10, column=1)
        self.lbl6 = Label(self.root, text="Awaiting Inputs...", width=50)
        self.lbl6.grid(row=10, column=2, columnspan=5)
        
        #Ok and close button
        
        okButton = Button(self.root, text="Run", command=lambda: self.mainProgram(MVUsername.get(), MVPassword.get(), RSQUsername.get(), RSQPassword.get(),
                    self.assetFile, trial_type.get(), self.admin_access.get(), self.test_case.get(), self.email_results.get(), self.cross_checker.get()))
        okButton.grid(row=11, column=2, sticky="WE")
        closeButton = Button(self.root, text="Close", command=sys.exit) #use quit to close window
        closeButton.grid(row=11, column=3, sticky="E")

        #Copyright Line
        lbl8 = Label(self.root, text="© Ryan Mooney Industries", width=10)
        lbl8.grid(row=12, column=1, columnspan=3, sticky="WE")

        self.centerWindow()

    def mainProgram(self, MVUsername, MVPassword, RSQUsername, RSQPassword, assetfile, trial_type, admin_access, test_case, email_results, cross_checker): 
        #Checks credentials before continuing
        credentials_correct = checkCredentials(MVUsername, MVPassword, RSQUsername, RSQPassword, cross_checker, self.root, self.lbl6)
        if credentials_correct=="NO":
            return()

        #Compiles a dictionary with each asset and its descriptors
        self.lbl6.config(text='Finding Assets...')
        self.root.update()

        #Determines how to find asset locations
        if test_case==1:
            trial_type='TEST'
            assetList, floor_counter, floor_list=get_asset_locations_test(MVUsername, MVPassword, assetListCreator(assetfile, trial_type), self.root, self.lbl6)
        elif admin_access==1:
            assetList, floor_counter, floor_list=get_asset_locations_admin(MVUsername, MVPassword, assetListCreator(assetfile, trial_type), self.root, self.lbl6)
        else:
            assetList, floor_counter, floor_list=get_asset_locations_nonadmin(MVUsername, MVPassword, assetListCreator(assetfile, trial_type), self.root, self.lbl6)
        
        #Saves each new asset data point to database with unique 'trial' number
        self.lbl6.config(text='Saving Data...')
        self.root.update()
        connection=connect()
        assetList, trial=assign_trial_number(assetList, connection, trial_type, test_case)
        save_to_db(assetList, self.root, self.lbl6, connection)
        connection.close()

        #If only active PMs are wanted, we cross check them with an RSQ list and return only the assets with active PMs
        if cross_checker==1:
            self.lbl6.config(text='Cross checking for active PMs...')
            self.root.update()
            activeAssets=crossCheckAssets(assetList, RSQUsername, RSQPassword, self.root, self.lbl6)
            trial_type=trial_type+' Active PMs Only'
        else:
            activeAssets='None'
        
        #Creates and Exports Data to Excel
        self.lbl6.config(text='Exporting to Excel...')
        self.root.update()
        file=exportToExcel(assetList, trial, floor_counter, floor_list, trial_type, activeAssets)

        #Sends results to email addresses if requested
        if email_results==1:
            self.lbl6.config(text='Sending Emails...')
            self.root.update()
            email_file='./ResultsMailingList.txt'
            send_results(file, email_file)
            
        self.lbl6.config(text='Completed!')
        self.root.update()
        os.startfile(file)

        #Emails results, if possible
        
    def centerWindow(self):
      
        w = 475
        h = 325

        sw = self.master.winfo_screenwidth()
        sh = self.master.winfo_screenheight()
        
        x = (sw - w)/2
        y = (sh - h)/2
        self.master.geometry('%dx%d+%d+%d' % (w, h, x, y))

    def onOpen(self):

          filename=filedialog.askopenfilename()
          self.lbl4.config(text=os.path.basename(filename))

          self.assetFile=filename            

    def readFile(self, filename):

        with open(filename, "rb") as f:
            text = f.read()
            
        return text

    def run(self):
        print('I am running')
        
def main():
  
    root = Tk()                
    app = MainDialog(root)
    root.mainloop()

    def quit(self):
        self.root.destroy()
    
if __name__ == '__main__':
    main()
