import os
from PIL import Image, ImageTk
from tkinter import Tk, RIGHT, BOTH, RAISED, Text, X, N, LEFT, Menu, END, filedialog
from tkinter.ttk import Frame, Label, Style, Button, Entry
from test import *

class Example(Frame):

    def __init__(self):
        super().__init__()
        self.assetFile=('./DefaultFile.xls')
        self.initUI()
        
    def initUI(self):

        #Each time a .pack is called, it creates a new frame
        self.master.title("MobileView Assistant")

        self.rowconfigure(1, pad=10)
        self.rowconfigure(2, pad=10)
        self.rowconfigure(3, pad=10)
        self.rowconfigure(4, pad=10)
        self.rowconfigure(5, pad=10)

        self.columnconfigure(1, pad=5)
        self.columnconfigure(2, pad=5)

        #Username area
        
        lbl1 = Label(self, text="Username:", width=10)
        lbl1.grid(row=1, column=1)           
       
        entry1 = Entry(self)
        entry1.grid(row=1, column=2)

        #Password area
        
        lbl2 = Label(self, text="Password:", width=10)
        lbl2.grid(row=2, column=1)        

        entry2 = Entry(self, show="*")
        entry2.grid(row=2, column=2)

        #File area        

        lbl3 = Label(self, text="Asset File:", width=10)
        lbl3.grid(row=3, column=1)
        assetButton = Button(self, text="Browse", command=self.onOpen)
        assetButton.grid(row=3, column=2, sticky="w")
        self.lbl4 = Label(self, text=os.path.basename(self.assetFile))
        self.lbl4.grid(row=4, column=2, columnspan=2, sticky="w")
        
        #self.txt = Text(self)
        #self.txt.pack(padx=5, expand=False, anchor="w")        

        #Task menu
        menubar = Menu(self.master)
        self.master.config(menu=menubar)
        
        fileMenu = Menu(menubar)
        fileMenu.add_command(label="Open", command=self.onOpen)
        fileMenu.add_command(label="Close", command=test)
        menubar.add_cascade(label="File", menu=fileMenu)
        
        #Ok and close button
        self.pack(fill=BOTH, expand=True)
        
        okButton = Button(self, text="Run", command=lambda: self.mainProgram(entry1.get(), entry2.get(), self.assetFile))
        okButton.grid(row=5, column=2, sticky="WE")
        closeButton = Button(self, text="Close", command=quit) #use quit to close window
        closeButton.grid(row=5, column=3, sticky="E")

        self.pack(fill=BOTH, expand=1)
        self.centerWindow()

    def mainProgram(self, username, password, assetfile):
        print(username)
        print(password)
        print(assetfile)
        
    def centerWindow(self):
      
        w = 300
        h = 200

        sw = self.master.winfo_screenwidth()
        sh = self.master.winfo_screenheight()
        
        x = (sw - w)/2
        y = (sh - h)/2
        self.master.geometry('%dx%d+%d+%d' % (w, h, x, y))

    def onOpen(self):

          filename=filedialog.askopenfilename()
          self.lbl4.config(text=os.path.basename(filename))

          self.assetFile=filename
##        ftypes = [('Python files', '*.py'), ('All files', '*')]
##        dlg = filedialog.Open(self, filetypes = ftypes)
##        fl = dlg.show()
##        print(f1)
##        
##        if fl != '':
##            text = self.readFile(fl)
##            self.txt.insert(END, text)
            

    def readFile(self, filename):

        with open(filename, "rb") as f:
            text = f.read()
            
        return text

    def run(self):
        print('I am running')
        
def main():
  
    root = Tk()                
    app = Example()
    root.mainloop()

    def quit(self):
        self.root.destroy()
    
if __name__ == '__main__':
    main()
