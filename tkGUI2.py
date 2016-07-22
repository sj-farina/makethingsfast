from tkinter import *
from tkinter.filedialog import *
from tkinter.messagebox import showerror
from tkinter import ttk
from batt_data_parse8 import get_data


filelist = []

class MyFrame(Frame):  
    def __init__(self):
        Frame.__init__(self)
        # Create window
        self.master.title("Battery Testing")
        self.master.rowconfigure(5, weight = 1)
        self.master.columnconfigure(5, weight = 1)
        self.grid(sticky = W+E+N+S)
        # Browse for files
        self.b1  =  Button(self, text = "Browse", command = self.load_file, width = 10)
        self.b1.grid(row = 1, column = 1, sticky = W)
        self.l1 = Label(self, text = "No Files Selected")
        self.l1.grid(row = 1, column = 2, sticky = W)
        # Radio buttons to select cell number, defaults to single
        self.v = IntVar()
        self.v.set("1")
        self.r1 = Radiobutton(self, text = "Single Cell", variable = self.v, value = True)
        self.r1.grid(row = 1, column = 3, sticky = W)
        self.r2 = Radiobutton(self, text = "Double Cell", variable = self.v, value = False)
        self.r2.grid(row = 2, column = 3, sticky = W)
        # Calculate button
        self.b2 = Button(self, text = "Calculate", command = self.calculate, width = 10)
        self.b2.grid(row = 3, column = 3, sticky = W)
        # Clear button
        self.b2 = Button(self, text = "Clear", command = self.clear_files, width = 10)
        self.b2.grid(row = 2, column = 1, sticky = W)

    def calculate(self):
        get_data(filelist, self.v.get())
        print("File Ready")


    def clear_files(self):
        filelist = []
        self.l1.config(text="Files Removed")


    def load_file(self):
        # Clear files incase filelist is already full
        self.clear_files()
        file_names_raw = askopenfiles(mode = 'r', filetypes = (("Excel files", "*.xls"), 
                                                                ("All files", "*.*") ))
        for i in range(len(file_names_raw)):
            filelist.append(file_names_raw[i].name)
        self.l1.config(text="Files Selected!")



if __name__ == "__main__":
    MyFrame().mainloop()