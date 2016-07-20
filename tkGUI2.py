from tkinter import *
from tkinter.filedialog import *
from tkinter.messagebox import showerror
from tkinter import ttk
from batt_data_parse7 import get_data

filelist = []

class MyFrame(Frame):  
    def __init__(self):
        Frame.__init__(self)

        self.master.title("Battery Testing")
        self.master.rowconfigure(5, weight = 1)
        self.master.columnconfigure(5, weight = 1)
        self.grid(sticky = W+E+N+S)

        self.button  =  Button(self, text = "Browse", command = self.load_file, width = 10)
        self.button.grid(row = 1, column = 1, sticky = W)

        self.button = Button(self, text = "Calculate", command = self.do_thingz, width = 10)
        self.button.grid(row = 2, column = 2, sticky = W)
    
    def do_thingz(self):
        get_data(filelist)
        print("beep")

    def load_file(self):
        file_names_raw = askopenfiles(mode = 'r', filetypes = (("Excel files", "*.xls"), ("All files", "*.*") ))
        for i in range(len(file_names_raw)):
            filelist.append(file_names_raw[i].name)
        #self.label = Label(self, text="Files Selected").grid(column=2, row=1, sticky=W)



if __name__ == "__main__":
    MyFrame().mainloop()