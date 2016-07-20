from tkinter import *
from tkinter.filedialog import *
from tkinter.messagebox import showerror
from tkinter import ttk
from batt_data_parse8 import get_data


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
        self.button.grid(row = 2, column = 2, sticky = E)
        

        # var = IntVar()
        # self.radio = Radiobutton(self, text="Option 1", variable = var, value=1, command=self.sel)
        # self.radio.grid(row = 1, column = 5, sticky = W )
        # self.radio = Radiobutton(self, text="Option 2", variable = var, value=2, command=self.sel)
        # self.radio.grid(row = 2, column = 5, sticky = W )

    def sel(self):
        print (str(var.get()))
       #selection = "You selected the option " + str(var.get())
       #label.config(text = selection)
    
    def do_thingz(self):
        get_data(filelist)
        print("beep")

    def clear_files(self):
        filelist = []
        #self.label = Label(self, text="No Files Selected").grid(column=2, row=1, sticky=W)


    def load_file(self):
        #clear files incase filelist is already full
        self.clear_files()
        file_names_raw = askopenfiles(mode = 'r', filetypes = (("Excel files", "*.xls"), ("All files", "*.*") ))
        for i in range(len(file_names_raw)):
            filelist.append(file_names_raw[i].name)
        #self.label = Label(self, text="Files Selected").grid(column=2, row=1, sticky=W)



if __name__ == "__main__":
    MyFrame().mainloop()