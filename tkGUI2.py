from tkinter import *
from tkinter.filedialog import *
from tkinter.messagebox import showerror
from tkinter import ttk
import batt_data_parse8 


filelist = []

class MyFrame(Frame):  
    def __init__(self):
        Frame.__init__(self)
        # Create window
        self.master.title("Battery Testing")
        self.master.rowconfigure(5, weight = 1)
        self.master.columnconfigure(5, weight = 1)
        self.grid(sticky = W+E+N+S)
        # Info text
        self.l1 = Label(self, text = "No Files Selected", width = 20)
        self.l1.grid(row = 1, column = 3, sticky = W)
        self.l2 = Label(self, text = " ", width = 20, wraplength = 140, justify = "center")
        self.l2.grid(row = 4, column = 2, sticky = W)
        self.l3 = Label(self, text = "Save output file as", width = 20)
        self.l3.grid(row = 3, column = 1, sticky = W)
        # Radio buttons to select cell number, defaults to single
        self.v = IntVar()
        self.v.set("1")
        self.r1 = Radiobutton(self, text = "Single Cell", variable = self.v, value = True)
        self.r1.grid(row = 2, column = 1, sticky = W)
        self.r2 = Radiobutton(self, text = "Double Cell", variable = self.v, value = False)
        self.r2.grid(row = 2, column = 2, sticky = W)
        # Browse button
        self.b1  =  Button(self, text = "Add files", command = self.load_file, width = 20)
        self.b1.grid(row = 1, column = 1, sticky = W)
        # Calculate button
        self.b2 = Button(self, text = "Process", command = self.calculate, width = 20)
        self.b2.grid(row = 3, column = 3, sticky = W)
        # Clear files button
        self.b2 = Button(self, text = "Clear selected files", command = self.clear_files)
        self.b2.grid(row = 1, column = 2, sticky = W)        
        # File name entry box
        self.e = Entry(self)
        self.e.grid(row = 3, column = 2, sticky = W)
        self.e.delete(0, END)
        self.e.insert(0, "output.xls")

    def calculate(self):
        if self.e.get() != '':
            fileout = self.e.get()
        else:
            fileout = "output.xls"
        # Make sure the file name is  at least reasonably valid, doesn't guarantee .xls 
        bads = re.compile('[^a-zA-Z0-9_.-]')
        if bads.search(fileout):
             self.l2.config(text="ERROR: Invalid characters", fg = "red")
        else:
            if filelist != []:
                self.l2.config(text="Processing...", fg = "black")
                self.update_idletasks()
                try:
                    batt_data_parse8.get_data(filelist, self.v.get(), fileout)
                    self.l2.config(text="Finished!", fg = "black")
                except:
                    self.l2.config(text="ERROR: Can't write to file Is it in use elsewhere?", fg = "red")

            else:
                self.l2.config(text="ERROR: No files selected", fg = "red")

    def clear_files(self):
        self.l2.config(text="", fg = "black")
        del filelist[:] 
        self.l1.config(text="No Files Selected")


    def load_file(self):
        self.l2.config(text="", fg = "black")

        file_names_raw = askopenfiles(mode = 'r', filetypes = (("Excel files", "*.xls"), 
                                                                ("All files", "*.*") ))
        for i in range(len(file_names_raw)):
            filelist.append(file_names_raw[i].name)
        if len(file_names_raw) != 0:
            self.l1.config(text="%i Files Selected" % len(filelist))



if __name__ == "__main__":
    MyFrame().mainloop()