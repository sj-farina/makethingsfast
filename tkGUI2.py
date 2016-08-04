from tkinter import *
from tkinter.filedialog import *
from tkinter.messagebox import showerror
from tkinter import ttk
import batt_data_parse8 
import webbrowser


filelist = []

ABOUT_TEXT = """Arbin Excel Data Scraper

Created and maintained by Janey Farina
Last Updated 8-4-2016

For source code, bug reporting, feature requests, 
or additional information please visit """

class MyFrame(Frame):  
    def __init__(self):
        Frame.__init__(self)
        # Create window
        self.master.title("Battery Testing")
        self.master.rowconfigure(5, weight = 1)
        self.master.columnconfigure(5, weight = 1)
        self.grid(sticky = W+E+N+S)

        # Menu bar
        self.master.title("Simple menu")
        menubar = Menu(self.master)
        self.master.config(menu=menubar)
        menubar.add_cascade(label="Info", command=self.popup) 

        # Dynamic Labels
        # # of files selected
        self.l1 = Label(self, text = "No Files Selected", width = 20)
        self.l1.grid(row = 1, column = 2, sticky = W)
        # Error message
        self.l2 = Label(self, text = " ", width = 20, wraplength = 140, justify = "center")
        self.l2.grid(row = 9, column = 2, sticky = W)

        # Static Labels
        self.l3 = Label(self, text = "Save output file as", width = 20)
        self.l3.grid(row = 8, column = 1, sticky = W)
        self.l4 = Label(self, text = "Options:", width = 20)
        self.l4.grid(row = 4, column = 2, sticky = W)
        self.l5 = Label(self, text = "cell before transitions", width = 20)
        self.l5.grid(row = 5, column = 3, sticky = E)
        self.l6 = Label(self, text = "cell before capacity    ", width = 20)
        self.l6.grid(row = 6, column = 3, sticky = E)
        self.l7 = Label(self, text = "", width = 20)
        self.l7.grid(row = 7, column = 2, sticky = W)


        # Radio buttons, select cell number
        self.v1 = IntVar()
        self.v1.set("1")
        self.r1 = Radiobutton(self, text = "Single Cell", variable = self.v1, value = True)
        self.r1.grid(row = 5, column = 1, sticky = W)
        self.r2 = Radiobutton(self, text = "Double Cell", variable = self.v1, value = False)
        self.r2.grid(row = 6, column = 1, sticky = W)

        # Check buttons, print names and temp
        self.v2 = IntVar()
        self.v3 = IntVar()
        self.v2.set("1")
        self.v3.set("1")
        self.cb1 = Checkbutton(self, text = "File Names", variable = self.v2)
        self.cb1.grid(row = 5, column = 2, sticky = W)
        self.cb2 = Checkbutton(self, text = "Cell Temp", variable = self.v3)
        self.cb2.grid(row = 6, column = 2, sticky = W)


        # Buttons
        # Browse button
        self.b1  =  Button(self, text = "Add files", command = self.load_file, width = 20)
        self.b1.grid(row = 1, column = 1, sticky = W)
        # Calculate button
        self.b2 = Button(self, text = "Process", command = self.calculate, width = 20)
        self.b2.grid(row = 8, column = 3, sticky = W)
        # Clear files button
        self.b2 = Button(self, text = "Clear selected files", command = self.clear_files, width = 20)
        self.b2.grid(row = 2, column = 1, sticky = W) 

        # Text boxes
        # File name entry box
        self.e1 = Entry(self)
        self.e1.grid(row = 8, column = 2)
        self.e1.delete(0, END)
        self.e1.insert(0, "output.xls")
        # Cell padding option 1
        self.e2 = Entry(self, width = 2)
        self.e2.grid(row = 5, column = 3, sticky = W)
        self.e2.delete(0, END)
        self.e2.insert(0, "1")
        # Cell padding option 2
        self.e3 = Entry(self, width = 2)
        self.e3.grid(row = 6, column = 3, sticky = W)
        self.e3.delete(0, END)
        self.e3.insert(0, "7")

    def calculate(self):
        if self.e1.get() != '':
            fileout = self.e1.get()
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
                    batt_data_parse8.get_data(filelist, self.v1.get(), fileout, self.v2.get(), self.v3.get(), int(self.e2.get()), int(self.e3.get()) )
                    self.l2.config(text="Finished!", fg = "black")
                except:
                    self.l2.config(text="ERROR: Could not print to file", fg = "red")

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
 
    def callback(self, event):
        webbrowser.open_new(r"https://github.com/sj-farina/makethingsfast")

    def popup(self):
        toplevel = Toplevel()
        label1 = Label(toplevel, text=ABOUT_TEXT, height=0, width=50)
        label1.pack()
        link = Label(toplevel, text="https://github.com/sj-farina/makethingsfast", fg="blue", cursor="hand2")
        link.pack()
        link.bind("<Button-1>", self.callback)

if __name__ == "__main__":
    MyFrame().mainloop()