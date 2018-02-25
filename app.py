from tkinter import *
from stripdata import stripdata

class App:

    def __init__(self, master):
        frame = Frame(master)
        frame.pack()
        frame.option_add("*font", "Times 16")

        self.m1 = Message(master, width="480", text="Click OK to copy all the .docx files in the chosen folder and strip all the metadata from the copies.")
        self.m1.pack()

        self.m2 = Message(master, anchor="w", width="480", text="The copies go inside a folder called 'stripped' placed inside the chosen folder.")
        self.m2.pack(fill="x")

        self.m3 = Message(master, anchor="w", width="480", text="The originals will be unchanged.", justify="left")
        self.m3.pack(fill="x")

        self.e1 = Entry(master, width="40")
        self.e1.pack()
        self.e1.delete(0, END)
        self.e1.insert(0, "V:\\")

        self.ok = Button(text="OK", command=self.strip_data)
        self.ok.pack(side=RIGHT)

        self.cancel = Button(text="Cancel", command=frame.quit)
        self.cancel.pack(side=RIGHT)

    def strip_data(self):
        print(self.e1.get())
        stripdata(self.e1.get())
        root.destroy() #close window when finished

root = Tk()
root.wm_title("Docx Data Stripper")

app = App(root)

root.mainloop()
