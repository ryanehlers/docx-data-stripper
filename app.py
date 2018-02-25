from tkinter import *
from stripdata import stripdata

class App:

    def __init__(self, master):
        frame = Frame(master)
        frame.pack()
        frame.option_add("*font", "Times 16")

        self.w = Message(master, width="480", text="Click OK to copy all the .docx files in the chosen folder and strip all the metadata from the copies.")
        self.w.pack()

        self.x = Message(master, anchor="w", width="480", text="The copies go inside a folder called 'stripped' placed inside the chosen folder.")
        self.x.pack(fill="x")

        self.x = Message(master, anchor="w", width="480", text="The originals will be unchanged.", justify="left")
        self.x.pack(fill="x")

        self.e = Entry(master, width="40")
        self.e.pack()
        self.e.delete(0, END)
        self.e.insert(0, "V:\\")

        self.hi_there = Button(text="OK", command=self.strip_data)
        self.hi_there.pack(side=RIGHT)

        self.button = Button(text="Cancel", command=frame.quit)
        self.button.pack(side=RIGHT)

    def strip_data(self):
        print(self.e.get())
        stripdata(self.e.get())
        root.destroy() #close window when finished

root = Tk()
root.wm_title("Docx Data Stripper")

app = App(root)

root.mainloop()
