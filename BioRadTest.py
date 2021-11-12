import tkinter as tk
import tkinter.font as tkFont

class App:
    def __init__(self, root):
        #setting title
        root.title("A1c Batch Data")
        #setting window size
        width=303
        height=254
        screenwidth = root.winfo_screenwidth()
        screenheight = root.winfo_screenheight()
        alignstr = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        root.geometry(alignstr)
        root.resizable(width=False, height=False)

        GButton_924=tk.Button(root)
        GButton_924["activebackground"] = "#009688"
        GButton_924["activeforeground"] = "#2e3445"
        GButton_924["bg"] = "#009688"
        ft = tkFont.Font(family='Times',size=10)
        GButton_924["font"] = ft
        GButton_924["fg"] = "#000000"
        GButton_924["justify"] = "center"
        GButton_924["text"] = "Read BioRad"
        GButton_924.place(x=20,y=30,width=260,height=40)
        GButton_924["command"] = self.GButton_924_command

        GButton_617=tk.Button(root)
        GButton_617["bg"] = "#009688"
        ft = tkFont.Font(family='Times',size=10)
        GButton_617["font"] = ft
        GButton_617["fg"] = "#000000"
        GButton_617["justify"] = "center"
        GButton_617["text"] = "Process Excel Input"
        GButton_617.place(x=20,y=150,width=260,height=40)
        GButton_617["command"] = self.GButton_617_command

        GButton_680=tk.Button(root)
        GButton_680["bg"] = "#87a987"
        ft = tkFont.Font(family='Times',size=10)
        GButton_680["font"] = ft
        GButton_680["fg"] = "#000000"
        GButton_680["justify"] = "center"
        GButton_680["text"] = "Help"
        GButton_680.place(x=20,y=210,width=70,height=26)
        GButton_680["command"] = self.GButton_680_command

        GButton_116=tk.Button(root)
        GButton_116["bg"] = "#009688"
        ft = tkFont.Font(family='Times',size=10)
        GButton_116["font"] = ft
        GButton_116["fg"] = "#000000"
        GButton_116["justify"] = "center"
        GButton_116["text"] = "Open Excel for Input"
        GButton_116.place(x=20,y=90,width=260,height=40)
        GButton_116["command"] = self.GButton_116_command

    def GButton_924_command(self):
        print("command")


    def GButton_617_command(self):
        print("command")


    def GButton_680_command(self):
        print("command")


    def GButton_116_command(self):
        print("command")

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
