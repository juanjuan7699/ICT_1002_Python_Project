from tkinter import *
import webbrowser

root = Tk()

new = 1
url = "https://www.google.com"

def openweb():
    webbrowser.open(url,new=new)

Btn = Button(root, text = "This opens Google",command=openweb)
Btn.pack()

root.mainloop()