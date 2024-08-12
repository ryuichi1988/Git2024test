from tkinter import *
from tkinter import messagebox
from tkinter import ttk
root = Tk()
frm = ttk.Frame(root, padding=10,borderwidth=30)
frm.grid()

def msg_show(message):
    messagebox.showinfo("メッセージ", message)


ttk.Label(frm, text="Hello World!").grid(column=0, row=0)

ttk.Button(frm, text="SHOW ME", command=lambda:msg_show("you are great")).grid(column=1, row=0,pady=10)
ttk.Button(frm, text="testbutton1", command=lambda:msg_show("you are great")).grid(column=1, row=2,pady=10)
ttk.Button(frm, text="testbutton2", command=lambda:msg_show("you are great")).grid(column=1, row=3,pady=10)
ttk.Button(frm, text="Quit", command=root.destroy).grid(column=1, row=4,pady=10)
def center_window(width=300, height=200):
    # get screen width and height
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()

    # calculate position x and y coordinates
    x = (screen_width/2) - (width/2)
    y = (screen_height/2) - (height/2)
    root.geometry('%dx%d+%d+%d' % (width, height, x, y))

center_window(400, 300)
root.mainloop()