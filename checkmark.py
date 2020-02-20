from tkinter import * 
from tkinter import ttk 
  
# root 
root = Tk()
root.geometry("400x200")
check_var = IntVar()
check_var.set(1)
# This will depict the features of Simple Checkbutton 
Label(root, text ='Simple Checkbutton').grid(row=0, column=0)
chkbtn1 = Checkbutton(root, text ='Checkbutton1', variable=check_var).grid(row=1, column=0)
chkbtn2 = Checkbutton(root, text ='Checkbutton2', variable=check_var).grid(row=2, column=0)
  
  
# This will depict the features of ttk.Checkbutton 
Label(root, text ='ttk.Checkbutton').grid(row=3, column=0)
chkbtn1 = ttk.Checkbutton(root, text ='Checkbutton1', variable=check_var).grid(row=4, column=0)
chkbtn2 = ttk.Checkbutton(root, text ='Checkbutton2', variable=check_var).grid(row=5, column=0)
  
root.mainloop() 