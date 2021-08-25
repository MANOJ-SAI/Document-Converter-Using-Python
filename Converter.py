#!/usr/bin/env python
# coding: utf-8

# # PDF TO DOCX AND DOCX TO PDF CONVERTER IN GUI IN PYTHON 

# In[5]:


#Import Required Packages 
from tkinter import *
import tkinter as tk
from PIL import ImageTk, Image
import pathlib
import time
from tkinter import messagebox  
from pdf2docx import Converter
from docx2pdf import convert as ct
import tkinter.ttk as ttk
from tkinter.filedialog import askopenfilename,asksaveasfile,askdirectory



top=Tk()
top.wm_iconbitmap(r"C:\Users\MS\Desktop\mini-projects\web-scarp\ms.ico")
top.title("File Converter")
top.configure(bg='red')
top.geometry('600x600')
top.resizable(0,0)
photo_p =Image.open(r'C:\Users\MS\Desktop\mini-projects\web-scarp\hacker.jpg')
img=ImageTk.PhotoImage(photo_p)
panel=Label(image=img)
panel.configure(width=600,height=600)
panel.place(x=0,y=0)
k=str()
sl=StringVar()
file=StringVar()
sk = ttk.Style()
sk.configure("black.Horizontal.TProgressbar", foreground='green', background='#000000')
c=Canvas(top,width=300,height=100,bg="#000000")
c.place(x=230,y=390)
te=c.create_text(150,50,fill="white",font=('Square Sans Serif 7',20,'bold'),text="")
load=ttk.Progressbar(c,orient=HORIZONTAL,length=300,mode='determinate')


# Open-file is a function used  for open the given file
def open_file(n,e):
    btn2.config(state="normal")
    file= askopenfilename(filetypes =[(n,e)])
    if file is not None:
        fe.delete(0,END)
        fe.insert(0,file)
        if pathlib.Path(fe.get()).suffix == '.pdf' or pathlib.Path(fe.get()).suffix=='.docx' :
            pass
        else:
            messagebox.showerror("error","Invalid-extension file")
    else:
        file=None
        messagebox.showerror("error","No file") 
        
# Savedl is a function used for after conversion of given file which location we save the file
def savedl():
    sl=askdirectory()
    if sl is not None:
        sf.delete(0,END)
        sf.insert(0,sl)
        btn2.place(x=270,y=350)
        
##Step function is used Loading Bar
def step():
    if load['value']==100:
        load['value']=0
        c.delete('all')
    c.create_text(150,50,fill="white",font=('Square Sans Serif 7',20,'bold'),text="Loading...")
    for x in range(5):
        load['value']+=20
        top.update_idletasks()
        time.sleep(1)
        
## Stop function is used for stop the Loading Bar
def stop():
    load.stop()
    
## Convert function is used for Convert the file into pdf to word either pdf to word
def convert(f,s):
    try:
    
        if clicked.get()=="PDF TO DOCX":
            ##For PDF TO DOCX Conversion
            pdf_file = f.get()
            docx_file = s.get()+"/"+pathlib.Path(f.get()).stem+".docx"
            cv = Converter(pdf_file)
            cv.convert(docx_file, start=0, end=None)
            cv.close()
        elif clicked.get() == "DOCX TO PDF":
            ##For DOCX TO PDF Conversion
            fp=f.get()
            a=s.get()+"/"+pathlib.Path(f.get()).stem+".pdf"
            print("a="+a)
            ct(fp, a)
        
        load.place(x=0,y=0)
        step()
        c.delete('all')
        c.create_text(150,50,fill="white",font=('Square Sans Serif 7',20,'bold'),text="Completed!!")
        fe.delete(0,END)
        sf.delete(0,END)
        btn2.configure(state='disabled')
    except Exception as e:
        messagebox.showerror("error",e)

#show function is used for Diaplay the Converter name
def show(s):
    label.config( text = s )
    
#Change_dropdown is used for Choosing the options in drop-down list 
def change_dropdown(*args):
    c.delete('all')
    if clicked.get() ==  "PDF TO DOCX":
        show("PDF to DOCX CONVERTER")
        name="Pdf files"
        ext="*.pdf"
    elif clicked.get() == "DOCX TO PDF":
        show("DOCX TO PDF CONVERTER")
        name="word file"
        ext="*.docx"
    b = Button(top, text='Browse',bg="#2F93E1",activebackground="green",font=('arial',12,'bold'),command=lambda:open_file(name,ext))
    b.config(width=10)
    b.place(x=320,y=230)
    
    

#Converter label
l1 = Label( top ,width=9, text = "Converter:",bg='#5B6458',fg='#8EF700',font=('arial',12,'bold'))
l1.place(x=60,y=150)
l2=Label(top, width=12,text="File Location :",bg='#5B6458',fg='#8EF700',font=('arial',12,'bold'))
l2.place(x=60,y=200)


#Saved Location Label
l3=Label(top,width=15,text="Saved Location :",bg='#5B6458',fg='#8EF700',font=('arial',12,'bold'))
l3.place(x=60,y=270)

#give the file path
fe=Entry(top,textvariable=file,bg="black",fg="red",font=('calibri',12,'bold'),width=40)
fe.place(x=230,y=200)

#give the folder path
sf=Entry(top,textvariable=sl,bg="black",fg="green",font=('calibri',12,'bold'),width=40)
sf.place(x=230,y=270)

##button-Browse the file
btn1= Button(top, text='Browse',bg="#2F93E1",activebackground="green",font=('arial',12,'bold'),command=lambda:savedl())
btn1.config(width=10)
btn1.place(x=320,y=300)

##Drop-down list options
options=['Select',
         'PDF TO DOCX',
         'DOCX TO PDF']
clicked = StringVar()
clicked.set("Select")
drop = OptionMenu( top , clicked , *options )
drop.config(width=40)
drop.place(x=230,y=150)
clicked.trace('w', change_dropdown)



label = tk.Label( top , text = "",bg="#000000",fg='red',font=('Square Sans Serif 7',20,'bold'))
label.place(x=90,y=20)

#button-Convert
btn2= Button(top, text='convert',bg="#0BF6A8",activebackground="green",font=('arial',12,'bold'),command=lambda:convert(fe,sf))
btn2.config(state='normal',width=20)




top.mainloop()

