from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd
import numpy as np


def browse_btn_excel():
    global xlfilename
    try:
        xlfilename = filedialog.askopenfilename(initialdir="/", title="Select file", filetypes=(("excel files", "*.xlsx"), ("all files", "*.*")))
        check = "xlsx" in xlfilename
        if not check:
            messagebox.showerror("error","please select an excel '.xlsx' file")
    except:
        messagebox.showerror("error","please select a file")


def read_sheet(loc,sheet_name,start,end,rnamecol,pcol,ncol):
    pcol = pcol-1
    ncol = ncol-1
    rnamecol=rnamecol-1
    start = start-2
    df = pd.read_excel(loc,sheet_name,skiprows=start,header=None)
    df = df[:end-start]
    df = df.iloc[:, [rnamecol,pcol,ncol]]
    df = df.T
    return df


def get_cagr(df,yrdiff):
    ndarr=df.iloc[[1,2],1:].astype('int64').values
    tminus1yr=ndarr[0]
    tyr=ndarr[1]
    out=np.divide(tyr,tminus1yr)
    exponent=1.0/yrdiff
    out_with_power=(np.power(out,exponent)-1)*100
    data=['CAGR']
    for item in out_with_power:
        data.append(item)
    data_df=pd.DataFrame(data)
    df = pd.concat([df.T, data_df], axis=1)
    return df


root =Tk()
root.geometry('400x275')
textn = Label(root, text="                              ")
textn.grid(row=0)
text1 = Label(root, text="Select Excel file")
text1.grid(row=1)
btn1 = Button(root, text="Browse excel file", command=browse_btn_excel)
btn1.grid(row=1, column=1)
sheetnamelabel = Label(root,text="Enter Sheet Name")
sheetnamelabel.grid(row=2)
sheetnameentry = Entry(root)
sheetnameentry.grid(row=2,column=1)
startrowlabel = Label(root,text="Enter Start row")
startrowlabel.grid(row=3)
startrowentry = Entry(root)
startrowentry.grid(row=3,column=1)
endrowlabel = Label(root, text='Enter End row')
endrowlabel.grid(row=4)
endrowentry = Entry(root)
endrowentry.grid(row=4,column=1)
rnmcollabel = Label(root, text='Enter Column containing row names')
rnmcollabel.grid(row=5)
rnmcolentry = Entry(root)
rnmcolentry.grid(row=5,column=1)
pcollabel = Label(root,text="Enter First Survey Data Column")
pcollabel.grid(row=6)
pcolentry = Entry(root)
pcolentry.grid(row=6,column=1)
ncollabel = Label(root, text='Enter Next Survey Data Column')
ncollabel.grid(row=7)
ncolentry = Entry(root)
ncolentry.grid(row=7,column=1)
yrL= Label(root, text='Enter Year Difference b/w Survey')
yrL.grid(row=8)
yrE=Entry(root)
yrE.grid(row=8,column=1)
outputl= Label(root, text="Enter Output file Name")
outputl.grid(row=9)
outputentry = Entry(root)
outputentry.grid(row=9,column=1)


def submit():
    global startrow
    global endrow
    global yrdiff
    global rnmcol
    global pcol
    global ncol
    global sheetname
    global outputname
    outputname=outputentry.get()
    sheetname=sheetnameentry.get()
    try:
        startrow = int(startrowentry.get())
    except:
        messagebox.showerror("error", "the start row must be an integer")
    try:
        endrow = int(endrowentry.get())
    except:
        messagebox.showerror("error", "the end row must be an integer")
    try:
        yrdiff = int(yrE.get())
    except:
        messagebox.showerror("error", "the Year difference must be an integer")
    try:
        pcol = int(pcolentry.get())
    except:
        messagebox.showerror("error", "the First survey column must be an integer")
    try:
        ncol = int(ncolentry.get())
    except:
        messagebox.showerror("error", "the Next survey column must be an integer")
    try:
        rnmcol = int(rnmcolentry.get())
    except:
        messagebox.showerror("error", "the Row names column must be an integer")
    df=get_cagr(read_sheet(xlfilename,sheetname,startrow,endrow,rnmcol,pcol,ncol),yrdiff)
    with pd.ExcelWriter(outputname + '.xlsx') as writer:  # doctest: +SKIP
        df.to_excel(writer, sheet_name='Sheet1', header=False, index=False)
    root.destroy()


text0 = Label(root, text ="                                  ")
text0.grid(row =10)
submit = Button(root, text="Submit", command=submit)
submit.grid(row=11, column=1)
root.mainloop()


# get_cagr(read_sheet('testfile.xlsx','Sheet1',2,5,1,2,3),5)

