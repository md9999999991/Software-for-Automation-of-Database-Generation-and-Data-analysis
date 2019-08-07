from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd
import matplotlib.pyplot as plt


# -------------- input final organised dataframe ----------------#
def get_plot(df,plot_type):
    if plot_type=='Pie Chart':
        df.plot.pie(subplots=True,autopct='%.2f')
    elif plot_type=='Bar Plot':
        df.plot.bar()
    elif plot_type=='Line Graph':
        df.plot.line()
    plt.show()


def ret_num_df(df,rcdata):
    row_dat=rcdata[0]
    col_dat=rcdata[1]
    # ----------- continuous rows and columns ---------------#
    if row_dat[0]==1 and col_dat[0]==1:
        rows = df.iloc[row_dat[2]:row_dat[3], [row_dat[1]]]
        rows = rows.values
        rnames = []
        for val in rows:
            rnames.append(val[0].strip())
        cols = df.iloc[[col_dat[1]],col_dat[2]:col_dat[3]]
        cols=cols.T
        cols=cols.values
        cnames = []
        for val in cols:
            cnames.append(val[0].strip())
        start_row=row_dat[2]
        end_row=row_dat[3]
        start_col=col_dat[2]
        end_col=col_dat[3]
        df = df.iloc[start_row:end_row,start_col:end_col]
        df =df.astype('int64')
        df.index=rnames
        df.columns =cnames
    # ------- continuous cols and discrete rows----#
    elif row_dat[0]==0 and col_dat[0]==1:
        rows=df.iloc[row_dat[2],row_dat[1]]
        rows = rows.values
        rnames = []
        for val in rows:
            rnames.append(val.strip())
        cols = df.iloc[[col_dat[1]], col_dat[2]:col_dat[3]]
        cols = cols.T
        cols = cols.values
        cnames = []
        for val in cols:
            cnames.append(val[0].strip())
        start_col = col_dat[2]
        end_col = col_dat[3]
        df = df.iloc[row_dat[2], start_col:end_col]
        df = df.astype('int64')
        df.index = rnames
        df.columns = cnames
    # -------- discrete col and rows -------#
    elif row_dat[0]==0 and col_dat[0]==0:
        rows = df.iloc[row_dat[2], row_dat[1]]
        rows = rows.values
        rnames = []
        for val in rows:
            rnames.append(val.strip())
        print(rnames)
        cols = df.iloc[col_dat[1], col_dat[2]]
        cols = cols.values
        cnames = []
        for val in cols:
            cnames.append(val.strip())
        print(cnames)
        df = df.iloc[row_dat[2], col_dat[2]]
        df = df.astype('int64')
        df.index = rnames
        df.columns = cnames
    # ----------- discrete cols and continuous rows-------------#
    elif row_dat[0]==1 and col_dat[0]==0:
        rows = df.iloc[row_dat[2]:row_dat[3], [row_dat[1]]]
        rows = rows.values
        rnames = []
        for val in rows:
            rnames.append(val[0].strip())
        cols = df.iloc[col_dat[1], col_dat[2]]
        cols = cols.values
        cnames = []
        for val in cols:
            cnames.append(val.strip())
        start_row = row_dat[2]
        end_row = row_dat[3]
        df = df.iloc[start_row:end_row, col_dat[2]]
        df = df.astype('int64')
        df.index = rnames
        df.columns = cnames
    print(df)
    return df


def browse_btn_excel():
    global xlfilename
    try:
        xlfilename = filedialog.askopenfilename(initialdir="/", title="Select file",
                                                filetypes=(("excel files", "*.xlsx"), ("all files", "*.*")))
        check = "xlsx" in xlfilename
        if not check:
            messagebox.showerror("error", "please select an excel '.xlsx' file")
    except:
        messagebox.showerror("error", "please select a file")


root = Tk()
OPTIONS=['Discrete Rows and Discrete Columns','Discrete Rows and Continuous Columns','Continuous Rows and Discrete Columns','Continous Rows and Continuous Columns']
variable=StringVar(root)
variable.set(OPTIONS[3])
root.geometry('450x150')
textn = Label(root, text="                              ")
textn.grid(row=0)
loc_sheetsL = Label(root, text='Select Sheet')
loc_sheetsL.grid(row=1)
loc_sheetsB= Button(root,text='Browse', command=browse_btn_excel)
loc_sheetsB.grid(row=1,column=1)
sheet_namL=Label(root,text='Sheet Name')
sheet_namL.grid(row=2)
sheet_namE=Entry(root)
sheet_namE.grid(row=2,column=1)
type_label= Label(root, text=' Row and Column Selection type')
type_label.grid(row=3)
type_entry=OptionMenu(root,variable,*OPTIONS)
type_entry.grid(row=3,column=1)


def submit():
    root.geometry('450x250')
    try:
        print(xlfilename)
    except:
        messagebox.showerror('error','please select a file')
        root.destroy()
    global sheetname
    sheetname = sheet_namE.get()
    chk=sheetname==''
    if chk:
        messagebox.showerror('error','Sheet name cannot be null')
        root.destroy()
    df = pd.read_excel(xlfilename,sheetname, header=None)
    print(df)
    plotoptions = ['Bar Plot', 'Pie Chart','Line Graph']
    var = StringVar(root)
    var.set(plotoptions[0])
    loc_sheetsB.destroy()
    loc_sheetsL.destroy()
    type_entry.destroy()
    type_label.destroy()
    b.destroy()
    text0.destroy()
    sheet_namE.destroy()
    sheet_namL.destroy()
    # ------------------continuous rows and continuous column------------------------#
    if variable.get() == OPTIONS[3]:
        txt=Label(root,text='                      ')
        txt.grid(row=0)
        col_withrow_names=Label(root,text='Column which contains row names')
        col_withrow_names.grid(row=1)
        col_withrow_names_ent=Entry(root)
        col_withrow_names_ent.grid(row=1,column=1)
        srl=Label(root,text='Row number where row names start')
        srl.grid(row=2)
        sre=Entry(root)
        sre.grid(row=2,column=1)
        erl=Label(root,text='Row number where row names end')
        erl.grid(row=3)
        ere=Entry(root)
        ere.grid(row=3,column=1)
        row_withcol_names=Label(root,text='Row which contains column names')
        row_withcol_names.grid(row=4)
        row_withcol_names_ent=Entry(root)
        row_withcol_names_ent.grid(row=4,column=1)
        scl=Label(root,text='Column number where column names start')
        scl.grid(row=5)
        sce=Entry(root)
        sce.grid(row=5,column=1)
        ecl=Label(root,text='Column number where column names end')
        ecl.grid(row=6)
        ece=Entry(root)
        ece.grid(row=6,column=1)
        plt_type_label = Label(root, text=' Plot Type')
        plt_type_label.grid(row=7)
        plot_type = OptionMenu(root, var, *plotoptions)
        plot_type.grid(row=7, column=1)

        def intrnl_submitcnc():
            try:
                cnm_row=int(row_withcol_names_ent.get())-1
            except:
                messagebox.showerror("Error",'Entries must be integer')
                root.destroy()
            try:
                cnm_cs=int(sce.get())-1
            except:
                messagebox.showerror("Error",'Entries must be integer')
                root.destroy()
            try:
                cnm_ce=int(ece.get())
            except:
                messagebox.showerror("Error",'Entries must be integer')
                root.destroy()
            try:
                rnm_col=int(col_withrow_names_ent.get())-1
            except:
                messagebox.showerror("Error",'Entries must be integer')
                root.destroy()
            try:
                rnm_rs=int(sre.get())-1
            except:
                messagebox.showerror("Error",'Entries must be integer')
                root.destroy()
            try:
                rnm_re=int(ere.get())
            except:
                messagebox.showerror("Error",'Entries must be integer')
                root.destroy()
            col_index = [1, cnm_row, cnm_cs, cnm_ce]
            row_index = [1, rnm_col, rnm_rs, rnm_re]
            row_col_dat=[row_index,col_index]
            dfp = ret_num_df(df, row_col_dat)
            get_plot(dfp, var.get())

        txt = Label(root, text='                      ')
        txt.grid(row=8)
        btn=Button(root,text='submit',command=intrnl_submitcnc)
        btn.grid(row=9,column=1)
    # ------------------continuous rows and discrete column------------------------#
    elif variable.get() == OPTIONS[2]:
        txt = Label(root, text='                      ')
        txt.grid(row=0)
        col_withrow_names = Label(root, text='Column which contains row names')
        col_withrow_names.grid(row=1)
        col_withrow_names_ent = Entry(root)
        col_withrow_names_ent.grid(row=1, column=1)
        srl = Label(root, text='Row number where row names start')
        srl.grid(row=2)
        sre = Entry(root)
        sre.grid(row=2, column=1)
        erl = Label(root, text='Row number where row names end')
        erl.grid(row=3)
        ere = Entry(root)
        ere.grid(row=3, column=1)
        row_withcol_names = Label(root, text='Row which contains column names')
        row_withcol_names.grid(row=4)
        row_withcol_names_ent = Entry(root)
        row_withcol_names_ent.grid(row=4, column=1)
        cl = Label(root, text='Column Numbers')
        cl.grid(row=5)
        ce = Entry(root)
        ce.grid(row=5, column=1)
        plt_type_label = Label(root, text=' Plot Type')
        plt_type_label.grid(row=6)
        plot_type = OptionMenu(root, var, *plotoptions)
        plot_type.grid(row=6, column=1)

        def intrnl_submitcnd():
            try:
                cnm_row=int(row_withcol_names_ent.get())-1
            except:
                messagebox.showerror("Error",'Entries must be integer')
                root.destroy()
            try:
                rnm_col=int(col_withrow_names_ent.get())-1
            except:
                messagebox.showerror("Error",'Entries must be integer')
                root.destroy()
            try:
                rnm_rs=int(sre.get())-1
            except:
                messagebox.showerror("Error",'Entries must be integer')
                root.destroy()
            try:
                rnm_re=int(ere.get())
            except:
                messagebox.showerror("Error",'Entries must be integer')
                root.destroy()
            list_of_cols=ce.get().split(',')
            k=0
            while k<len(list_of_cols):
                try:
                    list_of_cols[k]=int(list_of_cols[k])-1
                except:
                    messagebox.showerror('error','Please enter Integral values of Columns')
                    root.destroy()
                k+=1
            col_index = [0, cnm_row, list_of_cols]
            row_index = [1, rnm_col, rnm_rs, rnm_re]
            row_col_dat=[row_index,col_index]
            dfp = ret_num_df(df, row_col_dat)
            get_plot(dfp,var.get())

        txt = Label(root, text='                      ')
        txt.grid(row=7)
        btn = Button(root, text='submit',command=intrnl_submitcnd)
        btn.grid(row=8, column=1)
    # ---------------discrete rows and continuous col------------------------#
    elif variable.get() == OPTIONS[1]:
        txt = Label(root, text='                      ')
        txt.grid(row=0)
        col_withrow_names = Label(root, text='Column which contains row names')
        col_withrow_names.grid(row=1)
        col_withrow_names_ent = Entry(root)
        col_withrow_names_ent.grid(row=1, column=1)
        rl = Label(root, text='Row Numbers')
        rl.grid(row=2)
        re = Entry(root)
        re.grid(row=2, column=1)
        row_withcol_names = Label(root, text='Row which contains column names')
        row_withcol_names.grid(row=3)
        row_withcol_names_ent = Entry(root)
        row_withcol_names_ent.grid(row=3, column=1)
        scl = Label(root, text='Column number where column names start')
        scl.grid(row=4)
        sce = Entry(root)
        sce.grid(row=4, column=1)
        ecl = Label(root, text='Column number where column names end')
        ecl.grid(row=5)
        ece = Entry(root)
        ece.grid(row=5, column=1)
        plt_type_label = Label(root, text=' Plot Type')
        plt_type_label.grid(row=6)
        plot_type = OptionMenu(root, var, *plotoptions)
        plot_type.grid(row=6, column=1)

        def intrnl_submitdnc():
            try:
                cnm_row=int(row_withcol_names_ent.get())-1
            except:
                messagebox.showerror("Error",'Entries must be integer')
                root.destroy()
            try:
                cnm_cs=int(sce.get())-1
            except:
                messagebox.showerror("Error",'Entries must be integer')
                root.destroy()
            try:
                cnm_ce=int(ece.get())
            except:
                messagebox.showerror("Error",'Entries must be integer')
                root.destroy()
            try:
                rnm_col=int(col_withrow_names_ent.get())-1
            except:
                messagebox.showerror("Error",'Entries must be integer')
                root.destroy()
            list_of_rows = re.get().split(',')
            k = 0
            while k < len(list_of_rows):
                try:
                    list_of_rows[k] = int(list_of_rows[k]) - 1
                except:
                    messagebox.showerror('error', 'Please enter Integral values of Columns')
                    root.destroy()
                k += 1
            col_index = [1, cnm_row, cnm_cs, cnm_ce]
            row_index = [0, rnm_col, list_of_rows]
            row_col_dat=[row_index,col_index]
            dfp = ret_num_df(df, row_col_dat)
            if var.get() == plotoptions[0]:
                dfp.plot.bar()
                root.destroy()
                plt.show()

            elif var.get() == plotoptions[1]:
                df.plot.pie(subplots=True, autopct='%.2f')
                root.destroy()
                plt.show()

        txt = Label(root, text='                      ')
        txt.grid(row=7)
        btn = Button(root, text='submit',command=intrnl_submitdnc)
        btn.grid(row=8, column=1)
    # ---------------discrete rows and discrete col------------------------#
    elif variable.get() == OPTIONS[0]:
        txt = Label(root, text='                      ')
        txt.grid(row=0)
        col_withrow_names = Label(root, text='Column which contains row names')
        col_withrow_names.grid(row=1)
        col_withrow_names_ent = Entry(root)
        col_withrow_names_ent.grid(row=1, column=1)
        rl = Label(root, text='Row Numbers')
        rl.grid(row=2)
        re = Entry(root)
        re.grid(row=2, column=1)
        row_withcol_names = Label(root, text='Row which contains column names')
        row_withcol_names.grid(row=3)
        row_withcol_names_ent = Entry(root)
        row_withcol_names_ent.grid(row=3, column=1)
        cl = Label(root, text='Column Numbers')
        cl.grid(row=4)
        ce = Entry(root)
        ce.grid(row=4, column=1)
        plt_type_label = Label(root, text=' Plot Type')
        plt_type_label.grid(row=5)
        plot_type = OptionMenu(root, var, *plotoptions)
        plot_type.grid(row=5, column=1)

        def intrnl_submitdnd():
            try:
                cnm_row=int(row_withcol_names_ent.get())-1
            except:
                messagebox.showerror("Error",'Entries must be integer')
                root.destroy()
            try:
                rnm_col=int(col_withrow_names_ent.get())-1
            except:
                messagebox.showerror("Error",'Entries must be integer')
                root.destroy()
            list_of_rows = re.get().split(',')
            k = 0
            while k < len(list_of_rows):
                try:
                    list_of_rows[k] = int(list_of_rows[k]) - 1
                except:
                    messagebox.showerror('error', 'Please enter Integral values of Columns')
                    root.destroy()
                k += 1
            list_of_cols = ce.get().split(',')
            k = 0
            while k < len(list_of_cols):
                try:
                    list_of_cols[k] = int(list_of_cols[k]) - 1
                except:
                    messagebox.showerror('error', 'Please enter Integral values of Columns')
                    root.destroy()
                k += 1
            col_index = [0, cnm_row, list_of_cols]
            row_index = [0, rnm_col, list_of_rows]
            row_col_dat=[row_index,col_index]
            dfp = ret_num_df(df, row_col_dat)
            get_plot(dfp, var.get())

        txt = Label(root, text='                      ')
        txt.grid(row=6)
        btn = Button(root, text='submit',command=intrnl_submitdnd)
        btn.grid(row=7, column=1)


text0 = Label(root, text="                                  ")
text0.grid(row=4)
b = Button(root, text="submit", command=submit)
b.grid(row=5, column=1)
root.mainloop()