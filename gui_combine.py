from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd


# ---------- returning sliced dataframes for 1st sheet data-----------#
def make_firstdataframe(sheet_dat, rows, cols):
    df = pd.read_excel(sheet_dat[0][0], sheet_dat[0][1],header=None)
    if rows[0]=='c' and cols[0]=='c':
        startrow =int(rows[1])
        endrow = int(rows[2])+1
        startcol = int(cols[1])
        endcol= int(cols[2])+1
        df =df.iloc[startrow:endrow,startcol:endcol]
    elif rows[0]=='c' and cols[0]=='d':
        startrow = int(rows[1])
        endrow = int(rows[2]) + 1
        col=[]
        for value in cols[1]:
            col.append(int(value))
        df = df.iloc[startrow:endrow,col]
    elif rows[0] == 'd' and cols[0] == 'c':
        startcol = int(cols[1])
        endcol = int(cols[2]) + 1
        row = []
        for value in rows[1]:
            row.append(int(value))
        df = df.iloc[row, startcol:endcol]
    elif rows[0] == 'd' and cols[0] == 'd':
        row = []
        for value in rows[1]:
            row.append(int(value))
        col = []
        for value in cols[1]:
            col.append(int(value))
        df = df.iloc[row,col]
    return df


# ----------- data frames for next sheets -----------------------#
def make_othrdataframes(sheet_dat, rows, cols,num):
    df = pd.read_excel(sheet_dat[num][0], sheet_dat[num][1],header=None)
    print(df)
    if rows[0]=='c' and cols[0]=='c':
        startrow =int(rows[1])
        endrow = int(rows[2])+1
        startcol = int(cols[1])
        endcol= int(cols[2])+1
        df =df.iloc[startrow:endrow,startcol:endcol]
    elif rows[0]=='c' and cols[0]=='d':
        startrow = int(rows[1])
        endrow = int(rows[2]) + 1
        col=[]
        for value in cols[1]:
            col.append(int(value))
            print(value)
        df = df.iloc[startrow:endrow,col]
    elif rows[0] == 'd' and cols[0] == 'c':
        startcol = int(cols[1])
        endcol = int(cols[2]) + 1
        row = []
        for value in rows[1]:
            row.append(int(value))
        df = df.iloc[row, startcol:endcol]
    elif rows[0] == 'd' and cols[0] == 'd':
        row = []
        for value in rows[1]:
            row.append(int(value))
        col = []
        for value in cols[1]:
            col.append(int(value))
        df = df.iloc[row,col]

    return df


OPTIONS=['Discrete Rows and Discrete Columns','Discrete Rows and Continuous Columns','Continuous Rows and Discrete Columns','Continous Rows and Continuous Columns']
root = Tk()

variable=StringVar(root)
variable.set(OPTIONS[3])
root.geometry('450x150')
textn = Label(root, text="                              ")
textn.grid(row=0)
num_sheetsL = Label(root, text=' number of sheets      ')
num_sheetsL.grid(row=1)
num_sheetsE= Entry(root)
num_sheetsE.grid(row=1,column=1)
type_label= Label(root, text=' Row and Column Selection type')
type_label.grid(row=2)
type_entry=OptionMenu(root,variable,*OPTIONS)
type_entry.grid(row=2,column=1)


def submit():

    try:
        num_of_sheets= int(num_sheetsE.get())
    except:
        messagebox.showerror("error","the number of sheets must be an integer")
    print(num_of_sheets)
    num_sheetsE.destroy()
    text0.destroy()
    num_sheetsL.destroy()
    b.destroy()
    type_label.destroy()
    type_entry.destroy()

    sheet_locs=[]
    row_dat=[]
    col_dat=[]
    sheet_dat=[]
    root.geometry('300x200')

    def browse_btn_excel():
        try:
            xlfilename = filedialog.askopenfilename(initialdir="/", title="Select file",
                                                    filetypes=(("excel files", "*.xlsx"), ("all files", "*.*")))
            check = "xlsx" in xlfilename
            if not check:
                messagebox.showerror("error", "please select an excel '.xlsx' file")
            sheet_locs.append(xlfilename)
        except:
            messagebox.showerror("error", "please select a file")

# ------------------continuous rows and continuous column------------------------#
    if variable.get()==OPTIONS[3]:
        i = 0
        label_names = []
        btn_names = []
        srl_names = []
        sre_names = []
        erl_names = []
        ere_names = []
        scl_names = []
        sce_names = []
        ecl_names = []
        ece_names = []
        sheet_nameL = []
        sheet_nameE = []
        while i < num_of_sheets:
            label_names.append('l' + str(i + 1))
            btn_names.append('b' + str(i + 1))
            sheet_nameL.append('snl' + str(i + 1))
            sheet_nameE.append('sne' + str(i + 1))
            srl_names.append('srl' + str(i + 1))
            sre_names.append('sre' + str(i + 1))
            erl_names.append('erl' + str(i + 1))
            ere_names.append('ere' + str(i + 1))
            scl_names.append('scl' + str(i + 1))
            sce_names.append('sce' + str(i + 1))
            ecl_names.append('ecl' + str(i + 1))
            ece_names.append('ece' + str(i + 1))
            i += 1

        j = 0
        while j < num_of_sheets:
            label_names[j] = Label(root,text='Select Sheet '+str(j+1))
            label_names[j].grid(row=1+7*j)
            btn_names[j] = Button(root, text='browse', command=browse_btn_excel)
            btn_names[j].grid(row=1+7*j, column=1)
            sheet_nameL[j]=Label(root,text='Sheet Name '+str(j+1))
            sheet_nameL[j].grid(row=2+7*j)
            sheet_nameE[j]=Entry(root)
            sheet_nameE[j].grid(row=2+7*j,column=1)
            srl_names[j]=Label(root,text='Start Row')
            srl_names[j].grid(row=3+7*j)
            sre_names[j]=Entry(root)
            sre_names[j].grid(row=3+7*j, column=1)
            erl_names[j]=Label(root,text='End Row')
            erl_names[j].grid(row=4+7*j)
            ere_names[j]=Entry(root)
            ere_names[j].grid(row=4+7*j,column=1)
            scl_names[j]=Label(root,text='Start Column')
            scl_names[j].grid(row=5+7*j)
            sce_names[j]=Entry(root)
            sce_names[j].grid(row=5+7*j,column=1)
            ecl_names[j]=Label(root,text='End Column')
            ecl_names[j].grid(row=6+7*j)
            ece_names[j]=Entry(root)
            ece_names[j].grid(row=6+7*j,column=1)
            text = Label(root, text="                              ")
            text.grid(row=7+7*j)
            j += 1
        out_lab = Label(root, text='output file name')
        out_lab.grid(row=7 * j + 1)
        out_ent = Entry(root)
        out_ent.grid(row=7 * j + 1, column=1)

        def final_subcnc():
            i=0
            while i<num_of_sheets:
                chk2 = out_ent.get() == ''
                if chk2:
                    messagebox.showerror('error', 'output sheet name is blank')
                    root.destroy()
                else:
                    output = out_ent.get()
                chk=sheet_nameE[i].get().strip()=='' or sheet_locs[i]==''
                if chk:
                    messagebox.showerror('error', 'sheet name or sheet location ' + str(i+1) + ' is blank')
                    root.destroy()
                else:
                    sheet_dat.append([sheet_locs[i],sheet_nameE[i].get().strip()])
                try:
                    sr=int(sre_names[i].get())-1
                except:
                    messagebox.showerror("error", "Start row must be an integer")
                    root.destroy()
                try:
                    er=int(ere_names[i].get())-1
                except:
                    messagebox.showerror("error", "end row must be an integer")
                    root.destroy()
                try:
                    sc=int(sce_names[i].get())-1
                except:
                    messagebox.showerror("error", "Start column must be an integer")
                    root.destroy()
                try:
                    ec=int(ece_names[i].get())-1
                except:
                    messagebox.showerror("error", "End column must be an integer")
                    root.destroy()
                row_dat.append(['c',str(sr),str(er),str(er-sr+1)])
                col_dat.append(['c',str(sc),str(ec),str(ec-sc+1)])
                i+=1
            print(sheet_dat)
            print(row_dat)
            print(col_dat)
            if num_of_sheets > 1:
                rows = row_dat[0]
                nr = int(rows[-1])
                columns = col_dat[0]
                nc = int(columns[-1])
                df = make_firstdataframe(sheet_dat, rows, columns)
                row_indices = []
                if nr != 0:
                    n = 0
                    while n < nr:
                        row_indices.append(n)
                        n += 1
                    df.index = row_indices
                i = 1
                while i < num_of_sheets:
                    rows = row_dat[i]
                    chkr = int(rows[-1]) != nr
                    if chkr:
                        messagebox.showerror('error','the number of rows should be same')
                    columns = col_dat[i]
                    df1 = make_othrdataframes(sheet_dat, rows, columns, i)
                    if nr != 0:
                        df1.index = row_indices
                    print(df1)
                    df = pd.concat([df, df1], axis=1)
                    i += 1

            else:
                rows = row_dat[0]
                columns = col_dat[0]
                df = make_firstdataframe(sheet_dat, rows, columns)

            dfw = df.copy()
            print(dfw)
            with pd.ExcelWriter(output+'.xlsx') as writer:  # doctest: +SKIP
                dfw.to_excel(writer, sheet_name='Sheet1', header=False, index=False)
            root.destroy()
        text=Label(root,text='                                            ')
        text.grid(row=7*j+2)
        btn_final=Button(root,text='submit', command=final_subcnc)
        btn_final.grid(row=7*j+3,column=1)
# ------------------continuous rows and discrete column------------------------#
    elif variable.get()==OPTIONS[2]:
        i = 0
        label_names = []
        btn_names = []
        sheet_nameL = []
        sheet_nameE = []
        srl_names = []
        sre_names = []
        erl_names = []
        ere_names = []
        cl_names=[]
        ce_names=[]
        while i < num_of_sheets:
            label_names.append('l' + str(i + 1))
            btn_names.append('b' + str(i + 1))
            sheet_nameL.append('snl' + str(i + 1))
            sheet_nameE.append('sne' + str(i + 1))
            srl_names.append('srl' + str(i + 1))
            sre_names.append('sre' + str(i + 1))
            erl_names.append('erl' + str(i + 1))
            ere_names.append('ere' + str(i + 1))
            cl_names.append('cl' + str(i + 1))
            ce_names.append('ce' + str(i + 1))
            i += 1
        j = 0
        while j < num_of_sheets:
            label_names[j] = Label(root, text='Select Sheet ' + str(j + 1))
            label_names[j].grid(row=1 + 6 * j)
            btn_names[j] = Button(root, text='browse', command=browse_btn_excel)
            btn_names[j].grid(row=1 + 6 * j, column=1)
            sheet_nameL[j] = Label(root, text='Sheet Name ' + str(j + 1))
            sheet_nameL[j].grid(row=2 + 6 * j)
            sheet_nameE[j] = Entry(root)
            sheet_nameE[j].grid(row=2 + 6 * j, column=1)
            srl_names[j] = Label(root, text='Start Row')
            srl_names[j].grid(row=3 + 6 * j)
            sre_names[j] = Entry(root)
            sre_names[j].grid(row=3 + 6 * j, column=1)
            erl_names[j] = Label(root, text='End Row')
            erl_names[j].grid(row=4 + 6 * j)
            ere_names[j] = Entry(root)
            ere_names[j].grid(row=4 + 6 * j, column=1)
            cl_names[j] = Label(root, text='Column Numbers')
            cl_names[j].grid(row=5 + 6 * j)
            ce_names[j] = Entry(root)
            ce_names[j].grid(row=5 + 6 * j, column=1)
            text = Label(root, text="                              ")
            text.grid(row=6 + 6 * j)
            j += 1
        out_lab = Label(root, text='output file name')
        out_lab.grid(row=6 * j + 1)
        out_ent = Entry(root)
        out_ent.grid(row=6 * j + 1, column=1)

        def final_subcnd():
            i=0
            while i<num_of_sheets:
                chk2 = out_ent.get() == ''
                if chk2:
                    messagebox.showerror('error', 'output sheet name is blank')
                    root.destroy()
                else:
                    output = out_ent.get()
                chk=sheet_nameE[i].get().strip()=='' or sheet_locs[i]==''
                if chk:
                    messagebox.showerror('error', 'sheet name or sheet location ' + str(i+1) + ' is blank')
                    root.destroy()
                else:
                    sheet_dat.append([sheet_locs[i],sheet_nameE[i].get().strip()])
                try:
                    sr=int(sre_names[i].get())-1
                except:
                    messagebox.showerror("error", "Start row must be an integer")
                    root.destroy()
                try:
                    er=int(ere_names[i].get())-1
                except:
                    messagebox.showerror("error", "End row must be an integer")
                    root.destroy()
                col=ce_names[i].get().split(',')
                k=0
                while k<len(col):
                   try:
                       c=int(col[k])-1
                       col[k]=str(c)
                   except:
                       messagebox.showerror('error',str(k+1)+' column entry of sheet '+str(i+1)+' is not integer')
                       root.destroy()
                   k+=1
                row_dat.append(['c',str(sr),str(er),str(er-sr+1)])
                col_dat.append(['d',col,str(len(col))])
                i+=1
            print(sheet_dat)
            print(row_dat)
            print(col_dat)
            if num_of_sheets > 1:
                rows = row_dat[0]
                nr = int(rows[-1])
                columns = col_dat[0]
                nc = int(columns[-1])
                df = make_firstdataframe(sheet_dat, rows, columns)
                row_indices = []
                if nr != 0:
                    n = 0
                    while n < nr:
                        row_indices.append(n)
                        n += 1
                    df.index = row_indices
                i = 1
                while i < num_of_sheets:
                    rows = row_dat[i]
                    chkr = int(rows[-1]) != nr
                    if chkr:
                        messagebox.showerror('error','the number of rows should be same')
                    columns = col_dat[i]
                    df1 = make_othrdataframes(sheet_dat, rows, columns, i)
                    if nr != 0:
                        df1.index = row_indices
                    print(df1)
                    df = pd.concat([df, df1], axis=1)
                    i += 1

            else:
                rows = row_dat[0]
                columns = col_dat[0]
                df = make_firstdataframe(sheet_dat, rows, columns)

            dfw = df.copy()
            print(dfw)
            with pd.ExcelWriter(output+'.xlsx') as writer:  # doctest: +SKIP
                dfw.to_excel(writer, sheet_name='Sheet1', header=False, index=False)
            root.destroy()
        text=Label(root,text='                               ')
        text.grid(row=6*j+2)
        btn_final=Button(root,text='submit', command=final_subcnd)
        btn_final.grid(row=6*j+3,column=1)
# ---------------discrete rows and continuous col------------------------#
    elif variable.get()==OPTIONS[1]:
        i = 0
        label_names = []
        btn_names = []
        sheet_nameL = []
        sheet_nameE = []
        scl_names = []
        sce_names = []
        ecl_names = []
        ece_names = []
        rl_names=[]
        re_names=[]
        while i < num_of_sheets:
            label_names.append('l' + str(i + 1))
            btn_names.append('b' + str(i + 1))
            sheet_nameL.append('snl' + str(i + 1))
            sheet_nameE.append('sne' + str(i + 1))
            rl_names.append('rl' + str(i + 1))
            re_names.append('re' + str(i + 1))
            scl_names.append('scl' + str(i + 1))
            sce_names.append('sce' + str(i + 1))
            ecl_names.append('ecl' + str(i + 1))
            ece_names.append('ece' + str(i + 1))
            i += 1
        j = 0
        while j < num_of_sheets:
            label_names[j] = Label(root, text='Select Sheet ' + str(j + 1))
            label_names[j].grid(row=1 + 6 * j)
            btn_names[j] = Button(root, text='browse', command=browse_btn_excel)
            btn_names[j].grid(row=1 + 6 * j, column=1)
            sheet_nameL[j] = Label(root, text='Sheet Name ' + str(j + 1))
            sheet_nameL[j].grid(row=2 + 6 * j)
            sheet_nameE[j] = Entry(root)
            sheet_nameE[j].grid(row=2 + 6 * j, column=1)
            rl_names[j] = Label(root, text='Row Numbers')
            rl_names[j].grid(row=3 + 6 * j)
            re_names[j] = Entry(root)
            re_names[j].grid(row=3 + 6 * j, column=1)
            scl_names[j] = Label(root, text='Start Column')
            scl_names[j].grid(row=4 + 6 * j)
            sce_names[j] = Entry(root)
            sce_names[j].grid(row=4 + 6 * j, column=1)
            ecl_names[j] = Label(root, text='End Column')
            ecl_names[j].grid(row=5 + 6 * j)
            ece_names[j] = Entry(root)
            ece_names[j].grid(row=5 + 6 * j, column=1)
            text = Label(root, text="                              ")
            text.grid(row=6 + 6 * j)
            j += 1
        out_lab = Label(root, text='output file name')
        out_lab.grid(row=6 * j + 1)
        out_ent = Entry(root)
        out_ent.grid(row=6 * j + 1, column=1)

        def final_subdnc():
            i = 0
            while i < num_of_sheets:
                chk2 = out_ent.get() == ''
                if chk2:
                    messagebox.showerror('error', 'output sheet name is blank')
                    root.destroy()
                else:
                    output = out_ent.get()
                chk = sheet_nameE[i].get().strip() == '' or sheet_locs[i] == ''
                if chk:
                    messagebox.showerror('error', 'sheet name or sheet location ' + str(i + 1) + ' is blank')
                    root.destroy()
                else:
                    sheet_dat.append([sheet_locs[i], sheet_nameE[i].get().strip()])
                row = re_names[i].get().split(',')
                k = 0
                while k < len(row):
                    try:
                        r = int(row[k])-1
                        row[k]=str(r)
                    except:
                        messagebox.showerror('error', str(k+1)+' row entry of sheet '+str(i+1)+' is not integer')
                        root.destroy()
                    k += 1
                try:
                    sc = int(sce_names[i].get())-1
                except:
                    messagebox.showerror("error", "Start column must be an integer")
                    root.destroy()
                try:
                    ec = int(ece_names[i].get())-1
                except:
                    messagebox.showerror("error", "End column must be an integer")
                    root.destroy()
                row_dat.append(['d', row, str(len(row))])
                col_dat.append(['c', str(sc), str(ec), str(ec - sc + 1)])
                i += 1
            print(sheet_dat)
            print(row_dat)
            print(col_dat)
            if num_of_sheets > 1:
                rows = row_dat[0]
                nr = int(rows[-1])
                columns = col_dat[0]
                nc = int(columns[-1])
                df = make_firstdataframe(sheet_dat, rows, columns)
                row_indices = []
                if nr != 0:
                    n = 0
                    while n < nr:
                        row_indices.append(n)
                        n += 1
                    df.index = row_indices
                i = 1
                while i < num_of_sheets:
                    rows = row_dat[i]
                    chkr = int(rows[-1]) != nr
                    if chkr:
                        messagebox.showerror('error','the number of rows should be same')
                    columns = col_dat[i]
                    df1 = make_othrdataframes(sheet_dat, rows, columns, i)
                    if nr != 0:
                        df1.index = row_indices
                    print(df1)
                    df = pd.concat([df, df1], axis=1)
                    i += 1

            else:
                rows = row_dat[0]
                columns = col_dat[0]
                df = make_firstdataframe(sheet_dat, rows, columns)

            dfw = df.copy()
            print(dfw)
            with pd.ExcelWriter(output+'.xlsx') as writer:  # doctest: +SKIP
                dfw.to_excel(writer, sheet_name='Sheet1', header=False, index=False)
            root.destroy()

        text = Label(root, text='                                    ')
        text.grid(row=6 * j + 2)
        btn_final = Button(root, text='submit', command=final_subdnc)
        btn_final.grid(row=6 * j + 3, column=1)

    elif variable.get() == OPTIONS[0]:  # discrete rows and discrete col
        i = 0
        label_names = []
        btn_names = []
        sheet_nameL = []
        sheet_nameE = []
        rl_names = []
        re_names = []
        cl_names = []
        ce_names = []
        while i < num_of_sheets:
            label_names.append('l' + str(i + 1))
            btn_names.append('b' + str(i + 1))
            sheet_nameL.append('snl' + str(i + 1))
            sheet_nameE.append('sne' + str(i + 1))
            rl_names.append('rl' + str(i + 1))
            re_names.append('re' + str(i + 1))
            cl_names.append('cl' + str(i + 1))
            ce_names.append('ce' + str(i + 1))
            i += 1
        j=0
        while j < num_of_sheets:
            label_names[j] = Label(root, text='Select Sheet ' + str(j + 1))
            label_names[j].grid(row=1 + 5 * j)
            btn_names[j] = Button(root, text='browse', command=browse_btn_excel)
            btn_names[j].grid(row=1 + 5 * j, column=1)
            sheet_nameL[j] = Label(root, text='Sheet Name ' + str(j + 1))
            sheet_nameL[j].grid(row=2 + 5 * j)
            sheet_nameE[j] = Entry(root)
            sheet_nameE[j].grid(row=2 + 5 * j, column=1)
            rl_names[j] = Label(root, text='Row numbers')
            rl_names[j].grid(row=3 + 5 * j)
            re_names[j] = Entry(root)
            re_names[j].grid(row=3 + 5 * j, column=1)
            cl_names[j] = Label(root, text='Column numbers')
            cl_names[j].grid(row=4 + 5 * j)
            ce_names[j] = Entry(root)
            ce_names[j].grid(row=4 + 5 * j, column=1)
            text = Label(root, text="                              ")
            text.grid(row=5 + 5 * j)
            j += 1
        out_lab=Label(root, text='output file name')
        out_lab.grid(row=5*j+1)
        out_ent=Entry(root)
        out_ent.grid(row=5*j+1,column=1)

        def final_subdnd():
            i = 0
            while i < num_of_sheets:

                chk = sheet_nameE[i].get().strip() == '' or sheet_locs[i] == ''
                if chk:
                    messagebox.showerror('error', 'sheet name or sheet location ' + str(i + 1) + ' is blank')
                    root.destroy()
                else:
                    sheet_dat.append([sheet_locs[i], sheet_nameE[i].get().strip()])

                chk2=out_ent.get()==''
                if chk2:
                    messagebox.showerror('error', 'output sheet name is blank')
                    root.destroy()
                else:
                    output=out_ent.get()
                row = re_names[i].get().split(',')
                k = 0
                while k < len(row):
                    try:
                        r = int(row[k])-1
                        row[k]=str(r)
                    except:
                        messagebox.showerror('error', str(k+1)+' row entry of sheet '+str(i+1)+' is not integer')
                        root.destroy()
                    k += 1
                col = ce_names[i].get().split(',')
                k = 0
                while k < len(col):
                    try:
                        c = int(col[k])-1
                        col[k]=str(c)
                    except:
                        messagebox.showerror('error', str(k+1)+' column entry of sheet '+str(i+1)+' is not integer')
                        root.destroy()
                    k += 1
                col_dat.append(['d', col, str(len(col))])
                row_dat.append(['d', row, str(len(row))])
                i += 1
            print(sheet_dat)
            print(row_dat)
            print(col_dat)
            if num_of_sheets > 1:
                rows = row_dat[0]
                nr = int(rows[-1])
                columns = col_dat[0]
                nc = int(columns[-1])
                df = make_firstdataframe(sheet_dat, rows, columns)
                row_indices = []
                if nr != 0:
                    n = 0
                    while n < nr:
                        row_indices.append(n)
                        n += 1
                    df.index = row_indices
                i = 1
                while i < num_of_sheets:
                    rows = row_dat[i]
                    chkr = int(rows[-1]) != nr
                    if chkr:
                        messagebox.showerror('error','the number of rows should be same')
                    columns = col_dat[i]
                    df1 = make_othrdataframes(sheet_dat, rows, columns, i)
                    if nr != 0:
                        df1.index = row_indices
                    print(df1)
                    df = pd.concat([df, df1], axis=1)
                    i += 1

            else:
                rows = row_dat[0]
                columns = col_dat[0]
                df = make_firstdataframe(sheet_dat, rows, columns)

            dfw = df.copy()
            print(dfw)
            with pd.ExcelWriter(output+'.xlsx') as writer:  # doctest: +SKIP
                dfw.to_excel(writer, sheet_name='Sheet1', header=False, index=False)
            root.destroy()
        text=Label(root,text='                             ')
        text.grid(row=5*j+2)
        btn_final = Button(root, text='submit', command=final_subdnd)
        btn_final.grid(row=5 * j + 3, column=1)


text0 = Label(root, text="                                  ")
text0.grid(row=3)
b = Button(root, text="submit", command=submit)
b.grid(row=4, column=1)

root.mainloop()