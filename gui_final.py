from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd
import matplotlib.pyplot as plt
import xlsxwriter
import setuptools._vendor
import numpy as np

def gui_getxl():
    def calc_col_num(start_pos, max_num_of_col, starts_max_col):
        z = 0
        start_pos = int(start_pos)
        diff_arr = []
        while z < max_num_of_col:
            diff_arr.append(abs(start_pos - starts_max_col[z]))
            z += 1

        return diff_arr.index(min(diff_arr))

    def checkpt(matrix_data, curr_row, curr_col):
        if curr_row > 0:
            chk2 = False
            i = 0
            while i < len(matrix_data[curr_row - 1]):
                if curr_col <= matrix_data[curr_row - 1][i][1]:
                    chk2 = True
                i += 1
            return chk2
        else:
            return False

    def recheck_col(my_lst, matrix, cr, cc):
        diff_arr = []
        i = 0
        while i < len(matrix[cr - 1]):
            diff_arr.append(abs(int(my_lst[cr - 1][i][1]) - int(my_lst[cr][cc][1])))
            i += 1
        j = diff_arr.index(min(diff_arr))
        return matrix[cr - 1][j][1]

    def browse_btn_notepad():
        global npfilename
        try:
            npfilename = filedialog.askopenfilename(initialdir="/", title="Select file",
                                                    filetypes=(("text files", "*.txt"), ("all files", "*.*")))
        except:
            messagebox.showerror("error", "please select a file")
        # print(npfilename)

    root = Tk()
    root.title('Notepad To Excel')
    root.geometry('250x150')
    text0 = Label(root, text="                                  ")
    text0.grid(row=0)
    # for the icon use: root.iconbitmap(r'location')
    text2 = Label(root, text="Select notepad file")
    text2.grid(row=1)
    btn2 = Button(root, text='Browse', command=browse_btn_notepad)
    btn2.grid(row=1, column=1)

    outputlabel = Label(root, text="Output file Name")
    outputlabel.grid(row=2)
    outputentry = Entry(root)
    outputentry.grid(row=2, column=1)

    def close_window():
        global outputname
        outputname = outputentry.get()
        flopen = open(npfilename, 'r+')
        lines = flopen.readlines()
        ptr = 0
        flopen.seek(0)
        for line in lines:
            for char in line:
                if char is '\n':
                    flopen.write(line[:len(line) - 1] + '     \n')

        # ---------- figure out the words and their positions---------#
        # --- make the lines to be taken as dynamic to obtain specific tables------#
        # ---- give instruction to align the data file before hand for proper alignment ----#
        oldstring = ""
        flopen.seek(0)
        count = 0
        line_no = 0
        len_calc = 0
        row = 0
        col_nums = []
        elem = 0
        col_count = 0
        dict2 = []
        i = 0
        for line in lines:
            count = 0
            line_no += 1
            len_calc = 0
            col_count = 0
            oldstring = ''
            for word in line:
                if word is not " ":
                    count = 0
                    oldstring += word
                else:
                    count += 1
                    if count < 2:
                        oldstring += word
                    else:
                        if len(oldstring.strip()) != 0:
                            if '------' not in oldstring:
                                startpos = len_calc - count - len(oldstring.strip()) + 1
                                if startpos < 0:
                                    startpos = 0
                                lst2 = [oldstring.strip(), str(startpos), str(len(oldstring.strip()) + startpos)]
                                # print(lst2)
                                dict2.append(lst2)
                                col_count += 1
                        oldstring = ''
                len_calc += 1

            if col_count != 0:
                col_nums.append(col_count)

        # print(col_nums)

        lst = []
        w = 0
        k = 0
        while k < len(col_nums):
            lst.append([])
            n = 0
            while n < col_nums[k]:
                lst[k].append(dict2[w])
                w += 1
                n += 1
            k += 1

        max_num_of_col = max(col_nums)
        lin_max_col = col_nums.index(max(col_nums))

        # --------------------------column calculations--------------------#
        starts_max_col = []
        w = 0
        while w < max_num_of_col:
            starts_max_col.append(int(lst[lin_max_col][w][1]))
            w += 1

        row = 0
        row_col_dat = []
        while row < len(col_nums):
            row_col_dat.append([])
            if col_nums[row] < max_num_of_col:
                n = 0
                while n < col_nums[row]:
                    row_col_dat[row].append([row, calc_col_num(lst[row][n][1], max_num_of_col, starts_max_col)])
                    if col_nums[row] >= 2 and n > 0:
                        if row_col_dat[row][n - 1][1] == row_col_dat[row][n][1]:
                            row_col_dat[row][n][1] = row_col_dat[row][n][1] + 1
                            if checkpt(row_col_dat, row, n):
                                row_col_dat[row][n][1] = recheck_col(lst, row_col_dat, row, n)
                    n += 1
            else:
                m = 0
                while m < max_num_of_col:
                    row_col_dat[row].append([row, m])
                    m += 1
            row += 1

        # -------- writing data to excel------------#
        workbook = xlsxwriter.Workbook(outputname + '.xlsx')
        worksheet = workbook.add_worksheet()
        row = 0

        while row < len(col_nums):
            col = 0
            while col < col_nums[row]:
                worksheet.write(row_col_dat[row][col][0], row_col_dat[row][col][1], lst[row][col][0])
                col += 1
            row += 1
        workbook.close()
        flopen.close()
        root.destroy()

    text0 = Label(root, text="                                  ")
    text0.grid(row=11)
    submit = Button(root, text="Submit", command=close_window)
    submit.grid(row=12, column=1)
    root.mainloop()


def gui_combine():
    # ---------- returning sliced dataframes for 1st sheet data-----------#
    def make_firstdataframe(sheet_dat, rows, cols):
        df = pd.read_excel(sheet_dat[0][0], sheet_dat[0][1], header=None)
        if rows[0] == 'c' and cols[0] == 'c':
            startrow = int(rows[1])
            endrow = int(rows[2]) + 1
            startcol = int(cols[1])
            endcol = int(cols[2]) + 1
            df = df.iloc[startrow:endrow, startcol:endcol]
        elif rows[0] == 'c' and cols[0] == 'd':
            startrow = int(rows[1])
            endrow = int(rows[2]) + 1
            col = []
            for value in cols[1]:
                col.append(int(value))
            df = df.iloc[startrow:endrow, col]
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
            df = df.iloc[row, col]
        return df

    # ----------- data frames for next sheets -----------------------#
    def make_othrdataframes(sheet_dat, rows, cols, num):
        df = pd.read_excel(sheet_dat[num][0], sheet_dat[num][1], header=None)
        # print(df)
        if rows[0] == 'c' and cols[0] == 'c':
            startrow = int(rows[1])
            endrow = int(rows[2]) + 1
            startcol = int(cols[1])
            endcol = int(cols[2]) + 1
            df = df.iloc[startrow:endrow, startcol:endcol]
        elif rows[0] == 'c' and cols[0] == 'd':
            startrow = int(rows[1])
            endrow = int(rows[2]) + 1
            col = []
            for value in cols[1]:
                col.append(int(value))
                # print(value)
            df = df.iloc[startrow:endrow, col]
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
            df = df.iloc[row, col]

        return df

    OPTIONS = ['Discrete Rows and Discrete Columns', 'Discrete Rows and Continuous Columns',
               'Continuous Rows and Discrete Columns', 'Continous Rows and Continuous Columns']
    root = Tk()
    root.title('Combine Sheets')
    variable = StringVar(root)
    variable.set(OPTIONS[3])
    root.geometry('450x150')
    textn = Label(root, text="                              ")
    textn.grid(row=0)
    num_sheetsL = Label(root, text=' number of sheets      ')
    num_sheetsL.grid(row=1)
    num_sheetsE = Entry(root)
    num_sheetsE.grid(row=1, column=1)
    type_label = Label(root, text=' Row and Column Selection type')
    type_label.grid(row=2)
    type_entry = OptionMenu(root, variable, *OPTIONS)
    type_entry.grid(row=2, column=1)

    def submit():

        try:
            num_of_sheets = int(num_sheetsE.get())
        except:
            messagebox.showerror("error", "the number of sheets must be an integer")
        # print(num_of_sheets)
        num_sheetsE.destroy()
        text0.destroy()
        num_sheetsL.destroy()
        b.destroy()
        type_label.destroy()
        type_entry.destroy()

        sheet_locs = []
        row_dat = []
        col_dat = []
        sheet_dat = []
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
        if variable.get() == OPTIONS[3]:
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
                label_names[j] = Label(root, text='Select Sheet ' + str(j + 1))
                label_names[j].grid(row=1 + 7 * j)
                btn_names[j] = Button(root, text='browse', command=browse_btn_excel)
                btn_names[j].grid(row=1 + 7 * j, column=1)
                sheet_nameL[j] = Label(root, text='Sheet Name ' + str(j + 1))
                sheet_nameL[j].grid(row=2 + 7 * j)
                sheet_nameE[j] = Entry(root)
                sheet_nameE[j].grid(row=2 + 7 * j, column=1)
                srl_names[j] = Label(root, text='Start Row')
                srl_names[j].grid(row=3 + 7 * j)
                sre_names[j] = Entry(root)
                sre_names[j].grid(row=3 + 7 * j, column=1)
                erl_names[j] = Label(root, text='End Row')
                erl_names[j].grid(row=4 + 7 * j)
                ere_names[j] = Entry(root)
                ere_names[j].grid(row=4 + 7 * j, column=1)
                scl_names[j] = Label(root, text='Start Column')
                scl_names[j].grid(row=5 + 7 * j)
                sce_names[j] = Entry(root)
                sce_names[j].grid(row=5 + 7 * j, column=1)
                ecl_names[j] = Label(root, text='End Column')
                ecl_names[j].grid(row=6 + 7 * j)
                ece_names[j] = Entry(root)
                ece_names[j].grid(row=6 + 7 * j, column=1)
                text = Label(root, text="                              ")
                text.grid(row=7 + 7 * j)
                j += 1
            out_lab = Label(root, text='output file name')
            out_lab.grid(row=7 * j + 1)
            out_ent = Entry(root)
            out_ent.grid(row=7 * j + 1, column=1)

            def final_subcnc():
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
                    try:
                        sr = int(sre_names[i].get()) - 1
                    except:
                        messagebox.showerror("error", "Start row must be an integer")
                        root.destroy()
                    try:
                        er = int(ere_names[i].get()) - 1
                    except:
                        messagebox.showerror("error", "end row must be an integer")
                        root.destroy()
                    try:
                        sc = int(sce_names[i].get()) - 1
                    except:
                        messagebox.showerror("error", "Start column must be an integer")
                        root.destroy()
                    try:
                        ec = int(ece_names[i].get()) - 1
                    except:
                        messagebox.showerror("error", "End column must be an integer")
                        root.destroy()
                    row_dat.append(['c', str(sr), str(er), str(er - sr + 1)])
                    col_dat.append(['c', str(sc), str(ec), str(ec - sc + 1)])
                    i += 1
                # print(sheet_dat)
                # print(row_dat)
                # print(col_dat)
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
                            messagebox.showerror('error', 'the number of rows should be same')
                        columns = col_dat[i]
                        df1 = make_othrdataframes(sheet_dat, rows, columns, i)
                        if nr != 0:
                            df1.index = row_indices
                        # print(df1)
                        df = pd.concat([df, df1], axis=1)
                        i += 1

                else:
                    rows = row_dat[0]
                    columns = col_dat[0]
                    df = make_firstdataframe(sheet_dat, rows, columns)

                dfw = df.copy()
                # print(dfw)
                with pd.ExcelWriter(output + '.xlsx') as writer:  # doctest: +SKIP
                    dfw.to_excel(writer, sheet_name='Sheet1', header=False, index=False)
                root.destroy()

            text = Label(root, text='                                            ')
            text.grid(row=7 * j + 2)
            btn_final = Button(root, text='submit', command=final_subcnc)
            btn_final.grid(row=7 * j + 3, column=1)
        # ------------------continuous rows and discrete column------------------------#
        elif variable.get() == OPTIONS[2]:
            i = 0
            label_names = []
            btn_names = []
            sheet_nameL = []
            sheet_nameE = []
            srl_names = []
            sre_names = []
            erl_names = []
            ere_names = []
            cl_names = []
            ce_names = []
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
                    try:
                        sr = int(sre_names[i].get()) - 1
                    except:
                        messagebox.showerror("error", "Start row must be an integer")
                        root.destroy()
                    try:
                        er = int(ere_names[i].get()) - 1
                    except:
                        messagebox.showerror("error", "End row must be an integer")
                        root.destroy()
                    col = ce_names[i].get().split(',')
                    k = 0
                    while k < len(col):
                        try:
                            c = int(col[k]) - 1
                            col[k] = str(c)
                        except:
                            messagebox.showerror('error', str(k + 1) + ' column entry of sheet ' + str(
                                i + 1) + ' is not integer')
                            root.destroy()
                        k += 1
                    row_dat.append(['c', str(sr), str(er), str(er - sr + 1)])
                    col_dat.append(['d', col, str(len(col))])
                    i += 1
                # print(sheet_dat)
                # print(row_dat)
                # print(col_dat)
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
                            messagebox.showerror('error', 'the number of rows should be same')
                        columns = col_dat[i]
                        df1 = make_othrdataframes(sheet_dat, rows, columns, i)
                        if nr != 0:
                            df1.index = row_indices
                        # print(df1)
                        df = pd.concat([df, df1], axis=1)
                        i += 1

                else:
                    rows = row_dat[0]
                    columns = col_dat[0]
                    df = make_firstdataframe(sheet_dat, rows, columns)

                dfw = df.copy()
                # print(dfw)
                with pd.ExcelWriter(output + '.xlsx') as writer:  # doctest: +SKIP
                    dfw.to_excel(writer, sheet_name='Sheet1', header=False, index=False)
                root.destroy()

            text = Label(root, text='                               ')
            text.grid(row=6 * j + 2)
            btn_final = Button(root, text='submit', command=final_subcnd)
            btn_final.grid(row=6 * j + 3, column=1)
        # ---------------discrete rows and continuous col------------------------#
        elif variable.get() == OPTIONS[1]:
            i = 0
            label_names = []
            btn_names = []
            sheet_nameL = []
            sheet_nameE = []
            scl_names = []
            sce_names = []
            ecl_names = []
            ece_names = []
            rl_names = []
            re_names = []
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
                            r = int(row[k]) - 1
                            row[k] = str(r)
                        except:
                            messagebox.showerror('error',
                                                 str(k + 1) + ' row entry of sheet ' + str(i + 1) + ' is not integer')
                            root.destroy()
                        k += 1
                    try:
                        sc = int(sce_names[i].get()) - 1
                    except:
                        messagebox.showerror("error", "Start column must be an integer")
                        root.destroy()
                    try:
                        ec = int(ece_names[i].get()) - 1
                    except:
                        messagebox.showerror("error", "End column must be an integer")
                        root.destroy()
                    row_dat.append(['d', row, str(len(row))])
                    col_dat.append(['c', str(sc), str(ec), str(ec - sc + 1)])
                    i += 1
                # print(sheet_dat)
                # print(row_dat)
                # print(col_dat)
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
                            messagebox.showerror('error', 'the number of rows should be same')
                        columns = col_dat[i]
                        df1 = make_othrdataframes(sheet_dat, rows, columns, i)
                        if nr != 0:
                            df1.index = row_indices
                        # print(df1)
                        df = pd.concat([df, df1], axis=1)
                        i += 1

                else:
                    rows = row_dat[0]
                    columns = col_dat[0]
                    df = make_firstdataframe(sheet_dat, rows, columns)

                dfw = df.copy()
                # print(dfw)
                with pd.ExcelWriter(output + '.xlsx') as writer:  # doctest: +SKIP
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
            j = 0
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
            out_lab = Label(root, text='output file name')
            out_lab.grid(row=5 * j + 1)
            out_ent = Entry(root)
            out_ent.grid(row=5 * j + 1, column=1)

            def final_subdnd():
                i = 0
                while i < num_of_sheets:

                    chk = sheet_nameE[i].get().strip() == '' or sheet_locs[i] == ''
                    if chk:
                        messagebox.showerror('error', 'sheet name or sheet location ' + str(i + 1) + ' is blank')
                        root.destroy()
                    else:
                        sheet_dat.append([sheet_locs[i], sheet_nameE[i].get().strip()])

                    chk2 = out_ent.get() == ''
                    if chk2:
                        messagebox.showerror('error', 'output sheet name is blank')
                        root.destroy()
                    else:
                        output = out_ent.get()
                    row = re_names[i].get().split(',')
                    k = 0
                    while k < len(row):
                        try:
                            r = int(row[k]) - 1
                            row[k] = str(r)
                        except:
                            messagebox.showerror('error',
                                                 str(k + 1) + ' row entry of sheet ' + str(i + 1) + ' is not integer')
                            root.destroy()
                        k += 1
                    col = ce_names[i].get().split(',')
                    k = 0
                    while k < len(col):
                        try:
                            c = int(col[k]) - 1
                            col[k] = str(c)
                        except:
                            messagebox.showerror('error', str(k + 1) + ' column entry of sheet ' + str(
                                i + 1) + ' is not integer')
                            root.destroy()
                        k += 1
                    col_dat.append(['d', col, str(len(col))])
                    row_dat.append(['d', row, str(len(row))])
                    i += 1
                # print(sheet_dat)
                # print(row_dat)
                # print(col_dat)
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
                            messagebox.showerror('error', 'the number of rows should be same')
                        columns = col_dat[i]
                        df1 = make_othrdataframes(sheet_dat, rows, columns, i)
                        if nr != 0:
                            df1.index = row_indices
                        # print(df1)
                        df = pd.concat([df, df1], axis=1)
                        i += 1

                else:
                    rows = row_dat[0]
                    columns = col_dat[0]
                    df = make_firstdataframe(sheet_dat, rows, columns)

                dfw = df.copy()
                # print(dfw)
                with pd.ExcelWriter(output + '.xlsx') as writer:  # doctest: +SKIP
                    dfw.to_excel(writer, sheet_name='Sheet1', header=False, index=False)
                root.destroy()

            text = Label(root, text='                             ')
            text.grid(row=5 * j + 2)
            btn_final = Button(root, text='submit', command=final_subdnd)
            btn_final.grid(row=5 * j + 3, column=1)

    text0 = Label(root, text="                                  ")
    text0.grid(row=3)
    b = Button(root, text="submit", command=submit)
    b.grid(row=4, column=1)
    root.mainloop()


def gui_plotting():
    # -------------- input final organised dataframe ----------------#
    def get_plot(df, plot_type):
        if plot_type == 'Pie Chart':
            df.plot.pie(subplots=True, autopct='%.2f')
        elif plot_type == 'Bar Plot':
            df.plot.bar()
        elif plot_type=='Line Graph':
            df.plot.line()
        plt.show()

    def ret_num_df(df, rcdata):
        row_dat = rcdata[0]
        col_dat = rcdata[1]
        # ----------- continuous rows and columns ---------------#
        if row_dat[0] == 1 and col_dat[0] == 1:
            rows = df.iloc[row_dat[2]:row_dat[3], [row_dat[1]]]
            rows = rows.values
            rnames = []
            for val in rows:
                rnames.append(val[0].strip())
            cols = df.iloc[[col_dat[1]], col_dat[2]:col_dat[3]]
            cols = cols.T
            cols = cols.values
            cnames = []
            for val in cols:
                cnames.append(val[0].strip())
            start_row = row_dat[2]
            end_row = row_dat[3]
            start_col = col_dat[2]
            end_col = col_dat[3]
            df = df.iloc[start_row:end_row, start_col:end_col]
            df = df.astype('float64')
            df.index = rnames
            df.columns = cnames
        # ------- continuous cols and discrete rows----#
        elif row_dat[0] == 0 and col_dat[0] == 1:
            rows = df.iloc[row_dat[2], row_dat[1]]
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
            df = df.astype('float64')
            df.index = rnames
            df.columns = cnames
        # -------- discrete col and rows -------#
        elif row_dat[0] == 0 and col_dat[0] == 0:
            rows = df.iloc[row_dat[2], row_dat[1]]
            rows = rows.values
            rnames = []
            for val in rows:
                rnames.append(val.strip())
            # print(rnames)
            cols = df.iloc[col_dat[1], col_dat[2]]
            cols = cols.values
            cnames = []
            for val in cols:
                cnames.append(val.strip())
            # print(cnames)
            df = df.iloc[row_dat[2], col_dat[2]]
            df = df.astype('float64')
            df.index = rnames
            df.columns = cnames
        # ----------- discrete cols and continuous rows-------------#
        elif row_dat[0] == 1 and col_dat[0] == 0:
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
            df = df.astype('float64')
            df.index = rnames
            df.columns = cnames
        # print(df)
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
    root.title('PlotIT')
    OPTIONS = ['Discrete Rows and Discrete Columns', 'Discrete Rows and Continuous Columns',
               'Continuous Rows and Discrete Columns', 'Continous Rows and Continuous Columns']
    variable = StringVar(root)
    variable.set(OPTIONS[3])
    root.geometry('450x150')
    textn = Label(root, text="                              ")
    textn.grid(row=0)
    loc_sheetsL = Label(root, text='Select Sheet')
    loc_sheetsL.grid(row=1)
    loc_sheetsB = Button(root, text='Browse', command=browse_btn_excel)
    loc_sheetsB.grid(row=1, column=1)
    sheet_namL = Label(root, text='Sheet Name')
    sheet_namL.grid(row=2)
    sheet_namE = Entry(root)
    sheet_namE.grid(row=2, column=1)
    type_label = Label(root, text=' Row and Column Selection type')
    type_label.grid(row=3)
    type_entry = OptionMenu(root, variable, *OPTIONS)
    type_entry.grid(row=3, column=1)

    def submit():
        root.geometry('450x250')
        try:
            print(xlfilename)
        except:
            messagebox.showerror('error', 'please select a file')
            root.destroy()
        global sheetname
        sheetname = sheet_namE.get()
        chk = sheetname == ''
        if chk:
            messagebox.showerror('error', 'Sheet name cannot be null')
            root.destroy()
        df = pd.read_excel(xlfilename, sheetname, header=None)
        # print(df)
        plotoptions = ['Bar Plot', 'Pie Chart',"Line Graph"]
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
            scl = Label(root, text='Column number where column names start')
            scl.grid(row=5)
            sce = Entry(root)
            sce.grid(row=5, column=1)
            ecl = Label(root, text='Column number where column names end')
            ecl.grid(row=6)
            ece = Entry(root)
            ece.grid(row=6, column=1)
            plt_type_label = Label(root, text=' Plot Type')
            plt_type_label.grid(row=7)
            plot_type = OptionMenu(root, var, *plotoptions)
            plot_type.grid(row=7, column=1)

            def intrnl_submitcnc():
                try:
                    cnm_row = int(row_withcol_names_ent.get()) - 1
                except:
                    messagebox.showerror("Error", 'Entries must be integer')
                    root.destroy()
                try:
                    cnm_cs = int(sce.get()) - 1
                except:
                    messagebox.showerror("Error", 'Entries must be integer')
                    root.destroy()
                try:
                    cnm_ce = int(ece.get())
                except:
                    messagebox.showerror("Error", 'Entries must be integer')
                    root.destroy()
                try:
                    rnm_col = int(col_withrow_names_ent.get()) - 1
                except:
                    messagebox.showerror("Error", 'Entries must be integer')
                    root.destroy()
                try:
                    rnm_rs = int(sre.get()) - 1
                except:
                    messagebox.showerror("Error", 'Entries must be integer')
                    root.destroy()
                try:
                    rnm_re = int(ere.get())
                except:
                    messagebox.showerror("Error", 'Entries must be integer')
                    root.destroy()
                col_index = [1, cnm_row, cnm_cs, cnm_ce]
                row_index = [1, rnm_col, rnm_rs, rnm_re]
                row_col_dat = [row_index, col_index]
                dfp = ret_num_df(df, row_col_dat)
                get_plot(dfp, var.get())

            txt = Label(root, text='                      ')
            txt.grid(row=8)
            btn = Button(root, text='submit', command=intrnl_submitcnc)
            btn.grid(row=9, column=1)
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
                    cnm_row = int(row_withcol_names_ent.get()) - 1
                except:
                    messagebox.showerror("Error", 'Entries must be integer')
                    root.destroy()
                try:
                    rnm_col = int(col_withrow_names_ent.get()) - 1
                except:
                    messagebox.showerror("Error", 'Entries must be integer')
                    root.destroy()
                try:
                    rnm_rs = int(sre.get()) - 1
                except:
                    messagebox.showerror("Error", 'Entries must be integer')
                    root.destroy()
                try:
                    rnm_re = int(ere.get())
                except:
                    messagebox.showerror("Error", 'Entries must be integer')
                    root.destroy()
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
                row_index = [1, rnm_col, rnm_rs, rnm_re]
                row_col_dat = [row_index, col_index]
                dfp = ret_num_df(df, row_col_dat)
                get_plot(dfp, var.get())

            txt = Label(root, text='                      ')
            txt.grid(row=7)
            btn = Button(root, text='submit', command=intrnl_submitcnd)
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
                    cnm_row = int(row_withcol_names_ent.get()) - 1
                except:
                    messagebox.showerror("Error", 'Entries must be integer')
                    root.destroy()
                try:
                    cnm_cs = int(sce.get()) - 1
                except:
                    messagebox.showerror("Error", 'Entries must be integer')
                    root.destroy()
                try:
                    cnm_ce = int(ece.get())
                except:
                    messagebox.showerror("Error", 'Entries must be integer')
                    root.destroy()
                try:
                    rnm_col = int(col_withrow_names_ent.get()) - 1
                except:
                    messagebox.showerror("Error", 'Entries must be integer')
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
                row_col_dat = [row_index, col_index]
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
            btn = Button(root, text='submit', command=intrnl_submitdnc)
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
                    cnm_row = int(row_withcol_names_ent.get()) - 1
                except:
                    messagebox.showerror("Error", 'Entries must be integer')
                    root.destroy()
                try:
                    rnm_col = int(col_withrow_names_ent.get()) - 1
                except:
                    messagebox.showerror("Error", 'Entries must be integer')
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
                row_col_dat = [row_index, col_index]
                dfp = ret_num_df(df, row_col_dat)
                get_plot(dfp, var.get())

            txt = Label(root, text='                      ')
            txt.grid(row=6)
            btn = Button(root, text='submit', command=intrnl_submitdnd)
            btn.grid(row=7, column=1)

    text0 = Label(root, text="                                  ")
    text0.grid(row=4)
    b = Button(root, text="submit", command=submit)
    b.grid(row=5, column=1)
    root.mainloop()


def gui_cagr():
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

    def read_sheet(loc, sheet_name, start, end, rnamecol, pcol, ncol):
        pcol = pcol - 1
        ncol = ncol - 1
        rnamecol = rnamecol - 1
        start = start - 2
        df = pd.read_excel(loc, sheet_name, skiprows=start, header=None)
        df = df[:end - start]
        df = df.iloc[:, [rnamecol, pcol, ncol]]
        df = df.T
        return df

    def get_cagr(df, yrdiff):
        ndarr = df.iloc[[1, 2], 1:].astype('float64').values
        tminus1yr = ndarr[0]
        tyr = ndarr[1]
        out = np.divide(tyr, tminus1yr)
        exponent = 1.0 / yrdiff
        out_with_power = (np.power(out, exponent) - 1) * 100
        data = ['CAGR']
        for item in out_with_power:
            data.append(item)
        data_df = pd.DataFrame(data)
        df = pd.concat([df.T, data_df], axis=1)
        return df

    root = Tk()
    root.geometry('400x275')
    textn = Label(root, text="                              ")
    textn.grid(row=0)
    text1 = Label(root, text="Select Excel file")
    text1.grid(row=1)
    btn1 = Button(root, text="Browse excel file", command=browse_btn_excel)
    btn1.grid(row=1, column=1)
    sheetnamelabel = Label(root, text="Enter Sheet Name")
    sheetnamelabel.grid(row=2)
    sheetnameentry = Entry(root)
    sheetnameentry.grid(row=2, column=1)
    startrowlabel = Label(root, text="Enter Start row")
    startrowlabel.grid(row=3)
    startrowentry = Entry(root)
    startrowentry.grid(row=3, column=1)
    endrowlabel = Label(root, text='Enter End row')
    endrowlabel.grid(row=4)
    endrowentry = Entry(root)
    endrowentry.grid(row=4, column=1)
    rnmcollabel = Label(root, text='Enter Column containing row names')
    rnmcollabel.grid(row=5)
    rnmcolentry = Entry(root)
    rnmcolentry.grid(row=5, column=1)
    pcollabel = Label(root, text="Enter First Survey Data Column")
    pcollabel.grid(row=6)
    pcolentry = Entry(root)
    pcolentry.grid(row=6, column=1)
    ncollabel = Label(root, text='Enter Next Survey Data Column')
    ncollabel.grid(row=7)
    ncolentry = Entry(root)
    ncolentry.grid(row=7, column=1)
    yrL = Label(root, text='Enter Year Difference b/w Survey')
    yrL.grid(row=8)
    yrE = Entry(root)
    yrE.grid(row=8, column=1)
    outputl = Label(root, text="Enter Output file Name")
    outputl.grid(row=9)
    outputentry = Entry(root)
    outputentry.grid(row=9, column=1)

    def submit():
        global startrow
        global endrow
        global yrdiff
        global rnmcol
        global pcol
        global ncol
        global sheetname
        global outputname
        outputname = outputentry.get()
        sheetname = sheetnameentry.get()
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
        df = get_cagr(read_sheet(xlfilename, sheetname, startrow, endrow, rnmcol, pcol, ncol), yrdiff)
        with pd.ExcelWriter(outputname + '.xlsx') as writer:  # doctest: +SKIP
            df.to_excel(writer, sheet_name='Sheet1', header=False, index=False)
        root.destroy()

    text0 = Label(root, text="                                  ")
    text0.grid(row=10)
    submit = Button(root, text="Submit", command=submit)
    submit.grid(row=11, column=1)
    root.mainloop()

mainframe=Tk()
OPTIONS=['Get Excel sheet from Notepad File','Combine multiple Excel Files',"Get Plots From Excel Sheet",'Calculate CAGR']
variable=StringVar(mainframe)
variable.set(OPTIONS[0])
mainframe.geometry('450x150')
textn = Label(mainframe, text="                              ")
textn.grid(row=0)
sheet_namL=Label(mainframe,text='Select your Purpose')
sheet_namL.grid(row=1,column=1)
type_label= Label(mainframe, text=' Operation type')
type_label.grid(row=3)
type_entry=OptionMenu(mainframe,variable,*OPTIONS)
type_entry.grid(row=3,column=1)


def submit():
    if variable.get()==OPTIONS[0]:
        mainframe.destroy()
        gui_getxl()

    elif variable.get()==OPTIONS[1]:
        mainframe.destroy()
        gui_combine()

    elif variable.get()==OPTIONS[2]:
        mainframe.destroy()
        gui_plotting()
    elif variable.get()==OPTIONS[3]:
        mainframe.destroy()
        gui_cagr()

text0 = Label(mainframe, text="                                  ")
text0.grid(row=4)
b = Button(mainframe, text="submit", command=submit)
b.grid(row=5, column=1)
mainframe.mainloop()