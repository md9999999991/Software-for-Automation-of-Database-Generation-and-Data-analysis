from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import xlsxwriter


def calc_col_num(start_pos,max_num_of_col,starts_max_col):
    z=0
    start_pos=int(start_pos)
    diff_arr =[]
    while z<max_num_of_col:
        diff_arr.append(abs(start_pos-starts_max_col[z]))
        z+=1

    return diff_arr.index(min(diff_arr))


def checkpt(matrix_data,curr_row,curr_col):
    if curr_row>0:
        chk2=False
        i=0
        while i < len(matrix_data[curr_row-1]):
            if curr_col<=matrix_data[curr_row-1][i][1]:
                chk2=True
            i+=1
        return chk2
    else:
        return False


def recheck_col(my_lst,matrix,cr,cc):
    diff_arr=[]
    i=0
    while i<len(matrix[cr-1]):
        diff_arr.append(abs(int(my_lst[cr-1][i][1])-int(my_lst[cr][cc][1])))
        i+=1
    j= diff_arr.index(min(diff_arr))
    return matrix[cr-1][j][1]


def browse_btn_notepad():
    global npfilename
    try:
        npfilename = filedialog.askopenfilename(initialdir="/", title="Select file", filetypes=(("text files", "*.txt"), ("all files", "*.*")))
    except:
        messagebox.showerror("error", "please select a file")
    #print(npfilename)


root = Tk()
root.title('Notepad To Excel')
root.geometry('250x150')
text0 = Label(root, text ="                                  ")
text0.grid(row =0)
# for the icon use: root.iconbitmap(r'location')
text2 = Label(root, text="Select notepad file")
text2.grid(row=1)
btn2 = Button(root, text='Browse', command=browse_btn_notepad)
btn2.grid(row=1,column=1)

outputlabel= Label(root, text="Output file Name")
outputlabel.grid(row=2)
outputentry = Entry(root)
outputentry.grid(row=2,column=1)


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
                if count < 3:
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
    workbook = xlsxwriter.Workbook(outputname+'.xlsx')
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


text0 = Label(root, text ="                                  ")
text0.grid(row =11)
submit = Button(root, text="Submit", command=close_window)
submit.grid(row=12, column=1)
root.mainloop()