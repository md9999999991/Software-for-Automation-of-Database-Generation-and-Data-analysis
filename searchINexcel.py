import pandas as pd
import xlsxwriter

# --------------------- custom rows------------------#


def getrows(num):
    row_sel=input('for sheet '+str(num+1)+' discrete, continuous or all rows please enter D/C/A?\n')
    chk2 = row_sel.lower()== 'd'
    chk1=row_sel.lower()=='c'
    chk3 = row_sel.lower()=='a'
    chk = chk1 or chk2 or chk3
    while not chk:
        row_sel = input('Invalid response! please enter D/C/A?\n')
        chk2 = row_sel.lower() == 'd'
        chk3 = row_sel.lower() == 'a'
        chk1 = row_sel.lower() == 'c'
        chk = chk1 or chk2 or chk3

    if chk2:
        print('discrete')
        num_of_discR = input('please enter number of rows u want:\n' )
        checkpt = num_of_discR.isdigit() and not int(num_of_discR)==0
        while not checkpt:
            num_of_discR = input('enter a non zero integer value for number of rows:\n')
            checkpt = num_of_discR.isdigit() and not int(num_of_discR)==0
        num_of_discR = int(num_of_discR)
        i=0
        print("enter values of discrete rows")
        rows = []
        while i<num_of_discR:
            dv = input('enter next value:\n')
            chk=dv.isdigit()
            while not chk:
                dv=input('enter integer value:\n')
                chk=dv.isdigit()
            rows.append(str(int(dv)-1))
            i+=1
        print(rows)
        return ['d',rows,str(num_of_discR)]
    elif chk1:
        print('continuous')
        start_row=input('please enter number of row to start with:\n')
        checkpt = start_row.isdigit()
        while not checkpt:
            start_row = input('enter an integer value for number of rows:\n')
            checkpt = start_row.isdigit()
        start_row=int(start_row)-1
        end_row=input('please enter number of row to end with:\n')
        checkpt = end_row.isdigit() and int(end_row)>start_row
        while not checkpt:
            end_row = input('enter an integer value for number of rows:\n')
            checkpt = end_row.isdigit() and int(end_row)>start_row
        end_row=int(end_row)-1
        print(start_row)
        print(end_row)
        num_contr=int(end_row)-int(start_row)+1
        return ['c', start_row, end_row,str(num_contr)]
    else:
        return ['a']
# ------ runs successfully------#


# ------------------------ custom columns --------------------#
def getcolumns(num):
    col_sel=input('for sheet '+str(num+1)+'discrete, continuous or all columns please enter D/C/A?\n')
    chk2 = col_sel.lower()== 'd'
    chk1=col_sel.lower()=='c'
    chk3 = col_sel.lower() == 'a'
    chk = chk1 or chk2 or chk3
    while not chk:
        col_sel = input('Invalid response! please enter D/C/A?\n')
        chk2 = col_sel.lower() == 'd'
        chk1 = col_sel.lower() == 'c'
        chk3 =col_sel.lower()=='a'
        chk = chk1 or chk2 or chk3

    if chk2:
        print('discrete')
        num_of_discC = input('please enter number of columns u want:\n')
        checkpt = num_of_discC.isdigit() and not int(num_of_discC)==0
        while not checkpt:
            num_of_discC = input('enter a non zero integer value for number of columns:\n')
            checkpt = num_of_discC.isdigit() and not int(num_of_discC)==0
        num_of_discC = int(num_of_discC)
        i=0
        print("enter values of discrete columns")
        cols = []
        while i<num_of_discC:
            dv = input('enter next value:\n')
            chk=dv.isdigit()
            while not chk:
                dv=input('enter integer value:\n')
                chk=dv.isdigit()
            cols.append(str(int(dv)-1))
            i+=1
        print(cols)
        return ['d',cols,str(num_of_discC)]
    elif chk1:
        print('continuous')
        start_col=input('please enter number of columns to start with:\n')
        checkpt = start_col.isdigit()
        while not checkpt:
            start_col = input('enter an integer value for number of columns:\n')
            checkpt = start_col.isdigit()
        start_col=int(start_col)-1
        end_col=input('please enter number of column to end with:\n')
        checkpt = end_col.isdigit() and int(end_col)>start_col
        while not checkpt:
            end_col = input('enter an integer value for number of columns:\n')
            checkpt = end_col.isdigit() and int(end_col)>start_col
        end_col=int(end_col)-1
        print(start_col)
        print(end_col)
        num_contC=int(end_col)-int(start_col)+1
        return ['c',start_col,end_col,str(num_contC)]
    else:
        return ['a']
# ------ runs successfully------#

# ----------- input excel files -------- #
def get_sheets():
    num_of_excel =input('input the number of sheets you want:\n')
    checkpt = num_of_excel.isdigit()
    while not checkpt:
        num_of_excel = input('enter an integer value for number of sheets:\n')
        checkpt = num_of_excel.isdigit()
    num_of_excel = int(num_of_excel)
    sheet_loc_and_name=[]
    i=0
    while i<num_of_excel:
        sheet_loc=input('enter the sheet '+str(i+1)+'\'s location:\n')
        sheet_name=input('enter the sheet '+str(i+1)+'\'s name:\n')
        sheet_loc_and_name.append([sheet_loc,sheet_name])
        i+=1
    return sheet_loc_and_name
# ------ runs successfully------#


# print(get_sheets())
# print(getrows())
# print(getcolumns())

# ---------- returning sliced dataframes for 1st sheet data-----------#
def make_firstdataframe(sheet_dat, rows, cols):
    df = pd.read_excel(sheet_dat[0][0], sheet_dat[0][1],header=None)
    if rows[0]=='c' and cols[0]=='c':
        startrow =rows[1]
        endrow = rows[2]+1
        startcol = cols[1]
        endcol= cols[2]+1
        df =df.iloc[startrow:endrow,startcol:endcol]
    elif rows[0]=='c' and cols[0]=='d':
        startrow = rows[1]
        endrow = rows[2] + 1
        col=[]
        for value in cols[1]:
            col.append(int(value))
        df = df.iloc[startrow:endrow,col]
    elif rows[0] == 'd' and cols[0] == 'c':
        startcol = cols[1]
        endcol = cols[2] + 1
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
    elif rows[0]=='a' and cols[0]=='c':
        startcol = cols[1]
        endcol = cols[2] + 1
        df = df.iloc[:, startcol:endcol]
    elif rows[0]=='a' and cols[0]=='d':
        col = []
        for value in cols[1]:
            col.append(int(value))
        df = df.iloc[:, col]
    elif rows[0]=='d' and cols[0]=='a':
        row = []
        for value in rows[1]:
            row.append(int(value))
        df = df.iloc[row,:]
    elif rows[0]=='c' and cols[0]=='a':
        startrow = rows[1]
        endrow = rows[2] + 1
        df = df.iloc[startrow:endrow,:]
    return df


# ----------- data frames for next sheets -----------------------#
def make_othrdataframes(sheet_dat, rows, cols,num):
    df = pd.read_excel(sheet_dat[num][0], sheet_dat[num][1],header=None)
    print(df)
    if rows[0]=='c' and cols[0]=='c':
        startrow =rows[1]
        endrow = rows[2]+1
        startcol = cols[1]
        endcol= cols[2]+1
        df =df.iloc[startrow:endrow,startcol:endcol]
    elif rows[0]=='c' and cols[0]=='d':
        startrow = rows[1]
        endrow = rows[2] + 1
        col=[]
        for value in cols[1]:
            col.append(int(value))
            print(value)
        df = df.iloc[startrow:endrow,col]
    elif rows[0] == 'd' and cols[0] == 'c':
        startcol = cols[1]
        endcol = cols[2] + 1
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
    elif rows[0] == 'a' and cols[0] == 'c':
        startcol = cols[1]
        endcol = cols[2] + 1
        df = df.iloc[:, startcol:endcol]
    elif rows[0] == 'a' and cols[0] == 'd':
        col = []
        for value in cols[1]:
            col.append(int(value))
        df = df.iloc[:, col]
    elif rows[0] == 'd' and cols[0] == 'a':
        row = []
        for value in rows[1]:
            row.append(int(value))
        df = df.iloc[row, 1:]
    elif rows[0] == 'c' and cols[0] == 'a':
        startrow = rows[1]
        endrow = rows[2] + 1
        df = df.iloc[startrow:endrow, 1:]
    return df


#print(make_dataframes(['p67_1.xlsx','Sheet1'],['c',3,79],['d',['0','5','6']]))
sheets_dat=get_sheets()
print(sheets_dat)
num_of_sheets=len(sheets_dat)
print(num_of_sheets)
# rows=getrows(1)
# cols=getcolumns(1)
if num_of_sheets>1:
    rows = getrows(0)
    if rows[0] is not 'a':
        nr=int(rows[-1])
        row_type=rows[0]
    else:
        row_type='a'
        nr =0
    columns = getcolumns(0)
    if columns[0]is not 'a':
        nc=int(columns[-1])
        col_type=columns[0]
    else:
        col_type='a'
    df = make_firstdataframe(sheets_dat, rows, columns)
    row_indices = []
    if nr!=0:
        n = 0
        while n<nr:
            row_indices.append(n)
            n+=1
        df.index = row_indices
    i=1
    while i<num_of_sheets:
        rows = getrows(i)
        chk=rows[0]==row_type
        chkr =True
        while not chk and chkr:
            print('the row selection must match with the first sheet')
            rows = getrows(i)
            if row_type is 'd'or 'c':
                chkr = int(rows[-1]) != nr
                if chkr:
                    print('the number of rows should be same')
            chk = rows[0] == row_type
        columns = getcolumns(i)
        df1 = make_othrdataframes(sheets_dat, rows, columns, i)
        if nr!=0:
            df1.index = row_indices
        print(df1)
        df=pd.concat([df,df1],axis=1)
        i+=1

else:
    rows = getrows(0)
    columns = getcolumns(0)
    df = make_firstdataframe(sheets_dat, rows, columns)


dfw=df.copy()
print(dfw)
with pd.ExcelWriter('output.xlsx') as writer:  # doctest: +SKIP
    dfw.to_excel(writer, sheet_name='Sheet_name_1', header = False,index = False)
# print(make_othrdataframes(sheets_dat,rows,cols,2))


