import pandas as pd
import matplotlib.pyplot as plt


# -------------- input final organised dataframe ----------------#
def get_plot(df):
    out=input('what plot do you need bar or pie?\n'+' please enter B or P:\n')
    chk=out is 'p'or 'b'
    while not chk:
        out=input(' please enter B or P:\n')
        chk = out is 'p' or 'b'
    if out.lower()=='p':
        df.plot.pie(subplots=True,autopct='%.2f')
    elif out.lower()=='b':
        df.plot.bar()
    plt.show()


# -------------- input raw dataframe ----------------#
def get_row_and_colnames():
    print('#----------------------collection of row data------------------------#')
    type_of_rowdat = input('mention your data arrangement type:\ndiscrete or continuous please enter D/C?\n')
    chk2 = type_of_rowdat.lower() == 'd'
    chk1 = type_of_rowdat.lower() == 'c'
    chk = chk1 or chk2
    row_index = []
    while not chk:
        type_of_rowdat = input('Invalid response! please enter D/C?\n')
        chk2 = type_of_rowdat.lower() == 'd'
        chk1 = type_of_rowdat.lower() == 'c'
        chk = chk1 or chk2
    if chk1:
        rnm_col = input('enter the column which contains row names:\n')
        chk = rnm_col.isdigit()
        while not chk:
            rnm_col = input('enter an integer:\n')
            chk = rnm_col.isdigit()
        rnm_col = int(rnm_col) - 1
        rnm_rs = input('enter the row number where row names start:\n')
        chk = rnm_rs.isdigit()
        while not chk:
            rnm_rs = input('enter an integer:\n')
            chk = rnm_rs.isdigit()
        rnm_rs = int(rnm_rs) - 1
        rnm_re = input('enter the row number where row names end:\n')
        chk = rnm_re.isdigit()
        while not chk:
            rnm_re = input('enter an integer:\n')
            chk = rnm_re.isdigit()
        rnm_re = int(rnm_re)
        row_index = [1, rnm_col, rnm_rs, rnm_re]
    # ------------------ '1' means continuous row data------------------------#
    # ------------------ '0' means discrete row data -------------------------#
    elif chk2:
        rnm_col = input('enter the column which contains row names:\n')
        chk = rnm_col.isdigit()
        while not chk:
            rnm_col = input('enter an integer:\n')
            chk = rnm_col.isdigit()
        rnm_col = int(rnm_col) - 1
        num_row = input('enter the number of row selections(integer):\n')
        chk = num_row.isdigit()
        while not chk:
            num_row = input('enter an integer:\n')
            chk = num_row.isdigit()
        num_row = int(num_row)
        i = 0
        list_of_rows = []
        while i < num_row:
            row_num = input("enter the number of " + str(i+1) + 'th row in your selection:\n')
            chk = row_num.isdigit()
            while not chk:
                row_num = input('enter an integer:\n')
                chk = row_num.isdigit()
            row_num = int(row_num) - 1
            list_of_rows.append(row_num)
            i += 1
        row_index = [0, rnm_col, list_of_rows]
# ----------------- column data -----------------------#
    print('#-----------------collection of column data------------------------#')
    type_of_coldat=input('mention your data arrangement type:\ndiscrete or continuous please enter D/C?\n')
    chk2 = type_of_coldat.lower() == 'd'
    chk1 = type_of_coldat.lower() == 'c'
    chk = chk1 or chk2
    col_index=[]
    while not chk:
        type_of_coldat = input('Invalid response! please enter D/C?\n')
        chk2 = type_of_coldat.lower() == 'd'
        chk1 = type_of_coldat.lower() == 'c'
        chk = chk1 or chk2
    if chk1:
        cnm_row=input('enter the row which contains column names:\n')
        chk = cnm_row.isdigit()
        while not chk:
            cnm_row=input('enter an integer:\n')
            chk=cnm_row.isdigit()
        cnm_row=int(cnm_row)-1
        cnm_cs=input('enter the column number where column names start:\n')
        chk=cnm_cs.isdigit()
        while not chk:
            cnm_cs = input('enter an integer:\n')
            chk = cnm_cs.isdigit()
        cnm_cs = int(cnm_cs) - 1
        cnm_ce = input('enter the column number where column names end:\n')
        chk = cnm_ce.isdigit()
        while not chk:
            cnm_ce = input('enter an integer:\n')
            chk = cnm_ce.isdigit()
        cnm_ce = int(cnm_ce)
        col_index=[1,cnm_row,cnm_cs,cnm_ce]
    # ------------------ '1' means continuous column data------------------------#
    # ------------------ '0' means discrete column data -------------------------#
    elif chk2:
        cnm_row = input('enter the row which contains column names:\n')
        chk = cnm_row.isdigit()
        while not chk:
            cnm_row = input('enter an integer:\n')
            chk = cnm_row.isdigit()
        cnm_row = int(cnm_row) - 1
        num_col = input('enter the number of column selections(integer):\n')
        chk = num_col.isdigit()
        while not chk:
            num_col = input('enter an integer:\n')
            chk = num_col.isdigit()
        num_col=int(num_col)
        i=0
        list_of_cols=[]
        while i<num_col:
            col_num=input("enter the number of "+str(i+1)+'th column in your selection:\n')
            chk = col_num.isdigit()
            while not chk:
                col_num = input('enter an integer:\n')
                chk = col_num.isdigit()
            col_num = int(col_num)-1
            list_of_cols.append(col_num)
            i+=1
        col_index=[0,cnm_row,list_of_cols]
    return [row_index, col_index]


def ret_raw_df():
    flname = input('please enter the address of your excel file:\n')
    df = pd.read_excel(flname, header=None)
    print(df)
    return df


# -------------- return numeric data ----------#
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
    return df


dfp = ret_raw_df()
rcdat = get_row_and_colnames()
print(rcdat)
dfp=ret_num_df(dfp,rcdat)
print(dfp)
get_plot(dfp)