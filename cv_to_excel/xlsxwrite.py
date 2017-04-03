def out(rowheights, columnwidths, values) :

    # import the python math module necessary for certain arithmetic computations
    import math
    import re
    import xlsxwriter

    def num2col(num):
        q,r=divmod(num-1,26)
        return num2col(q)+chr(r+ord('A')) if num!=0 else ''

    def toIdx(y):
        s,p=0,1
        for c in y[::-1]:
            s+=p*(int(c,36)-9)
            p*=26
        return s

    # When given a list of cells to merge, xlsxwriter requires that I provide the top-left cell and the bottom-right cell.
    # This method is used to compute those and works by sorting the cell values based on their columns and rows.
    # The row is multiplied by the maximum number of columns
    def char2num(char):
        mn_c=num2col(min(toIdx(re.split('(\D+)',c)[1]) for c in char))
        mx_c=num2col(max(toIdx(re.split('(\D+)',c)[1]) for c in char))
        mn_r=str(min(re.split('(\D+)',c)[2] for c in char))
        mx_r=str(max(re.split('(\D+)',c)[2] for c in char))
        return mn_c+mn_r+":"+mx_c+mx_r

    # Create an new Excel file and add a worksheet.
    workbook = xlsxwriter.Workbook('demo.xlsx')
    worksheet = workbook.add_worksheet()

    default_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter'
    })


    for i in range(len(rowheights)) :
        worksheet.set_row(i, rowheights[i])

    for j in range(len(columnwidths)) :
        worksheet.set_column(str(num2col(j+1))+":"+str(num2col(j+1)), columnwidths[j])

    for l in range(len(values)) :
        if len(values[l][0]) == 1:
            worksheet.write(values[l][0][0],values[l][1], default_format)
        else :
            worksheet.merge_range(char2num(values[l][0]), values[l][1], default_format)

    workbook.close()



##col = [20,25,30,35,40,45,50,55,60,65,70,75]
##row = [30,40,50,60,70,80,90,100,110,120,130]
##hello = [(['A1','B2','A3','C4','A4','C2','B1','C3','B3','C1','B4','A2','D1',"D2","D3","D4"],"First entry"), (['D5'],"Second entry"), (['E7','F8','E8','F7'],"second merge")]
##out(row, col, hello)
