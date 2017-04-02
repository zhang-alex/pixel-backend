##############################################################################
#
# A simple example of some of the features of the XlsxWriter Python module.
#
# Copyright 2013-2016, John McNamara, jmcnamara@cpan.org
#
import math
base26 = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z']

def find_powers(num) :
    i = 1
    counter = 1
    while num >= i:
        i += 26**counter
        counter += 1
    digits = counter - 1
    for i in range(digits) :
        num -= 26**i
    cell_is_A=0
    if num==0 : cell_is_A=1
    base10_26 = []
    for i in range(digits) :
        base10_26.append(0)
        temp_quotient = int(num/(26**(digits-1-i)))
        num -= (26**(digits-1-i)) * temp_quotient
        base10_26[i] += temp_quotient
    cell = ""
    for i in range(len(base10_26)) :
        cell += base26[base10_26[i]]

    return cell + ":" + cell

import xlsxwriter

temp = ""

# Create an new Excel file and add a worksheet.
workbook = xlsxwriter.Workbook('demo.xlsx')
worksheet = workbook.add_worksheet()


col = [54,21,168,74]
row = [27,39,17]

def out(rowheights, columnwidths) :

    for i in range(len(rowheights)) :
        worksheet.set_row(i, rowheights[i])

    for j in range(len(columnwidths)) :
        print find_powers(j+1)
        worksheet.set_column(find_powers(j+1), columnwidths[j])



# Write some simple text.
# worksheet.write('A1', 'Hello')

# Text with formatting.
# worksheet.write('A2', 'World')

# Write some numbers, with row/column notation.
# worksheet.write(2, 0, 123)
# worksheet.write(3, 0, 123.456)

# merge_format = workbook.add_format({
#     'align': 'center',
#     'valign': 'vcenter',
# })

# Merge 3 cells.
# worksheet.merge_range('B4:E9', 'Merged Range', merge_format)

out(row, col)

workbook.close()

#rows[]

#worksheet.set_column('A:A',col[0])
