##############################################################################
#
# A simple example of some of the features of the XlsxWriter Python module.
#
# Copyright 2013-2016, John McNamara, jmcnamara@cpan.org
#
import math
base26 = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z']

def num2char(num) :
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

def char2num(char) :

    # xlsxwriter supports excel's worksheet limits of 1,048,576 rows by 16,384 columns.

    value = 0

    cols = ""
    rows = ""

    counter = 0

    while cols == "" and rows == "" :
        if char[counter] == '0' or char[counter] == '1' or char[counter] == '2' or char[counter] == '3' or char[counter] == '4' or char[counter] == '5' or char[counter] == '6' or char[counter] == '7' or char[counter] == '8' or char[counter] == '9' :
            cols = char.split(char[counter],1)[0]
            rows = char[counter] + char.split(char[counter],1)[1]
        counter += 1

    print cols
    print rows

    #max excel length
    value += (2**20) * int(rows)

    def f(x):
        return {
            'A': 1,
            'B': 2,
            'C': 3,
            'D': 4,
            'E': 5,
            'F': 6,
            'G': 7,
            'H': 8,
            'I': 9,
            'J': 10,
            'K': 11,
            'L': 12,
            'M': 13,
            'N': 14,
            'O': 15,
            'P': 16,
            'Q': 17,
            'R': 18,
            'S': 19,
            'T': 20,
            'U': 21,
            'V': 22,
            'W': 23,
            'X': 24,
            'Y': 25,
            'Z': 26
        }[x]

    for i in range(len(cols)) :
        value += (26**i)*f(cols[len(cols)-i-1])

    print value

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
        worksheet.set_column(num2char(j+1), columnwidths[j])





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

char2num('ABHA183')

out(row, col)

workbook.close()

#rows[]

#worksheet.set_column('A:A',col[0])
