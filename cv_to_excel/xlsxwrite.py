# import the python math module necessary for certain arithmetic computations
import math

# an ordered list of characters from the English alphabet
base26 = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z']

# this method will convert a number to its corresponding excel column value.
# It's essentially a base-26 conversion, but with a few modifications
def num2col(num) :
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

# When given a list of cells to merge, xlsxwriter requires that I provide the top-left cell and the bottom-right cell.
# This method is used to compute those and works by sorting the cell values based on their columns and rows.
# The row is multiplied by the maximum number of columns
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

    #max excel length
    value += (2**14) * int(rows)

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

    return value

import xlsxwriter

temp = ""

# Create an new Excel file and add a worksheet.
workbook = xlsxwriter.Workbook('demo.xlsx')
worksheet = workbook.add_worksheet()


default_format = workbook.add_format({
     'align': 'center',
     'valign': 'vcenter',
})

cell_vals = []

def out(rowheights, columnwidths, values) :

    cell_vals[:] = []

    for i in range(len(rowheights)) :
        worksheet.set_row(i, rowheights[i])

    for j in range(len(columnwidths)) :
        worksheet.set_column(num2col(j+1), columnwidths[j])




    for l in range(len(values)) :

        cell_vals[:] = []

        print ""
        if len(values[l]) == 2:
            worksheet.write(values[l][0],values[l][1], default_format)
        else :
            for k in range(len(values[l])-1) : #0,1,2
                cell_vals.append(char2num(values[l][k]))

            print "sorted"
            print  cell_vals

            new_cell_vals = sorted(cell_vals)

            smallest = 0
            largest = 0

            print new_cell_vals

            smallest = new_cell_vals[0]
            largest = new_cell_vals[len(new_cell_vals)-1]

            print "smallest:", smallest
            print "largest:", largest

            for i in range(len(cell_vals)) :
                if smallest == cell_vals[i] :
                    smallest = i
                elif largest == cell_vals[i] :
                    largest = i

            print smallest
            print largest

            asdf = str(values[l][smallest]) + ":" + str(values[l][largest])

            worksheet.merge_range(asdf, values[l][len(values[l])-1], default_format)


col = [20,25,30,35,40,45,50,55,60,65,70,75]
row = [30,40,50,60,70,80,90,100,110,120,130]
hello = [('A1','B2','A3','C4','A4','C2','B1','C3','B3','C1','B4','A2',"First entry"), ('D5',"Second entry"), ('E7','F8','E8','F7',"second merge")]
out(row, col, hello)



workbook.close()





# Write some simple text.
# worksheet.write('A1', 'Hello')

# Text with formatting.
# worksheet.write('A2', 'World')

# Write some numbers, with row/column notation.
# worksheet.write(2, 0, 123)
# worksheet.write(3, 0, 123.456)



# Merge 3 cells.
# worksheet.merge_range('B4:E9', 'Merged Range', merge_format)

#rows[]

#worksheet.set_column('A:A',col[0])
