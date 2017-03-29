import math

base26 = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z']

def find_powers(num) :

    cell = ""
    formatted_cell = ""

    i = 1
    counter = 1

    while num >= i:
        i += 26**counter
        print num,"is at least",counter,"digits"
        counter += 1

    digits = counter - 1

    print num,"is",digits,"digits"
    print math.log(num,26)

    for i in range(digits) :
        num -= 26**i

    print num

    #cell_is_A=0
    #if num==0 : cell_is_A=1


    ##num is now a regular integer.

    base10_26 = []

    for i in range(digits) :

        base10_26.append(0)

        print 'power of 26 we are dividing by: 26^',digits - 1 - i

        temp_quotient = int(num/(26**(digits-1-i)))
        print 'temporary quotient',temp_quotient

        num -= (26**(digits-1-i)) * temp_quotient
        print 'number after subtraction:',num

        base10_26[i] += temp_quotient

        ##num -= (26**(((new_digits-1)-i))) * base10_26[i]

    print 'array values:'
    for i in base10_26 :
        print i

    for i in range(len(base10_26)) :
        cell += base26[base10_26[i]]

    print "cell:", cell

    formatted_cell = cell + ":" + cell
    print formatted_cell


#---------------------------------------------
    import xlsxwriter

    # Create an new Excel file and add a worksheet.
    workbook = xlsxwriter.Workbook('demo.xlsx')
    worksheet = workbook.add_worksheet()

    # Widen the first column to make the text clearer.
    worksheet.set_column(formatted_cell, 200)

    worksheet.set_row("1:1",400)

    # Add a bold format to use to highlight cells.
    bold = workbook.add_format({'bold': True})

    # Write some simple text.
    worksheet.write('A1', 'Hello')

    # Text with formatting.
    worksheet.write('A2', 'World', bold)

    # Write some numbers, with row/column notation.
    worksheet.write(1, 1, 123)
    worksheet.write(3, 1, 123.456)

    merge_format = workbook.add_format({
        'bold': 1,
        'border': 2,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': 'yellow'})

    # Merge 3 cells.
    worksheet.merge_range('B4:E9', 'Merged Range', merge_format)

    workbook.close()

    #rows[]

    #worksheet.set_column('A:A',col[0])

    #col = [54,21,168,74]
    #row = [27,39,17]

    def write_values(values) :

        #given a 2-dimensional array "values"

        for i in range(len(values)) :
            for j in range(len(values[i]))
                worksheet.write(i,j,values[i,j])


#-----------------------------------------------------------

find_powers(1)
