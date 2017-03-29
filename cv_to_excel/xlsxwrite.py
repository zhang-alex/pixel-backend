##############################################################################
#
# A simple example of some of the features of the XlsxWriter Python module.
#
# Copyright 2013-2016, John McNamara, jmcnamara@cpan.org
#
import xlsxwriter


# Create an new Excel file and add a worksheet.
workbook = xlsxwriter.Workbook('demo.xlsx')
worksheet = workbook.add_worksheet()

# Widen the first column to make the text clearer.
worksheet.set_column('A:A', 200)

worksheet.set_row("1:1",400)

# Add a bold format to use to highlight cells.
bold = workbook.add_format({'bold': True})

# Write some simple text.
worksheet.write('A1', 'Hello')

# Text with formatting.
worksheet.write('A2', 'World', bold)

# Write some numbers, with row/column notation.
worksheet.write(2, 0, 123)
worksheet.write(3, 0, 123.456)

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
