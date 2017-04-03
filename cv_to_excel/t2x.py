import sys
import os
import math
import re
import cv2
import numpy as np
from PIL import Image
import pytesseract
import xlsxwriter

def table_to_excel():

	# try to load source image from sys argv
    source = cv2.imread(sys.argv[1], cv2.IMREAD_COLOR)

    # check if image was loaded.  we like to work with actual images and not
    # bad paths.
    if source is None:
        print("Image failed to load with given path {}.".format(sys.argv[1]))
        return -1

    # resize array to 800 x 900 for more reasonable computation times
    resized = cv2.resize(source, (800, 900))

    # convert to grayscale
    grayscale = resized if len(resized.shape) == 2 else cv2.cvtColor(resized, cv2.COLOR_BGR2GRAY)

    # binarize
    binarize = cv2.adaptiveThreshold(cv2.bitwise_not(grayscale), 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY, 15, -2)

    # these images will be used to hold horizontal and vertical extracted
    # lines.
    horizontal = binarize.copy()
    vertical = binarize.copy()

    # scale is the amount of lines to be detected.  set to 100, arbitrarily. To be honest, I don't know what this does
    # to quality or speed.  Haven't benchmarked.
    scale = 100

    hSize = horizontal.shape[1] / scale
    vSize = vertical.shape[0] / scale

    # structuring elements for the two images
    horizontalStructure = cv2.getStructuringElement(cv2.MORPH_RECT, (hSize, 1))
    verticalStructure = cv2.getStructuringElement(cv2.MORPH_RECT, (1, vSize))

    # do morphological operations.
    # horizontal lines in table
    horizontal = cv2.dilate(cv2.erode(horizontal, horizontalStructure, (-1, -1)), horizontalStructure, (-1, -1))
    #vertical lines in table
    vertical = cv2.dilate(cv2.erode(vertical, verticalStructure, (-1, -1)), verticalStructure, (-1, -1))

    # create the mask: the sum of the horizontal and vertical lines. should look like the white outlines of tables.
    mask = horizontal + vertical

    # joints is, instead of sum of horizontal and vertical lines, the bitwise_and.  so you can see the joints of the table.
    joints = cv2.bitwise_and(horizontal, vertical)

    # find some contours to work with
    _,contours, hierarchy = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    # these two lists will be used to hold extracted Regions of Interest and Images, respectively.
    # the criteria for sorting is to count the number of joints.  If it is four or less joints, it is an image.
    # if it is more than 4 joints, then we assume it is a table.
    roi = []
    nroi = []
    img = []

    # the below code will sort images and ROIs based on the contours.
    for contour in contours:

        # filter out those contours that have an area less than 100 (set arbitrarily).
        # if the area is too small, we safely assume that it must be a bit of noise that isn't a table.
        if cv2.contourArea(contour) <= 100:
            continue

        # Approximate the shape of the contour.
        contour_poly = cv2.approxPolyDP(contour, 3, True)

        # if it's not a 4-sided figure, we don't want to work with it.  Filter pentagonal and hexagonal stuff out.
        if not len(contour_poly) == 4:
            continue

        # get bounding rectangle of our quadrilateral approximated polygon.
        x,y,w,h = cv2.boundingRect(contour_poly)

        # use this bounding rectangle to cut out the approximated polygon from the previously bitwise_and'ed
        # joints image.
        interestingRegion = joints[y:y+h,x:x+w]

        # find contours of this subimage.  This gives us the number of joints in the subimage.
        _, joints_contours, _ = cv2.findContours(interestingRegion, cv2.RETR_CCOMP, cv2.CHAIN_APPROX_SIMPLE)

        # if less than four points, then it must be a picture or something.
        if len(joints_contours) <= 4:
            # so, append it to the list of images, and keep going.
            img.append(np.copy(resized[y:y+h,x:x+w]))
            continue

        # if we are here, then the joints_contours must be more than four.  Add it to the list of tables.
        roi.append(np.copy(mask[y:y+h,x:x+w]))
        nroi.append(np.copy(grayscale[y:y+h,x:x+w]))
        cv2.rectangle(mask, (x,y), (x+w,y+h), (0, 255, 0), 1, 8, 0);

    # By now, we have taken the original document and extracted all pictures and tables from it.  Then, we sorted
    # them into two lists.  The pictures need no further work, they are to be exported as-is.

    # Now, we will work with the list of ROIs and begin the process of exporting them to the Excel file format.

    # perform HoughLinesP on each image to extract the horizontal and vertical lines while ignoring text.
    #for idx,table in enumerate(roi):

    # Extract vertical lines using the probabilistic hough transformlat
    vlines = cv2.HoughLinesP(roi[0], 1, np.pi, threshold=50, minLineLength=50, maxLineGap=1)

    # Extract horizontal lines using the probabilistic hough transform
    hlines = cv2.HoughLinesP(roi[0], 1, np.pi / 2, threshold=50, minLineLength=50, maxLineGap=1)

    # get height and width of the table, store them in some convenience variables
    (height, width) = roi[0].shape

    plain_table = np.zeros((height,width), np.uint8)
    for vline in vlines:
        for x1,y1,x2,y2 in vline:
            plain_table = cv2.line(plain_table, (x1,y1),(x2,y2),(255,255,255))
    for hline in hlines:
        for x1,y1,x2,y2 in hline:
            plain_table = cv2.line(plain_table, (x1,y1),(x2,y2),(255,255,255))

    # create an empty, black, binary image with the same dimensions as the ROI.
    vlined = np.zeros((height,width), np.uint8)

    # a really dumb array to hack through a bug later in the code.
    vlinehack = []

    # iterate through the vertical lines extracted by Hough and draw them out on the vlined image.
    for vline in vlines:
        for x1,y1,x2,y2 in vline:

            # extend all lines to the edge of the image.
            # since this is a vertical line, we extend y1 to the bottom of the image,
            # and y2 to the top.
            y1 = 0
            y2 = height
            vlinehack.append( [[x1,y1,x2,y2]] )
            vlined = cv2.line(vlined, (x1,y1), (x2,y2), (255,255,255))

    # create an empty, black, binary image with the same dimensions as the ROI.
    hlined = np.zeros((height,width), np.uint8)

    # yet another Really Dumb Array
    hlinehack = []

    # iterate through horizontal Hough lines and draw.
    for hline in hlines:
        for x1,y1,x2,y2 in hline:

            # extend all lines to the edge of image.

            # BUG: you will notice on the hline assignment line that the second coordinate
            # is (x2,y1) instead of (x2,y2).  This is because during the hough transform,
            # a few vertical "nubs" were extracted (and I am not sure why this is so).
            # if we were to use (x2,y2), we would get crazy diagonal lines.
            # TODO: fix this.  this is like a quick and dirty way to do it.
            x1 = 0
            x2 = width
            hlinehack.append( [[x1,y1,x2,y1]] )
            hlined = cv2.line(hlined, (x1,y1), (x2,y1), (255,255,255))

    cv2.imwrite("out/vlined.jpg", vlined)
    cv2.imwrite("out/hlined.jpg", hlined)

    # by now, we have two images, hlined and vlined, that represent the horizontal and vertical
    # lines within the ROI, respectively.

    # Next, we will choose a 1-pixel horizontal line within the vlined image, and based off of this,
    # count the number of columns that will be needed in the final Excel table, and calculate the widths
    # of each column.

    # minimum cell width or height. set to 10px, again, arbitrarily.
    mindim = 10

    # 1-pixel line is set to a horizontal line 10px above the bottom of the image, arbitrarily.
    # it shouldn't matter, as long as the line is within the image.
    y = 10

    # counter variable for checking and measuring column widths.
    c = 0

    # column widths stored in this list.
    cols = []

    # stores actual locations of columns.  will be used in conjunction with row_loc to produce
    # cell_loc, a universal identifier of cells based on cell location.
    col_loc = []

    # convenience variable
    start_loc = 0

    # iterate through the horizontal line of pixels
    for x in xrange(0, width-1):

        # get pixel at that index
        pixel = vlined[y,x]

        if pixel == 0:

            # if the pixel is black and the counter variable is equal to zero, then we must be at the start
            # of a cell. store the starting location in start_loc.
            if c == 0:
                start_loc = x

            # increment counter
            c += 1

        else:

            # if pixel is white, check value of counter variable.

            # c must be greater than or equal to the minimum dimension. otherwise, we assume that we have
            # just encountered a double-line or some other mistake and keep going.
            # this also handles possible white lines on/next to the left edge of the image.
            # this also handles white lines greater than one pixel wide.
            if c < mindim:
                continue

            # if c is less than mindim away from the right edge, skip it.
            if ((width-1)-c) < mindim:
                continue

            # the detected line has made it through all of the checks, so it must be a valid line.  add the
            # current value of c to the cols list, append the finalized tuple to col_loc, and reset counter.
            cols.append(c)
            col_loc.append( (start_loc, x-1) )
            c = 0

    # 1-pixel line is set to a vertical line 10px to the right of the left edge of image, arbitrarily.
    # it shouldn't matter, as long as the line is within the image.
    x = 10

    # counter variable for checking and measuring row heights.
    c = 0

    # row heights stored in this list.
    rows = []

    # stores actual locations of rows.  will be used in conjunction with col_loc to produce
    # cell_loc, a universal identifier of cells based on cell location.
    row_loc = []

    # convenience variable
    start_loc = 0

    # iterate through the vertical line of pixels
    for y in xrange(0, height-1):

        # get pixel at that index
        pixel = hlined[y,x]

        if pixel == 0:

            # if the pixel is black and the counter variable is equal to zero, then we must be at the start
            # of a cell. store the starting location in start_loc.
            if c == 0:
                start_loc = y

            # increment counter
            c += 1

        else:

            # if pixel is white, check value of counter variable.

            # c must be greater than or equal to the minimum dimension. otherwise, we assume that we have
            # just encountered a double-line or some other mistake and keep going.
            # this also handles possible white lines on/next to the bottom of the image.
            # this also handles white lines greater than one pixel wide.
            if c < mindim:
                continue

            # if c is less than mindim away from the top, skip it.
            if ((width-1)-c) < mindim:
                continue

            # the detected line has made it through all of the checks, so it must be a valid line.  add the
            # current value of c to the rows list, append the finalized tuple to row_loc, and reset counter.
            rows.append(c)
            row_loc.append( (start_loc, y-1) )
            c = 0

    # by now, we have two lists, cols and rows, that hold the widths and heights of the columns and rows, respectively, respectively.
    # also, we have two lists col_loc and row_loc which hold the locations of columns and rows.
    # next, we need to find which cells need merging.  the first step is to compute cell_loc based on col_loc and row_loc.

    # helper inner func to convert index to excel column.
    # rule table:
    #   0 -> A
    #   1 -> B
    # etc.
    def toExcelCol(x):
        quotient,remainder = divmod(x-1,26)
        return toExcelCol(quotient)+chr(remainder+ord('A')) if x!=0 else ''

    cell_loc = {}
    for i in xrange(0, len(col_loc)):
        for k in xrange(0, len(row_loc)):
            cell_loc[ (col_loc[i][0],row_loc[k][0],col_loc[i][1],row_loc[k][1]) ] = toExcelCol(i+1) + str(k+1)

    # cell_loc now holds the position of cells within the image.  each cell is represented by a tuple (x1,y1,x2,y2), which represent
    # two points denoting opposite corners of a rectangle.

    # calculate the midpoint of each tuple in cell_loc and store this in cell_loc_pt.
    cell_loc_pt = {}
    for key in cell_loc:
        newx = (key[0] + key[2])/2
        newy = (key[1] + key[3])/2
        cell_loc_pt[ (newx,newy) ] = cell_loc[key]

    # invert extracted table
    inv = cv2.bitwise_not(plain_table)
    plain_table = cv2.cvtColor(plain_table, cv2.COLOR_GRAY2BGR)
    _,contours, hierarchy = cv2.findContours(inv, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    datatupes = []
    for contour in contours:
        if cv2.contourArea(contour) <= mindim**2:
            continue
        peri = cv2.arcLength(contour, True)

        # Approximate the shape of the contour.
        contour_poly = cv2.approxPolyDP(contour, 0.04 * peri, True)

        # get bounding rectangle of our quadrilateral approximated polygon.
        x,y,w,h = cv2.boundingRect(contour_poly)
        plain_table = cv2.rectangle(plain_table, (x,y),(x+w,y+h),(0,255,0))

        mergecells = []
        for key in cell_loc_pt:
            if ((x < key[0]) and (key[0] < x+w) and (y < key[1]) and (key[1] < y+h)):
                mergecells.append(cell_loc_pt[key])

        tessimg = Image.fromarray(np.copy(nroi[0][y:y+h,x:x+w]))
        nh = 0
        nw = 0
        if h > w:
            nw = 101
            nh = (101/w)*h
        else:
            nh = 101
            nw = (101/h)*w

        size = nw, nh

        tessimg = tessimg.resize(size, Image.BILINEAR)
        tesstext = pytesseract.image_to_string(tessimg)
        datatupes.append( (mergecells, tesstext))

    # renaming some variables...
    values = datatupes
    rowheights = rows
    columnwidths = cols

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
        mn_c=toExcelCol(min(toIdx(re.split('(\D+)',c)[1]) for c in char))
        mx_c=toExcelCol(max(toIdx(re.split('(\D+)',c)[1]) for c in char))
        mn_r=str(min(re.split('(\D+)',c)[2] for c in char))
        mx_r=str(max(re.split('(\D+)',c)[2] for c in char))
        return mn_c+mn_r+":"+mx_c+mx_r

    # Create an new Excel file and add a worksheet.
    workbook = xlsxwriter.Workbook(sys.argv[2])
    worksheet = workbook.add_worksheet()

    default_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter'
    })


    for i in range(len(rowheights)) :
        worksheet.set_row(i, rowheights[i])

    for j in range(len(columnwidths)) :
        worksheet.set_column(str(toExcelCol(j+1))+":"+str(toExcelCol(j+1)), columnwidths[j])

    for l in range(len(values)) :
        if len(values[l][0]) == 1:
            worksheet.write(values[l][0][0],values[l][1], default_format)
        else :
            worksheet.merge_range(char2num(values[l][0]), values[l][1], default_format)

    workbook.close()

    return 0

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print "Usage: python ", sys.argv[0], " <path/to/file> <output_filename>.xlsx"
    else:
        table_to_excel()
