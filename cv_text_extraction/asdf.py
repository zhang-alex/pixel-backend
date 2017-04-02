# import the opencv-python module
import cv2

# equivalent to image = cv2.imread(digits.png, cv2.IMREAD_GRAYSCALE)
image = cv2.imread('digits.png',0)

# display window with title and image
cv2.imwrite('title of image',image)





if y1 > x1 :
    x2 = 100
    y2 = (100/x1)*y1

else :
    y2 = 100
    x1 = (100/y1)*x1
