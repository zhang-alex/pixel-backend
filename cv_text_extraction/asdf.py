# import the opencv-python module
import cv2 as ocv

# equivalent to image = cv2.IMREAD_GRAYSCALE(digits.png)
image = ocv.imread('digits.png',0)

# display window with title and image
ocv.imshow('title of image',image)
