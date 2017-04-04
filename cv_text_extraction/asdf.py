import cv2
import numpy as np
from matplotlib import pyplot as plt

file_path = ""

for i in range(2000) :

    img = cv2.imread('tesla.png',0)
    edges = cv2.Canny(img,(i/20)**2,1.5*((i/20)**2))

    file_path = 'canny_tesla_%d.png' % i

    cv2.imwrite(file_path,edges)

    print i
