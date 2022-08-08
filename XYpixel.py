import cv2
import numpy as np

# gets (x, y) point of where mouse clicks on image and plots it on the image

def find_coord(event, x, y, flags, param):
    if event==cv2.EVENT_FLAG_LBUTTON:
        print(x, ',', y)
        font = cv2.FONT_HERSHEY_PLAIN
        cv2.putText(img,str(x) +','+ str(y), (x,y), font, 1,(255,0,0))
        cv2.imshow("image", img)

if __name__=="__main__":
    img = cv2.imread('C:/Users/admin/Downloads/image.JPG')
    p = 0.4
    w = int(img.shape[1] * p)
    h = int(img.shape[0] * p)
    img = cv2.resize(img, (w, h))
    cv2.imshow("image", img)
    cv2.setMouseCallback("image", find_coord)
    cv2.waitKey(0)
    cv2.destroyAllWindows()
