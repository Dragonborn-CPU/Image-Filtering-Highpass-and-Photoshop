import cv2
from win32com.client import Dispatch, GetActiveObject, GetObject
from photoshop import Session
import subprocess
import os
import psutil
import time
import matplotlib.pyplot as plt
import numpy as np
from math import sqrt
from mayavi import mlab

''' By: Ethan S for MLI
Program compatible with SpinView and Imaging Edge Remote. Sends two images through high pass filter, finds image 
difference, then finds all contours and their areas. Can also enable finding (X, Y) cords of all contours in 
find_contour(). Furthermore, after getting image diff with contours, can left click on image to get 2D graph of all 
pixel signals that have same x and same y as (x, y) cord clicked. Right-clicking on image will show a 3D 9x9 
neighborhood heatmap of pixel signals of (x, y) cord clicked.
Before Starting:
1. Must have Imaging Edge Remote/SpinView OR disable sony_camera() and SpinView() in main()
2. If using Imaging Edge Remote/SpinView, the two pictures taken MUST have the same filepath as filename and filename2
3. Photoshop must be CLOSED before starting program or replace Dispatch function with GetActiveObject in Photoshop() 
which requires photoshop to be opened before starting program.
Steps when running program (If not using SpinView/ImagingEdge then just click run - you don't have to do anything else):
1. Click Run
2. Imaging Edge Remote or SpinView will open and click camera that you want to take pictures with
3. Take  two pictures in imaging edge remote
4. Close app
5. After app is closed, program will continue to run
'''

filename = 'C:/Users/admin/Downloads/DSC00048.JPG'  # first image path, same path as picture taken from Imaging/SpinView
filename2 = 'C:/Users/admin/Downloads/DSC00049.JPG'  # second image path, same path as picture taken from Imaging/Spin
hp = 15  # set highpass filter scale for Photoshop() here (range is 0-2000)
resize = .3  # set resize value of image (in case image is too big). Range is 0 < x <= 1
font = cv2.FONT_HERSHEY_COMPLEX
name, name2 = filename.split(".")[0] + '_PPP.JPG', filename2.split(".")[0] + '_PPP.JPG'


def checkIfProcessRunning(processName):   # checks if process is running
    print('Checking if application is running...')
    for proc in psutil.process_iter():
        try:
            if processName.lower() in proc.name().lower():
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass
    return False


def sony_camera():
    subprocess.Popen(r'C:\Program Files\Sony\Imaging Edge\\Remote')
    time.sleep(10)
    ImagingEdgeRunning = True
    while ImagingEdgeRunning:
        if checkIfProcessRunning('Remote'):
            print('Imaging Edge Remote is open. Try again in 5 seconds')
            time.sleep(5)  # Wait 5 seconds and try again
            print('Checking again...')
        else:
            print('Imaging Edge Remote is closed. Acquiring images, opening Photoshop')
            ImagingEdgeRunning = False  # Sets SpinViewRunning False to exit loop


def SpinView():
    os.startfile(r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Spinnaker SDK (64bit)\SpinView.lnk")
    time.sleep(10)
    SpinViewRunning = True
    while SpinViewRunning:
        if checkIfProcessRunning('SpinView'):
            print('SpinView is open. Try again in 5 seconds')
            time.sleep(5)  # Wait 5 seconds and try again
            print('Checking again...')
        else:
            print('SpinView is closed. Acquiring images, opening Photoshop')
            SpinViewRunning = False  # Sets SpinViewRunning False to exit loop


def Photoshop():
    with Session(action=None) as ps:   # first image
        app = Dispatch('Photoshop.Application')
        # app = GetActiveObject("Photoshop.Application")
        docRef = app.Open(filename)
        docRef.ActiveLayer.ApplyHighPass(hp)   # Apply High Pass Filter
        docRef.Selection.Deselect()

        options = ps.JPEGSaveOptions(quality=5)   # save image
        ps.active_document.saveAs(name, options, asCopy=True)
    with Session(action=None) as ps:    # second image
        app = Dispatch('Photoshop.Application')
        # app = GetActiveObject("Photoshop.Application")
        docRef = app.Open(filename2)
        docRef.ActiveLayer.ApplyHighPass(hp)   # apply High Pass Filter
        docRef.Selection.Deselect()

        options = ps.JPEGSaveOptions(quality=5)   # save image
        ps.active_document.saveAs(name2, options, asCopy=True)
        app.Quit()


def set_up():
    img, img2 = cv2.imread(name), cv2.imread(name2)
    p = resize
    w, h = int(img.shape[1] * p), int(img.shape[0] * p)
    img, img2 = cv2.resize(img, (w, h)), cv2.resize(img2, (w, h))
    img, img2 = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY), cv2.cvtColor(img2, cv2.COLOR_BGR2GRAY)
    imgdiff = cv2.subtract(img, img2)
    cv2.imwrite(filename.split(".")[0] + '_Diff.JPG', imgdiff)


def find_contour():
    contours = cv2.findContours(edged, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    contours = contours[0] if len(contours) == 2 else contours[1]
    index = 1
    cluster_count = 0
    for cntr in contours:
        areas = str(cv2.contourArea(cntr))
        if 10000 > cv2.contourArea(cntr) > -1:  # change area value for threshold
            # draw contours
            cv2.drawContours(diff, [cntr], -1, (0, 0, 255), 3)
            cluster_count = cluster_count + 1

            # all of this is for getting area LISTED above the contour
            n = cntr.ravel()
            i = 0
            for j in n:
                if i % 2 == 0:
                    x = n[i]
                    y = n[i + 1]
                    if i == 0:
                        cv2.putText(diff, areas, (x, y),
                                    font, 1, (0, 255, 0))
                    else:  # if you want to put coordinates on other vertices of the contours
                        pass
                        # text on remaining co-ordinates.
                        # cv2.putText(img, string, (x, y),
                        # font, 0.5, (0, 255, 0))
            # i = i + 1
    index = index + 1
    print('number of particles detected:', cluster_count)


def find_coord(event, x, y, flags, param):  # mouse functions
    if event == cv2.EVENT_FLAG_LBUTTON:  # 2D Signal Graphs
        print('(', x, ',', y, ')')  # print pixel coordinate
        Pixel_XValue = []
        X = np.arange(0, 1272)  # X Dimensions of Image
        r, g, b = diff[y, X, 2], diff[y, X, 1], diff[y, X, 0]
        numx_count = 0
        for num in X:
            mean = sqrt(0.241 * (r[numx_count] ** 2) + 0.691 * (g[numx_count] ** 2) + 0.068 * (b[numx_count] ** 2))
            Pixel_XValue.append(mean)
            numx_count += 1
        plt.scatter(X, Pixel_XValue)
        plt.gca().update(dict(title='X-Cord Signal', xlabel='X', ylabel='Signal', xlim=None, ylim=None))  # set labels
        plt.show()  # for 2D X Signal graph

        Pixel_YValue = []
        Y = np.arange(0, 849)  # Y Dimensions of Image
        r, g, b = diff[Y, x, 2], diff[Y, x, 1], diff[Y, x, 0]
        numy_count = 0
        for num in Y:
            mean = sqrt(0.241 * (r[numy_count] ** 2) + 0.691 * (g[numy_count] ** 2) + 0.068 * (b[numy_count] ** 2))
            Pixel_YValue.append(mean)
            numy_count += 1
        plt.scatter(Y, Pixel_YValue)
        plt.gca().update(dict(title='Y-Cord Signal', xlabel='Y', ylabel='Signal', xlim=None, ylim=None))  # set labels
        plt.show()  # for 2D Y Signal graph

    elif event == cv2.EVENT_FLAG_RBUTTON:  # find pixel intensity in 9x9 kernel w/ 3D graph
        print('(', x, ',', y, ')')  # print pixel coordinate
        Y = ([y + 4], [y + 4], [y + 4], [y + 4], [y + 4], [y + 4], [y + 4], [y + 4], [y + 4],  # kernels
             [y + 3], [y + 3], [y + 3], [y + 3], [y + 3], [y + 3], [y + 3], [y + 3], [y + 3],
             [y + 2], [y + 2], [y + 2], [y + 2], [y + 2], [y + 2], [y + 2], [y + 2], [y + 2],
             [y + 1], [y + 1], [y + 1], [y + 1], [y + 1], [y + 1], [y + 1], [y + 1], [y + 1],
             [y], [y], [y], [y], [y], [y], [y], [y], [y],
             [y - 1], [y - 1], [y - 1], [y - 1], [y - 1], [y - 1], [y - 1], [y - 1], [y - 1],
             [y - 2], [y - 2], [y - 2], [y - 2], [y - 2], [y - 2], [y - 2], [y - 2], [y - 2],
             [y - 3], [y - 3], [y - 3], [y - 3], [y - 3], [y - 3], [y - 3], [y - 3], [y - 3],
             [y - 4], [y - 4], [y - 4], [y - 4], [y - 4], [y - 4], [y - 4], [y - 4], [y - 4])
        X = ([x - 4], [x - 3], [x - 2], [x - 1], [x], [x + 1], [x + 2], [x + 3], [x + 4],
             [x - 4], [x - 3], [x - 2], [x - 1], [x], [x + 1], [x + 2], [x + 3], [x + 4],
             [x - 4], [x - 3], [x - 2], [x - 1], [x], [x + 1], [x + 2], [x + 3], [x + 4],
             [x - 4], [x - 3], [x - 2], [x - 1], [x], [x + 1], [x + 2], [x + 3], [x + 4],
             [x - 4], [x - 3], [x - 2], [x - 1], [x], [x + 1], [x + 2], [x + 3], [x + 4],
             [x - 4], [x - 3], [x - 2], [x - 1], [x], [x + 1], [x + 2], [x + 3], [x + 4],
             [x - 4], [x - 3], [x - 2], [x - 1], [x], [x + 1], [x + 2], [x + 3], [x + 4],
             [x - 4], [x - 3], [x - 2], [x - 1], [x], [x + 1], [x + 2], [x + 3], [x + 4],
             [x - 4], [x - 3], [x - 2], [x - 1], [x], [x + 1], [x + 2], [x + 3], [x + 4])
        r, g, b = diff[Y, X, 2], diff[Y, X, 1], diff[Y, X, 0]  # get RGB values
        X, Y = np.array(X).flatten(), np.array(Y).flatten()  # convert set of lists into array
        num_count = 0
        Pixel_Intensity = []
        for num in Y:
            if num_count == 80:
                # print('(' + str(r[num_count])[1:-1] + ',' + str(g[num_count])[1:-1] + ',' +
                # str(b[num_count])[1:-1] + ')')
                r_int, g_int, b_int = int((str(r[num_count])[1:-1])), int((str(g[num_count])[1:-1])), \
                                      int((str(b[num_count])[1:-1]))
                mean = sqrt(0.241 * (r_int ** 2) + 0.691 * (g_int ** 2) + 0.068 * (b_int ** 2))
                Pixel_Intensity.append(mean)
                # print('Pixel Intensity: ' + str(mean))
                # print('')
                Z = np.array(Pixel_Intensity)
                pts = mlab.points3d(X, Y, Z, Z)
                mesh = mlab.pipeline.delaunay2d(pts)
                pts.remove()
                surf = mlab.pipeline.surface(mesh)
                mlab.xlabel("X")
                mlab.ylabel("Y")
                mlab.zlabel("Signal")
                mlab.show()  # 3D heatmap for 9x9 kernel
                return
            if num_count < 80:
                # print('(' + str(r[num_count])[1:-1] + ',' + str(g[num_count])[1:-1] + ',' + str(b[num_count])[1:-1] + ')')
                r_int, g_int, b_int = int((str(r[num_count])[1:-1])), int((str(g[num_count])[1:-1])), \
                                      int((str(b[num_count])[1:-1]))
                mean = sqrt(0.241 * (r_int ** 2) + 0.691 * (g_int ** 2) + 0.068 * (b_int ** 2))
                Pixel_Intensity.append(mean)
                # print('Pixel Intensity: ' + str(mean))
                num_count += 1


if __name__ == '__main__':
    sony_camera()
    # SpinView()
    time.sleep(5)
    Photoshop()
    set_up()
    diff = cv2.imread(filename.split(".")[0] + '_Diff.JPG')
    gray = cv2.cvtColor(diff, cv2.COLOR_BGR2GRAY)
    edged = cv2.Canny(diff, 30, 200)
    find_contour()
    cv2.imshow("image", diff)
    cv2.setMouseCallback("image", find_coord)
    cv2.waitKey(0)
    cv2.destroyAllWindows()
