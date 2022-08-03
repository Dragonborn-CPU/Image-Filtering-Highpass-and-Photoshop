import cv2
from win32com.client import Dispatch, GetActiveObject, GetObject
from photoshop import Session
import subprocess
import os
import psutil
import time
import numpy as np
from math import sqrt
from mayavi import mlab
import matplotlib.pyplot as plt
import rawpy
import imageio
import glob

''' By: Ethan S for MLI
Program compatible with SpinView and Imaging Edge Remote. Sends two images through high pass filter, finds image 
difference, then finds all contours and their areas. Prints out location, area, perimeter, aspect ratio, and extent of 
each contour. Can also enable finding (X, Y) cords of all contours in find_contour(). Furthermore, after getting image 
diff with contours, can left click on image to get 2D graph of all pixel signals that have same x and same y as (x, y) 
cord clicked. Right-clicking on image will show a 3D 9x9 neighborhood heatmap of pixel signals of (x, y) cord clicked.
Before Starting:
1. Must have Imaging Edge Remote/SpinView OR disable sony_camera() and SpinView() in main()
2. If using Imaging Edge Remote/SpinView, the two pictures taken MUST have the same filepath as folder path
3. Photoshop must be CLOSED before starting program or replace Dispatch function with GetActiveObject in Photoshop() 
which requires photoshop to be opened before starting program.
Steps when running program (If not using SpinView/ImagingEdge then just click run - you don't have to do anything else):
1. Click Run
2. Imaging Edge Remote or SpinView will open and click camera that you want to take pictures with
3. Take  two pictures in imaging edge remote
4. Close app
5. After app is closed, program will continue to run
'''

filename_save = 'path'
'''Determines which save path to follow:
path : If set to path, must set paths for images
no path: don't set path for images being used, but must set path for folder where raw images are stored. Gets latest 2
images added to folder where raw images are stored.
Must disable which path you are not using right below.
'''
# For filename_save = 'path', disable section if using 'no path'
path = 'C:/Users/admin/Downloads/DSC00002.ARW'  # first image path
path2 = 'C:/Users/admin/Downloads/DSC00003.ARW'  # second image path

# For filename_save = 'no path', disable section if using 'path'
# path = r'C:/Users/admin/OneDrive/Desktop/Raw Images/*'  # folder where raw images are stored in
# path2 = r'C:/Users/admin/Downloads/'   # folder where edited images are stored


# Other settings
hp = 15  # set highpass filter scale for Photoshop() here (range is 0-2000)
resize = .3  # set resize value of image (in case image is too big). Range is 0 < x <= 1
highpass = 'True'
''' Set highpass to either True or False:
True : images are sent through highpass in OpenCV
False : images are not sent through highpass in OpenCV
'''
high_pass_filter = 'sobel'
'''high_pass_filter options (set high_pass_filter to one of the following):
bilateral, box, canny, filter2D, sepfilter2D, gaussian, laplacian, log, median, prewitt, scharr, sobel
NOTE: All highpass filters degree/effect can be changed in high_pass_CV() 
'''


def convertRAW():  # converts images from RAW to jpg. Furthermore, specifies filename_save paths.
    global filename
    global filename2
    if filename_save == 'path':
        rgb = rawpy.imread(path).postprocess()
        imageio.imsave(path.split(".")[0] + ".jpg", rgb)
        rgb = rawpy.imread(path2).postprocess()
        imageio.imsave(path2.split(".")[0] + ".jpg", rgb)
        filename, filename2 = path.split(".")[0] + ".jpg", path2.split(".")[0] + ".jpg"
    if filename_save == 'no path':
        list_of_files = glob.glob(path)  # * means all if you need specific format then *.csv
        sorted_files = sorted(list_of_files, key=os.path.getctime)
        rgb = rawpy.imread(sorted_files[-1]).postprocess()
        imageio.imsave(path2 + os.path.basename(sorted_files[-1]).split(".")[0] + '.jpg', rgb)
        rgb = rawpy.imread(sorted_files[-2]).postprocess()
        imageio.imsave(path2 + os.path.basename(sorted_files[-2]).split(".")[0] + '.jpg', rgb)
        filename = path2 + os.path.basename(sorted_files[-1]).split(".")[0] + ".jpg"
        filename2 = path2 + os.path.basename(sorted_files[-1]).split(".")[0] + ".jpg"


def checkIfProcessRunning(processName):  # checks if process is running
    print('Checking if application is running...')
    for proc in psutil.process_iter():
        try:
            if processName.lower() in proc.name().lower():
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass
    return False


def sony_camera():  # Opens Imaging Edge Remote and pauses program until app is closed
    subprocess.Popen(r'C:\Program Files\Sony\Imaging Edge\\Remote')
    time.sleep(10)
    ImagingEdgeRunning = True
    while ImagingEdgeRunning:
        if checkIfProcessRunning('Remote'):
            print('Imaging Edge Remote is open. Try again in 5 seconds')
            time.sleep(5)  # Wait 5 seconds and try again
            print('Checking again...')
        else:
            print('Imaging Edge Remote is closed. Acquiring images, deploying high pass filter...')
            ImagingEdgeRunning = False  # Sets ImagingEdgeRunning False to exit loop


def SpinView():  # Opens SpinView and pauses program until app is closed
    os.startfile(r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Spinnaker SDK (64bit)\SpinView.lnk")
    time.sleep(10)
    SpinViewRunning = True
    while SpinViewRunning:
        if checkIfProcessRunning('SpinView'):
            print('SpinView is open. Try again in 5 seconds')
            time.sleep(5)  # Wait 5 seconds and try again
            print('Checking again...')
        else:
            print('SpinView is closed. Acquiring images, deploying high pass filter...')
            SpinViewRunning = False  # Sets SpinViewRunning False to exit loop


def Photoshop():  # sends images through Photoshop highpass filter
    with Session(action=None) as ps:  # first image
        app = Dispatch('Photoshop.Application')
        # app = GetActiveObject("Photoshop.Application")
        docRef = app.Open(filename)
        docRef.ActiveLayer.ApplyHighPass(hp)  # apply High Pass Filter
        docRef.Selection.Deselect()

        options = ps.JPEGSaveOptions(quality=5)  # save image
        ps.active_document.saveAs(name, options, asCopy=True)
    with Session(action=None) as ps:  # second image
        app = Dispatch('Photoshop.Application')
        # app = GetActiveObject("Photoshop.Application")
        docRef = app.Open(filename2)
        docRef.ActiveLayer.ApplyHighPass(hp)  # apply High Pass Filter
        docRef.Selection.Deselect()

        options = ps.JPEGSaveOptions(quality=5)  # save image
        ps.active_document.saveAs(name2, options, asCopy=True)
        app.Quit()


def save(src, src2):  # saves images from high_pass_CV
    if high_pass_filter == ("prewitt" or "sobel" or "scharr"):
        cv2.imwrite(name, src)
        cv2.imwrite(name2, src2)
    else:
        cv2.imwrite(name, src)
        cv2.imwrite(name2, src2)


def high_pass_CV():  # high pass filters in OpenCV, ddepth should always be -1
    if highpass == 'True':
        read = cv2.cvtColor(cv2.imread(filename), cv2.COLOR_BGR2GRAY)
        read2 = cv2.cvtColor(cv2.imread(filename2), cv2.COLOR_BGR2GRAY)
        if high_pass_filter == 'bilateral':
            bib = read - cv2.bilateralFilter(read, 15, 80, 80) + 8   # img, ksize, sigmaColor, sigmaSpace
            bib2 = read2 - cv2.bilateralFilter(read2, 15, 80, 80) + 8
            save(bib, bib2)
        if high_pass_filter == 'box':
            box = read - cv2.boxFilter(read, -1, (8, 8)) + 8  # img, ddepth, ksize (as tuple)
            box2 = read2 - cv2.boxFilter(read2, -1, (8, 8)) + 8
            save(box, box2)
        if high_pass_filter == 'canny':
            can = cv2.Canny(read, 80, 120)  # img, threshold1, threshold2
            can2 = cv2.Canny(read2, 80, 120)
            save(can, can2)
        if high_pass_filter == 'filter2D':
            kernel = np.array([[0.0, -1.0, 0.0],  # create custom mask here
                               [-1.0, 4.0, -1.0],
                               [0.0, -1.0, 0.0]])
            kernel = kernel / (np.sum(kernel) if np.sum(kernel) != 0 else 1)
            fil = cv2.filter2D(read, -1, kernel)  # img, ddepth, mask/filter
            fil2 = cv2.filter2D(read2, -1, kernel)
            save(fil, fil2)
        if high_pass_filter == 'sepfilter2D':
            m = np.array([-1, 2, -1])   # custom filter using sepFilter2D, change here
            g = np.array([-1, 2, -1])
            sep = cv2.sepFilter2D(read, -1, m, g)   # img, ddepth, mask/filter
            sep2 = cv2.sepFilter2D(read2, -1, m, g)
            save(sep, sep2)
        if high_pass_filter == 'gaussian':
            gau = read - cv2.GaussianBlur(read, (13, 13), 3) + 8  # img, height, width, standard dev
            gau2 = read2 - cv2.GaussianBlur(read2, (13, 13), 3) + 8
            save(gau, gau2)
        if high_pass_filter == 'laplacian':
            lap = cv2.Laplacian(read, -1)  # img, ddepth
            lap2 = cv2.Laplacian(read2, -1)
            save(lap, lap2)
        if high_pass_filter == 'log':
            log = cv2.Laplacian(cv2.GaussianBlur(read, (5, 5), 0), -1)   # Gaussian Blur, ddepth
            log2 = cv2.Laplacian(cv2.GaussianBlur(read2, (5, 5), 0), -1)
            save(log, log2)
        if high_pass_filter == 'median':
            med = read - cv2.medianBlur(read, 15) + 8   # img, ksize
            med2 = read2 - cv2.medianBlur(read2, 15) + 8
            save(med, med2)
        if high_pass_filter == 'prewitt':
            kernelX = np.array([[1, 1, 1], [0, 0, 0], [-1, -1, -1]])   # prewitt filter using filter2D
            kernelY = np.array([[-1, 0, 1], [-1, 0, 1], [-1, 0, 1]])
            preX, preY = cv2.filter2D(read, -1, kernelX), cv2.filter2D(read, -1, kernelY)
            preX2, preY2 = cv2.filter2D(read2, -1, kernelX), cv2.filter2D(read2, -1, kernelY)
            save(preX + preY, preX2 + preY2)
        if high_pass_filter == 'scharr':
            schx = cv2.Scharr(read, -1, 1, 0)  # img, ddepth, dx, dy
            schy = cv2.Scharr(read, -1, 0, 1)
            schx2 = cv2.Scharr(read2, -1, 1, 0)
            schy2 = cv2.Scharr(read2, -1, 0, 1)
            save(schx + schy, schx2 + schy2)
        if high_pass_filter == 'sobel':
            sobx = cv2.Sobel(read, -1, 1, 0)  # img, ddepth, dx, dy
            soby = cv2.Sobel(read, -1, 0, 1)
            sobx2 = cv2.Sobel(read2, -1, 1, 0)
            soby2 = cv2.Sobel(read2, -1, 0, 1)
            save(sobx + soby, sobx2 + soby2)
    if highpass == 'False':
        pass


def set_up():  # set-up images: import into OpenCV, resize, change to grayscale, and get image difference
    if highpass == 'True':
        img, img2 = cv2.imread(name), cv2.imread(name2)
        w, h = int(img.shape[1] * resize), int(img.shape[0] * resize)
        img, img2 = cv2.resize(img, (w, h)), cv2.resize(img2, (w, h))
        img, img2 = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY), cv2.cvtColor(img2, cv2.COLOR_BGR2GRAY)
        cv2.imshow('Image 1', img)
        cv2.imshow('Image 2', img2)
        imgdiff = cv2.subtract(img, img2)
        cv2.imwrite(filename.split(".")[0] + '_Diff.JPG', imgdiff)
    if highpass == 'False':
        img, img2 = cv2.imread(filename), cv2.imread(filename2)
        w, h = int(img.shape[1] * resize), int(img.shape[0] * resize)
        img, img2 = cv2.resize(img, (w, h)), cv2.resize(img2, (w, h))
        img, img2 = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY), cv2.cvtColor(img2, cv2.COLOR_BGR2GRAY)
        cv2.imshow('Image 1', img)
        cv2.imshow('Image 2', img2)
        imgdiff = cv2.subtract(img, img2)
        cv2.imwrite(filename.split(".")[0] + '_Diff.JPG', imgdiff)


def find_contour():  # find contours in image difference
    contours = cv2.findContours(edged, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    contours = contours[0] if len(contours) == 2 else contours[1]
    index = 1
    cluster_count = 0
    for cntr in contours:
        area = cv2.contourArea(cntr)  # area of contour
        # round function is used to round following values to 5 decimal places
        perimeter = round(cv2.arcLength(cntr, True), 5)  # perimeter of contour
        AspectRatio = round((cv2.boundingRect(cntr)[2]) / (cv2.boundingRect(cntr)[3]), 5)
        # Aspect Ratio is width to height ratio of bounding box. If contour is a perfect circle, aspect ratio = 1.
        extent = round((cv2.contourArea(cntr)) / (cv2.boundingRect(cntr)[2] * cv2.boundingRect(cntr)[3]), 5)
        # extent is contour area to bounding box area. If contour is a perfect circle, then extent = .7854
        if area < 5000 and perimeter < 3000 and .3 < AspectRatio < 3 and .2 < extent:  # set limits for each variable
            # draw contours
            cv2.drawContours(diff, [cntr], -1, (0, 0, 255), 3)
            cluster_count = cluster_count + 1

            # all of this is for getting area LISTED above the contour
            n = cntr.ravel()
            i = 0
            for j in n:
                cords = []
                if i % 2 == 0:
                    x, y = n[i], n[i + 1]
                    if i == 0:
                        cv2.putText(diff, str(area), (x, y), font, 1, (0, 255, 0))
                        cords.append((x, y))
            print("Coordinates:", cords[0], " Area:", str(area), " Perimeter:", str(perimeter),
                  " Aspect Ratio:", str(AspectRatio), " Extent:", str(extent))
    index = index + 1
    print('Number of particles detected:', cluster_count)


def find_coord(event, x, y, flags, param):  # mouse functions: left-click to get 2D graph, right-click to get 3D graph
    if event == cv2.EVENT_FLAG_LBUTTON:  # 2D Signal Graphs
        print('(', x, ',', y, ')')  # print pixel coordinate
        Pixel_XValue = []
        X = np.arange(0, diff.shape[1])  # X Dimensions of Image
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
        Y = np.arange(0, diff.shape[0])  # Y Dimensions of Image
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
                # print('('+str(r[num_count])[1:-1]+','+str(g[num_count])[1:-1]+','+str(b[num_count])[1:-1]+')')
                r_int, g_int, b_int = int((str(r[num_count])[1:-1])), int((str(g[num_count])[1:-1])), \
                                      int((str(b[num_count])[1:-1]))
                mean = sqrt(0.241 * (r_int ** 2) + 0.691 * (g_int ** 2) + 0.068 * (b_int ** 2))
                Pixel_Intensity.append(mean)
                # print('Pixel Intensity: ' + str(mean))
                num_count += 1


if __name__ == '__main__':
    # sony_camera()
    # SpinView()
    convertRAW()
    global font
    global name
    global name2
    font = cv2.FONT_HERSHEY_COMPLEX
    name, name2 = filename.split(".")[0] + '_PPP.jpg', filename2.split(".")[0] + '_PPP.jpg'
    # Photoshop()
    high_pass_CV()
    set_up()
    diff = cv2.imread(filename.split(".")[0] + '_Diff.JPG')
    edged = cv2.Canny(diff, 30, 200)
    find_contour()
    cv2.imshow("image difference", diff)
    cv2.setMouseCallback("image difference", find_coord)
    cv2.waitKey(0)
    cv2.destroyAllWindows()
