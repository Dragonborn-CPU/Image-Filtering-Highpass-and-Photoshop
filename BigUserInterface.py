"""
By: Ethan S for MlI
Date: 7/08/2022
Lots of buttons do lots of different things. Most importantly, the Photoshop high pass filter button sends last two
images in a folder through a highpass filter in photoshop. Smart sharpen is another filter in photoshop, so when
clicked images are sent through photoshop. User interface is pretty messy and so is the code. Must import
python-photoshop-api or get at https://github.com/lohriialo/photoshop-scripting-python. Check README file before trying to use this program.
"""

from tkinter import *
import subprocess
import numpy as np
from PIL import Image, ImageChops
import glob
import os
import cv2
from win32com.client import Dispatch, GetActiveObject
from photoshop import Session
import tkinter.font as tkFont
import matplotlib.pyplot as plt

root = Tk()
root.title("Image Processing")


def path1():  # open program
    subprocess.Popen(r'C:\Program Files\Sony\Imaging Edge\\Remote')


def path2():  # Smart Sharpen -> Image Sharpening with Photoshop

    with Session(action="new_document") as ps:
        path = r'C:/Users/admin/OneDrive/Desktop/Imaging Edge Remote Photos/*'
        list_of_files = glob.glob(path)  # * means all if you need specific format then *.csv
        sorted_files = sorted(list_of_files, key=os.path.getctime)
        app = Dispatch('Photoshop.Application')
        # app = GetActiveObject("Photoshop.Application")

        fileName = sorted_files[-1]
        docRef = app.Open(fileName)

        doc = ps.active_document

        # nLayerSets = docRef.LayerSets
        # nArtLayers = docRef.LayerSets.Item(len(nLayerSets)).ArtLayers

        # docRef.ActiveLayer = docRef.LayerSets.Item(len(nLayerSets)).ArtLayers.Item(len(nArtLayers))

        def SmartSharpen(inAmount, inRadius, inNoise):
            idsmartSharpenID = app.StringIDToTypeID("smartSharpen")
            desc37 = Dispatch('Photoshop.ActionDescriptor')

            idpresetKind = app.StringIDToTypeID("presetKind")
            idpresetKindType = app.StringIDToTypeID("presetKindType")
            idpresetKindCustom = app.StringIDToTypeID("presetKindCustom")
            desc37.PutEnumerated(idpresetKind, idpresetKindType, idpresetKindCustom)

            idAmnt = app.CharIDToTypeID("Amnt")
            idPrc = app.CharIDToTypeID("#Prc")
            desc37.PutUnitDouble(idAmnt, idPrc, inAmount)

            idRds = app.CharIDToTypeID("Rds ")
            idPxl = app.CharIDToTypeID("#Pxl")
            desc37.PutUnitDouble(idRds, idPxl, inRadius)

            idnoiseReduction = app.StringIDToTypeID("noiseReduction")
            idPrc = app.CharIDToTypeID("#Prc")
            desc37.PutUnitDouble(idnoiseReduction, idPrc, inNoise)

            idblur = app.CharIDToTypeID("blur")
            idblurType = app.StringIDToTypeID("blurType")
            idGsnB = app.CharIDToTypeID("GsnB")
            desc37.PutEnumerated(idblur, idblurType, idGsnB)

            # now execute the action
            app.ExecuteAction(idsmartSharpenID, desc37)

        SmartSharpen(300, 2, 20)

        options = ps.JPEGSaveOptions(quality=5)
        jpg = 'C:/Users/admin/OneDrive/Desktop/Processed Images/' \
              + os.path.basename(sorted_files[-1]).split(".")[0] + '_PS.JPG'
        doc.saveAs(jpg, options, asCopy=True)

    with Session(action="new_document") as ps:
        app = GetActiveObject("Photoshop.Application")
        fileName = sorted_files[-2]
        docRef = app.Open(fileName)

        doc = ps.active_document

        # nLayerSets = docRef.LayerSets
        # nArtLayers = docRef.LayerSets.Item(len(nLayerSets)).ArtLayers
        # docRef.ActiveLayer = docRef.LayerSets.Item(len(nLayerSets)).ArtLayers.Item(len(nArtLayers))

        def SmartSharpen(inAmount, inRadius, inNoise):
            idsmartSharpenID = app.StringIDToTypeID("smartSharpen")
            desc37 = Dispatch('Photoshop.ActionDescriptor')

            idpresetKind = app.StringIDToTypeID("presetKind")
            idpresetKindType = app.StringIDToTypeID("presetKindType")
            idpresetKindCustom = app.StringIDToTypeID("presetKindCustom")
            desc37.PutEnumerated(idpresetKind, idpresetKindType, idpresetKindCustom)

            idAmnt = app.CharIDToTypeID("Amnt")
            idPrc = app.CharIDToTypeID("#Prc")
            desc37.PutUnitDouble(idAmnt, idPrc, inAmount)

            idRds = app.CharIDToTypeID("Rds ")
            idPxl = app.CharIDToTypeID("#Pxl")
            desc37.PutUnitDouble(idRds, idPxl, inRadius)

            idnoiseReduction = app.StringIDToTypeID("noiseReduction")
            idPrc = app.CharIDToTypeID("#Prc")
            desc37.PutUnitDouble(idnoiseReduction, idPrc, inNoise)

            idblur = app.CharIDToTypeID("blur")
            idblurType = app.StringIDToTypeID("blurType")
            idGsnB = app.CharIDToTypeID("GsnB")
            desc37.PutEnumerated(idblur, idblurType, idGsnB)

            # now execute the action
            app.ExecuteAction(idsmartSharpenID, desc37)

        SmartSharpen(300, 2, 20)  # 300, 2, 20

        options = ps.JPEGSaveOptions(quality=5)
        jpg = 'C:/Users/admin/OneDrive/Desktop/Processed Images/' \
              + os.path.basename(sorted_files[-2]).split(".")[0] + '_PS.JPG'
        doc.saveAs(jpg, options, asCopy=True)
        app.Quit()


def path3():  # High Pass With Photoshop
    path = r'C:/Users/admin/OneDrive/Desktop/Imaging Edge Remote Photos/*'
    list_of_files = glob.glob(path)  # * means all if you need specific format then *.csv
    sorted_files = sorted(list_of_files, key=os.path.getctime)
    img1 = cv2.imread(sorted_files[-1])
    # img1 = cv2.resize(img1, (680, 520), interpolation=cv2.INTER_CUBIC)
    img2 = cv2.imread(sorted_files[-2])
    # img2 = cv2.resize(img2, (680, 520), interpolation=cv2.INTER_CUBIC)

    # subtract the original image with the blurred image
    # after subtracting add 127 to the total result
    hpf1 = img1 - cv2.GaussianBlur(img1, (127, 127), 3) + 127
    hpf2 = img2 - cv2.GaussianBlur(img2, (127, 127), 3) + 127

    # display both original image and filtered image
    # cv2.imshow("High Pass Filter 1", hpf1)
    # cv2.imshow("High Pass Filter 2", hpf2)

    cv2.imwrite(
        'C:/Users/admin/OneDrive/Desktop/Processed Images/' + os.path.basename(sorted_files[-1]).split(".")[0]
        + '_HP.JPG', hpf1)
    cv2.imwrite(
        'C:/Users/admin/OneDrive/Desktop/Processed Images/' + os.path.basename(sorted_files[-2]).split(".")[0]
        + '_HP.JPG', hpf2)

    # if you provide 1000 instead of 0 then
    # image will close in 1sec
    # you pass in millisecond
    cv2.waitKey(0)
    cv2.destroyAllWindows()


def path4():
    with Session(action="new_document") as ps:
        path = r'C:/Users/admin/OneDrive/Desktop/Imaging Edge Remote Photos/*'
        list_of_files = glob.glob(path)  # * means all if you need specific format then *.csv
        sorted_files = sorted(list_of_files, key=os.path.getctime)
        app = Dispatch('Photoshop.Application')
        # app = GetActiveObject("Photoshop.Application")

        fileName = sorted_files[-1]
        docRef = app.Open(fileName)

        doc = ps.active_document
        psDisplayNoDialogs = 3  # from enum PsDialogModes
        app.displayDialogs = psDisplayNoDialogs

        psPixels = 1
        strtRulerUnits = app.Preferences.RulerUnits
        if strtRulerUnits is not psPixels:
            app.Preferences.RulerUnits = psPixels

        active_layer = docRef.ActiveLayer
        active_layer.ApplyHighPass(15)
        docRef.Selection.Deselect()

        options = ps.JPEGSaveOptions(quality=5)
        jpg = 'C:/Users/admin/OneDrive/Desktop/Processed Images/' \
              + os.path.basename(sorted_files[-1]).split(".")[0] + '_HPPS.JPG'
        doc.saveAs(jpg, options, asCopy=True)

    with Session(action="new_document") as ps:

        app = GetActiveObject("Photoshop.Application")
        fileName = sorted_files[-2]
        docRef = app.Open(fileName)

        doc = ps.active_document
        psDisplayNoDialogs = 3  # from enum PsDialogModes
        app.displayDialogs = psDisplayNoDialogs

        psPixels = 1
        strtRulerUnits = app.Preferences.RulerUnits
        if strtRulerUnits is not psPixels:
            app.Preferences.RulerUnits = psPixels

        active_layer = docRef.ActiveLayer
        active_layer.ApplyHighPass(15)
        docRef.Selection.Deselect()

        options = ps.JPEGSaveOptions(quality=5)
        jpg = 'C:/Users/admin/OneDrive/Desktop/Processed Images/' \
              + os.path.basename(sorted_files[-2]).split(".")[0] + '_HPPS.JPG'
        doc.saveAs(jpg, options, asCopy=True)
        app.Quit()


def path5():  # show latest 2 images
    path = r'C:/Users/admin/OneDrive/Desktop/Processed Images/*'
    list_of_files = glob.glob(path)
    sorted_files = sorted(list_of_files, key=os.path.getctime)
    img1 = cv2.imread(sorted_files[-1])
    img1 = cv2.resize(img1, (680, 520),
                      interpolation=cv2.INTER_CUBIC)
    img2 = cv2.imread(sorted_files[-2])
    img2 = cv2.resize(img2, (680, 520),
                      interpolation=cv2.INTER_CUBIC)
    cv2.imshow(os.path.basename(sorted_files[-1]), img1)
    cv2.imshow(os.path.basename(sorted_files[-2]), img2)

    cv2.waitKey(0)
    cv2.destroyAllWindows()


def path6():  # delete latest 2 Images
    path = r'C:/Users/admin/OneDrive/Desktop/Processed Images/*'
    list_of_files = glob.glob(path)
    sorted_files = sorted(list_of_files, key=os.path.getctime)
    os.remove(sorted_files[-1])
    os.remove(sorted_files[-2])


def path7():  # plot Image Difference
    path = r'C:/Users/admin/OneDrive/Desktop/Processed Images/*'
    list_of_files = glob.glob(path)
    sorted_files = sorted(list_of_files, key=os.path.getctime)
    img1 = Image.open(sorted_files[-1])
    img2 = Image.open(sorted_files[-2])

    diff = ImageChops.difference(img1, img2)

    if diff.getbbox():
        diff.show()

    diff = diff.save('C:/Users/admin/Onedrive/Desktop/Image Difference/Image Difference.jpg')


def path8():
    # path = r'C:/Users/admin/OneDrive/Desktop/Image Difference/*'
    # list_of_files = glob.glob(path)
    # sorted_files = sorted(list_of_files, key=os.path.getctime)
    # image = cv2.imread(sorted_files[-1])
    image = cv2.imread('C:/Users/admin/Downloads/DSC00068.JPG')

    params = cv2.SimpleBlobDetector_Params()

    # Threshold
    params.minThreshold = 0
    params.maxThreshold = 200
    # Area
    params.filterByArea = True
    params.minArea = 11  # by pixels
    params.maxArea = 395
    # Color
    params.filterByColor = True
    params.blobColor = 255  # 0 = black color, 255 = light
    # Circularity
    params.filterByCircularity = True
    params.minCircularity = 0.2  # values between 0 and 1
    params.maxCircularity = 0.6
    # Inertia
    params.filterByInertia = True
    params.minInertiaRatio = .01  # values between 0 and 1
    params.maxInertiaRatio = 1
    # Convexity
    params.filterByConvexity = True
    params.minConvexity = .6  # values between 0 and 1
    params.maxConvexity = .9
    # Min Distance Between Blobs
    params.minDistBetweenBlobs = 0

    detector = cv2.SimpleBlobDetector_create(params)
    keypoints = detector.detect(image)
    print("Number of blobs detected: ", len(keypoints))

    # resize image
    scale_percent = 30  # percent of original size
    width = int(image.shape[1] * scale_percent / 100)
    height = int(image.shape[0] * scale_percent / 100)
    dim = (width, height)

    img_with_blobs = cv2.drawKeypoints(image, keypoints, np.array([]),
                                       (0, 0, 255), cv2.DRAW_MATCHES_FLAGS_DRAW_RICH_KEYPOINTS)
    resized = cv2.resize(img_with_blobs, dim, interpolation=cv2.INTER_AREA)
    plt.imshow(img_with_blobs)
    cv2.imshow("Blob Detection", resized)
    cv2.waitKey(0)
    cv2.destroyAllWindows()
    # cv2.imwrite('C:/Users/admin/OneDrive/Desktop/Image Difference/' + os.path.basename(sorted_files[-1]).split(
    # ".")[0] + '_BLOB.JPG', img_with_blobs)
    cv2.imwrite('C:/Users/admin/Downloads/Image Difference.jpg', img_with_blobs)


def path9():
    path = r'C:/Users/admin/OneDrive/Desktop/Imaging Edge Remote Photos/*'
    list_of_files = glob.glob(path)
    sorted_files = sorted(list_of_files, key=os.path.getctime)
    os.remove(sorted_files[-1])
    os.remove(sorted_files[-2])


def path10():
    path = r'C:/Users/admin/OneDrive/Desktop/Image Difference/*'
    list_of_files = glob.glob(path)
    sorted_files = sorted(list_of_files, key=os.path.getctime)
    os.remove(sorted_files[-1])
    os.remove(sorted_files[-2])


def path11():
    path = r'C:/Users/admin/OneDrive/Desktop/Imaging Edge Remote Photos/*'
    list_of_files = glob.glob(path)  # * means all if you need specific format then *.csv
    sorted_files = sorted(list_of_files, key=os.path.getctime)
    with Session(action="new_document") as ps:
        app = Dispatch('Photoshop.Application')
        # app = GetActiveObject("Photoshop.Application")

        fileName = sorted_files[-1]
        docRef = app.Open(fileName)

        doc = ps.active_document
        psDisplayNoDialogs = 3  # from enum PsDialogModes
        app.displayDialogs = psDisplayNoDialogs

        psPixels = 1
        strtRulerUnits = app.Preferences.RulerUnits
        if strtRulerUnits is not psPixels:
            app.Preferences.RulerUnits = psPixels

        active_layer = docRef.ActiveLayer
        active_layer.ApplySharpenEdges()
        docRef.Selection.Deselect()

        options = ps.JPEGSaveOptions(quality=5)
        jpg = 'C:/Users/admin/OneDrive/Desktop/Processed Images/' \
              + os.path.basename(sorted_files[-1]).split(".")[0] + '_HPPS.JPG'
        doc.saveAs(jpg, options, asCopy=True)

    with Session(action="new_document") as ps:

        app = GetActiveObject("Photoshop.Application")
        fileName = sorted_files[-2]
        docRef = app.Open(fileName)

        doc = ps.active_document
        psDisplayNoDialogs = 3  # from enum PsDialogModes
        app.displayDialogs = psDisplayNoDialogs

        psPixels = 1
        strtRulerUnits = app.Preferences.RulerUnits
        if strtRulerUnits is not psPixels:
            app.Preferences.RulerUnits = psPixels

        active_layer = docRef.ActiveLayer
        active_layer.ApplySharpenEdges()
        docRef.Selection.Deselect()

        options = ps.JPEGSaveOptions(quality=5)
        jpg = 'C:/Users/admin/OneDrive/Desktop/Processed Images/' \
              + os.path.basename(sorted_files[-2]).split(".")[0] + '_HPPS.JPG'
        doc.saveAs(jpg, options, asCopy=True)
        app.Quit()


# interface

frame = LabelFrame(root, padx=20, pady=20)
frame.pack(padx=10, pady=10)

fontStyle = tkFont.Font(family="Times New Roman", size=15)

myLabel1 = Label(frame, text="Filters", font=fontStyle)
myLabel1.grid(row=0, column=0)
myLabel2 = Label(frame, text="Image Display", font=fontStyle)
myLabel2.grid(row=0, column=1)
myLabel3 = Label(frame, text="Other", font=fontStyle)
myLabel3.grid(row=0, column=2)

myButton1 = Button(frame, text="Open Imaging Edge Remote", bg="gray88", width=20, height=0, padx=50, pady=15,
                   command=path1)  # fg=? text, bg=? background
myButton1.grid(row=1, column=2, padx=5, pady=5)
myButton2 = Button(frame, text="Smart Sharpen Filter", bg="gray88", width=20, height=0, padx=50, pady=15,
                   command=path2)
myButton2.grid(row=3, column=0, padx=5, pady=5)
myButton3 = Button(frame, text="High Pass Filter (No Photoshop)", bg="gray88", width=20, height=0, padx=50, pady=15,
                   command=path3)
myButton3.grid(row=2, column=0, padx=5, pady=5)
myButton4 = Button(frame, text="High Pass Filter", bg="gray88", width=20, height=0, padx=50, pady=15,
                   command=path4)
myButton4.grid(row=1, column=0, padx=5, pady=5)
myButton5 = Button(frame, text="Display Latest 2 Images in Processed Folder", bg="gray88", width=20, height=0, padx=50,
                   pady=15,
                   command=path5)
myButton5.grid(row=2, column=1, padx=5, pady=5)
myButton6 = Button(frame, text="Delete Latest 2 Images in Processed Folder", bg="gray88", width=20, height=0, padx=50,
                   pady=15,
                   command=path6)
myButton6.grid(row=3, column=2, padx=5, pady=5)
myButton9 = Button(frame, text="Delete Latest 2 Images in Raw Folder", bg="gray88", width=20, height=0, padx=50,
                   pady=15,
                   command=path9)
myButton9.grid(row=2, column=2, padx=5, pady=5)
myButton10 = Button(frame, text="Delete Latest 2 Images in Image Diff Folder", bg="gray88", width=20, height=0, padx=50,
                    pady=15,
                    command=path10)
myButton10.grid(row=4, column=2, padx=5, pady=5)
myButton7 = Button(frame, text="Find Image Difference", bg="gray88", width=20, height=0, padx=50, pady=15,
                   command=path7)
myButton7.grid(row=1, column=1, padx=5, pady=5)
myButton8 = Button(frame, text="Blob Detection", bg="gray88", width=20, height=0, padx=50, pady=15,
                   command=path8)
myButton8.grid(row=3, column=1, padx=5, pady=5)
myButton9 = Button(frame, text="Edge Filter", bg="gray88", width=20, height=0, padx=50, pady=15,
                   command=path11)
myButton9.grid(row=4, column=0, pady=5)

button_quit = Button(frame, text="Exit Program", bg="gray88", fg="red",
                     command=root.quit)
button_quit.grid(row=4, column=1, pady=20)

# button_setting = Button(frame, text="Settings", bg="gray88", command=path11)
# button_setting.grid(row=4, column=1, pady=20)

root.mainloop()
