"""
By: Ethan S for MlI
Date: 7/20/2022
For High Pass Filter Button: Finds latest two images in folder, sends them through high pass filter in photoshop,
saves them to a new folder before displaying the images and then finding image difference.
The Imaging Edge Remote Button opens up the Imaging Edge Remote App, disable if you don't have. Code is pretty messy,
must import photoshop-python-api or go to https://github.com/lohriialo/photoshop-scripting-python. Check README file before trying to use this program.
"""

from tkinter import *
import subprocess
from PIL import Image, ImageChops
import glob
import os
import cv2
from win32com.client import Dispatch, GetActiveObject
from photoshop import Session
import tkinter.font as tkFont

root = Tk()
root.title("Image Processing")


def path1():  # open program
    subprocess.Popen(r'C:\Program Files\Sony\Imaging Edge\\Remote')


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

    cv2.waitKey(5000)
    cv2.destroyAllWindows()

    path = r'C:/Users/admin/OneDrive/Desktop/Processed Images/*'
    list_of_files = glob.glob(path)
    sorted_files = sorted(list_of_files, key=os.path.getctime)
    img1 = Image.open(sorted_files[-1])
    img2 = Image.open(sorted_files[-2])

    diff = ImageChops.difference(img1, img2)

    if diff.getbbox():
        diff.show()

    diff = diff.save('C:/Users/admin/Onedrive/Desktop/Image Difference/Image Difference.jpg')


# interface

frame = LabelFrame(root, padx=20, pady=20)
frame.pack(padx=10, pady=10)

fontStyle = tkFont.Font(family="Times New Roman", size=15)

myLabel1 = Label(frame, text="Filters", font=fontStyle)
myLabel1.grid(row=0, column=0)
myLabel3 = Label(frame, text="Other", font=fontStyle)
myLabel3.grid(row=0, column=2)

myButton1 = Button(frame, text="Open Imaging Edge Remote", bg="gray88", width=20, height=0, padx=50, pady=15,
                   command=path1)  # fg=? text, bg=? background
myButton1.grid(row=1, column=2, padx=5, pady=5)
myButton4 = Button(frame, text="High Pass Filter", bg="gray88", width=20, height=0, padx=50, pady=15,
                   command=path4)
myButton4.grid(row=1, column=0, padx=5, pady=5)


button_quit = Button(frame, text="Exit Program", bg="gray88", fg="red",
                     command=root.quit)
button_quit.grid(row=4, column=1, pady=20)

# button_setting = Button(frame, text="Settings", bg="gray88", command=path11)
# button_setting.grid(row=4, column=1, pady=20)

root.mainloop()
