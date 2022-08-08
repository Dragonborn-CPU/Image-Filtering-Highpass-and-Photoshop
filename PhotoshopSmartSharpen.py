import glob
import os
import win32com
from win32com.client import Dispatch, GetActiveObject, GetObject
import photoshop.api as ps
from photoshop import Session

# Sends an image though Photoshop Smart Sharpen Filter

path = r'C:/Users/admin/OneDrive/Desktop/Imaging Edge Remote Photos/*'
list_of_files = glob.glob(path)  # * means all if need specific format then *.csv
sorted_files = sorted(list_of_files, key=os.path.getctime)   # get date added of files to folder in order
with Session(action="new_document") as ps:  # create new document

    app = GetActiveObject("Photoshop.Application")
    fileName = sorted_files[-1]  # get the most recent file added to the folder
    docRef = app.Open(fileName)  # open image in photoshop

    doc = ps.active_document

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

    SmartSharpen(300, 2.0, 20)  # edit smarsharpen

    options = ps.JPEGSaveOptions(quality=5)
    jpg = 'C:/Users/admin/OneDrive/Desktop/Processed Images/' \
          + os.path.basename(sorted_files[-1]).split(".")[0] + '_PS.JPG'
    doc.saveAs(jpg, options, asCopy=True)  # save image
    app.Quit()


