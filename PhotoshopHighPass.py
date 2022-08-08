from win32com.client import Dispatch, GetActiveObject, GetObject
from photoshop import Session

# must import photoshop.api

filename = 'C:/Users/admin/Downloads/DSC00068.jpg'  # image file path
name = 'C:/Users/admin/OneDrive/Desktop/PhotoshopImages' # where to save image after photoshop

def Photoshop():
    with Session(action=None) as ps:  # second image
        app = Dispatch('Photoshop.Application')  # open Photoshop if Photoshop is Not active
        # app = GetActiveObject("Photoshop.Application")  # open Photoshop if Photoshop is already active
        docRef = app.Open(filename)  # open image
        docRef.ActiveLayer.ApplyHighPass(15)  # apply High Pass Filter and amount applied
        docRef.Selection.Deselect()

        ps.active_document.saveAs(name, ps.JPEGSaveOptions(quality=5), asCopy=True)  # save image
        app.Quit()
        
if __name__ == '__main__':
    Photoshop()
