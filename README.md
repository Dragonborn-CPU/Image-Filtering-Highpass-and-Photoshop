# Image-Filtering-Photoshop
Three different interfaces that process images through photoshop and displays them. The best interface to use is the GreatestInterface.py. The other programs were pre-tests for GreatestInterface.py and are very messy.

When using any of the interfaces, one problem that may arise is that whatever software you are using to run the program does not acquire all the python libraries necessary. The main libraries that softwares have trouble importing is:
1. mayavi
Needed for creating 3D heatmap of 9x9 pixel neighborhood signals in GreatestInterface.py
2. PyQt3D
Needed for creating 3D heatmap of 9x9 pixel neighborhood signals in GreatestInterface.py
3. opencv-python
Needed to use OpenCV (cv2) in all 3 programs
5. pywin32
Needed to open photoshop application
7. photoshop-python-api (https://github.com/lohriialo/photoshop-scripting-python)
Needed to use photoshop in all 3 programs.

You may have to manually install the libraries in your project interpreter. 
Photoshop must be installed on your computer or disable photoshop commands in the program. For GreatestInterface.py, SpinView and Imaging Edge Remote must be installed on your computer and the filepaths must align with the code. However, you can disable Imaging Edge Remote and SpinView in SpinView() and ImagingEdge() in main(). Finally for any of the code, if Photoshop is used at all photoshop-python-api must be imported into program.
