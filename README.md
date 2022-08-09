# Image-Filtering-Photoshop
Three different interfaces that process images through photoshop and displays them. ONLY use the interface labeled GreatestInterface.py. The other programs were pre-tests for GreatestInterface.py and are very messy with many problems and lacking most features.

When using any of the interfaces, one problem that may arise is that whatever software you are using to run the program does not acquire all the python libraries necessary. The main libraries that softwares have trouble importing is:
1. mayavi : 
needed for creating 3D heatmap of 9x9 pixel neighborhood signals in GreatestInterface.py
2. PyQt3D : 
needed for creating 3D heatmap of 9x9 pixel neighborhood signals in GreatestInterface.py
3. opencv-python : 
needed to use OpenCV (cv2) in all programs
4. pywin32 : 
needed to open photoshop application
5. photoshop-python-api (https://github.com/lohriialo/photoshop-scripting-python) : 
needed to use filters (highpass, smartsharpen) in photoshop
6. psutil : needed for python program to determine whether Imaging Edge or SpinView is running
7. Rawpy : needed to convert RAW files (.arw) into readable files for OpenCV
8. ImageIO : needed to convert RAW files (.arw) into readable files for OpenCV

You may have to manually install the libraries in your project interpreter. 
Photoshop must be installed on your computer or disable photoshop commands in the program. For GreatestInterface.py, SpinView and Imaging Edge Remote are disabled but can be enabled for SpinView in SpinView() and ImagingEdge in ImagingEdge() in main(). Finally for any of the code, if Photoshop is used at all photoshop-python-api must be imported into program.

ISSUES:
1. Program can't load : 
It is common for the first time trying to run the program, the program will take quite a long time (about 20-30 seconds) to open Imaging Edge or SpinView. Thus, wait
for about half a minute when running the program for the first time. After running the program once, it will take about 3-5 seconds for the program to open Imaging Edge or SpinView for all other times running the program.
2. Sony Camera can't connect to Imaging Edge Remote : 
Sony Camera may have difficulties connecting to Imaging Edge Remote. This is not the fault of the python program but the fault of the computer, Imaging Edge, and camera. The issue cannot be fixed. If you are experiencing the problem, wait a few seconds before trying to reconnect the camera to Imaging Edge. If the problem continues, restart the program.
3. Images undergoing highpass filter are completely black : 
First, try changing the highpass filter settings in the highpass filter function in the code - high_pass-CV(). If the problem continues, then change canny function in
main(): edged = cv2.Canny(diff, 30, 200). Either change Canny values or switch to a different edge detector. The point of this line is to exemplify the differences and lines so that the program can detect the contours/blobs/particles.
