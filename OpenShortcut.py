import os

# Open a shortcut with python

def SpinView():  # Opens SpinView and pauses program until app is closed
    os.startfile(r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Spinnaker SDK (64bit)\SpinView.lnk")  # folder path and .lnk shortcut

if __name__ == '__main__':
    SpinView()
