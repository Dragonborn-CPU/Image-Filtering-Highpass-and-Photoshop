import subprocess

# opens application through python

def sony_camera():  # Opens Imaging Edge Remote and pauses program until app is closed
    subprocess.Popen(r'C:\Program Files\Sony\Imaging Edge\\Remote')  # folder path with double slash before application name (in this case "Remote)
    
if __name__ == '__main__':
    sony_camera()
