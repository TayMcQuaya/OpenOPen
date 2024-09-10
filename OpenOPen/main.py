import ctypes
import sys
import os
from PyQt5.QtWidgets import QApplication
from PyQt5.QtGui import QIcon
from editor import Editor

def main():
    app = QApplication(sys.argv)

    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))

    icon_path = os.path.join(base_path, 'app_icon_multi.ico')
    print(f"Icon path: {icon_path}")
    print(f"Icon exists: {os.path.exists(icon_path)}")

    if os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))
        print("Icon set successfully")

        # Force set the taskbar icon using Windows API (for taskbar icon issue)
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID('mycompany.myproduct.subproduct.version')
        hwnd = ctypes.windll.user32.GetActiveWindow()
        ctypes.windll.user32.SendMessageW(hwnd, 0x0080, 0, icon_path)
    else:
        print("Warning: Icon file not found!")

    ex = Editor()
    ex.show()
    
    return app.exec_()

if __name__ == '__main__':
    sys.exit(main())
