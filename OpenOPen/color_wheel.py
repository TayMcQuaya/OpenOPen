from PyQt5.QtWidgets import QColorDialog

class ColorWheel(QColorDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setOption(QColorDialog.ColorDialogOption.DontUseNativeDialog)
        self.setOption(QColorDialog.ColorDialogOption.ShowAlphaChannel, False)
