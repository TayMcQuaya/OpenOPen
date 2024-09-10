from PyQt5.QtWidgets import QDialog, QVBoxLayout, QLineEdit, QPushButton

class FindDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle('Find')
        self.setGeometry(100, 100, 300, 100)
        layout = QVBoxLayout()
        self.input = QLineEdit(self)
        layout.addWidget(self.input)
        self.findBtn = QPushButton('Find Next', self)
        layout.addWidget(self.findBtn)
        self.setLayout(layout)