from PyQt5.QtWidgets import QGraphicsView, QTextEdit, QGraphicsScene
from PyQt5.QtGui import QFont, QTransform, QWheelEvent
from PyQt5.QtCore import Qt, pyqtSignal
from constants import DEFAULT_FONT, DEFAULT_FONT_SIZE

class ZoomableTextEdit(QGraphicsView):
    zoomChanged = pyqtSignal(int)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.textEdit = QTextEdit()
        self.textEdit.setFont(QFont(DEFAULT_FONT, DEFAULT_FONT_SIZE))
        self.textEdit.setFixedSize(1240, 1754)  # A4 size at 96 DPI
        
        self.textEdit.setStyleSheet("""
            QTextEdit {
                padding: 50px;
                margin: 20px;
                border: 1px solid #ccc;
                border-radius: 5px;
            }
        """)
        
        self.scene = QGraphicsScene(self)
        self.setScene(self.scene)
        self.scene.addWidget(self.textEdit)
        
        self.zoomFactor = 100
        self.zoomStep = 5
        self.maxZoom = 1000  # 1000%
        
        self.setTransformationAnchor(QGraphicsView.AnchorUnderMouse)
        self.setResizeAnchor(QGraphicsView.AnchorUnderMouse)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.setFrameShape(QGraphicsView.NoFrame)

    def wheelEvent(self, event: QWheelEvent):
        if event.modifiers() & Qt.ControlModifier:
            if event.angleDelta().y() > 0:
                self.zoomIn()
            else:
                self.zoomOut()
            event.accept()
        else:
            super().wheelEvent(event)

    def zoomIn(self):
        self.zoom(self.zoomFactor + self.zoomStep)

    def zoomOut(self):
        self.zoom(self.zoomFactor - self.zoomStep)

    def zoom(self, factor: int):
        self.zoomFactor = max(min(factor, self.maxZoom), self.zoomStep)
        scaleFactor = self.zoomFactor / 100.0
        self.setTransform(QTransform().scale(scaleFactor, scaleFactor))
        self.zoomChanged.emit(self.zoomFactor)
    
    def setZoomFactor(self, factor):
        self.zoom(factor)

    def document(self):
        return self.textEdit.document()

    def setDocument(self, document):
        self.textEdit.setDocument(document)

    # Proxy methods for QTextEdit functionality
    def setFont(self, font):
        self.textEdit.setFont(font)

    def setFontPointSize(self, size):
        self.textEdit.setFontPointSize(size)

    def setFontWeight(self, weight):
        self.textEdit.setFontWeight(weight)

    def setFontItalic(self, italic):
        self.textEdit.setFontItalic(italic)

    def setFontUnderline(self, underline):
        self.textEdit.setFontUnderline(underline)

    def setTextColor(self, color):
        self.textEdit.setTextColor(color)

    def setAlignment(self, alignment):
        self.textEdit.setAlignment(alignment)

    def undo(self):
        self.textEdit.undo()

    def redo(self):
        self.textEdit.redo()

    def setText(self, text):
        self.textEdit.setText(text)

    def toHtml(self):
        return self.textEdit.toHtml()

    def toPlainText(self):
        return self.textEdit.toPlainText()

    def print_(self, printer):
        self.textEdit.print_(printer)