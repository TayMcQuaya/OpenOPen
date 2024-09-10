from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.colors import Color
from PyQt5.QtGui import QTextDocument, QTextCursor

def export_to_pdf(text_edit, filename):
    doc = SimpleDocTemplate(filename, pagesize=letter)
    styles = getSampleStyleSheet()
    flowables = []

    cursor = text_edit.textCursor()
    for block_num in range(text_edit.document().blockCount()):
        block = text_edit.document().findBlockByNumber(block_num)
        cursor.setPosition(block.position())
        cursor.movePosition(QTextCursor.EndOfBlock, QTextCursor.KeepAnchor)
        
        text = cursor.selectedText()
        if not text:  # Skip empty paragraphs
            continue

        style = ParagraphStyle('Custom')
        char_format = cursor.charFormat()
        
        style.fontName = char_format.font().family() or 'Times-Roman'
        style.fontSize = char_format.font().pointSize() or 12
        style.textColor = Color(char_format.foreground().color().redF(),
                                char_format.foreground().color().greenF(),
                                char_format.foreground().color().blueF())
        
        if char_format.font().bold():
            style.bold = True
        if char_format.font().italic():
            style.italic = True
        if char_format.font().underline():
            style.underline = True
        
        para = Paragraph(text, style)
        flowables.append(para)

    doc.build(flowables)