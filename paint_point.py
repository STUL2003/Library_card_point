from PyQt6 import uic, QtWidgets
from PyQt6.QtWidgets import QMessageBox
import sys
import win32com.client
import os

if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS
else:
    base_path = os.path.dirname(os.path.abspath(__file__))

ui_file_path = os.path.join(base_path, "libr.ui")
Form, _ = uic.loadUiType(ui_file_path)

class Ui(QtWidgets.QDialog, Form):
    def __init__(self):
        super(Ui, self).__init__()
        self.setupUi(self)
        self.pushButton.clicked.connect(self.on_button_clicked)
    
    def on_button_clicked(self):
        lib_dict = self.gen_dict()
        self.PaintPoint(lib_dict)
        
    def gen_dict(self):
        text = self.textEdit.toPlainText()
        items = text.replace(" ", "").split(",")
        
        lib_dict = {}
        
        for item in items:
            key, value = item.split("(")
            key = key.replace("ф", "")  
            value = value.replace(")", "")  
            lib_dict[key] = int(value)
        
        return lib_dict
    
    def realPaintPoint(self, n):
        return "." * n
            
    def PaintPoint(self, lib_dict):
        f = os.path.join(base_path, "setka.doc")

        word = win32com.client.Dispatch("Word.Application")
        word.visible = True 

        doc = word.Documents.Open(f, ReadOnly=False)
        tables = doc.Tables
        
        for table in tables:
            for row in table.Rows:
                for cell in row.Cells:
                    cell_text = cell.Range.Text.replace('\r\x07', '').replace('\r\a', '').strip()
                    
                    if "н" in cell_text:
                        cell_text = cell_text.replace("н", "")
                    if cell_text in lib_dict:
                        cell.Range.Text = cell.Range.Text + self.realPaintPoint(lib_dict[cell_text])

        # doc.Save()

        os.startfile(f)

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    w = Ui()
    w.show()
    sys.exit(app.exec())