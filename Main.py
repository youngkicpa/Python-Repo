import sys
from PySide6.QtWidgets import QApplication, QMainWindow, QFileDialog

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = QMainWindow()
    
    options = QFileDialog.Options()
    file_name, _ = QFileDialog.getOpenFileName(None, "Open File", "", "All Files (*)", options=options)

    print(file_name)
    window.show()
    sys.exit(app.exec())