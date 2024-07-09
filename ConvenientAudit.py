import sys
import os
from PySide6.QtWidgets import QApplication, QMainWindow, QPushButton, QFileDialog, QTextEdit, QVBoxLayout, QWidget, QHBoxLayout
import win32com.client as win

class FolderSelectorApp(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Folder Selector")
        self.setGeometry(100, 100, 1600, 800)

        self.text_edit = QTextEdit()
        self.text_edit.setReadOnly(True)

        self.select_button = QPushButton("Name 변경하기")
        self.select_button.clicked.connect(self.change_names)

        self.exit_button = QPushButton("Exit")
        self.exit_button.clicked.connect(self.close)

        button_layout = QVBoxLayout()
        button_layout.addWidget(self.select_button)
        button_layout.addWidget(self.exit_button)
        button_layout.addStretch()

        main_layout = QHBoxLayout()
        main_layout.addWidget(self.text_edit)
        main_layout.addLayout(button_layout)

        container = QWidget()
        container.setLayout(main_layout)
        self.setCentralWidget(container)

    def check_name(self, wb, name):
        for defined_name in wb.Names:
            if defined_name.Name == name:
                return True
        return False

    def change_names(self):
        folder_path = QFileDialog.getExistingDirectory(self, "Select Folder")
        files = []
        names = {
            "검토자":     "작성자:                 서명:                      작성일: 2024  / 07 /",
            "작성자":     "작성자:     김 영생     서명:                      작성일: 2024  / 07 /",
            "작성자기말": "작성자:     김 영생     서명:                      작성일: 2024  / 07 /",
            "기말날짜":   "2024-12-31",
            "전기말날짜": "2023-12-31",
            "당기말":     "2024-06-30",
            "전기말":     "2023-06-30",
            "회계연도":   "회계연도: 제 31 기 - 2024 년 1 월 1 일부터   2024  년 12 월 31 일까지",            
            "회사명":     "회사명: (주)아모텍"
        }
        xl = win.gencache.EnsureDispatch("Excel.Application")
        xl.Visible = False        
        
        if folder_path:
            self.text_edit.clear()
            files = self.get_files_list(folder_path)
            for filename in files:
                if filename.split('.')[1] == "xlsx":
                    wb = xl.Workbooks.Open(os.path.join(folder_path, filename))
                    for key, value in names.items():
                        if self.check_name(wb, key):
                            wb.Names.Item(key).RefersTo = value
                        else:
                            wb.Names.Add(Name=key, RefersTo=value)
                else:
                    continue

        wb.Save()
        wb.Close()
        xl.Quit()
        self.text_edit.append("Name 변경하기가 종료되었습니다.")

    def get_files_list(self, folder_path):
        self.text_edit.append(f"Folder: {os.path.basename(folder_path)}")
        
        items = os.listdir(folder_path)
        folders = [item for item in items if os.path.isdir(os.path.join(folder_path, item))]
        files = [item for item in items if os.path.isfile(os.path.join(folder_path, item))]
        
        self.text_edit.append("Folders:")
        for folder in folders:
            self.text_edit.append(f"    {folder}")
        self.text_edit.append("Files:")
        for file in files:
            self.text_edit.append(f"    {file}")

        return files

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = FolderSelectorApp()
    window.show()
    sys.exit(app.exec())

