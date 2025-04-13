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
            "검토자":     "검토자:                 서명:                      작성일: 2025  / 03 /",
            "작성자":     "작성자:     김 영생     서명:                      작성일: 2025  / 03 /",
            "작성자기말": "작성자:     김 영생     서명:                      작성일: 2025  / 03 /",
            "기말날짜":   "2024-12-31",
            "전기말날짜": "2023-12-31",
            "당기말":     "2024-12-31",
            "전기말":     "2023-12-31",
            "회계연도":   "회계연도: 제 38 기 - 2024 년 1 월 1 일부터   2024  년 12 월 31 일까지",            
            "회사명":     "회사명: 알파(주)"
        }
        
        xl = win.gencache.EnsureDispatch("Excel.Application")  
        xl.Visible = False  

        if folder_path:
            self.text_edit.clear()
            files = self.get_files_list(folder_path)
            
            for filename in files:
                if filename.split('.')[-1] == "xlsx":  
                    file_path = os.path.join(folder_path, filename)
                    wb = xl.Workbooks.Open(file_path)

                    # 기존 이름 삭제 (예외 처리 추가)
                    if wb.Names.Count > 0:  # 기존 이름이 있을 때만 실행
                        for name in list(wb.Names):
                            if name is None or name.Name is None:
                                continue  # None 값이 있으면 건너뜀
                            try:
                                name_str = str(name.Name)
                                print(f"Deleting: {name_str}")
                                name.Delete()
                                print(f"Deleted: {name_str}")
                            except Exception as e:
                                print(f"Error deleting name {name_str}: {e}")

                    # 새로운 이름 추가
                    for key, value in names.items():
                        if self.check_name(wb, key):
                            wb.Names.Item(key).RefersTo = value
                        else:
                            wb.Names.Add(Name=key, RefersTo=value)

                    wb.Save()  
                    wb.Close()  

        xl.Quit()

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

