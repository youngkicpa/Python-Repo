import sys
import os
import win32com.client
from PySide6.QtWidgets import QApplication, QMainWindow, QFileDialog, QTableWidget, QTableWidgetItem
from PySide6.QtGui import QAction

class ExcelReaderApp(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Excel Reader")
        self.setGeometry(100, 100, 800, 600)

        self.table_widget = QTableWidget()
        self.setCentralWidget(self.table_widget)

        self.statusBar().showMessage("Ready")

        self.create_menu()

    def create_menu(self):
        menu_bar = self.menuBar()
        file_menu = menu_bar.addMenu("File")

        open_action = QAction("Open File", self)
        open_action.triggered.connect(self.open_excel_file)
        file_menu.addAction(open_action)

        exit_action = QAction("Exit", self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

    def open_excel_file(self):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(self, "Open Excel File", "", "Excel Files (*.xls *.xlsx);;All Files (*)", options=options)

        if file_name:
            try:
                excel_app = win32com.client.Dispatch("Excel.Application")
                excel_app.Visible = False
                workbook = excel_app.Workbooks.Open(file_name)
                sheet = workbook.Sheets(1)

                data = sheet.UsedRange.Value
                
                workbook.Close(False)
                excel_app.Quit()
                excel_app = None  # Properly release COM object

                if data:
                    self.table_widget.setRowCount(len(data))
                    self.table_widget.setColumnCount(len(data[len(data)-1]))
                    for i, row in enumerate(data):
                        for j, value in enumerate(row):
                            item = QTableWidgetItem()
                            if value is not None:
                                if isinstance(value, (int, float)):
                                    item.setData(0, f"{value:,.0f}")
                                    item.setTextAlignment(2)
                                else:
                                    item.setData(0, str(value))
                            else:
                                item.setData(0, "")
                            self.table_widget.setItem(i, j, item)
                    self.statusBar().showMessage(f"File: {os.path.basename(file_name)}, Sheet: {sheet.Name}")
                else:
                    self.statusBar().showMessage("No data found in the selected sheet.")
            except Exception as e:
                self.statusBar().showMessage(f"Failed to open file: {str(e)}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelReaderApp()
    window.show()
    sys.exit(app.exec_())
