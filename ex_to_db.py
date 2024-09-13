import sys
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QLabel, QFileDialog, QTextEdit, QVBoxLayout, QWidget
import os

class ExcelToDirectoryApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.excel_path = None
        self.directory_path = None

    def initUI(self):
        self.setWindowTitle('Excel to Directory Checker')
        self.setGeometry(100, 100, 600, 400)

        # Layout
        layout = QVBoxLayout()

        # Button to select Excel file
        self.excel_btn = QPushButton('Select Excel File', self)
        self.excel_btn.clicked.connect(self.select_excel_file)
        layout.addWidget(self.excel_btn)

        # Button to select Directory
        self.dir_btn = QPushButton('Select Directory', self)
        self.dir_btn.clicked.connect(self.select_directory)
        layout.addWidget(self.dir_btn)

        # Button to check files
        self.check_btn = QPushButton('Check Files in Directory', self)
        self.check_btn.clicked.connect(self.check_files)
        layout.addWidget(self.check_btn)

        # Text Area to show results
        self.result_text = QTextEdit(self)
        layout.addWidget(self.result_text)

        # Set the layout
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def select_excel_file(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        file_name, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx);;All Files (*)", options=options)
        if file_name:
            self.excel_path = file_name
            self.result_text.append(f"Selected Excel File: {file_name}")

    def select_directory(self):
        directory = QFileDialog.getExistingDirectory(self, "Select Directory")
        if directory:
            self.directory_path = directory
            self.result_text.append(f"Selected Directory: {directory}")

    def check_files(self):
        if not self.excel_path or not self.directory_path:
            self.result_text.append("Please select both Excel file and directory.")
            return

        try:
            # Read Excel File
            df = pd.read_excel(self.excel_path)
            self.result_text.append(f"Loaded Excel file: {self.excel_path}")

            # Check if the Excel file has the expected column for file names
            if 'file_column' not in df.columns:
                self.result_text.append("Excel file must contain 'file_column' column.")
                return

            # Get the file names from the Excel file
            file_names_from_excel = df['file_column'].astype(str).str.strip().str.lower().tolist()

            # List files in the selected directory
            dir_content = [f.strip().lower() for f in os.listdir(self.directory_path)]

            # Display Comparison Results
            results = []
            for file_name in file_names_from_excel:
                if file_name not in dir_content:
                    results.append(f'{file_name} is not in the Dir.')
                else:
                    results.append(f'{file_name} is in the Dir.')

            self.result_text.append("\n".join(results))

        except Exception as e:
            self.result_text.append(f"An error occurred: {str(e)}")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = ExcelToDirectoryApp()
    ex.show()
    sys.exit(app.exec_())
