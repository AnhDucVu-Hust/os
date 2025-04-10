import sys
import os
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QLabel, QPushButton, QFileDialog, QMessageBox)
from PyQt5.QtCore import Qt
import pandas as pd
import roman
import numpy
from styleframe import StyleFrame, Styler, utils
from datetime import datetime
import warnings
warnings.filterwarnings("ignore")

class DropArea(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setMinimumSize(300, 200)
        self.setStyleSheet("""
            QWidget {
                border: 2px dashed #aaa;
                border-radius: 10px;
                background-color: #f8f8f8;
            }
        """)
        
        layout = QVBoxLayout()
        self.label = QLabel("Drag and drop Excel file here\nor click to select file")
        self.label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.label)
        self.setLayout(layout)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        files = [url.toLocalFile() for url in event.mimeData().urls()]
        if files:
            self.parent().process_file(files[0])

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("OS Check Processor")
        self.setMinimumSize(400, 300)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        self.drop_area = DropArea(self)
        layout.addWidget(self.drop_area)

        self.select_button = QPushButton("Select Excel File")
        self.select_button.clicked.connect(self.select_file)
        layout.addWidget(self.select_button)

        self.status_label = QLabel("")
        self.status_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.status_label)

    def select_file(self):
        file_name, _ = QFileDialog.getOpenFileName(
            self, "Select Excel File", "", "Excel Files (*.xlsx)"
        )
        if file_name:
            self.process_file(file_name)

    def get_time(self, year, month):
        return "03/06/2024 đến 28/06/2024"  # You might want to make this dynamic

    def process_file(self, file_path):
        try:
            if not file_path.endswith('.xlsx'):
                QMessageBox.warning(self, "Error", "Please select an Excel (.xlsx) file")
                return

            self.status_label.setText("Processing...")
            QApplication.processEvents()

            # Read the Excel file
            df = StyleFrame.read_excel(file_path, sheet_name="Sheet0")
            hop_dong = list(set(df["Hợp đồng"]))
            df_hd = {}
            df_chuan = {}
            
            for hd in hop_dong:
                self.status_label.setText(f"Processing contract: {hd}")
                QApplication.processEvents()
                
                df_chuan[hd] = pd.DataFrame(columns=['Id','Row label','Bảo trì','Nâng cấp','Total'])
                df_hd[hd] = df.loc[df["Hợp đồng"]==hd]
                he_thongs = list(set(df_hd[hd]["Hệ thống&CTKT"]))
                
                # Your existing processing logic here
                # ... (keeping the same logic as in os_check.py)
                
                # Save the processed file
                output_path = f'./my_excel_{str(hd).replace("/","_")}.xlsx'
                df_bbnv = self.process_contract(df_hd[hd], df_chuan[hd], hd)
                ew = StyleFrame.ExcelWriter(output_path)
                df_bbnv.to_excel(ew)
                ew.save()

            self.status_label.setText("Processing complete!\nAll files generated successfully!")
            QMessageBox.information(self, "Success", "All files processed successfully!")

        except Exception as e:
            self.status_label.setText("Error occurred!")
            QMessageBox.critical(self, "Error", f"An error occurred:\n{str(e)}")

    def process_contract(self, df_hd, df_chuan, hd):
        # Your existing contract processing logic here
        # This should be the same logic from your original os_check.py
        # Return the processed df_bbnv
        pass

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_()) 