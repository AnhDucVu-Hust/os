import sys
import os
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QLabel, QPushButton, QFileDialog, QMessageBox)
from PyQt5.QtCore import Qt
import numpy as np
import pandas as pd
from iteround import saferound

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
        self.setWindowTitle("Excel Rounder")
        self.setMinimumSize(400, 300)

        # Create central widget and layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # Create drop area
        self.drop_area = DropArea(self)
        layout.addWidget(self.drop_area)

        # Create select file button
        self.select_button = QPushButton("Select Excel File")
        self.select_button.clicked.connect(self.select_file)
        layout.addWidget(self.select_button)

        # Create status label
        self.status_label = QLabel("")
        self.status_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.status_label)

    def select_file(self):
        file_name, _ = QFileDialog.getOpenFileName(
            self, "Select Excel File", "", "Excel Files (*.xlsx)"
        )
        if file_name:
            self.process_file(file_name)

    def process_file(self, file_path):
        try:
            if not file_path.endswith('.xlsx'):
                QMessageBox.warning(self, "Error", "Please select an Excel (.xlsx) file")
                return

            self.status_label.setText("Processing...")
            QApplication.processEvents()

            # Read the Excel file
            df = pd.read_excel(file_path, sheet_name="Sheet0")
            new_groups = []
            groups = df.groupby('Mã story')

            for idx, group in groups:
                tong = group["ULNL story"].max()
                group['ULNL task'] = group['ULNL task'] * 100
                group['Task round'] = saferound(list(group['ULNL task']), places=0)
                group['Task round'] /= 100
                group['ULNL task'] = group['ULNL task'] / 100
                gap = round((tong - sum(group['Task round'])) / 0.01)
                
                if np.abs(gap) > 0.01:
                    for index, row in group.iterrows():
                        if gap > 0:
                            if row["Task round"] < row["ULNL task"]:
                                row["Task round"] = (row["Task round"] * 100 + 1) / 100
                                break
                        else:
                            if row["Task round"] > row["ULNL task"]:
                                row["Task round"] = (row["Task round"] * 100 - 1) / 100
                                break
                new_groups.append(group)

            # Create output file
            df_new = pd.concat(new_groups)
            output_path = os.path.splitext(file_path)[0] + "_new.xlsx"
            df_new.sort_index().to_excel(output_path, index=None)
            
            self.status_label.setText(f"Processing complete!\nOutput saved to:\n{output_path}")
            QMessageBox.information(self, "Success", f"File processed successfully!\nOutput saved to:\n{output_path}")

        except Exception as e:
            self.status_label.setText("Error occurred!")
            QMessageBox.critical(self, "Error", f"An error occurred:\n{str(e)}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_()) 