import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QLabel, QPushButton, QFileDialog, QMessageBox
from PyQt5.QtCore import Qt
import pandas as pd
from iteround import saferound
import numpy as np
import os

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Rounder")
        self.setMinimumSize(400, 200)  # Reduced size

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        self.select_button = QPushButton("Select Excel File")
        self.select_button.clicked.connect(self.select_file)
        layout.addWidget(self.select_button)

        self.status_label = QLabel("")
        self.status_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.status_label)

    def select_file(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx)")
        if file_name:
            self.process_file(file_name)

    def process_file(self, file_path):
        try:
            self.status_label.setText("Processing...")
            QApplication.processEvents()

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
                
                # Cập nhật giá trị cột ULNL task bằng với Task round
                group['ULNL task'] = group['Task round']

                # Xóa cột Task round
                group.drop(columns=['Task round'], inplace=True)

                new_groups.append(group)

            # Create output file path
            output_file_name = os.path.basename(file_path).replace(".xlsx", "_new.xlsx")  # Tên file mới
            output_path = os.path.join(os.path.dirname(file_path), output_file_name)  # Đường dẫn đầy đủ

            # Create output file
            df_new = pd.concat(new_groups)
            df_new.sort_index().to_excel(output_path, index=None)
            
            self.status_label.setText("Complete!")
            QMessageBox.information(self, "Success", f"File saved as:\n{output_path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_()) 