import sys
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QLabel, 
                            QPushButton, QFileDialog, QMessageBox, QHBoxLayout, QDateEdit)
from PyQt5.QtCore import Qt, QDate
import pandas as pd
import numpy
from styleframe import StyleFrame, Styler
import roman
from datetime import datetime

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("OS Document Processor")
        self.setMinimumSize(400, 300)  # Increased height for date pickers

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # Date selection section
        date_widget = QWidget()
        date_layout = QVBoxLayout(date_widget)
        
        # From date
        from_date_layout = QHBoxLayout()
        from_date_label = QLabel("From date:")
        self.from_date = QDateEdit()
        self.from_date.setDate(QDate.currentDate())
        self.from_date.setCalendarPopup(True)
        from_date_layout.addWidget(from_date_label)
        from_date_layout.addWidget(self.from_date)
        date_layout.addLayout(from_date_layout)

        # To date
        to_date_layout = QHBoxLayout()
        to_date_label = QLabel("To date:")
        self.to_date = QDateEdit()
        self.to_date.setDate(QDate.currentDate())
        self.to_date.setCalendarPopup(True)
        to_date_layout.addWidget(to_date_label)
        to_date_layout.addWidget(self.to_date)
        date_layout.addLayout(to_date_layout)

        layout.addWidget(date_widget)

        # File selection button
        self.select_button = QPushButton("Select Excel File")
        self.select_button.clicked.connect(self.select_file)
        layout.addWidget(self.select_button)

        # Status label
        self.status_label = QLabel("Select an Excel file to process")
        self.status_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.status_label)

    def select_file(self):
        file_name, _ = QFileDialog.getOpenFileName(
            self, "Select Excel File", "", "Excel Files (*.xlsx)"
        )
        if file_name:
            self.process_file(file_name)

    def get_time(self):
        from_date = self.from_date.date().toString("dd/MM/yyyy")
        to_date = self.to_date.date().toString("dd/MM/yyyy")
        return f"{from_date} đến {to_date}"

    def process_file(self, file_path):
        try:
            # Validate dates
            if self.from_date.date() > self.to_date.date():
                QMessageBox.warning(self, "Warning", "From date cannot be later than To date")
                return

            self.status_label.setText("Processing...")
            QApplication.processEvents()

            # Read the Excel file
            df = StyleFrame.read_excel(file_path, sheet_name="Sheet0")
            hop_dong = list(set(df["Hợp đồng"]))
            
            for hd in hop_dong:
                self.status_label.setText(f"Processing contract: {hd}")
                QApplication.processEvents()
                
                df_chuan = pd.DataFrame(columns=['Id','Row label','Bảo trì','Nâng cấp','Total'])
                df_hd = df.loc[df["Hợp đồng"]==hd]
                he_thongs = list(set(df_hd["Hệ thống&CTKT"]))
                
                # Process systems
                id_he_thong = 1
                tong_bao_tri = tong_nang_cap = 0
                
                for he_thong in he_thongs:
                    df_he_thong = df_hd.loc[df_hd["Hệ thống&CTKT"]==he_thong]
                    stories = list(set(df_he_thong["Tên story"]))
                    
                    # Calculate system totals
                    bao_tri = df_he_thong.loc[df_he_thong["Phân loại"]=='Bảo trì']["ULNL task"].sum()
                    nang_cap = df_he_thong.loc[df_he_thong["Phân loại"]=='Nâng cấp']["ULNL task"].sum()
                    total = bao_tri + nang_cap
                    tong_nang_cap += nang_cap
                    tong_bao_tri += bao_tri
                    
                    # Add system row
                    df_chuan = df_chuan.append({
                        'Id': roman.toRoman(id_he_thong),
                        'Row label': he_thong,
                        'Bảo trì': bao_tri,
                        "Nâng cấp": nang_cap,
                        'Total': total
                    }, ignore_index=True)
                    
                    # Process stories
                    self.process_stories(df_he_thong, df_chuan, stories, id_he_thong)
                    id_he_thong += 1

                # Add total row
                df_chuan = df_chuan.append({
                    'Id': '',
                    "Row label": 'Tổng',
                    'Bảo trì': tong_bao_tri,
                    "Nâng cấp": tong_nang_cap,
                    'Total': tong_bao_tri + tong_nang_cap
                }, ignore_index=True)

                # Format and save
                df_bbnv = self.format_output(df_chuan, hd)
                output_path = f'./OS_Document_{str(hd).replace("/","_")}.xlsx'
                ew = StyleFrame.ExcelWriter(output_path)
                df_bbnv.to_excel(ew)
                ew.save()

            self.status_label.setText("Processing complete!")
            QMessageBox.information(self, "Success", "All files processed successfully!")

        except Exception as e:
            self.status_label.setText("Error occurred!")
            QMessageBox.critical(self, "Error", str(e))

    def process_stories(self, df_he_thong, df_chuan, stories, id_he_thong):
        id_story = 1
        for story in stories:
            df_xet = df_he_thong.loc[df_he_thong["Tên story"]==story]
            bao_tri = df_xet.loc[df_xet["Phân loại"] == 'Bảo trì']["ULNL task"].sum()
            nang_cap = df_xet.loc[df_xet["Phân loại"] == 'Nâng cấp']["ULNL task"].sum()
            
            # Add story row
            df_chuan = df_chuan.append({
                'Id': f"{roman.toRoman(id_he_thong)}.{id_story}",
                "Row label": story,
                'Bảo trì': bao_tri,
                "Nâng cấp": nang_cap,
                'Total': bao_tri + nang_cap
            }, ignore_index=True)
            
            # Process tasks
            self.process_tasks(df_xet, df_chuan)
            id_story += 1

    def process_tasks(self, df_xet, df_chuan):
        tasks = list(set(df_xet["Summary"]))
        id_task = 1
        for task in tasks:
            df_xet2 = df_xet.loc[df_xet["Summary"]==task]
            bao_tri = df_xet2.loc[df_xet2["Phân loại"] == 'Bảo trì']["ULNL task"].sum()
            nang_cap = df_xet2.loc[df_xet2["Phân loại"] == 'Nâng cấp']["ULNL task"].sum()
            
            df_chuan = df_chuan.append({
                'Id': id_task,
                "Row label": task,
                'Bảo trì': bao_tri,
                "Nâng cấp": nang_cap,
                'Total': bao_tri + nang_cap
            }, ignore_index=True)
            id_task += 1

    def format_output(self, df_chuan, hd):
        df_bbnv = df_chuan.rename(columns={
            'Id': 'TT',
            'Row label': 'Nội dung công việc chi tiết',
            'Nâng cấp': 'Kết quả hoàn thành tương ứng nỗ lực nâng cấp (số MM)',
            'Bảo trì': 'Kết quả hoàn thành tương ứng nỗ lực bảo trì (số MM)'
        })
        
        # Add additional columns
        df_bbnv["Kết quả hoàn thành đánh giá theo phần trăm (%)"] = '100%'
        df_bbnv["Thời gian hoàn thành"] = self.get_time()
        df_bbnv["Kết quả hoàn thành tương ứng nỗ lực xây mới (số MM)"] = 0

        # Reorder columns
        df_bbnv = df_bbnv[[
            "TT", "Nội dung công việc chi tiết", "Thời gian hoàn thành",
            "Kết quả hoàn thành tương ứng nỗ lực xây mới (số MM)",
            "Kết quả hoàn thành tương ứng nỗ lực nâng cấp (số MM)",
            "Kết quả hoàn thành tương ứng nỗ lực bảo trì (số MM)",
            "Kết quả hoàn thành đánh giá theo phần trăm (%)"
        ]]

        # Apply styling
        df_bbnv.replace(0, numpy.NAN, inplace=True)
        df_bbnv = StyleFrame(df_bbnv)
        self.apply_styling(df_bbnv)
        
        return df_bbnv

    def apply_styling(self, df_bbnv):
        # Set column widths
        column_widths = {
            "TT": 6.67,
            "Nội dung công việc chi tiết": 47.67,
            "Thời gian hoàn thành": 16.50,
            "Kết quả hoàn thành tương ứng nỗ lực xây mới (số MM)": 14.60,
            "Kết quả hoàn thành tương ứng nỗ lực nâng cấp (số MM)": 15.83,
            "Kết quả hoàn thành tương ứng nỗ lực bảo trì (số MM)": 22.83,
            "Kết quả hoàn thành đánh giá theo phần trăm (%)": 13.17
        }
        
        for col, width in column_widths.items():
            df_bbnv.set_column_width(col, width=width)

        # Apply fonts and alignment
        indexes_to_bold = df_bbnv[df_bbnv['TT'].apply(lambda x: not str(x).isnumeric())]
        df_bbnv.apply_style_by_indexes(
            cols_to_style='Nội dung công việc chi tiết',
            indexes_to_style=indexes_to_bold,
            styler_obj=Styler(horizontal_alignment='left', bold=True, font="Times New Roman"),
            complement_style=Styler(horizontal_alignment='left', font="Times New Roman")
        )

        # Apply right alignment to numeric columns
        right_aligned_cols = [
            "Kết quả hoàn thành tương ứng nỗ lực nâng cấp (số MM)",
            "Kết quả hoàn thành tương ứng nỗ lực bảo trì (số MM)"
        ]
        
        for col in right_aligned_cols:
            df_bbnv.apply_style_by_indexes(
                cols_to_style=col,
                indexes_to_style=df_bbnv.index,
                styler_obj=Styler(horizontal_alignment='right', font="Times New Roman")
            )

        # Apply header styling
        df_bbnv.apply_headers_style(styler_obj=Styler(bold=True, font="Times New Roman"))

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_()) 