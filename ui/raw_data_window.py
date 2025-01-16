import pandas as pd
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import QWidget, QVBoxLayout, QLabel, QPushButton, QTableWidget, QFileDialog, QMessageBox, \
    QTableWidgetItem


class Step1(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        layout = QVBoxLayout()
        self.label = QLabel("Step 1: Load Data")
        self.load_data_button = QPushButton("Load Raw Data from Local")
        self.load_data_button.clicked.connect(self.load_data_file)

        self.table = QTableWidget()
        self.table.setAlternatingRowColors(True)

        layout.addWidget(self.label)
        layout.addWidget(self.load_data_button)
        layout.addWidget(self.table)
        self.setLayout(layout)

        self.raw_data_columns = []
        self.data = None  # Store the pandas DataFrame

    def load_data_file_old(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Raw Data File", "", "Excel Files (*.xlsx *.xls)")
        if file_path:
            try:
                self.data = pd.read_excel(file_path)
                self.raw_data_columns = self.data.columns.tolist()
                self.populate_table(self.data)
                QMessageBox.information(self, "File Loaded", f"Raw data loaded: {file_path}")
                self.main_window.step2.update_raw_data_columns(self.raw_data_columns)
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to load file: {str(e)}")

    def load_data_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Raw Data File", "", "Excel Files (*.xlsx *.xls)")
        if file_path:
            try:
                self.data = pd.read_excel(file_path)

                for column in self.data.columns:
                    if pd.api.types.is_datetime64_any_dtype(self.data[column]):
                        if (self.data[column].dt.time != pd.Timestamp(0).time()).any():
                            continue
                        else:
                            self.data[column] = self.data[column].dt.date

                self.raw_data_columns = self.data.columns.tolist()
                self.populate_table(self.data)
                QMessageBox.information(self, "File Loaded", f"Raw data loaded: {file_path}")
                self.main_window.step2.update_raw_data_columns(self.raw_data_columns)
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to load file: {str(e)}")

    def populate_table(self, data):
        self.table.setRowCount(data.shape[0])
        self.table.setColumnCount(data.shape[1])
        self.table.setHorizontalHeaderLabels(data.columns)

        for row in range(data.shape[0]):
            for col in range(data.shape[1]):
                value = data.iloc[row, col]
                item = QTableWidgetItem(str(value) if pd.notna(value) else "")
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)  # Make cells read-only
                self.table.setItem(row, col, item)

        self.table.resizeColumnsToContents()
        self.table.resizeRowsToContents()
