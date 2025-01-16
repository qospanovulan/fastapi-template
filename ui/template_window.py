import os
import sys

import requests
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont, QColor
from PyQt6.QtWidgets import QWidget, QVBoxLayout, QSplitter, QLabel, QPushButton, QListWidget, QTableWidget, \
    QHBoxLayout, QListWidgetItem, QFileDialog, QDialog, QMessageBox, QTableWidgetItem, QProgressDialog, QInputDialog
from openpyxl.reader.excel import load_workbook

from reportlab.lib.pagesizes import letter, A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from openpyxl import load_workbook
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib import fonts
from fpdf import FPDF

if getattr(sys, 'frozen', False):
    bundle_dir = sys._MEIPASS
else:
    base_dir = os.path.dirname(os.path.abspath(__file__))
    bundle_dir = os.path.join(base_dir, "fonts")
    print(bundle_dir)

pdfmetrics.registerFont(TTFont('Arial', os.path.join(bundle_dir, 'ARIAL.TTF')))
pdfmetrics.registerFont(TTFont('Arial-Bold', os.path.join(bundle_dir, 'ARIALBD.TTF')))
pdfmetrics.registerFont(TTFont('Arial-Italic', os.path.join(bundle_dir, 'ARIALI.TTF')))
pdfmetrics.registerFont(TTFont('Calibri', os.path.join(bundle_dir, 'calibri.ttf')))


class TemplateTable(QTableWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.zoom_level = 100
        self.default_font_size = self.font().pointSize()

    def wheelEvent(self, event):
        if event.modifiers() == Qt.KeyboardModifier.ControlModifier:
            delta = event.angleDelta().y()
            if delta > 0:
                self.zoom_level += 10
            else:
                self.zoom_level -= 10
                if self.zoom_level < 10:
                    self.zoom_level = 10

            scale_factor = self.zoom_level / 100
            self.horizontalHeader().setDefaultSectionSize(int(100 * scale_factor))
            self.verticalHeader().setDefaultSectionSize(int(30 * scale_factor))

            new_font_size = int(self.default_font_size * scale_factor)
            self.update_table_font(new_font_size)

            event.accept()
        else:
            super().wheelEvent(event)

    def update_table_font_old(self, font_size):
        """Update the font size of all cells in the table."""
        font = QFont()
        font.setPointSize(font_size)

        for row in range(self.rowCount()):
            for col in range(self.columnCount()):
                item = self.item(row, col)
                if item:
                    item.setFont(font)

    def update_table_font(self, font_size):
        """Update the font size of all cells in the table while preserving existing font styles."""
        for row in range(self.rowCount()):
            for col in range(self.columnCount()):
                item = self.item(row, col)
                if item:
                    current_font = item.font()
                    current_font.setPointSize(font_size * 2)
                    item.setFont(current_font)


class Step2(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.backend_url = "http://2.56.178.70:8000"
        self.main_window = main_window

        left_layout = QVBoxLayout()
        right_layout = QVBoxLayout()

        splitter = QSplitter(Qt.Orientation.Horizontal)

        self.label = QLabel("Step 2: Load Template")

        self.load_local_button = QPushButton("Load Template from Local")
        self.load_local_button.clicked.connect(self.load_local_file)

        self.load_cloud_button = QPushButton("Load Template from Cloud")
        self.load_cloud_button.clicked.connect(self.load_cloud_file)

        self.raw_data_label = QLabel("Raw Data Columns:")
        self.raw_data_list = QListWidget()

        self.template_table = TemplateTable()
        self.template_table.itemChanged.connect(self.on_template_changed)

        self.button_layout = QHBoxLayout()
        self.save_button = QPushButton("Save Changes")
        self.save_button.clicked.connect(self.save_changes)
        self.save_button.setEnabled(False)

        self.generate_button = QPushButton("Generate Files")
        self.generate_button.clicked.connect(self.generate_files)
        self.generate_button.setEnabled(False)

        self.button_layout.addWidget(self.save_button)
        self.button_layout.addWidget(self.generate_button)

        left_panel = QWidget()
        left_layout.addWidget(self.raw_data_label)
        left_layout.addWidget(self.raw_data_list)
        left_panel.setLayout(left_layout)

        right_panel = QWidget()
        right_layout.addWidget(self.load_local_button)
        right_layout.addWidget(self.load_cloud_button)
        right_layout.addWidget(self.template_table)
        right_layout.addLayout(self.button_layout)
        right_panel.setLayout(right_layout)

        splitter.addWidget(left_panel)
        splitter.addWidget(right_panel)

        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 5)

        main_layout = QVBoxLayout()
        main_layout.addWidget(splitter)

        self.setLayout(main_layout)

        self.template_path = ""
        self.saved_template_path = ""
        self.merged_ranges = []
        self.workbook = None
        self.sheet = None
        self.changes = {}
        self.placeholders = {}

    def update_raw_data_columns(self, columns):
        self.raw_data_list.clear()
        for idx, column_name in enumerate(columns, start=1):
            item = QListWidgetItem(f"{idx}. {column_name}")
            self.raw_data_list.addItem(item)

    def on_template_changed(self, item):
        if item and self.template_path:
            self.save_button.setEnabled(True)
            if self.sheet:
                row = item.row()
                col = item.column()

                self.changes[(row + 1, col + 1)] = item.text()

    def detect_placeholders(self, cell_value, row, col):
        if cell_value and isinstance(cell_value, str):
            import re
            matches = re.findall(r"\{(\d+)\}", cell_value)
            if matches:
                self.placeholders[(row, col)] = {
                    'template': cell_value,
                    'column_ids': [int(m) - 1 for m in matches]
                }

    def load_local_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Excel Template", "", "Excel Files (*.xlsx *.xls)")
        if file_path:
            self.template_path = file_path
            self.display_template()
            self.save_button.setEnabled(True)

    def load_cloud_file(self):
        try:
            response = requests.get(f"{self.backend_url}/templates/")
            response.raise_for_status()
            templates = response.json()

            dialog = QDialog(self)
            dialog.setWindowTitle("Select Template from Cloud")
            layout = QVBoxLayout()
            list_widget = QListWidget()
            for template in templates:
                item = QListWidgetItem(template['name'])
                item.setData(Qt.ItemDataRole.UserRole, template['id'])
                list_widget.addItem(item)
            layout.addWidget(list_widget)

            load_button = QPushButton("Load")
            load_button.clicked.connect(lambda: self.fetch_template_from_cloud(list_widget, dialog))
            layout.addWidget(load_button)

            dialog.setLayout(layout)
            dialog.exec()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to fetch templates: {str(e)}")

    def fetch_template_from_cloud(self, list_widget, dialog):
        selected_item = list_widget.currentItem()
        if selected_item:
            template_name = selected_item.text()
            template_id = selected_item.data(Qt.ItemDataRole.UserRole)
            try:
                response = requests.get(f"{self.backend_url}/templates/{template_id}/")
                response.raise_for_status()

                with open(template_name, "wb") as f:
                    f.write(response.content)

                self.template_path = template_name
                self.display_template()
                dialog.accept()
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to load template: {str(e)}")

    def display_template(self):
        if self.template_path:
            try:
                self.workbook = load_workbook(self.template_path)
                self.sheet = self.workbook.active

                self.template_table.setRowCount(self.sheet.max_row)
                self.template_table.setColumnCount(self.sheet.max_column)

                self.merged_ranges = []
                self.placeholders = {}
                self.changes = {}

                for row_idx, row in enumerate(self.sheet.iter_rows()):
                    for col_idx, cell in enumerate(row):
                        value = cell.value
                        item = QTableWidgetItem(str(value) if value is not None else "")

                        if cell.font:
                            font = QFont()
                            if cell.font.name:
                                font.setFamily(cell.font.name)
                            if cell.font.size:
                                font.setPointSize(int(cell.font.size))
                            font.setBold(cell.font.bold or False)
                            font.setItalic(cell.font.italic or False)
                            item.setFont(font)

                        if cell.alignment and cell.alignment.horizontal:
                            if cell.alignment.horizontal == "center":
                                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                            elif cell.alignment.horizontal == "left":
                                item.setTextAlignment(Qt.AlignmentFlag.AlignLeft)
                            elif cell.alignment.horizontal == "right":
                                item.setTextAlignment(Qt.AlignmentFlag.AlignRight)

                        if cell.fill and cell.fill.start_color and cell.fill.start_color.rgb:
                            rgb = cell.fill.start_color.rgb
                            if isinstance(rgb, str) and len(rgb) >= 6:
                                color = QColor(f"#f0f0f0")
                                item.setBackground(color)

                        self.template_table.setItem(row_idx, col_idx, item)

                for merged_range in self.sheet.merged_cells.ranges:
                    min_row = merged_range.min_row - 1
                    min_col = merged_range.min_col - 1
                    row_span = merged_range.max_row - merged_range.min_row + 1
                    col_span = merged_range.max_col - merged_range.min_col + 1
                    self.template_table.setSpan(min_row, min_col, row_span, col_span)
                    self.merged_ranges.append((min_row, min_col, row_span, col_span))

                self.template_table.resizeColumnsToContents()
                self.template_table.resizeRowsToContents()

                if self.placeholders:
                    QMessageBox.information(
                        self,
                        "Placeholders Detected",
                        f"Found {len(self.placeholders)} cells with placeholders in the template."
                    )

            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to load template: {str(e)}")

    def save_changes(self):
        self.save_changes_local()

        reply = QMessageBox.question(
            self,
            "Upload to Cloud",
            "Do you want to replace this template on the server?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )

        if reply == QMessageBox.StandardButton.Yes:
            try:
                with open(self.saved_template_path, "rb") as f:
                    files = {"file": f}
                    response = requests.post(f"{self.backend_url}/templates/update/", files=files)
                    response.raise_for_status()
                    QMessageBox.information(self, "Success", "Template uploaded to the server.")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to upload template: {str(e)}")

    def save_changes_local(self):
        if not self.sheet:
            QMessageBox.warning(self, "Warning", "Please load a template first.")
            return

        try:
            for merged_range in self.sheet.merged_cells.ranges.copy():
                self.sheet.unmerge_cells(str(merged_range))

            for row in range(self.template_table.rowCount()):
                for col in range(self.template_table.columnCount()):
                    item = self.template_table.item(row, col)
                    if item:
                        cell = self.sheet.cell(row=row + 1, column=col + 1)
                        cell.value = self.changes.get((row + 1, col + 1), item.text())

                        self.detect_placeholders(cell.value, row + 1, col + 1)

            for min_row, min_col, row_span, col_span in self.merged_ranges:
                start_row = min_row + 1
                start_col = min_col + 1
                end_row = start_row + row_span - 1
                end_col = start_col + col_span - 1

                value = self.sheet.cell(row=start_row, column=start_col).value

                self.sheet.merge_cells(
                    start_row=start_row,
                    start_column=start_col,
                    end_row=end_row,
                    end_column=end_col
                )

                self.sheet.cell(row=start_row, column=start_col).value = value

            if self.template_path.endswith('_marked.xlsx'):
                self.saved_template_path = self.template_path
            else:
                self.saved_template_path = self.template_path.replace('.xlsx', '_marked.xlsx')
            self.workbook.save(self.saved_template_path)
            self.changes = {}
            QMessageBox.information(self, "Success", f"Template saved as: {self.saved_template_path}")
            self.generate_button.setEnabled(True)

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save template: {str(e)}")

    def generate_files_old(self):
        if not self.saved_template_path or not self.main_window.step1.data is not None:
            QMessageBox.warning(self, "Warning", "Please ensure template is saved and raw data is loaded.")
            return

        chosen_dir = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if not chosen_dir:
            return
        try:
            raw_data = self.main_window.step1.data

            import os
            output_dir = os.path.join(chosen_dir, os.path.basename(self.template_path).replace('.xlsx', '_filled'))
            os.makedirs(output_dir, exist_ok=True)

            total_files = len(raw_data)
            progress = QProgressDialog("Generating files...", None, 0, total_files, self)
            progress.setWindowModality(Qt.WindowModality.WindowModal)

            for idx, row in enumerate(raw_data.itertuples(), start=1):
                new_wb = load_workbook(self.saved_template_path)
                new_sheet = new_wb.active

                for (row_idx, col_idx), placeholder_info in self.placeholders.items():
                    template_text = placeholder_info['template']

                    for col_id in placeholder_info['column_ids']:
                        if col_id < len(raw_data.columns):
                            value = str(row[col_id + 1])
                            template_text = template_text.replace(f"{{{col_id + 1}}}", value)

                    new_sheet.cell(row=row_idx, column=col_idx).value = template_text

                output_path = os.path.join(output_dir, f'generated_{idx}.xlsx')
                new_wb.save(output_path)
                progress.setValue(idx)

            QMessageBox.information(
                self,
                "Success",
                f"Generated {total_files} files in folder: {output_dir}"
            )

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to generate files: {str(e)}")
            raise e

    def generate_files_old_2(self):
        if not self.saved_template_path or self.main_window.step1.data is None:
            QMessageBox.warning(self, "Warning", "Please ensure template is saved and raw data is loaded.")
            return

        chosen_dir = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if not chosen_dir:
            return

        # Asking the user for file type and name pattern
        file_type, ok = QInputDialog.getItem(self, "Select File Type", "Choose file format:", ["Excel", "PDF"], 0,
                                             False)
        if not ok:
            return

        name_pattern, ok = QInputDialog.getText(self, "File Naming Pattern",
                                                "Enter file naming pattern (use {1}, {9} for column index pairs):",
                                                text="{1}_{9}_gen")
        if not ok:
            return

        try:
            raw_data = self.main_window.step1.data

            output_dir = os.path.join(chosen_dir, os.path.basename(self.template_path).replace('.xlsx', '_filled'))
            os.makedirs(output_dir, exist_ok=True)

            total_files = len(raw_data)
            progress = QProgressDialog("Generating files...", None, 0, total_files, self)
            progress.setWindowModality(Qt.WindowModality.WindowModal)

            for idx, row in enumerate(raw_data.itertuples(), start=1):
                # if file_type == "Excel":
                    # Excel file generation
                new_wb = load_workbook(self.saved_template_path)
                new_sheet = new_wb.active

                for (row_idx, col_idx), placeholder_info in self.placeholders.items():
                    template_text = placeholder_info['template']

                    for col_id in placeholder_info['column_ids']:
                        if col_id < len(raw_data.columns):
                            value = str(row[col_id + 1])
                            template_text = template_text.replace(f"{{{col_id + 1}}}", value)

                    new_sheet.cell(row=row_idx, column=col_idx).value = template_text

                # Replace placeholders with actual column index values in filename
                file_name = name_pattern
                for col_idx, col_name in enumerate(raw_data.columns.to_list()):
                    print(f"{col_idx=}")
                    # Replace {col_idx} with actual data from the row
                    value = str(row[col_idx])  # Adjust for 1-based index of `itertuples()`
                    print(f"{value=}")
                    file_name = file_name.replace(f"{{{col_idx}}}", value)
                    print(f"{file_name=}")

                if file_type == "Excel":
                    output_path = os.path.join(output_dir, file_name.format(idx=idx)) + ".xlsx"
                    new_wb.save(output_path)

                if file_type == "PDF":
                    pdf = FPDF()
                    pdf.set_auto_page_break(auto=True, margin=15)
                    pdf.add_page()

                    ...

                progress.setValue(idx)

            QMessageBox.information(self, "Success", f"Generated {total_files} files in folder: {output_dir}")

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to generate files: {str(e)}")
            raise e

    def generate_files(self):
        if not self.saved_template_path or self.main_window.step1.data is None:
            QMessageBox.warning(self, "Warning", "Please ensure template is saved and raw data is loaded.")
            return

        chosen_dir = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if not chosen_dir:
            return

        # Asking the user for file type and name pattern
        file_type, ok = QInputDialog.getItem(
            self,
            "Select File Type",
            "Choose file format:",
            ["Excel", "PDF"],
            0,
            False
        )
        if not ok:
            return

        name_pattern, ok = QInputDialog.getText(
            self,
            "File Naming Pattern",
            "Enter file naming pattern (use {1}, {9} for column index pairs):",
            text="{1}_{9}_gen"
        )
        if not ok:
            return

        try:
            raw_data = self.main_window.step1.data

            output_dir = os.path.join(chosen_dir, os.path.basename(self.template_path).replace('.xlsx', '_filled'))
            os.makedirs(output_dir, exist_ok=True)

            total_files = len(raw_data)
            progress = QProgressDialog("Generating files...", None, 0, total_files, self)
            progress.setWindowModality(Qt.WindowModality.WindowModal)

            for idx, row in enumerate(raw_data.itertuples(index=False), start=1):
                # Generate filename from the name pattern
                file_name = name_pattern
                for col_idx, col_value in enumerate(row, start=1):  # Adjusting for 1-based index
                    file_name = file_name.replace(f"{{{col_idx}}}", str(col_value))

                # if file_type == "Excel":
                    # Generate Excel file
                new_wb = load_workbook(self.saved_template_path)
                new_sheet = new_wb.active

                for (row_idx, col_idx), placeholder_info in self.placeholders.items():
                    template_text = placeholder_info['template']
                    for col_id in placeholder_info['column_ids']:
                        if col_id < len(raw_data.columns):
                            value = str(row[col_id])
                            template_text = template_text.replace(f"{{{col_id + 1}}}", value)

                    new_sheet.cell(row=row_idx, column=col_idx).value = template_text

                if file_type == "Excel":
                    # Generate Excel file
                    output_path = os.path.join(output_dir, file_name) + ".xlsx"
                    new_wb.save(output_path)

                elif file_type == "PDF":
                    pdf_path = os.path.join(output_dir, file_name) + ".pdf"
                    self.convert_excel_to_pdf(new_sheet, pdf_path)
                    # self.convert_excel_to_pdf_with_borders(new_sheet, pdf_path)

                progress.setValue(idx)

            QMessageBox.information(
                self,
                "Success",
                f"Generated {total_files} files in folder: {output_dir}"
            )

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to generate files: {str(e)}")
            raise e

    @staticmethod
    def convert_excel_to_pdf(sheet, pdf_path):
        """
        Converts a structured Excel file with complex formatting to a PDF.
        """
        try:
            fontsss = set()

            # Extract data and formatting from the Excel sheet
            data = []
            styles = []  # Store per-cell styles
            for row_idx, row in enumerate(sheet.iter_rows()):
                row_data = []
                row_styles = []
                for col_idx, cell in enumerate(row):
                    # Extract cell value
                    value = cell.value if cell.value is not None else ""
                    row_data.append(str(value))
                    fontsss.add(cell.font.name)
                    # Extract cell styles (alignment, font, etc.)
                    cell_style = {
                        "font_name": cell.font.name if cell.font else "Arial",
                        # "font_name": "Arial",
                        "font_size": cell.font.size if cell.font and cell.font.size else 10,
                        "bold": cell.font.bold if cell.font else False,
                        "italic": cell.font.italic if cell.font else False,
                        "alignment": cell.alignment.horizontal if cell.alignment else "left",
                        # "bg_color": f"#f0f0f0" if cell.fill and cell.fill.start_color.rgb else None,
                        "bg_color": f"#f0f0f0",
                    }
                    row_styles.append(cell_style)
                data.append(row_data)
                styles.append(row_styles)

            # Handle merged cells
            merged_ranges = []
            for merged_range in sheet.merged_cells.ranges:
                min_row = merged_range.min_row - 1
                min_col = merged_range.min_col - 1
                max_row = merged_range.max_row - 1
                max_col = merged_range.max_col - 1
                merged_ranges.append((min_row, min_col, max_row, max_col))

            # Create a PDF document
            pdf = SimpleDocTemplate(pdf_path, pagesize=letter)
            elements = []

            # Generate a Table with styles
            table = Table(data)
            table_style = []

            # Apply styles and handle merged cells
            for row_idx, row in enumerate(styles):
                for col_idx, cell_style in enumerate(row):
                    start_color = cell_style["bg_color"]
                    alignment = cell_style["alignment"]
                    font_name = cell_style["font_name"]
                    font_size = cell_style["font_size"]

                    # Apply background color
                    if start_color:
                        if start_color != "#f0f0f0":
                            print(start_color)
                        table_style.append(
                            ("BACKGROUND", (col_idx, row_idx), (col_idx, row_idx), colors.HexColor(start_color)))

                    # Apply alignment
                    if alignment == "center":
                        table_style.append(("ALIGN", (col_idx, row_idx), (col_idx, row_idx), "CENTER"))
                    elif alignment == "right":
                        table_style.append(("ALIGN", (col_idx, row_idx), (col_idx, row_idx), "RIGHT"))
                    else:
                        table_style.append(("ALIGN", (col_idx, row_idx), (col_idx, row_idx), "LEFT"))

                    # Apply font styles
                    if cell_style["bold"]:
                        font_name = "Arial-Bold"
                    if cell_style["italic"]:
                        font_name += "Arial-Italic"

                    table_style.append(("FONTNAME", (col_idx, row_idx), (col_idx, row_idx), font_name))
                    table_style.append(("FONTSIZE", (col_idx, row_idx), (col_idx, row_idx), font_size))

            # Add merged cell spans
            for min_row, min_col, max_row, max_col in merged_ranges:
                table_style.append(("SPAN", (min_col, min_row), (max_col, max_row)))

            # Apply styles to the table
            table.setStyle(TableStyle(table_style))
            elements.append(table)
            print(fontsss)

            # Build the PDF
            pdf.build(elements)
            print(f"PDF successfully generated at: {pdf_path}")

        except Exception as e:
            print(f"Error converting Excel to PDF: {e}")
            raise

    @staticmethod
    def convert_excel_to_pdf_with_borders(sheet, pdf_path):
        try:
            from openpyxl import load_workbook

            # Load the workbook and active sheet
            # workbook = load_workbook(excel_path)
            # sheet = workbook.active

            data = []
            table_styles = []

            # Read the Excel data and extract styles
            for row_idx, row in enumerate(sheet.iter_rows()):
                row_data = []
                for col_idx, cell in enumerate(row):
                    value = cell.value
                    row_data.append(str(value) if value is not None else "")

                    # Extract border styles
                    if cell.border:
                        border_weight = 1  # Default border thickness
                        color = colors.black  # Default border color

                        if cell.border.top.style:
                            table_styles.append(
                                ("LINEABOVE", (col_idx, row_idx), (col_idx, row_idx), border_weight, color))
                        if cell.border.bottom.style:
                            table_styles.append(
                                ("LINEBELOW", (col_idx, row_idx), (col_idx, row_idx), border_weight, color))
                        if cell.border.left.style:
                            table_styles.append(
                                ("LINEBEFORE", (col_idx, row_idx), (col_idx, row_idx), border_weight, color))
                        if cell.border.right.style:
                            table_styles.append(
                                ("LINEAFTER", (col_idx, row_idx), (col_idx, row_idx), border_weight, color))

                data.append(row_data)

            # Create the PDF document
            pdf = SimpleDocTemplate(pdf_path, pagesize=A4)
            table = Table(data)

            # Add table styles
            table.setStyle(TableStyle([
                ("FONT", (0, 0), (-1, -1), "Helvetica", 10),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),  # Default grid for cells
                *table_styles  # Custom styles for borders
            ]))

            # Build PDF
            pdf.build([table])

            print(f"PDF saved at {pdf_path}")

        except Exception as e:
            print(f"Error converting Excel to PDF: {e}")