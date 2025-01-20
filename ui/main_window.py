from PyQt6.QtCore import Qt
from PyQt6.QtGui import QIcon, QPixmap
from PyQt6.QtWidgets import QMainWindow, QWidget, QStackedWidget, QHBoxLayout, QPushButton, QVBoxLayout, QLabel

from ui.raw_data_window import Step1
from ui.template_window import Step2


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("TemplateFiller")
        self.setGeometry(100, 100, 800, 600)

        self.setWindowIcon(QIcon("ui/logo.png"))

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        self.stacked_widget = QStackedWidget()

        self.step1 = Step1(self)
        self.step2 = Step2(self)

        self.stacked_widget.addWidget(self.step1)
        self.stacked_widget.addWidget(self.step2)

        self.nav_layout = QHBoxLayout()
        self.btn_prev = QPushButton("Пред.")
        self.btn_prev.setEnabled(True)
        self.btn_prev.clicked.connect(self.go_to_previous_step)

        self.btn_next = QPushButton("След.")
        self.btn_next.setEnabled(True)
        self.btn_next.clicked.connect(self.go_to_next_step)

        self.nav_layout.addWidget(self.btn_prev)
        self.nav_layout.addWidget(self.btn_next)

        main_layout = QVBoxLayout()
        main_layout.addWidget(self.stacked_widget)
        main_layout.addLayout(self.nav_layout)

        self.central_widget.setLayout(main_layout)

        self.setStyleSheet("""
                    QMainWindow {
                        background-color: #f0f0f0;
                        color: #000;
                    }
                    QWidget {
                        background-color: #f0f0f0;
                        color: #000;
                    }
                    QPushButton {
                        font-size: 16px;
                        padding: 8px 16px;
                    }
                    QPushButton:disabled {
                        background-color: #ccc;
                    }
                    QLabel {
                        font-size: 18px;
                        font-weight: bold;
                    }
                    QTableWidget {
                        gridline-color: #bbb;
                        background-color: #f0f0f0;
                    }

                    QTableWidgetItem {
                        background-color: #f0f0f0;
                    }

                """)

    def go_to_previous_step(self):
        current_index = self.stacked_widget.currentIndex()
        if current_index > 0:
            self.stacked_widget.setCurrentIndex(current_index - 1)
            self.update_navigation_buttons()

    def go_to_next_step(self):
        current_index = self.stacked_widget.currentIndex()
        if current_index < self.stacked_widget.count() - 1:
            self.stacked_widget.setCurrentIndex(current_index + 1)
            self.update_navigation_buttons()

    def update_navigation_buttons(self):
        current_index = self.stacked_widget.currentIndex()
        self.btn_prev.setEnabled(current_index > 0)
        self.btn_next.setEnabled(current_index < self.stacked_widget.count() - 1)