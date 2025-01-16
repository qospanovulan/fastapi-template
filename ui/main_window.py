from PyQt6.QtWidgets import QMainWindow, QWidget, QStackedWidget, QHBoxLayout, QPushButton, QVBoxLayout

from ui.raw_data_window import Step1
from ui.template_window import Step2


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Processing Tool")
        self.setGeometry(100, 100, 800, 600)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        # Stacked widget to hold different steps
        self.stacked_widget = QStackedWidget()

        # Create steps in correct order
        self.step1 = Step1(self)  # Pass main window reference
        self.step2 = Step2(self)  # Pass main window reference

        # Add widgets in the desired order
        self.stacked_widget.addWidget(self.step1)
        self.stacked_widget.addWidget(self.step2)

        # Navigation buttons
        self.nav_layout = QHBoxLayout()
        self.btn_prev = QPushButton("Prev")
        # self.btn_prev.setIcon(QIcon('icons/prev.png'))  # Add icon
        self.btn_prev.setEnabled(True)
        self.btn_prev.clicked.connect(self.go_to_previous_step)

        self.btn_next = QPushButton("Next")
        self.btn_next.setEnabled(True)
        self.btn_next.clicked.connect(self.go_to_next_step)

        self.nav_layout.addWidget(self.btn_prev)
        self.nav_layout.addWidget(self.btn_next)

        # Main layout
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