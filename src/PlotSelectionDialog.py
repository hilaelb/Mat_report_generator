import os
from PyQt5.QtWidgets import QMainWindow, QPushButton, QLabel, QScrollArea, QWidget, QVBoxLayout, QHBoxLayout, QCheckBox, QCheckBox, QApplication,QMessageBox
from PyQt5.QtCore import Qt, QSize
from PyQt5.QtGui import QFont, QIcon
from generate_results import generate_report

class PlotSelectionWindow(QMainWindow):
    def __init__(self, file_paths, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Select Plots for Each File")
        self.showMaximized()  # Make the window full-size

        # Create a central widget and set layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        layout = QVBoxLayout(central_widget)

        font = QFont()
        font.setPointSize(14)

        self.scroll = QScrollArea(self)
        self.scroll.setWidgetResizable(True)
        self.scroll_widget = QWidget()
        self.scroll_layout = QVBoxLayout(self.scroll_widget)
        self.scroll_layout.setAlignment(Qt.AlignTop)
        self.scroll_layout.setContentsMargins(10, 10, 10, 10)
        self.scroll_layout.setSpacing(5)

        self.file_widgets = {}
        self.checkboxes = {}

        # Top layout with back button
        top_layout = QHBoxLayout()

        # Create the return button with an arrow icon
        return_button = QPushButton()
        return_button.setIcon(QIcon("back_arrow.png"))  # Use an arrow icon, replace "back_arrow.png" with the actual path
        return_button.setFixedSize(40, 40)  # Set the size of the button
        return_button.setIconSize(QSize(50, 50))  # Increase the icon size
        return_button.setStyleSheet("background: none; border: none;")  # Transparent button background
        return_button.clicked.connect(self.return_to_main_window)

        top_layout.addWidget(return_button, alignment=Qt.AlignLeft)

        layout.addLayout(top_layout)  # Add the top layout to the main layout

        # Scrollable area for file selection
        for file_path in file_paths:
            file_name_label = QLabel(os.path.basename(file_path))
            file_name_label.setFont(font)
            file_name_label.setStyleSheet("color: blue; cursor: pointer;")
            file_name_label.mousePressEvent = lambda event, fp=file_path: self.toggle_checkboxes(fp)
            self.scroll_layout.addWidget(file_name_label)

            checkbox_container, checkboxes = self.create_checkboxes(font)
            self.file_widgets[file_path] = checkbox_container
            self.checkboxes[file_path] = checkboxes

            checkbox_container.setVisible(False)
            self.scroll_layout.addWidget(checkbox_container)

        self.scroll.setWidget(self.scroll_widget)
        layout.addWidget(self.scroll)

        # Buttons at the bottom
        button_layout = QHBoxLayout()
        choose_same_button = QPushButton("Choose Same for All")
        choose_same_button.setFont(font)
        choose_same_button.clicked.connect(self.choose_same_for_all)
        button_layout.addWidget(choose_same_button)

        ok_button = QPushButton("OK")
        ok_button.setFont(font)
        ok_button.clicked.connect(lambda: self.on_ok_clicked(file_paths))
        button_layout.addWidget(ok_button)

        layout.addLayout(button_layout)

    def return_to_main_window(self):
        self.close()  # This closes the current window and returns to the main window

    def create_checkboxes(self, font):
        checkbox_container = QWidget()
        checkbox_layout = QVBoxLayout(checkbox_container)

        plot_options = [
            "Lat/Long Plot", "Add a picture of the map", "Bar30 measurements", "Jet Rpm/Time Plot",
            "Jet Voltage/Time Plot", "Jet data", "Histogram of Acceleration", "Inertial data", "Gyro/Accel data",
            "BME front measurements", "BME back measurements", "Motor Guidance data", "Vn200 data", "Tiger Mode/Time Plot",
            "QGC Mode/Time Plot", "Servo Angle/Time Plot", "Servo Voltage/Time Plot", "Servo Temperature/Time Plot",
            "ADC measurements", "Depth/Time Plot", "Vn Histogram", "Servo Histogram", "Bar30 Histogram", "BME Histogram",
            "ADC Histogram", "Pid pitch Histogram", "Pid heading Histogram", "Jet Histogram",
            "Position & Velocity Uncertainty/Time Plot", "Fix & Num Sats/Time Plot", "Yaw & Status/Time Plot",
            "Roll & Status/Time Plot", "Depth & Pitch/Time Plot"
        ]

        checkboxes = {}
        select_all_checkbox = QCheckBox("Select All")
        select_all_checkbox.setFont(font)
        checkbox_layout.addWidget(select_all_checkbox)

        for option in plot_options:
            checkbox = QCheckBox(option)
            checkbox.setFont(font)
            checkbox_layout.addWidget(checkbox)
            checkboxes[option] = checkbox

        select_all_checkbox.stateChanged.connect(
            lambda state, checkboxes=self.checkboxes: self.toggle_select_all(state, checkboxes)
        )

        return checkbox_container, checkboxes

    def toggle_checkboxes(self, file_path):
        for fp, widget in self.file_widgets.items():
            widget.setVisible(fp == file_path)

    def toggle_select_all(self, state, checkboxes):
        for checkbox in checkboxes.values():
            checkbox.setChecked(state == Qt.Checked)

    def choose_same_for_all(self):
        first_file_checkboxes = list(self.checkboxes.values())[0]
        first_file_states = {k: cb.isChecked() for k, cb in first_file_checkboxes.items()}

        for checkboxes in self.checkboxes.values():
            for key, checkbox in checkboxes.items():
                checkbox.setChecked(first_file_states[key])

    def get_selected_plots(self):
        selected_plots = {}
        for file_path, checkboxes in self.checkboxes.items():
            selected_plots[file_path] = {k: cb.isChecked() for k, cb in checkboxes.items()}
        return selected_plots

    def on_ok_clicked(self,file_paths):
        selected_plots = self.get_selected_plots()

        # Check if no plots are selected
        if not any(any(plots.values()) for plots in selected_plots.values()):
            # Show a message box to inform the user to select at least one plot
            QMessageBox.warning(self, "No Plots Selected",
                                "Please select at least one plot before proceeding.")
            return  # Do not proceed further if no plots are selected

        generate_report(selected_plots, file_paths)
        self.close()