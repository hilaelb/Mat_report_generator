import os
import json
import re
from datetime import datetime
import folium
import time

from PyQt5.QtWidgets import QApplication, QDialog, QFileDialog, QVBoxLayout, QCheckBox, QPushButton, QLabel, QScrollArea, QWidget, QHBoxLayout,QListWidget
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont
class PlotSelectionDialog(QDialog):
    def __init__(self, file_paths):
        super().__init__()
        self.setWindowTitle("Select Plots for Each File")
        self.setGeometry(100, 100, 800, 600)
        layout = QVBoxLayout(self)

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

        button_layout = QHBoxLayout()
        choose_same_button = QPushButton("Choose Same for All")
        choose_same_button.setFont(font)
        choose_same_button.clicked.connect(self.choose_same_for_all)
        button_layout.addWidget(choose_same_button)

        ok_button = QPushButton("OK")
        ok_button.setFont(font)
        ok_button.clicked.connect(self.accept)
        button_layout.addWidget(ok_button)

        layout.addLayout(button_layout)
        self.setLayout(layout)

    def create_checkboxes(self, font):
        checkbox_container = QWidget()
        checkbox_layout = QVBoxLayout(checkbox_container)

        plot_options = [

            "Lat/Long Plot",
            "Add a picture of the map",
            "Bar30 measurements",
            "Jet Rpm/Time Plot",
            "Jet Voltage/Time Plot",
            "Jet data",
            "Histogram of Acceleration",
            "Inertial data",
            "Gyro/Accel data",
            "BME front measurements",
            "BME back measurements",
            "Motor Guidance data",
            "Vn200 data",
            "Tiger Mode/Time Plot",
            "QGC Mode/Time Plot",
            "Servo Angle/Time Plot",
            "Servo Voltage/Time Plot",
            "Servo Temperature/Time Plot",
            "ADC measurements",
            "Depth/Time Plot",
            "Vn Histogram",
            "Servo Histogram",
            "Bar30 Histogram",
            "BME Histogram",
            "ADC Histogram",
            "Pid pitch Histogram",
            "Pid heading Histogram",
            "Jet Histogram",
            "Position & Velocity Uncertainty/Time Plot",
            "Fix & Num Sats/Time Plot",
            "Yaw & Status/Time Plot",
            "Roll & Status/Time Plot",
            "Depth & Pitch/Time Plot"




            # Add more plot options here
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
            lambda state, checkboxes=checkboxes: self.toggle_select_all(state, checkboxes)
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