import sys
import os
import json
import re
from datetime import datetime
import folium
import time
import docx.opc.constants
import scipy.io
from pprint import pprint
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Inches,Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from PIL import Image
import numpy as np
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont, QIcon
from PyQt5.QtWidgets import QApplication,QMainWindow, QDialog, QFileDialog, QVBoxLayout, QCheckBox, QPushButton, QLabel, QScrollArea, QWidget, QHBoxLayout,QListWidget
from PlotSelectionDialog import PlotSelectionWindow
from plots import create_plot, create_kml_from_waypoints, kml_to_gmplot_image, create_histogram, \
    create_partitioned_plot, create_mixed_plot, plot_tiger_modes, create_3d_plot



class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Tiger Reports Generator")

        # Create a main widget to contain the layout
        main_widget = QWidget()
        self.setCentralWidget(main_widget)

        # Create a vertical layout
        main_layout = QVBoxLayout()

        # Create a label and increase the font size
        label = QLabel("Here we are going to write the description of the application")
        label.setAlignment(Qt.AlignCenter)
        font = QFont()
        font.setPointSize(14)  # Increase the font size
        label.setFont(font)

        # Add the label to the layout
        main_layout.addWidget(label)

        # Create buttons
        convert_button = QPushButton("Convert Rosbag to Mat file")
        report_button = QPushButton("Generate Tiger Report from mat files")
        report_button.clicked.connect(self.report_button_clicked)
        excel_button = QPushButton("Create Excel from mat files")

        # Add buttons to the layout
        main_layout.addWidget(convert_button)
        main_layout.addWidget(report_button)
        main_layout.addWidget(excel_button)

        # Set the layout on the main widget
        main_widget.setLayout(main_layout)

    def closeEvent(self, event):
        QApplication.quit()


    def report_button_clicked(self, s):
        try :
            file_paths, _ = QFileDialog.getOpenFileNames(None, "Select .mat Files", "", "MAT files (*.mat)")

            if not file_paths:
                return
            window = PlotSelectionWindow(file_paths,self)


        except Exception as e:
            print(f"An error occurred: {e}")





