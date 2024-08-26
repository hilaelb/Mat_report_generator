import os
import json
import re
from datetime import datetime
import folium
import time


import docx.opc.constants
import scipy.io
from pprint import pprint
from PyQt5.QtWidgets import QApplication, QDialog, QFileDialog, QVBoxLayout, QCheckBox, QPushButton, QLabel, QScrollArea, QWidget, QHBoxLayout,QListWidget
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont
import matplotlib.pyplot as plt
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Inches,Pt
import numpy as np
import gmplot
from simplekml import Kml, LookAt
from selenium import webdriver

import folium



def create_kml_from_waypoints(mat_file_path, table, lat_key, lon_key, output_path):
    # Load the .mat file
    data = scipy.io.loadmat(mat_file_path)

    # Extract the waypoint data
    field_data = data[table]

    # Extract the latitude and longitude arrays
    try:
        latitudes = field_data[0][0][lat_key][0]
        longitudes = field_data[0][0][lon_key][0]
    except (IndexError, KeyError):
        print(f"Error accessing data for '{lat_key}' or '{lon_key}' in table '{table}'.")
        return

    # Convert the data to numpy arrays and flatten them
    latitudes = np.array(latitudes).flatten()
    longitudes = np.array(longitudes).flatten()

    # Check if the data is valid
    if latitudes.size == 0 or longitudes.size == 0:
        print(f"No data available for '{lat_key}' or '{lon_key}' in table '{table}'.")
        return
    # Create a KML object
    kml = Kml()

    # Define the LookAt object to center the view on the first point
    lookat = LookAt(latitude=latitudes[0], longitude=longitudes[0], altitude=0,
                    range=500, tilt=0, heading=0)
    kml.document.lookat = lookat

    for i in range(len(latitudes)):
        if i==0:
            point = kml.newpoint(name=f'Start Point', coords=[(longitudes[i], latitudes[i])])
            point.coords = [(longitudes[i], latitudes[i])]

        elif i==len(latitudes)-1:
            point = kml.newpoint(name=f'End Point', coords=[(longitudes[i], latitudes[i])])

        else:
            point = kml.newpoint(name=f'Point {i + 1}', coords=[(longitudes[i], latitudes[i])])

        point.coords = [(longitudes[i], latitudes[i])]
        point.visibility = 1  # Make the point visible by default

        # Create a linestring for the path with red color
        line = kml.newlinestring(name="AUV Route")
        line.coords = list(zip(longitudes, latitudes))
        line.style.linestyle.color = 'ff0000ff'  # Red color in KML (aabbggrr format)
        line.style.linestyle.width = 4  # Width of the line
        line.visibility = 1  # Make the line visible by default

    # Save the KML to a file
    kml.save(output_path)

def plot_waypoints_on_map(mat_file_path, table, lat_key, lon_key, output_html):
    # Load the .mat file
    data = scipy.io.loadmat(mat_file_path)

    # Extract the waypoint data
    field_data = data[table]

    # Extract the latitude and longitude arrays
    try:
        latitudes = field_data[0][0][lat_key][0]
        longitudes = field_data[0][0][lon_key][0]
    except (IndexError, KeyError):
        print(f"Error accessing data for '{lat_key}' or '{lon_key}' in table '{table}'.")
        return

    # Convert the data to numpy arrays and flatten them
    latitudes = np.array(latitudes).flatten()
    longitudes = np.array(longitudes).flatten()

    # Check if the data is valid
    if latitudes.size == 0 or longitudes.size == 0:
        print(f"No data available for '{lat_key}' or '{lon_key}' in table '{table}'.")
        return

    # Create a Google Map plotter object centered around the first point
    gmap = gmplot.GoogleMapPlotter(latitudes[0], longitudes[0], 10)  # Adjust the zoom level as needed

    # Scatter points on the map
    gmap.scatter(latitudes, longitudes, color='red', size=50, marker=True)

    # Draw a line connecting the waypoints
    gmap.plot(latitudes, longitudes, color='blue', edge_width=2.5)

    # Save the map to an HTML file
    gmap.draw(output_html)


def load_units_config(config_file_path):
    with open(config_file_path, 'r') as file:
        return json.load(file)


def kml_to_gmplot_image(mat_file_path, table, lat_key, lon_key, output_image_path, map_center=None, zoom_level=15):
    # Load the .mat file
    data = scipy.io.loadmat(mat_file_path)

    # Extract the waypoint data
    field_data = data[table]

    # Extract the latitude and longitude arrays
    try:
        latitudes = field_data[0][0][lat_key][0]
        longitudes = field_data[0][0][lon_key][0]
    except (IndexError, KeyError):
        print(f"Error accessing data for '{lat_key}' or '{lon_key}' in table '{table}'.")
        return

    # Convert the data to numpy arrays and flatten them
    latitudes = np.array(latitudes).flatten()
    longitudes = np.array(longitudes).flatten()

    # Check if the data is valid
    if latitudes.size == 0 or longitudes.size == 0:
        print(f"No data available for '{lat_key}' or '{lon_key}' in table '{table}'.")
        return

    # Create a map centered around the average latitude and longitude
    map_center = [latitudes.mean(), longitudes.mean()]
    m = folium.Map(location=map_center, zoom_start=zoom_level)

    # Add points to the map with labels
    for i, (lat, lon) in enumerate(zip(latitudes, longitudes)):
        # Label with latitude and longitude
        label = f"({lat:.4f}, {lon:.4f})"
        # Special labels for the first and last points
        if i == 0:
            label = "Start: " + label
        elif i == len(latitudes) - 1:
            label = "End: " + label

        folium.Marker(location =[lat, lon], popup=label).add_to(m)

    # Draw lines between points
    folium.PolyLine(locations=list(zip(latitudes, longitudes)), color='blue').add_to(m)

    # Save map to an HTML file
    map_html_path = 'temp.html'
    m.save(map_html_path)

    # # Determine map center if not provided
    # if map_center is None:
    #     map_center = (latitudes[0], longitudes[0])
    # # Create a gmplot map object
    # gmap = gmplot.GoogleMapPlotter(map_center[0], map_center[1], zoom_level)
    #
    # # Plot the points
    # gmap.plot(latitudes, longitudes, 'blue', edge_width=2.5)
    #
    # # Save the map in an HTML file
    # map_html_path = 'temp_map.html'
    # gmap.draw(map_html_path)

    # Use Selenium to open the HTML file and take a screenshot
    options = webdriver.ChromeOptions()
    #options.add_argument('--headless')
    options.add_argument('--disable-gpu')
    options.add_argument('--no-sandbox')
    driver = webdriver.Chrome(options=options)
    # Use absolute path for file URL
    file_url = f'file:///{os.path.abspath(map_html_path)}'
    print(f"Attempting to open file: {file_url}")
    driver.get(file_url)

    # Adjust window size to capture everything
    driver.set_window_size(1024, 768)
    time.sleep(2)  # Wait for the map to render

    # Take a screenshot and save it
    driver.save_screenshot(output_image_path)
    driver.quit()
def create_3d_plot(output_plot_path,dict_value_items, units_config_path="units_config.json"):


    # Load units configuration
    units_config = load_units_config(units_config_path)
    axis_x = axis_y1 = axis_y2 = None
    x_unit = y1_unit = y2_unit = "unknown unit"
    x_label = y1_label = y2_label = "unknown"

    for key, values in dict_value_items.items():
        if key == 'timestamp':
            # Convert time from microseconds to minutes
            axis_x = np.array(values).flatten() / 1e6 / 60
            # Normalize the time to start from 0 minutes
            axis_x = axis_x - axis_x[0]
            x_label = 'time'
            x_unit = units_config.get(x_label, "unknown unit")

        elif key == 'latitude':
            axis_y1 = np.array(values).flatten()
            y1_label = 'latitude'
            y1_unit = units_config.get(y1_label, "unknown unit")

        elif key == 'longitude':
            axis_y2 = np.array(values).flatten()
            y2_label = 'longitude'
            y2_unit = units_config.get(y2_label, "unknown unit")

        elif key == 'altitude':
            axis_y2 = np.array(values).flatten()
            y2_label = 'depth'
            y2_unit = units_config.get(y2_label, "unknown unit")

    if axis_x is None or axis_y1 is None or axis_y2 is None:
        print("Error: Missing necessary data for plotting.")
        return

    fig, ax1 = plt.subplots()

    color = 'tab:red'
    ax1.plot(axis_x, axis_y1, color=color)
    ax1.set_xlabel(f'{x_label} [{x_unit}]')
    ax1.set_ylabel(f'{y1_label} [{y1_unit}]', color=color)
    ax1.tick_params(axis='y', labelcolor=color)

    ax2 = ax1.twinx()  # instantiate a second axes that shares the same x-axis
    color = 'tab:blue'
    ax2.plot(axis_x, axis_y2, color=color)
    ax2.set_ylabel(f'{y2_label} [{y2_unit}]', color=color)
    ax2.tick_params(axis='y', labelcolor=color)

    #plt.legend()
    fig.tight_layout()  # avoid overlap
    plt.savefig(output_plot_path)
    plt.close()

def create_plot(mat_file_path, output_plot_path, table, x, y, units_config_path="units_config.json"):
    data = scipy.io.loadmat(mat_file_path)
    pprint(data)

    if table not in data:
        print(f"Table '{table}' not found in {mat_file_path}. Skipping...")
        return

    field_data = data[table]

    # Check if the field_data array is empty
    if field_data.size == 0:
        print(f"No data found in table '{table}'. Skipping...")
        return

    try:
        axis_x = field_data[0][0][x][0]
        axis_y = field_data[0][0][y][0]
    except IndexError:
        print(f"Error accessing data for '{x}' or '{y}' in table '{table}'. Skipping...")
        return
    except KeyError:
        print(f"'{x}' or '{y}' not found in table '{table}'. Skipping...")
        return

    if axis_x.size == 0 or axis_y.size == 0:
        print(f"No data available for '{x}' or '{y}' in table '{table}'. Skipping...")
        return

    # Load units configuration
    units_config = load_units_config(units_config_path)


    if x == 'timestamp':
        # Convert time from microseconds to minutes
        axis_x = np.array(axis_x).flatten() / 1e6 / 60
        # Normalize the time to start from 0 minutes
        axis_x = axis_x - axis_x[0]
        x = 'time'

    else:
        axis_x = np.array(axis_x).flatten()
    axis_y = np.array(axis_y).flatten()

    # Get the units for x and y
    x_unit = units_config.get(x, "unknown unit")
    y_unit = units_config.get(y, "unknown unit")

    plt.figure(figsize=(10, 6))
    plt.plot(axis_x, axis_y, label=f'{y} over {x}')
    plt.xlabel(f'{x} [{x_unit}]')
    plt.ylabel(f'{y} [{y_unit}]')


    plt.legend()
    plt.savefig(output_plot_path)
    plt.close()



def extract_datetime_from_path(plot_path):
    # Define a regex pattern to match the date and time in the format: YYYY_MM_DD-HH_MM_SS
    pattern = r'(\d{4}_\d{2}_\d{2})-(\d{2}_\d{2}_\d{2})'

    # Search the plot_path for the pattern
    match = re.search(pattern, plot_path)

    if match:
        # Extract the date and time strings
        date_str = match.group(1)
        time_str = match.group(2)

        # Convert to datetime object
        datetime_obj = datetime.strptime(f'{date_str} {time_str}', '%Y_%m_%d %H_%M_%S')

        # Format to desired output
        formatted_datetime = datetime_obj.strftime('%A, %m/%d/%y, %H:%M:%S')
        return formatted_datetime
    else:
        return None


def create_word_document(plot_paths, doc_file_path,map_file_paths):
    doc = Document()

    # Add the cover page
    doc.add_heading('Experimental Report of Tiger', level=1).alignment = 1  # Centered heading

    # Add a page break for the table of contents
    doc.add_page_break()

    last_file_date = None
    experiment_counter = 1

    for index, (plot_path, description) in enumerate(plot_paths):
        # Extract the filename to check for changes
        current_file_date = extract_datetime_from_path(plot_path)

        if current_file_date != last_file_date:
            if index > 0:
                doc.add_page_break()
            doc.add_heading(f'Experiment {experiment_counter}:', level=1)
            doc.add_paragraph('Location:')
            doc.add_paragraph(f'Date and time: {current_file_date}')
            doc.add_paragraph('Comments:')
            last_file_date = current_file_date
            experiment_counter += 1

        try:
            # Add the plot image
            doc.add_picture(plot_path, width=Inches(6))
            doc.add_paragraph(description)

            # Check if there's a corresponding map file for this plot
            base_name = os.path.basename(plot_path).replace('.png', '')
            map_file_path = next((path for path in map_file_paths if base_name in os.path.basename(path)), None)

            if map_file_path:
                #TODO: add a picture of the map to the document
                doc.add_picture(map_file_path, width=Inches(6))




        except Exception as e:
            print(f"Error adding plot {plot_path} to the document: {e}")
            continue


    doc.save(doc_file_path)

def clear_folder(plots_folder):
    for file_name in os.listdir(plots_folder):
        file_path = os.path.join(plots_folder, file_name)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
        except Exception as e:
            print(f'Failed to delete {file_path}. Reason: {e}')


class PlotSelectionDialog(QDialog):
    def __init__(self, file_paths):
        super().__init__()
        self.setWindowTitle("Select Plots for Each File")
        self.setGeometry(100, 100, 800, 600)  # Set larger window size
        layout = QVBoxLayout(self)

        font = QFont()
        font.setPointSize(14)  # Larger font size for the dialog

        self.scroll = QScrollArea(self)
        self.scroll.setWidgetResizable(True)
        self.scroll_widget = QWidget()
        self.scroll_layout = QVBoxLayout(self.scroll_widget)

        # Align the layout to the top of the scroll area
        self.scroll_layout.setAlignment(Qt.AlignTop)
        # Reduce spacing between each file
        self.scroll_layout.setContentsMargins(10, 10, 10, 10)  # Set margins around the layout
        self.scroll_layout.setSpacing(5)  # Reduce spacing between widgets

        self.file_widgets = {}
        self.checkboxes = {}

        for file_path in file_paths:
            # File name label
            file_name_label = QLabel(os.path.basename(file_path))
            file_name_label.setFont(font)  # Larger font size for file names
            file_name_label.setStyleSheet("color: blue; cursor: pointer;")
            file_name_label.mousePressEvent = lambda event, fp=file_path: self.toggle_checkboxes(fp)
            self.scroll_layout.addWidget(file_name_label)

            # Create checkbox container widget
            checkbox_container = QWidget()
            checkbox_layout = QVBoxLayout(checkbox_container)

            depth_time_checkbox = QCheckBox("Depth/Time Plot")
            lat_long_checkbox = QCheckBox("Lat/Long Plot")
            lat_long_time_checkbox = QCheckBox("Lat-Long/Time Plot")
            lat_depth_time_checkbox = QCheckBox("Lat-Depth/Time Plot")
            select_all_checkbox = QCheckBox("Select All")

            # Larger font size for checkboxes
            depth_time_checkbox.setFont(font)
            lat_long_checkbox.setFont(font)
            lat_long_time_checkbox.setFont(font)
            lat_depth_time_checkbox.setFont(font)
            select_all_checkbox.setFont(font)

            select_all_checkbox.stateChanged.connect(
                lambda state, dt=depth_time_checkbox, ll=lat_long_checkbox, llt= lat_long_time_checkbox, ldt = lat_depth_time_checkbox: self.toggle_select_all(state, dt, ll,llt,ldt)

            )

            checkbox_layout.addWidget(depth_time_checkbox)
            checkbox_layout.addWidget(lat_long_checkbox)
            checkbox_layout.addWidget(lat_long_time_checkbox)
            checkbox_layout.addWidget(lat_depth_time_checkbox)
            checkbox_layout.addWidget(select_all_checkbox)

            # Store widgets and hide the checkbox container initially
            self.file_widgets[file_path] = checkbox_container
            self.checkboxes[file_path] = {
                "depth_time": depth_time_checkbox,
                "lat_long": lat_long_checkbox,
                "lat_long_time": lat_long_time_checkbox,
                "lat_depth_time": lat_depth_time_checkbox,
                "select_all": select_all_checkbox
            }
            checkbox_container.setVisible(False)
            self.scroll_layout.addWidget(checkbox_container)

        self.scroll.setWidget(self.scroll_widget)
        layout.addWidget(self.scroll)

        button_layout = QHBoxLayout()

        # "Choose Same for All" button
        choose_same_button = QPushButton("Choose Same for All")
        choose_same_button.setFont(font)
        choose_same_button.clicked.connect(self.choose_same_for_all)
        button_layout.addWidget(choose_same_button)

        ok_button = QPushButton("OK")
        ok_button.setFont(font)  # Larger font size for the button
        ok_button.clicked.connect(self.accept)
        button_layout.addWidget(ok_button)

        layout.addLayout(button_layout)
        self.setLayout(layout)

    def toggle_checkboxes(self, file_path):
        # Toggle visibility of the checkbox container for the selected file
        for fp, widget in self.file_widgets.items():
            widget.setVisible(fp == file_path)

    def toggle_select_all(self, state, depth_time_checkbox, lat_long_checkbox, lat_long_time_checkbox, lat_depth_time_checkbox):
        depth_time_checkbox.setChecked(state == Qt.Checked)
        lat_long_checkbox.setChecked(state == Qt.Checked)
        lat_long_time_checkbox.setChecked(state == Qt.Checked)
        lat_depth_time_checkbox.setChecked(state == Qt.Checked)

    def choose_same_for_all(self):
        # Get the settings of the first file (assuming you want to apply its settings to all)
        first_file_checkboxes = list(self.checkboxes.values())[0]
        depth_time_checked = first_file_checkboxes["depth_time"].isChecked()
        lat_long_checked = first_file_checkboxes["lat_long"].isChecked()
        lat_long_time_checked = first_file_checkboxes["lat_long_time"].isChecked()
        lat_depth_time_checked = first_file_checkboxes["lat_depth_time"].isChecked()

        # Apply these settings to all files
        for checkboxes in self.checkboxes.values():
            checkboxes["depth_time"].setChecked(depth_time_checked)
            checkboxes["lat_long"].setChecked(lat_long_checked)
            checkboxes["lat_long_time"].setChecked(lat_long_time_checked)
            checkboxes["lat_depth_time"].setChecked(lat_depth_time_checked)

    def get_selected_plots(self):
        selected_plots = {}
        for file_path, checkboxes in self.checkboxes.items():
            selected_plots[file_path] = {
                "depth_time": checkboxes["depth_time"].isChecked(),
                "lat_long": checkboxes["lat_long"].isChecked(),
                "lat_long_time": checkboxes["lat_long_time"].isChecked(),
                "lat_depth_time": checkboxes["lat_depth_time"].isChecked()

            }
        return selected_plots



def process_files():
    app = QApplication([])
    file_paths, _ = QFileDialog.getOpenFileNames(None, "Select .mat Files", "", "MAT files (*.mat)")

    if not file_paths:
        return

    dialog = PlotSelectionDialog(file_paths)
    if dialog.exec_() != QDialog.Accepted:
        return

    selected_plots = dialog.get_selected_plots()

    desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
    output_folder = os.path.join(desktop_path, 'MAT_Files_Report')
    plots_folder = os.path.join(output_folder, 'plots')
    maps_folder = os.path.join(output_folder, 'maps')
    kmls_folder = os.path.join(output_folder, 'kmls')
    os.makedirs(maps_folder, exist_ok=True)
    os.makedirs(plots_folder, exist_ok=True)
    os.makedirs(output_folder, exist_ok=True)
    os.makedirs(kmls_folder, exist_ok=True)
    clear_folder(plots_folder)
    clear_folder(maps_folder)
    clear_folder(kmls_folder)

    plot_paths = []
    map_paths = []
    kmls = []

    for mat_file_path in file_paths:
        base_name = os.path.basename(mat_file_path).replace('.mat', '')

        if selected_plots[mat_file_path]["depth_time"]:
            plot_file_name_1 = f'{base_name}_ba30_depth_vs_time.png'
            plot_file_path_1 = os.path.join(plots_folder, plot_file_name_1)
            create_plot(mat_file_path, plot_file_path_1, 'ba30', 'timestamp', 'depth')
            plot_paths.append((plot_file_path_1, 'This plot shows the depth as a function of time.'))

        if selected_plots[mat_file_path]["lat_long"]:
            plot_file_name_2 = f'{base_name}_waypoint_rosbag_lat_vs_long.png'
            plot_file_path_2 = os.path.join(plots_folder, plot_file_name_2)
            create_plot(mat_file_path, plot_file_path_2, 'waypoint_rosbag', 'latitude', 'longitude')
            plot_paths.append((plot_file_path_2, 'This plot shows the latitude as a function of longitude.'))

            kml_path = os.path.join(kmls_folder, f'{base_name}_waypoint_rosbag_lat_vs_long.kml')
            create_kml_from_waypoints(mat_file_path, 'waypoint_rosbag', 'latitude', 'longitude', kml_path)
            kmls.append(kml_path)
            map_file_path = os.path.join(maps_folder, f'{base_name}_waypoint_rosbag_lat_vs_long.png')
            kml_to_gmplot_image(mat_file_path, 'waypoint_rosbag', 'latitude', 'longitude', map_file_path)
            map_paths.append(map_file_path)

        if selected_plots[mat_file_path]["lat_long_time"]:
            plot_file_name_3 = f'{base_name}_waypoint_rosbag_lat_long_vs_time.png'
            plot_file_path_3 = os.path.join(plots_folder, plot_file_name_3)
            dict = {'waypoint_rosbag': ['latitude', 'longitude'],
                    'waypoint_reached': ['timestamp']}
            dict_value_items = extract_values_from_data(mat_file_path, dict)
            create_3d_plot(plot_file_path_3, dict_value_items)
            plot_paths.append((plot_file_path_3,''))

        if selected_plots[mat_file_path]["lat_depth_time"]:
            plot_file_name_4 = f'{base_name}_waypoint_rosbag_lat_depth_vs_time.png'
            plot_file_path_4 = os.path.join(plots_folder, plot_file_name_4)
            dict = {'waypoint_rosbag': ['latitude', 'altitude'],
                    'waypoint_reached': ['timestamp']}
            dict_value_items = extract_values_from_data(mat_file_path, dict)
            create_3d_plot(plot_file_path_4, dict_value_items)
            plot_paths.append((plot_file_path_4,''))


    doc_file_path = os.path.join(output_folder, 'report.docx')
    create_word_document(plot_paths, doc_file_path,map_paths)


    print(f"Processing complete. Files saved to {output_folder}")

def extract_values_from_data(mat_file_path, dict):
    data = scipy.io.loadmat(mat_file_path)
    dict_value_items = {}
    for key, values in dict.items():
        if key not in data:
            print(f"Key '{key}' not found in {mat_file_path}. Skipping...")
            return

        field_data = data[key]
        # Check if the field_data array is empty
        if field_data.size == 0:
            print(f"No data found in table '{key}'. Skipping...")
            return

        for value in values:
            try:
                items = field_data[0][0][value][0]
                dict_value_items[value] = items

            except (IndexError, KeyError):
                print(f"Error accessing data for '{value}' in table '{key}'. Skipping...")
                return





    return dict_value_items









if __name__ == "__main__":
    process_files()
