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
from matplotlib.ticker import FuncFormatter
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Inches,Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from PIL import Image
import numpy as np
import gmplot
from scipy.ndimage import label
from simplekml import Kml, LookAt
from selenium import webdriver

import folium


def format_decimal(value, pos):
    return f'{value:.4f}'



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
    driver.set_window_size(800, 600)
    time.sleep(2)  # Wait for the map to render

    # Take a screenshot and save it
    driver.save_screenshot(output_image_path)
    driver.quit()


def determine_figure_size(x_length, y_length, base_size=(12, 4), scale_factor=0.1):
    """
    Determines the size of the figure based on the length of the X and Y axes data.

    :param x_length: Length of the X-axis data.
    :param y_length: Length of the Y-axis data.
    :param base_size: The base size of the figure (width, height) in inches.
    :param scale_factor: Scaling factor to adjust the figure size based on the data length.
    :return: Tuple (width, height) for the figure size.
    """
    # Calculate scaling
    width = base_size[0] + scale_factor * (x_length ** 0.5)
    height = base_size[1] + scale_factor * (y_length ** 0.5)

    return width, height

def create_acceleration_histogram(table, output_plot_path):
    if table is None:
        print("Error: Missing table data.")
        return
    num_columns = len(table)

    # Create a figure with subplots for each column
    fig, axes = plt.subplots(num_columns, 1, figsize=(11, 5 * num_columns))

    if num_columns == 1:
        axes = [axes]  # Ensure axes is a list even if there's only one subplot

    for i, (key, data) in enumerate(table.items()):
        data = np.array(data)

        # Calculate statistics
        mean_val = np.mean(data)
        max_val = np.max(data)
        std_dev = np.std(data)

        # Create the histogram
        axes[i].hist(data, bins=30, alpha=0.75, color='blue', edgecolor='black')

        # Add statistical data to the plot
        axes[i].axvline(mean_val, color='red', linestyle='dashed', linewidth=2, label=f'Mean: {mean_val:.2f}')
        axes[i].axvline(max_val, color='green', linestyle='dashed', linewidth=2, label=f'Max: {max_val:.2f}')
        axes[i].axvline(mean_val - std_dev, color='orange', linestyle='dashed', linewidth=2,
                        label=f'Std Dev: {std_dev:.2f}')
        axes[i].axvline(mean_val + std_dev, color='orange', linestyle='dashed', linewidth=2)

        axes[i].set_xlabel(f'{key}')
        axes[i].set_ylabel('Frequency')
        axes[i].set_title(f'Histogram of {key} with Stats')
        axes[i].legend()

    # Adjust layout and save the figure
    plt.tight_layout()
    plt.savefig(output_plot_path)
    plt.close()

def create_partitioned_plot(output_plot_path, tables,title='', units_config_path="units_config.json",max_height=10, max_width=12):
    units_config = load_units_config(units_config_path)
    if tables is None:
        print("Error: Missing tables data.")
        return
    keys = list(tables.keys())

    x_key = keys[0]
    if x_key == 'timestamp':
        axis_x = np.array(tables[x_key]).flatten() / 1e6 / 60
        axis_x = axis_x - axis_x[0]
        x_label = 'time'

    else:
        axis_x = np.array(tables[x_key]).flatten()
        x_label = x_key

    x_unit = units_config.get(x_label, "unknown unit")

    num_plots = len(keys) - 1

    # Calculate the number of data points
    data_points = len(axis_x)

    # Calculate the height dynamically based on the number of data points
    # dynamic_height = min(max_height, 0.05 * data_points + 2)
    # fig_height = max(3 * num_plots, dynamic_height)
    fig_width, fig_height = determine_figure_size(len(axis_x), len(tables[keys[1]]),(10,12))



    fig, axes = plt.subplots(num_plots, 1, figsize=(fig_width, fig_height), sharex=True)
    fig.suptitle(title,fontsize=16)

    if num_plots == 1:
        axes = [axes]  # Ensure axes is always a list even if there's only one subplot



    for i, key in enumerate(keys[1:]):
        axis_y = np.array(tables[key]).flatten()
        y_label = key
        y_unit = units_config.get(y_label, "unknown unit")

        axes[i].plot(axis_x, axis_y, label=f'{y_label} [{y_unit}]')
        axes[i].set_ylabel(f'{y_label} [{y_unit}]')
        axes[i].grid(True)
        axes[i].legend(loc='upper right')

    # Set the x-axis label on the last subplot
    axes[-1].set_xlabel(f'{x_label} [{x_unit}]')

    fig.tight_layout(rect=[0, 0.03, 1, 0.95])
    plt.savefig(output_plot_path)
    plt.close()



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

    fig_width, fig_height = determine_figure_size(len(axis_x), len(axis_y1))
    fig, ax1 = plt.subplots(figsize=(fig_width, fig_height))

    color = 'tab:red'
    ax1.plot(axis_x, axis_y1, color=color)
    ax1.set_xlabel(f'{x_label} [{x_unit}]')
    ax1.set_ylabel(f'{y1_label} [{y1_unit}]', color=color)
    ax1.tick_params(axis='y', labelcolor=color)
    ax1.yaxis.set_major_formatter(FuncFormatter(format_decimal))  # Format y-axis labels

    ax2 = ax1.twinx()  # instantiate a second axes that shares the same x-axis
    color = 'tab:blue'
    ax2.plot(axis_x, axis_y2, color=color)
    ax2.set_ylabel(f'{y2_label} [{y2_unit}]', color=color)
    ax2.tick_params(axis='y', labelcolor=color)
    ax2.xaxis.set_major_formatter(FuncFormatter(format_decimal))  # Format x-axis labels
    ax2.yaxis.set_major_formatter(FuncFormatter(format_decimal))  # Format y-axis labels

    plt.plot(label=f'{y1_label} and {y2_label} over {x_label}')
    plt.title(f'{y1_label} and {y2_label} over {x_label}')
    plt.savefig(output_plot_path)
    plt.close()

def plot_tiger_modes(tables, output_plot_path, units_config_path="units_config.json"):
    # Define state labels and values
    state_labels = {}

    if list(tables.keys())[0] == 'tiger_mode':
        state_labels = {
            1: 'IDLE',
            2: 'JOYSTICK',
            4: 'FIXED_MANUVER',
            8: 'HOLD_POSITION',
            16: 'ABORTED_ONLY_MANUAL',
            32: 'ABORTED_LPM',
            64: 'CALIBRATION',
            128: 'GUIDED_MISSION'
        }
    else:
        state_labels = {
            0: 'PRE-FLIGHT',
            28: 'PLAY',
            64: 'STOP',
            128: 'PAUSE'
        }
    units_config = load_units_config(units_config_path)
    time_data = state_data = None

    for key, values in tables.items():
        if key == 'timestamp':
            # Convert time from microseconds to minutes
            time_data = np.array(values).flatten() / 1e6 / 60
            # Normalize the time to start from 0 minutes
            time_data = time_data - time_data[0]
            x_label = 'time'
            x_unit = units_config.get(x_label, "unknown unit")

        elif key == 'tiger_mode':
            y_label = 'Tiger State'
            state_data = np.array(values).flatten()
        else:
            y_label = 'QGC State'
            state_data = np.array(values).flatten()


    # Create a list of state labels based on the state data
    state_names = [state_labels.get(state, 'UNKNOWN') for state in state_data]


    # Create a unique list of states for plotting
    unique_states = list(state_labels.values())
    state_mapping = {state: i for i, state in enumerate(unique_states)}
    numeric_states = [state_mapping[state] for state in state_names]

    # # Create the plot
    # num_data_points = len(time_data)
    # width = max(22, min(20, num_data_points / 50))  # Adjust width dynamically
    # height = 10  # Fixed height for consistency

    # Calculate figure size based on data lengths
    fig_width, fig_height = determine_figure_size(len(time_data), len(state_data))

    plt.figure(figsize=(fig_width, fig_height))
    plt.step(time_data, numeric_states, where='post', label=y_label)
    plt.yticks(range(len(unique_states)), unique_states)
    plt.xlabel(f'Time [{x_unit}]')
    plt.ylabel('State')
    plt.title(f'{y_label} Over Time')
    plt.grid(True)
    plt.legend()
    plt.savefig(output_plot_path)
    plt.close()




def create_plot(output_plot_path, tables,title='', units_config_path="units_config.json"):

    # Load units configuration
    units_config = load_units_config(units_config_path)
    keys = list(tables.keys())

    if len(keys) != 2:
        print("Error: Expected 2 keys in the tables dictionary.")
        return

    y_key = keys[0]
    x_key = keys[1]

    if x_key == 'timestamp':
        # Convert time from microseconds to minutes
        axis_x = np.array(tables[x_key]).flatten() / 1e6 / 60
        # Normalize the time to start from 0 minutes
        axis_x = axis_x - axis_x[0]
        x_label = 'time'
    else:
        axis_x = np.array(tables[x_key]).flatten()
        x_label = x_key

    if y_key == 'v_in':
        y_label = 'voltage'

    else:
        y_label = y_key
    axis_y = np.array(tables[y_key]).flatten()

    # Get the units from the configuration
    x_unit = units_config.get(x_label, "unknown unit")
    y_unit = units_config.get(y_label, "unknown unit")

    # Determine figure size based on the number of data points
    # num_data_points = len(axis_x)
    # width = max(10, min(20, num_data_points / 50))  # Adjust width dynamically
    # height = 4  # Fixed height for consistency

    fig_width, fig_height = determine_figure_size(len(axis_x), len(axis_y))
    plt.figure(figsize=(fig_width, fig_height))
    plt.plot(axis_x, axis_y, label=f'{y_label} over {x_label}')
    plt.title(title)
    plt.xlabel(f'{x_label} [{x_unit}]')
    plt.ylabel(f'{y_label} [{y_unit}]')



    # Set the format for the x and y ticks to show 4 decimal places
    plt.gca().xaxis.set_major_formatter(plt.FuncFormatter(format_decimal))
    plt.gca().yaxis.set_major_formatter(plt.FuncFormatter(format_decimal))


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
    # Add the LAR.png image and center it
    lar_paragraph = doc.add_paragraph('\n \n \n')
    lar_run = lar_paragraph.add_run()
    lar_run.add_picture('LAR.png', width=Inches(2.5), height=Inches(2.5))
    lar_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Add a page break for the table of contents
    doc.add_page_break()

    last_file_date = None
    experiment_counter = 1

    # Define maximum dimensions for the images (in inches)
    max_width = Inches(6.5)  # Slightly less than the page width
    max_height = Inches(9)  # Slightly less than the page height

    for index, (plot_path, description) in enumerate(plot_paths):
        # Extract the filename to check for changes
        current_file_date = extract_datetime_from_path(plot_path)

        if current_file_date != last_file_date and current_file_date is not None:
            if index > 0:
                doc.add_page_break()
            doc.add_heading(f'Experiment {experiment_counter}:', level=1)
            doc.add_paragraph('Location:')
            doc.add_paragraph(f'Date and time: {current_file_date}')
            doc.add_paragraph('Comments:')
            last_file_date = current_file_date
            experiment_counter += 1

        try:



            doc.add_picture(plot_path, width=Inches(6))
            if description != '':
                # doc.add_page_break()
                doc.add_paragraph(description).alignment = 1  # Centered text
            # # Open the image using PIL to get its size
            # with Image.open(plot_path) as img:
            #     width, height = img.size
            #     aspect_ratio = width / height
            #
            #     # Calculate the new dimensions
            #     if width > height:
            #         new_width = min(max_width, Inches(width / 96))  # Convert pixels to inches
            #         new_height = new_width / aspect_ratio
            #     else:
            #         new_height = min(max_height, Inches(height / 96))
            #         new_width = new_height * aspect_ratio
            #
            #     # Add the plot image with the adjusted size
            #     doc.add_picture(plot_path, width=new_width, height=new_height)


            # Check if there's a corresponding map file for this plot
            base_name = os.path.basename(plot_path).replace('.png', '')
            map_file_path = next((path for path in map_file_paths if base_name in os.path.basename(path)), None)

            if map_file_path:
                doc.add_picture(map_file_path, width=Inches(6))
                # with Image.open(map_file_path) as img:
                #     width, height = img.size
                #     aspect_ratio = width / height
                #
                #     # Calculate the new dimensions
                #     if width > height:
                #         new_width = min(max_width, Inches(width / 96))
                #         new_height = new_width / aspect_ratio
                #     else:
                #         new_height = min(max_height, Inches(height / 96))
                #         new_width = new_height * aspect_ratio
                #
                #     # Add the map image with the adjusted size
                #     doc.add_picture(map_file_path, width=new_width, height=new_height)

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
            "Lat-Long/Time Plot",
            "Lat-Depth/Time Plot",
            "Bar30 measurements",
            "Jet Rpm/Time Plot",
            "Jet Voltage/Time Plot",
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
            "Depth/Time Plot"

            # Add more plot options here
        ]

        checkboxes = {}
        for option in plot_options:
            checkbox = QCheckBox(option)
            checkbox.setFont(font)
            checkbox_layout.addWidget(checkbox)
            checkboxes[option] = checkbox

        select_all_checkbox = QCheckBox("Select All")
        select_all_checkbox.setFont(font)
        select_all_checkbox.stateChanged.connect(
            lambda state, checkboxes=checkboxes: self.toggle_select_all(state, checkboxes)
        )
        checkbox_layout.addWidget(select_all_checkbox)

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

        if selected_plots[mat_file_path]["Depth/Time Plot"]:
            plot_file_name_1 = f'{base_name}_ba30_depth_vs_time.png'
            plot_file_path_1 = os.path.join(plots_folder, plot_file_name_1)
            dict = {'ba30': ['depth', 'timestamp']}
            dict_value_items = extract_values_from_data(mat_file_path, dict)
            create_plot(plot_file_path_1, dict_value_items)
            plot_paths.append((plot_file_path_1, ''))

        if selected_plots[mat_file_path]["Lat/Long Plot"]:
            plot_file_name_2 = f'{base_name}_waypoint_rosbag_lat_vs_long.png'
            plot_file_path_2 = os.path.join(plots_folder, plot_file_name_2)
            dict = {'waypoint': ['latitude', 'longitude']}
            dict_value_items = extract_values_from_data(mat_file_path, dict)
            pprint(dict_value_items)
            create_plot(plot_file_path_2, dict_value_items)
            plot_paths.append((plot_file_path_2, ''))

            kml_path = os.path.join(kmls_folder, f'{base_name}_waypoint_rosbag_lat_vs_long.kml')
            create_kml_from_waypoints(mat_file_path, 'waypoint', 'latitude', 'longitude', kml_path)
            kmls.append(kml_path)
            map_file_path = os.path.join(maps_folder, f'{base_name}_waypoint_rosbag_lat_vs_long.png')
            kml_to_gmplot_image(mat_file_path, 'waypoint', 'latitude', 'longitude', map_file_path)
            map_paths.append(map_file_path)


        if selected_plots[mat_file_path]["Lat-Long/Time Plot"]:

            plot_file_name_3 = f'{base_name}_waypoint_rosbag_lat_long_vs_time.png'
            plot_file_path_3 = os.path.join(plots_folder, plot_file_name_3)
            dict = {'waypoint': ['latitude', 'longitude','tool_speed']}
            dict_value_items = extract_values_from_data(mat_file_path, dict)
            time = calculate_travel_times(dict_value_items)
            del dict_value_items['tool_speed']
            dict_value_items['timestamp'] = time
            create_3d_plot(plot_file_path_3, dict_value_items)
            plot_paths.append((plot_file_path_3,''))

        if selected_plots[mat_file_path]["Lat-Depth/Time Plot"]:
            plot_file_name_4 = f'{base_name}_waypoint_rosbag_lat_depth_vs_time.png'
            plot_file_path_4 = os.path.join(plots_folder, plot_file_name_4)
            dict = {'waypoint': ['latitude','longitude', 'altitude','tool_speed']}
            dict_value_items = extract_values_from_data(mat_file_path, dict)
            time = calculate_travel_times(dict_value_items)
            del dict_value_items['tool_speed']
            del dict_value_items['longitude']
            dict_value_items['timestamp'] = time
            create_3d_plot(plot_file_path_4, dict_value_items)
            plot_paths.append((plot_file_path_4,''))


        if selected_plots[mat_file_path]["Jet Rpm/Time Plot"]:
            plot_file_name_5 = f'{base_name}_jet_rpm_vs_time.png'
            plot_file_path_5 = os.path.join(plots_folder, plot_file_name_5)
            dict = {'jet': ['rpm', 'timestamp']}
            dict_value_items = extract_values_from_data(mat_file_path, dict)
            create_plot(plot_file_path_5,dict_value_items,'Jet RPM')
            plot_paths.append((plot_file_path_5, ''))

        if selected_plots[mat_file_path]["Jet Voltage/Time Plot"]:
            plot_file_name_6 = f'{base_name}_jet_voltage_vs_time.png'
            plot_file_path_6 = os.path.join(plots_folder, plot_file_name_6)
            dict = {'jet': ['v_in', 'timestamp']}
            dict_value_items = extract_values_from_data(mat_file_path, dict)
            create_plot(plot_file_path_6,dict_value_items,'Jet Voltage')
            plot_paths.append((plot_file_path_6, ''))

        if selected_plots[mat_file_path]["Servo Angle/Time Plot"]:
            plot_file_name_8 = f'{base_name}_servo_angle_vs_time.png'
            plot_file_path_8 = os.path.join(plots_folder, plot_file_name_8)
            dict = {'servo': ['timestamp', 'srv1_angle', 'srv2_angle', 'srv3_angle', 'srv4_angle']}
            dict_value_items = extract_values_from_data(mat_file_path, dict)
            create_partitioned_plot(plot_file_path_8,dict_value_items,'Servos angle')
            plot_paths.append((plot_file_path_8, ''))

        if selected_plots[mat_file_path]["Servo Voltage/Time Plot"]:
            plot_file_name_9 = f'{base_name}_servo_voltage_vs_time.png'
            plot_file_path_9 = os.path.join(plots_folder, plot_file_name_9)
            dict = {'servo': ['timestamp', 'srv1_voltage', 'srv2_voltage', 'srv3_voltage', 'srv4_voltage']}
            dict_value_items = extract_values_from_data(mat_file_path, dict)
            create_partitioned_plot(plot_file_path_9,dict_value_items,'Servos Voltage')
            plot_paths.append((plot_file_path_9, ''))

        if selected_plots[mat_file_path]["Servo Temperature/Time Plot"]:
            plot_file_name_10 = f'{base_name}_servo_temperature_vs_time.png'
            plot_file_path_10 = os.path.join(plots_folder, plot_file_name_10)
            dict = {'servo': ['timestamp', 'srv1_temperature', 'srv2_temperature', 'srv3_temperature', 'srv4_temperature']}
            dict_value_items = extract_values_from_data(mat_file_path, dict)
            create_partitioned_plot(plot_file_path_10,dict_value_items,'Servos temperature')
            plot_paths.append((plot_file_path_10, ''))

        if selected_plots[mat_file_path]["Inertial data"]:
            plot_file_name_11 = f'{base_name}_inertial_data.png'
            plot_file_path_11 = os.path.join(plots_folder, plot_file_name_11)
            dict = {'insstatus': ['timestamp', 'ins_mode', 'ins_error', 'ins_fix']}
            dict_value_items = extract_values_from_data(mat_file_path, dict)
            create_partitioned_plot(plot_file_path_11,dict_value_items,'Inertial system data')
            plot_paths.append((plot_file_path_11, ''))

        if selected_plots[mat_file_path]["Gyro/Accel data"]:
            plot_file_name_12 = f'{base_name}_gyro_data.png'
            plot_file_path_12 = os.path.join(plots_folder, plot_file_name_12)
            dict = {'gyro_accel_data': ['timestamp', 'gyro_x', 'gyro_y', 'gyro_z', 'accel_x', 'accel_y', 'accel_z']}
            dict_value_items = extract_values_from_data(mat_file_path, dict)
            create_partitioned_plot(plot_file_path_12,dict_value_items,'Gyro/Accel data')
            plot_paths.append((plot_file_path_12, ''))


        if selected_plots[mat_file_path]["BME front measurements"]:
            plot_file_name_14 = f'{base_name}_bme_front_data.png'
            plot_file_path_14 = os.path.join(plots_folder, plot_file_name_14)
            dict = {'bme_front': ['timestamp', 'temperature', 'pressure', 'humidity']}
            dict_value_items = extract_values_from_data(mat_file_path, dict)
            create_partitioned_plot(plot_file_path_14,dict_value_items,'BME front measurements')
            plot_paths.append((plot_file_path_14, ''))

        if selected_plots[mat_file_path]["BME back measurements"]:
            plot_file_name_15 = f'{base_name}_bme_back_data.png'
            plot_file_path_15 = os.path.join(plots_folder, plot_file_name_15)
            dict = {'bme_back': ['timestamp', 'temperature', 'pressure', 'humidity']}
            dict_value_items = extract_values_from_data(mat_file_path, dict)
            create_partitioned_plot(plot_file_path_15,dict_value_items,'BME back measurements')
            plot_paths.append((plot_file_path_15, ''))

        if selected_plots[mat_file_path]["Bar30 measurements"]:
            plot_file_name_16 = f'{base_name}_bar30_data.png'
            plot_file_path_16 = os.path.join(plots_folder, plot_file_name_16)
            dict = {'ba30': ['timestamp', 'temperature', 'pressure', 'depth']}
            dict_value_items = extract_values_from_data(mat_file_path, dict)
            create_partitioned_plot(plot_file_path_16,dict_value_items,'Bar30 measurements')
            plot_paths.append((plot_file_path_16, ''))

        if selected_plots[mat_file_path]["Motor Guidance data"]:
            plot_file_name_17 = f'{base_name}_motor_guidance_data.png'
            plot_file_path_17 = os.path.join(plots_folder, plot_file_name_17)
            dict = {'motor_guidance_data': ['timestamp', 'srv_1_angle', 'srv_2_angle', 'srv_4_angle', 'srv_4_angle', 'jet_rpm']}
            dict_value_items = extract_values_from_data(mat_file_path, dict)
            create_partitioned_plot(plot_file_path_17,dict_value_items,'Motor guidance data')
            plot_paths.append((plot_file_path_17, ''))

        if selected_plots[mat_file_path]["Vn200 data"]:
            plot_file_name_18 = f'{base_name}_vn200_data.png'
            plot_file_path_18 = os.path.join(plots_folder, plot_file_name_18)
            dict = {'vn': ['timestamp', 'yaw', 'pitch', 'roll', 'poslla_x', 'poslla_y']}
            dict2 = {'vn': ['timestamp','poslla_z' ,'velned_x', 'velned_y', 'velned_z', 'posu','velu']}
            dict_value_items = extract_values_from_data(mat_file_path, dict)
            dict_value_items2 = extract_values_from_data(mat_file_path, dict2)
            create_partitioned_plot(plot_file_path_18, dict_value_items, 'Vn200 data')
            create_partitioned_plot(f'{plot_file_path_18[:-4]}_2.png', dict_value_items2)
            plot_paths.append((plot_file_path_18, ''))
            plot_paths.append((f'{plot_file_path_18[:-4]}_2.png',''))


        if selected_plots[mat_file_path]["Histogram of Acceleration"]:
            plot_file_name_19 = f'{base_name}_acceleration_histogram.png'
            plot_file_path_19 = os.path.join(plots_folder, plot_file_name_19)
            dict = {'gyro_accel_data': ['accel_x', 'accel_y', 'accel_z']}
            dict_value_items = extract_values_from_data(mat_file_path, dict)
            create_acceleration_histogram(dict_value_items, plot_file_path_19)
            plot_paths.append((plot_file_path_19, 'Histogram of Acceleration'))

        if selected_plots[mat_file_path]["Tiger Mode/Time Plot"]:
            plot_file_name_7 = f'{base_name}_tiger_mode_vs_time.png'
            plot_file_path_7 = os.path.join(plots_folder, plot_file_name_7)
            dict = {'tiger': ['tiger_mode', 'timestamp']}
            dict_value_items = extract_values_from_data(mat_file_path, dict)
            plot_tiger_modes(dict_value_items, plot_file_path_7)
            plot_paths.append((plot_file_path_7, ''))

        if selected_plots[mat_file_path]["QGC Mode/Time Plot"]:
            plot_file_name_20 = f'{base_name}_qgc_mode_vs_time.png'
            plot_file_path_20 = os.path.join(plots_folder, plot_file_name_20)
            dict = {'qgc': ['timestamp','mode']}
            dict_value_items = extract_values_from_data(mat_file_path, dict)
            plot_tiger_modes(dict_value_items, plot_file_path_20)
            plot_paths.append((plot_file_path_20, ''))

        if selected_plots[mat_file_path]["ADC measurements"]:
            plot_file_name_13 = f'{base_name}_adc_data.png'
            plot_file_path_13 = os.path.join(plots_folder, plot_file_name_13)
            dict = {'adc': ['timestamp', 'v_24', 'v_12', 'v_5', 'v_3', 'i_24']}
            dict2 = {'adc': ['timestamp', 'i_12', 'i_5', 'i_3', 'i_in', 'v_in']}
            dict3 = {'adc': ['timestamp', 'ain10', 'ain11']}
            dict_value_items = extract_values_from_data(mat_file_path, dict)
            dict_value_items2 = extract_values_from_data(mat_file_path, dict2)
            dict_value_items3 = extract_values_from_data(mat_file_path, dict3)
            create_partitioned_plot(plot_file_path_13, dict_value_items,'ADC measurements')
            create_partitioned_plot(f'{plot_file_path_13[:-4]}_2.png', dict_value_items2)
            create_partitioned_plot(f'{plot_file_path_13[:-4]}_3.png', dict_value_items3)
            plot_paths.append((plot_file_path_13, ''))
            plot_paths.append((f'{plot_file_path_13[:-4]}_2.png', ''))
            plot_paths.append((f'{plot_file_path_13[:-4]}_3.png', ''))



    doc_file_path = os.path.join(output_folder, 'report.docx')
    create_word_document(plot_paths, doc_file_path,map_paths)


    print(f"Processing complete. Files saved to {output_folder}")


def calculate_travel_times(data_dict):
    """
    Calculate travel times between waypoints based on latitude, longitude, and tool speed.

    Parameters:
        data_dict (dict): A dictionary containing 'latitude', 'longitude', and 'tool_speed'.

    Returns:
        np.array: Array of travel times between waypoints.
    """
    latitudes = np.array(data_dict['latitude'])
    longitudes = np.array(data_dict['longitude'])
    speeds = np.array(data_dict['tool_speed'])

    # Haversine formula to calculate distance between two points on the Earth
    def haversine(lat1, lon1, lat2, lon2):
        R = 6371.0  # Earth radius in kilometers

        dlat = np.radians(lat2 - lat1)
        dlon = np.radians(lon2 - lon1)
        a = np.sin(dlat / 2) ** 2 + np.cos(np.radians(lat1)) * np.cos(np.radians(lat2)) * np.sin(dlon / 2) ** 2
        c = 2 * np.arctan2(np.sqrt(a), np.sqrt(1 - a))
        return R * c  # Distance in kilometers

    # Calculate distances between consecutive waypoints
    distances = [
        haversine(latitudes[i], longitudes[i], latitudes[i+1], longitudes[i+1])
        for i in range(len(latitudes) - 1)
    ]

    # Convert speed from km/h to km/min by dividing by 60
    speeds_min = speeds[:-1] / 60.0

    # Calculate travel times (distance / speed)
    travel_times = np.array(distances) / speeds_min

    travel_times = np.insert(travel_times, 0, 0)  # Insert 0 for the first travel time

    return travel_times

def extract_values_from_data(mat_file_path, dict):
    data = scipy.io.loadmat(mat_file_path)
    pprint(data)
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
