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
from PyQt5.QtWidgets import QApplication, QDialog, QFileDialog, QVBoxLayout, QCheckBox, QPushButton, QLabel, QScrollArea, QWidget, QHBoxLayout,QListWidget
from PlotSelectionDialog import PlotSelectionDialog
from plots import create_plot, create_kml_from_waypoints, kml_to_gmplot_image, create_acceleration_histogram, create_partitioned_plot, create_mixed_plot, plot_tiger_modes



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

            # Check if there's a corresponding map file for this plot
            base_name = os.path.basename(plot_path).replace('.png', '')
            map_file_path = next((path for path in map_file_paths if base_name in os.path.basename(path)), None)

            if map_file_path:
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


def process_files():
    """Process selected files, generate plots, and create a Word document report."""
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
        for plot_type, plot_info in selected_plots[mat_file_path].items():
            if plot_info:
                plot_info_func = {
                    "Depth/Time Plot": (create_plot, {'ba30': ['depth', 'timestamp']}),
                    "Lat/Long Plot": (create_plot, {'waypoint': ['latitude', 'longitude']}),
                    "Add a picture of the map": (kml_to_gmplot_image, {'waypoint': ['latitude', 'longitude']}),
                    "Jet Rpm/Time Plot": (create_plot, {'jet': ['rpm', 'timestamp']}),
                    "Jet Voltage/Time Plot": (create_plot, {'jet': ['v_in', 'timestamp']}),
                    "Jet data": (create_partitioned_plot, {'jet': ['timestamp', 'avg_input_current', 'avg_motor_current','rpm','temp_fet','v_in']}),
                    "Servo Angle/Time Plot": (create_partitioned_plot, {'servo': ['timestamp', 'srv1_angle', 'srv2_angle', 'srv3_angle', 'srv4_angle']}),
                    "Servo Voltage/Time Plot": (create_partitioned_plot, {'servo': ['timestamp', 'srv1_voltage', 'srv2_voltage', 'srv3_voltage', 'srv4_voltage']}),
                    "Servo Temperature/Time Plot": (create_partitioned_plot, {'servo': ['timestamp', 'srv1_temperature', 'srv2_temperature', 'srv3_temperature', 'srv4_temperature']}),
                    "Inertial data": (create_mixed_plot, {'insstatus': ['timestamp', 'ins_mode', 'ins_error', 'ins_fix']}),
                    "Gyro/Accel data": (create_partitioned_plot, {'gyro_accel_data': ['timestamp', 'gyro_x', 'gyro_y', 'gyro_z', 'accel_x', 'accel_y', 'accel_z']}),
                    "BME front measurements": (create_partitioned_plot, {'bme_front': ['timestamp', 'temperature', 'pressure', 'humidity']}),
                    "BME back measurements": (create_mixed_plot, {'bme_back': ['timestamp', 'temperature', 'pressure', 'humidity']}),
                    "Bar30 measurements": (create_partitioned_plot, {'ba30': ['timestamp', 'temperature', 'pressure', 'depth']}),
                    "Motor Guidance data": (create_partitioned_plot, {'motor_guidance_data': ['timestamp', 'srv_1_angle', 'srv_2_angle', 'srv_4_angle', 'srv_4_angle', 'jet_rpm']}),
                    "Vn200 data": lambda p, d: process_vn200_data(p, d),
                    "Histogram of Acceleration": (create_acceleration_histogram, {'gyro_accel_data': ['accel_x', 'accel_y', 'accel_z']}),
                    "Tiger Mode/Time Plot": (plot_tiger_modes, {'tiger': ['tiger_mode', 'timestamp']}),
                    "QGC Mode/Time Plot": (plot_tiger_modes, {'qgc': ['timestamp','mode']}),
                    "ADC measurements": lambda p, d: process_adc_data(p, d),
                }

                if plot_type in plot_info_func:
                    plot_func, data_dict = plot_info_func[plot_type]
                    plot_file_name = f'{base_name}_{plot_type.lower().replace(" ", "_")}.png'
                    plot_file_path = os.path.join(plots_folder, plot_file_name)

                    if plot_type == "Add a picture of the map":
                        map_file_path = os.path.join(maps_folder, f'{base_name}_waypoint_rosbag_lat_vs_long.png')
                        plot_func(mat_file_path, 'waypoint', 'latitude', 'longitude', map_file_path)
                        map_paths.append(map_file_path)
                    else:
                        dict_value_items = extract_values_from_data(mat_file_path, data_dict)
                        plot_func(dict_value_items, plot_file_path)
                        plot_paths.append((plot_file_path, plot_type))

    # Create KML files
    kml_paths = []
    for mat_file_path in file_paths:
        base_name = os.path.basename(mat_file_path).replace('.mat', '')
        kml_file_path = os.path.join(kmls_folder, f'{base_name}.kml')
        create_kml_from_waypoints(mat_file_path, kml_file_path)
        kml_paths.append(kml_file_path)

    create_word_document(plot_paths, os.path.join(output_folder, 'report.docx'), map_paths)
    print(f'Report created at {os.path.join(output_folder, "report.docx")}')
def process_vn200_data(mat_file_path, plot_type):
    """Process and plot VN200 data."""
    # Add your specific VN200 data processing code here.
    pass


def process_adc_data(mat_file_path, plot_type):
    """Process and plot ADC data."""
    # Add your specific ADC data processing code here.
    pass

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
