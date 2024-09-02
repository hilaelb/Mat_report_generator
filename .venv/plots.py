import os
import json
import re
from datetime import datetime
import folium
import time


import docx.opc.constants
import scipy.io
from pprint import pprint

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
    width = base_size[0] + scale_factor * (x_length ** 0.45)
    height = base_size[1] + scale_factor * (y_length ** 0.45)

    return width, height

def turn_time_to_nomeric_dev(timestamp):
    # Convert microseconds to seconds
    time_in_seconds = np.array(timestamp) / 1e6

    # Calculate the difference between consecutive timestamps
    time_diffs = np.diff(time_in_seconds)

    return time_diffs


def create_histogram(table, output_plot_path,title=''):
    if table is None:
        print("Error: Missing table data.")
        return
    num_columns = len(table)

    # Create a figure with subplots for each column
    fig, axes = plt.subplots(num_columns, 1, figsize=(11, 5 * num_columns))

    if num_columns == 1:
        axes = [axes]  # Ensure axes is a list even if there's only one subplot

    x_label = None

    for i, (key, data) in enumerate(table.items()):
        if key == 'timestamp':
            data = turn_time_to_nomeric_dev(data)
            x_label = 'Time [sec]'
        else:
            data = np.array(data)
            x_label = key

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

        axes[i].set_xlabel(f'{x_label}')
        axes[i].set_ylabel('# of Occurrences')
        axes[i].set_title(f'Histogram of {title} {x_label} with Stats')
        axes[i].legend()

    # Adjust layout and save the figure
    plt.tight_layout()
    plt.savefig(output_plot_path)
    plt.close()


def create_mixed_plot(output_plot_path, tables, title='',y_names ='',units_config_path="units_config.json" ):
    if tables is None:
        print("Error: Missing table data.")
        return
    units_config = load_units_config(units_config_path)

    x_data = None
    y_data_list = []
    y_labels = []

    for key, values in tables.items():
        if key == 'timestamp':
            # Convert time from microseconds to minutes
            x_data = np.array(values).flatten() / 1e6 / 60
            # Normalize the time to start from 0 minutes
            x_data = x_data - x_data[0]
            x_label = 'Time'
            x_unit = units_config.get(x_label.lower(), "")
        else:
            y_data_list.append(np.array(values).flatten())
            y_labels.append(key)

    if x_data is None or not y_data_list:
        print("Error: Missing necessary data for plotting.")
        return

    # Create the plot
    fig_width, fig_height = determine_figure_size(len(x_data), len(y_data_list[0]))
    plt.figure(figsize=(fig_width, fig_height))

    for i, y_data in enumerate(y_data_list):
        plt.scatter(x_data, y_data, label=f'{y_labels[i]} [{units_config.get(y_labels[i], "")}]')

    plt.xlabel(f'{x_label} [{x_unit}]')
    plt.ylabel(y_names)
    plt.title(title)
    plt.legend()
    plt.grid(True)

    # Save the plot
    plt.tight_layout()
    plt.savefig(output_plot_path)
    plt.close()

def create_partitioned_plot(output_plot_path, tables,title='', units_config_path="units_config.json"):

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

    x_unit = units_config.get(x_label, "")

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
        y_unit = units_config.get(y_label, "")

        axes[i].scatter(axis_x, axis_y, label=f'{y_label} [{y_unit}]')
        axes[i].set_ylabel(f'{y_label} [{y_unit}]')
        axes[i].grid(True)
        axes[i].legend(loc='upper right')




    # Set the x-axis label on the last subplot
    axes[-1].set_xlabel(f'{x_label} [{x_unit}]')

    fig.tight_layout(rect=[0, 0.03, 1, 0.95])
    plt.savefig(output_plot_path)
    plt.close()



def create_3d_plot(output_plot_path,dict_value_items, units_config_path="units_config.json"):
    if dict_value_items is None:
        print("Error: Missing table data.")
        return
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
            x_unit = units_config.get(x_label, "")

        elif key == 'latitude':
            axis_y1 = np.array(values).flatten()
            y1_label = 'latitude'
            y1_unit = units_config.get(y1_label, "")

        elif key == 'longitude':
            axis_y2 = np.array(values).flatten()
            y2_label = 'longitude'
            y2_unit = units_config.get(y2_label, "")

        elif key == 'altitude':
            axis_y2 = np.array(values).flatten()
            y2_label = 'depth'
            y2_unit = units_config.get(y2_label, "")

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
    if tables is None:
        print("Error: Missing tables data.")
        return
    pprint(tables)
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
            32: 'ARM', # we need to check whats modes there are
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




def create_plot(output_plot_path, tables,title='', units_config_path="units_config.json",plot=False,mean_sd = False,range= (0,0)):

    if tables is None:
        return
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
    x_unit = units_config.get(x_label, "")
    y_unit = units_config.get(y_label, "")

    fig_width, fig_height = determine_figure_size(len(axis_x), len(axis_y))
    plt.figure(figsize=(fig_width, fig_height))
    if plot:
        plt.plot(axis_x, axis_y, label=f'{y_label} over {x_label}')
    else:
        plt.scatter(axis_x, axis_y, label=f'{y_label} over {x_label}')

    # Calculate mean and standard deviation if mean_sd is True
    if mean_sd:
        mean = np.mean(axis_y)
        std_dev = np.std(axis_y)
        plt.axhline(y=mean, color='r', linestyle='--', label=f'Mean: {mean:.2f}')
        plt.fill_between(axis_x, mean - std_dev, mean + std_dev, color='r', alpha=0.2, label=f'SD: {std_dev:.2f}')

    if range!= (0,0):
        plt.ylim(range)

    # Increase the font size of the title, labels, and ticks
    plt.title(title, fontsize=20)  # Larger title
    plt.xlabel(f'{x_label} [{x_unit}]', fontsize=16)  # Larger x-axis label
    plt.ylabel(f'{y_label} [{y_unit}]', fontsize=16)  # Larger y-axis label
    plt.xticks(fontsize=14)  # Larger x-axis tick labels
    plt.yticks(fontsize=14)  # Larger y-axis tick labels



    # Set the format for the x and y ticks to show 4 decimal places
    plt.gca().xaxis.set_major_formatter(plt.FuncFormatter(format_decimal))
    plt.gca().yaxis.set_major_formatter(plt.FuncFormatter(format_decimal))


    #plt.legend()
    plt.savefig(output_plot_path)
    plt.close()



