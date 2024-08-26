import os
import scipy.io
# import tkinter as tk
# from tkinter import filedialog
from pprint import pprint
from PyQt5.QtWidgets import QApplication, QFileDialog
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Inches
import numpy as np

def create_depth_plot(mat_file_path, output_plot_path):
    # Load the .mat file
    data = scipy.io.loadmat(mat_file_path)

    # Extract the relevant field
    field_data = data['ba30']

    # Extract timestamp and depth
    timestamps = field_data[0][0]['timestamp'][0]
    depths = field_data[0][0]['depth'][0]

    # Convert arrays to numpy arrays for easier manipulation
    timestamps = np.array(timestamps).flatten()
    depths = np.array(depths).flatten()

    # Create the plot
    plt.figure(figsize=(10, 6))
    plt.plot(timestamps, depths, label='Depth over Time')
    plt.xlabel('Time')
    plt.ylabel('Depth')
    plt.title(f'Depth as a Function of Time for {mat_file_path}')
    plt.legend()

    # Save the plot to a file
    plt.savefig(output_plot_path)
    plt.close()




def create_word_document(plot_paths, doc_file_path):
    doc = Document()

    for index, plot_path in enumerate(plot_paths):
        if index > 0:
            # Add a new page for each plot except the first one
            doc.add_page_break()
        doc.add_heading(f'Plot from file: {plot_path}', level=1)
        doc.add_picture(plot_path, width=Inches(6))
        doc.add_paragraph('This plot shows the depth as a function of time.')

    # Save the document
    doc.save(doc_file_path)

def clear_plots_folder(plots_folder):
    for file_name in os.listdir(plots_folder):
        file_path = os.path.join(plots_folder, file_name)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
        except Exception as e:
            print(f'Failed to delete {file_path}. Reason: {e}')

def process_files():

    app = QApplication([])
    file_paths, _ = QFileDialog.getOpenFileNames(None, "Select .mat Files", "", "MAT files (*.mat)")

    if not file_paths:
        return

    # Create output directory on the desktop
    desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
    output_folder = os.path.join(desktop_path, 'MAT_Files_Report')
    plots_folder = os.path.join(output_folder, 'plots')
    os.makedirs(plots_folder, exist_ok=True)
    os.makedirs(output_folder, exist_ok=True)

    # Clear the plots folder before adding new plots
    clear_plots_folder(plots_folder)

    plot_paths = []

    for mat_file_path in file_paths:
        plot_file_name = os.path.basename(mat_file_path).replace('.mat', '.png')
        plot_file_path = os.path.join(plots_folder, plot_file_name)
        create_depth_plot(mat_file_path, plot_file_path)
        plot_paths.append(plot_file_path)

    doc_file_path = os.path.join(output_folder, 'depth_vs_time_report.docx')
    create_word_document(plot_paths, doc_file_path)

    # Notify user
    print(f"Processing complete. Files saved to {output_folder}")



if __name__ == "__main__":
    process_files()