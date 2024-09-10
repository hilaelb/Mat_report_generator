from MainWindow import MainWindow
from PyQt5.QtWidgets import QApplication

def process_files():
    app = QApplication([])
    main_window = MainWindow()

    main_window.showMaximized()
    app.exec_()



if __name__ == "__main__":
    process_files()