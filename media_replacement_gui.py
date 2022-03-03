from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog
from webcam_feed import WebCam
from cv2 import VideoCapture, imread
from pptx import Presentation as PPoint
from sys import argv, exit
from os import mkdir, path

import win32com.client


class MainWindowUi:
    ACCEPTED_TYPES = ['.jpg', '.jpeg', '.png', '.mp4', '.pptx']

    WHITE = (255, 255, 255)
    YELLOW = (238, 210, 2)
    RED = (237, 67, 55)
    GREEN = (0, 128, 0)

    def set_up_ui(self, main_window) -> None:
        main_window.resize(407, 410)
        main_window.setFixedSize(407, 410)
        main_window.setWindowTitle("Media Background Replacement")
        main_window.setWindowIcon(QtGui.QIcon('media_replacement.ico'))

        font = QtGui.QFont()
        font.setPointSize(11)

        self.central_widget = QtWidgets.QWidget(main_window)
        self.central_widget.setStyleSheet("background-color: rgb(102, 51, 152);")

        self.gui_img = QtWidgets.QLabel(self.central_widget)
        self.gui_img.setGeometry(QtCore.QRect(0, 0, 401, 201))
        self.gui_img.setStyleSheet("background-image: url('gui_background.png');")
        self.gui_img.setAlignment(QtCore.Qt.AlignCenter)

        self.browser_file_button = QtWidgets.QPushButton(self.central_widget)
        self.browser_file_button.setGeometry(QtCore.QRect(50, 220, 161, 61))
        self.browser_file_button.setFont(font)
        self.browser_file_button.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.browser_file_button.setText("Browse File...")
        self.browser_file_button.clicked.connect(self.set_browse_file)

        self.launch_webcam_button = QtWidgets.QPushButton(self.central_widget)
        self.launch_webcam_button.setGeometry(QtCore.QRect(50, 290, 161, 61))
        self.launch_webcam_button.setFont(font)
        self.launch_webcam_button.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.launch_webcam_button.setText("Launch Webcam")
        self.launch_webcam_button.setEnabled(False)
        self.launch_webcam_button.clicked.connect(self.web_cam_opened_actions)

        self.status_color = QtWidgets.QTextBrowser(self.central_widget)
        self.status_color.setGeometry(QtCore.QRect(290, 250, 71, 71))
        self.status_color.setStyleSheet("background-color: rgb(255, 255, 255);")

        self.menu_bar = QtWidgets.QMenuBar(main_window)
        self.menu_bar.setGeometry(QtCore.QRect(0, 0, 407, 21))

        self.status_bar = QtWidgets.QStatusBar(main_window)

        main_window.setStatusBar(self.status_bar)
        main_window.setMenuBar(self.menu_bar)
        main_window.setCentralWidget(self.central_widget)

        QtCore.QMetaObject.connectSlotsByName(main_window)

    def set_browse_file(self) -> None:
        file_name = QFileDialog.getOpenFileName()
        self.file_path = file_name[0]
        last_part = path.splitext(self.file_path)[1].lower()
        if last_part in self.ACCEPTED_TYPES:
            if last_part == ".mp4":
                self.data_type = "mp4"
                self.media = VideoCapture(self.file_path)
            if last_part == '.jpg' or last_part == '.jpeg' or last_part == '.png':
                self.data_type = "img"
                self.media = imread(self.file_path)
            try:
                if last_part == '.pptx':
                    self.data_type = "ppt"
                    ppt_application = win32com.client.Dispatch('Powerpoint.Application')
                    presentation = ppt_application.presentations.Open(self.file_path)
                    presentation_count = PPoint(self.file_path)
                    slides = presentation_count.slides
                    img_folder = r'C:\Windows\Temp\powerpoint_imgs\\'
                    mkdir(img_folder)
                    self.media = []
                    for slide_index in range(len(slides)):
                        img = f'{img_folder}{slide_index}.jpg'
                        presentation.Slides[slide_index].Export(img, 'JPG')
                        self.media.append(imread(img))
                    ppt_application.Quit()
                self.launch_webcam_button.setEnabled(True)
                self.browser_file_button.setEnabled(False)
                self.status_color.setStyleSheet(f"background-color: rgb{self.YELLOW};")
            except Exception:
                self.status_color.setStyleSheet(f"background-color: rgb{self.RED};")
                error_dialog = QtWidgets.QErrorMessage()
                error_dialog.showMessage('Install PowerPoint onto the device and then try again.')
                error_dialog.setWindowTitle('Error')
                error_dialog.exec_()
        else:
            self.status_color.setStyleSheet(f"background-color: rgb{self.RED};")
            error_dialog = QtWidgets.QErrorMessage()
            error_dialog.showMessage(f"Error loading {last_part} file. Select an appropriate file type and try again.")
            error_dialog.setWindowTitle('Error')
            error_dialog.exec_()

    def web_cam_opened_actions(self) -> None:
        self.web_cam_window = WebCam(self.media, self.data_type)
        self.web_cam_window.start()
        self.web_cam_window.show()
        self.web_cam_window.window_closed.connect(self.web_cam_closed_actions)
        self.launch_webcam_button.setEnabled(False)
        self.status_color.setStyleSheet(f"background-color: rgb{self.GREEN};")

    def web_cam_closed_actions(self) -> None:
        self.status_color.setStyleSheet(f"background-color: rgb{self.WHITE};")
        self.browser_file_button.setEnabled(True)


if __name__ == "__main__":
    app = QtWidgets.QApplication(argv)
    window = QtWidgets.QMainWindow()
    ui = MainWindowUi()
    ui.set_up_ui(window)
    window.show()
    exit(app.exec_())
