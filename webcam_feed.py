from PyQt5.QtWidgets import QWidget
from PyQt5.QtGui import QImage, QPixmap
from PyQt5.QtCore import QTimer, pyqtSignal
from PyQt5.Qt import Qt
from PyQt5 import QtCore
from webcam_window import WebCamWindow
from shutil import rmtree
from numpy import ndarray, dstack, where

import cv2
import cvzone
import mediapipe as mp


class WebCam(QWidget):
    mp_selfie_segmentation = mp.solutions.selfie_segmentation
    segment = mp_selfie_segmentation.SelfieSegmentation()
    window_closed = pyqtSignal()

    def __init__(self, media, data_type):
        super().__init__()
        self.ui = WebCamWindow()
        self.ui.create_window(self)
        self.media = media
        self.data_type = data_type
        self.timer = QTimer()
        self.timer.timeout.connect(self.open_cv_scene)
        self.cap = cv2.VideoCapture(0, cv2.CAP_DSHOW)
        self.vid_count = 0
        self.index = 0
        self.arrows = cv2.imread('arrows.png', cv2.IMREAD_UNCHANGED)
        self.arrows = cv2.resize(self.arrows, (0, 0), None, 0.95, 0.95)

    def closeEvent(self, event) -> None:
        self.cap.release()
        self.setAttribute(QtCore.Qt.WA_DeleteOnClose)
        self.window_closed.emit()
        event.accept()
        if self.data_type == 'ppt':
            img_folder = r'C:\Windows\Temp\powerpoint_imgs\\'
            rmtree(img_folder)

    def keyPressEvent(self, event) -> None:
        if self.data_type == 'ppt':
            if event.key() == Qt.Key_Left and self.index != 0:
                self.index -= 1
            if event.key() == Qt.Key_Right and self.index < len(self.media) - 1:
                self.index += 1

    def modify_background(self, screen, threshold, media) -> ndarray:
        rgb_img = cv2.cvtColor(screen, cv2.COLOR_BGR2RGB)
        result = WebCam.segment.process(rgb_img)
        binary_mask = result.segmentation_mask > threshold
        binary_mask_3 = dstack((binary_mask, binary_mask, binary_mask))
        if self.data_type == 'mp4':
            if self.vid_count == media.get(cv2.CAP_PROP_FRAME_COUNT) - 1:
                self.vid_count = 0
                media.set(cv2.CAP_PROP_POS_FRAMES, 0)
            _, media = media.read()
            self.vid_count += 1
            self.vid = cv2.cvtColor(media, cv2.COLOR_BGR2RGB)
            background_img = cv2.resize(self.vid, (screen.shape[1], screen.shape[0]))
            output_img = where(binary_mask_3, screen, background_img)
        if self.data_type == 'img':
            bg = cv2.cvtColor(media, cv2.COLOR_BGR2RGB)
            bg = cv2.resize(bg, (1920, 1080))
            background_img = cv2.resize(bg, (screen.shape[1], screen.shape[0]))
            output_img = where(binary_mask_3, screen, background_img)
        if self.data_type == 'ppt':
            bg = cv2.cvtColor(media[self.index], cv2.COLOR_BGR2RGB)
            bg = cv2.resize(bg, (1920, 1080))
            background_img = cv2.resize(bg, (screen.shape[1], screen.shape[0]))
            output_img = where(binary_mask_3, screen, background_img)
        return output_img

    def open_cv_scene(self) -> None:
        try:
            success, frame = self.cap.read()
            frame = cv2.resize(frame, (640, 360))
            frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            flipped = cv2.flip(frame, 1)
            frame = self.modify_background(flipped, 0.65, self.media)
            height, width, channel = frame.shape
            step = channel * width
            if self.data_type == 'ppt':
                overlay_frame = cvzone.overlayPNG(frame, self.arrows, [10, 290])
                q_img = QImage(overlay_frame.data, width, height, step, QImage.Format_RGB888)
            else:
                q_img = QImage(frame.data, width, height, step, QImage.Format_RGB888)
            self.ui.image_label.setPixmap(QPixmap.fromImage(q_img))
        except cv2.error:
            self.cap.release()

    def start(self) -> None:
        self.timer.start(20)
