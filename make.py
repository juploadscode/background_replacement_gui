from cx_Freeze import setup, Executable
from PyQt5 import QtCore, QtGui, QtWidgets, QtCore
from PyQt5.QtWidgets import QFileDialog, QWidget
from PyQt5.QtGui import QImage, QPixmap
from PyQt5.QtCore import QTimer, pyqtSignal
from PyQt5.Qt import Qt
from webcam_feed import WebCam
from webcam_window import WebCamWindow
from os import mkdir, path

import win32com.client


from shutil import rmtree
from numpy import ndarray, dstack, where

import cv2
import cvzone
import mediapipe as mp
import pptx
import sys

build_exe_options = {"packages": ['PyQt5', 'webcam_feed', 'webcam_window', 'pptx', 'os', 'win32com',
                                  'shutil', 'numpy', 'cv2', 'cvzone', 'mediapipe', 'sys'],
                     "include_files": ["arrows.png", 'gui_background.png', 'media_replacement.ico']}

base = None
if sys.platform == 'win32':
    base = 'Win32GUI'

setup(
    name="Media Greenscreen Replacement",
    version='1.0',
    description='Upload media to be replaced as web cam background.',
    options={'build_exe': build_exe_options},
    executables=[Executable("media_replacement_gui.py", base=base, icon='media_replacement.ico')]
)
