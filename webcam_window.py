from PyQt5 import QtCore, QtWidgets, QtGui


class WebCamWindow:
    def create_window(self, form) -> None:
        form.setObjectName("form")
        form.resize(640, 360)
        form.setContentsMargins(0, 0, 0, 0)
        form.setFixedSize(640, 360)
        form.setWindowIcon(QtGui.QIcon('media_replacement.ico'))
        form.setWindowTitle("Webcam Window")
        self.vertical_layout = QtWidgets.QVBoxLayout(form)
        self.image_label = QtWidgets.QLabel(form)
        self.vertical_layout.addWidget(self.image_label)
        QtCore.QMetaObject.connectSlotsByName(form)
