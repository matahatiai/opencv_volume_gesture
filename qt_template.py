from PyQt5 import uic
from PyQt5 import QtGui, QtWidgets, uic
import sys
import cv2
from PyQt5.QtCore import QThread, Qt, pyqtSignal, pyqtSlot
from PyQt5.QtGui import QImage, QPixmap
from cvzone.HandTrackingModule import HandDetector
import math
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import numpy as np

# Get default audio device using PyCAW
devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(
    IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))
# Get current volume
currentVolumeDb = volume.GetMasterVolumeLevel()
# print(volume.GetVolumeRange())

detector = HandDetector(detectionCon=0.5, maxHands=1)

class Thread(QThread):
    changePixmap = pyqtSignal(QImage)

    def run(self):
        last_angle = None
        last_length = None

        minAngle = 0
        maxAngle = 180
        angle = 0
        angleBar = 400
        angleDeg = 0
        angleVol = 0
        minHand = 50  # 50
        maxHand = 300  # 300

        ################################
        wCam, hCam = 1280, 720
        ################################

        last_volume = 0
        deg_step = 180 / 50
        cap = cv2.VideoCapture(1)
        cap.set(3, wCam)
        cap.set(4, hCam)
        while True:
            success, img = cap.read()
            if success:
                # Find the hand and its landmarks
                img = detector.findHands(img)
                lmList, bboxInfo = detector.findPosition(img)
                if len(lmList) != 0:
                    # print(lmList[4], lmList[8])
                    x1, y1 = lmList[4][0], lmList[4][1]
                    x2, y2 = lmList[8][0], lmList[8][1]
                    cx, cy = (x1 + x2) // 2, (y1 + y2) // 2

                    cv2.circle(img, (x1, y1), 15, (0, 0, 255), cv2.FILLED)
                    cv2.circle(img, (x2, y2), 15, (0, 0, 255), cv2.FILLED)
                    cv2.line(img, (x1, y1), (x2, y2), (0, 0, 255), 3)
                    cv2.circle(img, (cx, cy), 15, (0, 0, 255), cv2.FILLED)

                    length = math.hypot(x2 - x1, y2 - y1)

                    # Hand range 50 - 300
                    angle = np.interp(length, [minHand, maxHand], [minAngle, maxAngle])
                    angleBar = np.interp(length, [minHand, maxHand], [400, 150])
                    angleDeg = np.interp(length, [minHand, maxHand], [0, 180])  # degree angle 0 - 180

                    # Convert angledeg to level volume
                    maxVol = -62.25
                    angleDegStep = maxVol / 180
                    valVol = maxVol - (angleDeg * angleDegStep)
                    # volume.SetMasterVolumeLevel(valVol, None)
                    angleVol = math.floor(angleDeg / (180 / 100))

                    last_angle = angle
                    last_length = length

                    if length < 50:
                        cv2.circle(img, (cx, cy), 15, (0, 255, 0), cv2.FILLED)

                cv2.rectangle(img, (50, 150), (85, 400), (255, 0, 0), 3)
                cv2.rectangle(img, (50, int(angleBar)), (85, 400), (255, 0, 0), cv2.FILLED)
                cv2.putText(img, f'Vol.{int(angleVol)}', (40, 90), cv2.FONT_HERSHEY_COMPLEX, 2, (0, 9, 255), 3)

                # https://stackoverflow.com/a/55468544/6622587
                rgbImage = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
                h, w, ch = rgbImage.shape
                bytesPerLine = ch * w
                convertToQtFormat = QImage(rgbImage.data, w, h, bytesPerLine, QImage.Format_RGB888)
                p = convertToQtFormat.scaled(509, 521 , Qt.KeepAspectRatio)
                self.changePixmap.emit(p)


class Ui(QtWidgets.QMainWindow):
    def __init__(self):
        super(Ui, self).__init__()
        uic.loadUi('template.ui', self)
        self.setWindowTitle('Hand Tracking Volume')

        entries = [
            '************************************************',
            'HAND TRACKING VOLUME BY FAJARLABS',
            '************************************************',
            'Posisikan telunjuk dan ibu jari',
            'mendekat atau menjauh.',
            '',
            'Ketika telunjuk dan ibu jari menjauh',
            'suara akan membesar.',
            '',
            'Ketika telunjuk dan ibu jari mendekat',
            'suara akan mengecil.'
        ]

        model = QtGui.QStandardItemModel()
        self.listView.setModel(model)

        for i in entries:
            item = QtGui.QStandardItem(i)
            model.appendRow(item)

        th = Thread(self)
        th.changePixmap.connect(self.setImage)
        th.start()

        # self.pushButton_2.clicked.connect(self.on_click)
        self.show()

    @pyqtSlot(QImage)
    def setImage(self, image):
        self.label.setPixmap(QPixmap.fromImage(image))

    @pyqtSlot()
    def on_click(self):
        print('PyQt5 button click')

app = QtWidgets.QApplication(sys.argv)
window = Ui()
app.exec_()