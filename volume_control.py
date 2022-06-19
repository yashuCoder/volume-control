import cv2 as cv
import mediapipe as mp
import numpy as np
import math
import time
import hand_tracking_module as htm
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

cap = cv.VideoCapture(0)
hand_detector = htm.detector(detection_con=0.75)

ptime = 0

devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(
    IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))
# volume.GetMute()
# volume.GetMasterVolumeLevel()
volRange = volume.GetVolumeRange()
max_vol = volRange[1]
min_vol = volRange[0]
vol = 0
volbar = 400
volper = 0
volume.SetMasterVolumeLevel(-10,None)
while True:
    success, frame = cap.read()
    frame = cv.resize(frame, (540, 420))
    frame = hand_detector.draw_hands(frame)
    lm_list = hand_detector.find_pos(frame, draw=False)

    if len(lm_list) != 0:
        x, y = lm_list[4][1], lm_list[4][2]
        x1, y1 = lm_list[8][1], lm_list[8][2]
        cx, cy = (x + x1) // 2, (y + y1) // 2
        cv.circle(frame, (x, y), 10, (255, 0, 255), cv.FILLED)
        cv.circle(frame, (x1, y1), 10, (255, 0, 255), cv.FILLED)
        cv.circle(frame, (cx, cy), 10, (255, 0, 225), cv.FILLED)
        cv.line(frame, (x, y), (x1, y1), (255, 0, 255), thickness=3)
        length = math.hypot(x1 - x, y1 - y)
        vol = np.interp(length, [30, 190], [min_vol, max_vol])
        volbar = np.interp(vol, [min_vol, max_vol], [400, 120])
        volper = np.interp(vol, [min_vol, max_vol], [0, 100])
        volume.SetMasterVolumeLevel(vol, None)
        if (length < 30):
            cv.circle(frame, (cx, cy), 10, (0, 255, 0), cv.FILLED)

    cv.rectangle(frame, (30, 120), (70, 400), (255, 0, 255), thickness=2)
    cv.rectangle(frame, (30, int(volbar)), (70, 400), (255, 0, 255), thickness=cv.FILLED)
    cv.putText(frame, f"{str(int(volper))}%", (30, 100), cv.FONT_HERSHEY_COMPLEX, 1, (255, 0, 255), thickness=2)

    ctime = time.time()
    fps = 1 / (ctime - ptime)
    ptime = ctime

    cv.putText(frame, f"FPS : {str(int(fps))}%", (20, 30), cv.FONT_HERSHEY_COMPLEX, 1, (255, 0, 255), thickness=2)
    cv.imshow("frame", frame)
    if cv.waitKey(1) & 0xFF == ord('d'):
        break
cap.release()
cv.destroyAllWindows()
