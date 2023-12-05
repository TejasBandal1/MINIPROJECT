from sklearn.neighbors import KNeighborsClassifier
import cv2
import pickle
import numpy as np
import os
import csv
import time
from datetime import datetime
from win32com.client import Dispatch

def speak(str1):
    speak = Dispatch(("SAPI.SpVoice"))
    speak.Speak(str1)

video = cv2.VideoCapture(0)
facedetect = cv2.CascadeClassifier('data/haarcascade_frontalface_default.xml')

with open('data/names.pkl', 'rb') as w:
    LABELS = pickle.load(w)
with open('data/faces_data.pkl', 'rb') as f:
    FACES = pickle.load(f)

print('Shape of Faces matrix -->', FACES.shape)

knn = KNeighborsClassifier(n_neighbors=5)
knn.fit(FACES, LABELS)

imgBackground = cv2.imread("background.png")

COL_NAMES = ['NAME', 'TIME']

while True:
    ret, frame = video.read()
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    faces = facedetect.detectMultiScale(gray, 1.3, 5)

    for (x, y, w, h) in faces:
        crop_img = frame[y:y + h, x:x + w, :]
        resized_img = cv2.resize(crop_img, (50, 50)).flatten().reshape(1, -1)
        output = knn.predict(resized_img)
        ts = time.time()
        date = datetime.fromtimestamp(ts).strftime("%d-%m-%Y")
        timestamp = datetime.fromtimestamp(ts).strftime("%H:%M-%S")
        exist = os.path.isfile("Attendance/Attendance_" + date + ".csv")
        cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 0, 255), 1)
        cv2.rectangle(frame, (x, y), (x + w, y + h), (50, 50, 255), 2)
        cv2.rectangle(frame, (x, y - 40), (x + w, y), (50, 50, 255), -1)
        cv2.putText(frame, str(output[0]), (x, y - 15), cv2.FONT_HERSHEY_COMPLEX, 1, (255, 255, 255), 1)
        cv2.rectangle(frame, (x, y), (x + w, y + h), (50, 50, 255), 1)
        attendance = [str(output[0]), str(timestamp)]

    resized_frame = cv2.resize(frame, (640, 480))
    y_offset = (imgBackground.shape[0] - resized_frame.shape[0]) // 2
    x_offset = (imgBackground.shape[1] - resized_frame.shape[1]) // 2
    imgBackground[y_offset:y_offset + resized_frame.shape[0], x_offset:x_offset + resized_frame.shape[1]] = resized_frame

    text1 = "PRESS 'O' TO TAKE ATTENDANCE"
    text2 = "PRESS 'Q' TO EXIT"
    text_position1 = (100, imgBackground.shape[0] - 50)
    text_position2 = (100, imgBackground.shape[0] - 10)
    font = cv2.FONT_HERSHEY_SIMPLEX
    font_scale = 0.7
    font_color = (128, 0, 128)  # Purple color
    font_thickness = 2
    cv2.putText(imgBackground, text1, text_position1, font, font_scale, font_color, font_thickness)
    cv2.putText(imgBackground, text2, text_position2, font, font_scale, font_color, font_thickness)

    heading_text = "AUTOMATED ATTENDANCE SYSTEM"
    heading_font_scale = 1.5
    heading_font_thickness = 3
    heading_font_color = (128, 0, 128)  # Purple color
    heading_text_size = cv2.getTextSize(heading_text, font, heading_font_scale, heading_font_thickness)[0]
    heading_text_position = ((imgBackground.shape[1] - heading_text_size[0]) // 2, heading_text_size[1] + 20)
    cv2.putText(imgBackground, heading_text, heading_text_position, font, heading_font_scale, heading_font_color, heading_font_thickness)

    cv2.imshow("Frame", imgBackground)

    k = cv2.waitKey(1)
    if k == ord('o') or k == ord('O'):
        speak("Attendance Taken..")
        time.sleep(5)
        if exist:
            with open("Attendance/Attendance_" + date + ".csv", "+a") as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(attendance)
        else:
            with open("Attendance/Attendance_" + date + ".csv", "+a") as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(COL_NAMES)
                writer.writerow(attendance)

    if k == ord('q') or k == ord('Q'):
        break

video.release()
cv2.destroyAllWindows() 
