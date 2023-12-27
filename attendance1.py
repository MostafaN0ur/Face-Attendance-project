import cv2
import numpy as np
import face_recognition
import os
from datetime import datetime,date
import xlrd
from xlutils.copy import copy as xl_copy
#import matplotlib.pyplot as plt
# from PIL import ImageGrab

path = 'known_people'
images = []
classNames = []
matches = []
faceLoc = []
faceDis = []
matchIndex = None
myList = os.listdir(path)
# print(myList)
for cl in myList:
    curImg = cv2.imread(f'{path}/{cl}')
    images.append(curImg)
    classNames.append(os.path.splitext(cl)[0])
# print(classNames)


def findEncodings(images):
    encodeList = []
    for img in images:
         img = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
         encode = face_recognition.face_encodings(img)[0]
         encodeList.append(encode)
    return encodeList


rb = xlrd.open_workbook('attendence_excel.xls', formatting_info=True)
wb = xl_copy(rb)
# myDataList = wb.readlines()
nameList = []
inp = input('Please give current subject lecture name')
sheet1 = wb.add_sheet(inp)
sheet1.write(0, 0, 'Name')
sheet1.write(0, 1, str(date.today()))
row = 1
col = 0
already_attendence_taken = ""

#  ### FOR CAPTURING SCREEN RATHER THAN WEBCAM
# def captureScreen(bbox=(300,300,690+300,530+300)):
#     capScr = np.array(ImageGrab.grab(bbox))
#     capScr = cv2.cvtColor(capScr, cv2.COLOR_RGB2BGR)
#     return capScr

encodeListKnown = findEncodings(images)
# print('Encoding Complete')

cap = cv2.VideoCapture(0)

while True:
    success, img = cap.read()
    # img = captureScreen()
    imgS = cv2.resize(img, (0, 0), None, 0.25, 0.25)
    imgS = cv2.cvtColor(imgS, cv2.COLOR_BGR2RGB)

    facesCurFrame = face_recognition.face_locations(imgS)
    encodesCurFrame = face_recognition.face_encodings(imgS, facesCurFrame)

    for encodeFace, faceLoc in zip(encodesCurFrame, facesCurFrame):
        matches = face_recognition.compare_faces(encodeListKnown, encodeFace)
        faceDis = face_recognition.face_distance(encodeListKnown, encodeFace)
        # print(faceDis)
        matchIndex = np.argmin(faceDis)
        # plt.imshow(imgS)
        # plt.show()

    if matches[matchIndex]:
        name = classNames[matchIndex].upper()
        # print(name)
        y1, x2, y2, x1 = faceLoc
        y1, x2, y2, x1 = y1 * 4, x2 * 4, y2 * 4, x1 * 4
        cv2.rectangle(img, (x1, y1), (x2, y2), (0, 255, 0), 2)
        cv2.rectangle(img, (x1, y2 - 35), (x2, y2), (0, 255, 0), cv2.FILLED)
        cv2.putText(img, name, (x1 + 1, y2 - 6), cv2.FONT_HERSHEY_DUPLEX, 1, (255, 255, 255), 1)
        if ((already_attendence_taken != name) and (name != "Unknown")):
            sheet1.write(row, col, name)
            col = col + 1
            now = datetime.now()
            dtString = now.strftime('%H:%M:%S')
            sheet1.write(row, col, dtString)
            row = row + 1
            col = 0
            print("attendence taken")
            wb.save('attendence_excel.xls')
            already_attendence_taken = name
        else:
            print("next student")

    if faceDis[matchIndex] < 0.50:
            name = classNames[matchIndex].upper()
            # plt.imshow(imgS)
            # plt.show()
            if ((already_attendence_taken != name) and (name != "Unknown")):
                sheet1.write(row, col, name)
                col = col + 1
                now = datetime.now()
                dtString = now.strftime('%H:%M:%S')
                sheet1.write(row, col, dtString)
                row = row + 1
                col = 0
                print("attendence taken")
                wb.save('attendence_excel.xls')
                already_attendence_taken = name
            else:
                print("next student")
    else:
        name = 'Unknown'
        # print(name)
        y1, x2, y2, x1 = faceLoc
        y1, x2, y2, x1 = y1 * 4, x2 * 4, y2 * 4, x1 * 4
        cv2.rectangle(img, (x1, y1), (x2, y2), (0, 255, 0), 2)
        cv2.rectangle(img, (x1, y2 - 35), (x2, y2), (0, 255, 0), cv2.FILLED)
        cv2.putText(img, name, (x1 + 1, y2 - 6), cv2.FONT_HERSHEY_DUPLEX, 1, (255, 255, 255), 1)

    cv2.imshow('Webcam', img)
    key = cv2.waitKey(1)
    if key == 27:  # esc
        break
