from datetime import datetime
import cv2
from pyzbar.pyzbar import decode
from oauth2client.service_account import ServiceAccountCredentials
import gspread
import time
from time import localtime
from PIL import ImageFont, ImageDraw, Image
import numpy as np

font = ImageFont.truetype(font="NotoSansKR-Regular.otf", size=48)

#명렬표 불러오기
scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive",
]

json_key_path1 = "qr.json" # JSON Key File Path
json_key_path2 = "qr_input.json"

credential1 = ServiceAccountCredentials.from_json_keyfile_name(json_key_path1, scope)
gc1 = gspread.authorize(credential1)
credential2 = ServiceAccountCredentials.from_json_keyfile_name(json_key_path2, scope)
gc2 = gspread.authorize(credential2)

spreadsheet_url1 = "https://docs.google.com/spreadsheets/d/1IMJK2dUZvAHWBEmCO2-U6w_i6BlMHPeeWbxfGDgORSo/edit?usp=sharing"
spreadsheet_url2 = "https://docs.google.com/spreadsheets/d/1FzeHKxmETsF_adx3hmeayIbFKMny4EI2OwBkeklw-EY/edit#gid=0"
doc1 = gc1.open_by_url(spreadsheet_url1)
doc2 = gc2.open_by_url(spreadsheet_url2)

#당일의 요일을 가져옴
weeknum = datetime.today().weekday()
weekday = ""

#d값을 정의함.
d = 811

if weeknum == 0:
    weekday = "월요일"
elif weeknum == 1:
    weekday = "화요일"
elif weeknum == 2:
    weekday = "수요일"
elif weeknum == 3:
    weekday = "목요일"
elif weeknum == 4:
    weekday = "금요일"
elif weeknum == 5:
    weekday = "토요일"
elif weeknum == 6:
    weekday = "일요일"

# 노트북 카메라 캡처
cap = cv2.VideoCapture(0)
flag = 0

sheet1 = doc1.worksheet(weekday)
sheet2 = doc2.worksheet("시트1")

# B열의 전체 데이터를 가져옴
column_b1 = sheet1.col_values(2)
column_a2 = sheet2.col_values(1)

while True:
    # 프레임 읽기
    ret, frame = cap.read()

    # QR 코드 디코딩
    decoded_data = decode(frame)
    

    # 디코딩된 QR 코드 정보 출력
    if decoded_data:
        if flag == 0:
            sheet1 = doc1.worksheet(weekday)
            sheet2 = doc2.worksheet("시트1")

            # B열의 전체 데이터를 가져옴
            column_b1 = sheet1.col_values(2)
            column_a2 = sheet2.col_values(1)
        flag = 1
        인식결과 = decoded_data[0].data.decode()
        개인정보 = 인식결과.split(' ')[0]
        시간 = float(인식결과.split(' ')[1])
        현재시간 = time.time()
        tm = localtime(현재시간)
        시간차이 = 현재시간-시간
        연도 = tm.tm_year
        달 = tm.tm_mon
        일 = tm.tm_mday
        시 = tm.tm_hour
        분 = tm.tm_min
        초 = tm.tm_sec
        print(시간, 현재시간, 시간차이)
        print(개인정보)
        print(인식결과)
        if 개인정보 in column_b1 :
            if 시간차이<=60 :
                if 개인정보 in column_a2 :
                    print('이미 출석하였습니다!')
                    text = "이미 출석하였습니다!"
                    img = frame
                    img = Image.fromarray(img)
                    draw = ImageDraw.Draw(img)
                    draw.text((100, 300), text, font=font, fill=(255, 0, 0))
                    img = np.array(img)
                    frame = img
                    cv2.imshow("QR code reader", frame)
                    cv2.waitKey(1000)
                    continue

                else:
                    flag = 0
                    print('통과입니다!')
                    text = "통과입니다!"
                    img = frame
                    img = Image.fromarray(img)
                    draw = ImageDraw.Draw(img)
                    draw.text((200, 300), text, font=font, fill=(255, 0, 0))
                    img = np.array(img)
                    frame = img
                    row = [개인정보, 연도, 달, 일, 시, 분, 초]
                    sheet2.append_row(row)
                    cv2.imshow("QR code reader", frame)
                    cv2.waitKey(1000)
                    continue
                
            else:
                print("만료된 QR코드입니다")
                text = "만료된 QR코드입니다"
                img = frame
                img = Image.fromarray(img)
                draw = ImageDraw.Draw(img)
                draw.text((200, 300), text, font=font, fill=(255, 0, 0))
                img = np.array(img)
                frame = img
                cv2.imshow("QR code reader", frame)
                cv2.waitKey(1000)
                continue
        else:
            print('학생이 아닙니다.')
            text = "학생이 아닙니다."
            img = frame
            img = Image.fromarray(img)
            draw = ImageDraw.Draw(img)
            draw.text((200, 300), text, font=font, fill=(255, 0, 0))
            img = np.array(img)
            frame = img
            cv2.imshow("QR code reader", frame)
            cv2.waitKey(1000)
            continue
    else:
        flag = 0
    # 이미지 창에 보여주기
    cv2.imshow("QR code reader", frame)
    #if flag:
        #time.sleep(1)

    # 'q' 키를 누르면 종료
    if cv2.waitKey(1) & 0xFF == ord('q'):
        break
    time.sleep(0.1)

# 종료
cap.release()
cv2.destroyAllWindows()