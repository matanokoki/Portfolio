import cv2
import numpy as np

def detect_led_roi(frame):
    #グレースケール
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)

    #二値化(部屋の明るさ200とする)
    _, thresh = cv2.threshold(gray, 200, 255, cv2.THRESH_BINARY)

    #輪郭抽出
    contours, _ = cv2.findContours(thresh, cv2.PETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    if not contours:
        return None
    
    #一番大きい光をLEDとみなす
    max_contour = max(contours, key=cv2.contourArea)
    x, y, w, h = cv2.boundingRect(max_contour)

    return (x,y,w,h)