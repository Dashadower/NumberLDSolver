# -*- coding:utf-8 -*-
import cv2, os, pytesseract, sys, logging
from screencap import MapleScreenCapturer
import win32api, win32con, numpy as np

def click(x,y):
    win32api.SetCursorPos((x,y))
    time.sleep(0.1)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,x,y,0,0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,x,y,0,0)


numbers = {
    1:0x02,
    2:0x03,
    3:0x04,
    4:0x05,
    5:0x06,
    6:0x07,
    7:0x08,
    8:0x09,
    9:0x0A,
    0:0x0B}
DIK_RETURN = 0x1C

import time, ctypes, threading

CONFIRM_BUTTON_COORDS = (741, 484)

SendInput = ctypes.windll.user32.SendInput
# C struct redefinitions
PUL = ctypes.POINTER(ctypes.c_ulong)

class KeyBdInput(ctypes.Structure):
    _fields_ = [("wVk", ctypes.c_ushort),
                ("wScan", ctypes.c_ushort),
                ("dwFlags", ctypes.c_ulong),
                ("time", ctypes.c_ulong),
                ("dwExtraInfo", PUL)]

class HardwareInput(ctypes.Structure):
    _fields_ = [("uMsg", ctypes.c_ulong),
                ("wParamL", ctypes.c_short),
                ("wParamH", ctypes.c_ushort)]

class MouseInput(ctypes.Structure):
    _fields_ = [("dx", ctypes.c_long),
                ("dy", ctypes.c_long),
                ("mouseData", ctypes.c_ulong),
                ("dwFlags", ctypes.c_ulong),
                ("time",ctypes.c_ulong),
                ("dwExtraInfo", PUL)]


class Input_I(ctypes.Union):
    _fields_ = [("ki", KeyBdInput),
                 ("mi", MouseInput),
                 ("hi", HardwareInput)]


class Input(ctypes.Structure):
    _fields_ = [("type", ctypes.c_ulong),
                ("ii", Input_I)]

def PressKey(hexKeyCode):
    extra = ctypes.c_ulong(0)
    ii_ = Input_I()
    ii_.ki = KeyBdInput(0, hexKeyCode, 0x0008, 0, ctypes.pointer(extra) )
    x = Input( ctypes.c_ulong(1), ii_)
    ctypes.windll.user32.SendInput(1, ctypes.pointer(x), ctypes.sizeof(x))


def ReleaseKey(hexKeyCode):
    extra = ctypes.c_ulong(0)
    ii_ = Input_I()
    ii_.ki = KeyBdInput(0, hexKeyCode, 0x0008 | 0x0002, 0, ctypes.pointer(extra))  # 0x0008: KEYEVENTF_SCANCODE
    x = Input( ctypes.c_ulong(1), ii_)
    ctypes.windll.user32.SendInput(1, ctypes.pointer(x), ctypes.sizeof(x))  # 0x0002: KEYEVENTF_KEYUP


class ImageProcessor:
    def __init__(self, image = None):
        """
        :param image: PIL ImageGrab image (RGB image)
        """
        if image:
            self.image = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)

        self.character_roi = [272, 390, 375, 420]  # x1, y1, x2, y2
        self.cropped_images = []
        self.result = []
        self.roi_cropped = None
        self.tesseract_dir = r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"
        pytesseract.pytesseract.tesseract_cmd = self.tesseract_dir
        self.tessdata = os.path.join(os.getcwd(), "tessdata")


    def load_external_image(self, filename):
        file = cv2.imread(filename, cv2.IMREAD_COLOR)
        self.image = file

    def crop_characters(self):
        roi_cropped = self.image[self.character_roi[1]:self.character_roi[3], self.character_roi[0]:self.character_roi[2]]
        self.roi_cropped = roi_cropped
        roi_hsv = cv2.cvtColor(roi_cropped, cv2.COLOR_BGR2HLS)
        cv2.imwrite(os.path.join(os.getcwd(), "cropped.png"), roi_cropped)

        total_crops = 0
        scanning_char = False
        start_x = 0
        end_x = roi_cropped.shape[1]
        start_y = 0
        end_y = roi_cropped.shape[0]

        for x in range(0, roi_cropped.shape[1]):
            detected_pixels = False
            for y in range(0, roi_cropped.shape[0]):
                if roi_hsv[y][x][1] <= 200:
                    # Is the pixel a different color than white?
                    detected_pixels = True
                    #cv2.line(roi_cropped, (x,0), (x,34), (0,255,0), 2)
                    #print(x, "detected", roi_hsv[y][x][1])
                    break

            #print("x coord", x, detected_pixels)
            if scanning_char:
                if not detected_pixels:
                    print("end char", x)
                    end_x = x

                    y_min_finished = False
                    for inner_y in range(0, roi_cropped.shape[0]):
                        for inner_x in range(start_x, end_x+1):
                            if roi_hsv[inner_y][inner_x][1] <= 200:
                                start_y = inner_y
                                y_min_finished = True
                                break

                        if y_min_finished:
                            break

                    y_min_finished = False
                    for iy in range(roi_cropped.shape[0]-1, -1, -1):
                        for ix in range(start_x, end_x+1):
                            if roi_hsv[iy][ix][1] <= 200:
                                end_y = iy + 1
                                y_min_finished = True
                                break

                        if y_min_finished:
                            break


                    scanning_char = False
                    result = roi_cropped[start_y:end_y, start_x:end_x]
                    borderpad = 5
                    bordered = cv2.copyMakeBorder(result,top=borderpad, bottom=borderpad, left=borderpad, right=borderpad,borderType=cv2.BORDER_CONSTANT, value=[255,255,255])
                    cv2.imwrite("%d.png"%(total_crops), bordered)
                    self.cropped_images.append(cv2.cvtColor(bordered, cv2.COLOR_BGR2GRAY))
                    total_crops += 1

            elif not scanning_char:
                if detected_pixels:
                    start_x = x
                    print("start char", x)
                    scanning_char = True

        return total_crops

    def run_ocr(self):
        for img in self.cropped_images:
            result = pytesseract.image_to_string(img, lang="digits", config="--psm 10")
            print("result", result)
            if result:
                if result == "g":
                    self.result.append(9)
                elif result == "l":
                    self.result.append(1)
                else:
                    self.result.append(int(result))

        return self.result


if __name__ == "__main__":
    print(os.path.join(os.getcwd(), "tessdata"))
    logging.basicConfig(filename="log.log", level=logging.INFO)
    scrp = MapleScreenCapturer()
    hwnd = scrp.ms_get_screen_hwnd()
    if not hwnd:
        logging.info("no ms window")
        sys.exit()
    rect = scrp.ms_get_screen_rect(hwnd)
    img = scrp.capture(rect=rect)
    print(rect)
    processor = ImageProcessor(img)
    chars = processor.crop_characters()
    print("chars", chars)
    logging.info("number of chars"+str(chars))
    solution = processor.run_ocr()
    print("solution", solution)
    logging.info("OCR result:" + str(solution))
    for char in solution:
        PressKey(numbers[char])
        time.sleep(0.05)
        ReleaseKey(numbers[char])
        time.sleep(0.1)

    click(rect[0]+CONFIRM_BUTTON_COORDS[0], rect[1]+CONFIRM_BUTTON_COORDS[1])
    logging.info("keypress complete")
    sys.exit(0)







