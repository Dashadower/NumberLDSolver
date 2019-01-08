# -*- coding:utf-8 -*-
import cv2, win32gui, time
from PIL import ImageGrab

MAPLESTORY_WINDOW_TITLE = "MapleStory"

class MapleScreenCapturer:
    """Container for capturing MS screen"""
    def __init__(self):
        self.hwnd = None

    def ms_get_screen_hwnd(self):
        window_hwnd = win32gui.FindWindow(0, MAPLESTORY_WINDOW_TITLE)
        if not window_hwnd:
            return 0
        else:
            return window_hwnd

    def ms_get_screen_rect(self, hwnd):
        return win32gui.GetWindowRect(hwnd)

    def capture(self, set_focus=True, hwnd=None, rect=None):
        """Returns Maplestory window screenshot handle(not np.array!)
        :param set_focus : True if MapleStory window is to be focusesd before capture, False if not
        :param hwnd : Default: None Win32API screen handle to use. If None, sets and uses self.hwnd
        :param hwnd : If defined, captures specificed ScreenRect area. Else, uses MS window ms_screen_rect.
        :return : returns Imagegrab of screen (PIL Image)"""
        if hwnd:
            self.hwnd = hwnd
        if not hwnd:
            self.hwnd = self.ms_get_screen_hwnd()
        if not rect:
            rect = self.ms_get_screen_rect(self.hwnd)
        if set_focus:
            win32gui.SetForegroundWindow(self.hwnd)
            time.sleep(0.1)
        img = ImageGrab.grab(rect)
        if img:
            return img
        else:
            return 0

dt = MapleScreenCapturer()
dt.capture().save("bb.png")