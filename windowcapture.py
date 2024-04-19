from threading import Thread, Lock

import numpy as np
import win32con
import win32gui
import win32ui
import win32api


class WindowCapture:
    stopped = True
    lock = None
    screenshot = None
    rect = (0, 0)  # in case we need to calculate absolute x, y for the click coordinates
    w = 0
    h = 0
    hwnd = None
    cropped_x = 0
    cropped_y = 0
    offset_x = 0
    offset_y = 0

    def __init__(self, window_name=None):
        # create a thread lock object
        self.lock = Lock()

        # find the handle for the window we want to capture.
        # if no window name is given, capture the entire screen
        if window_name is None:
            self.hwnd = win32gui.GetDesktopWindow()
        else:
            self.hwnd = win32gui.FindWindow(None, window_name)

            # win32gui.SetForegroundWindow(self.hwnd)
            if not self.hwnd:
                raise Exception('Window not found: {}'.format(window_name))

        # get the window size
        window_rect = win32gui.GetWindowRect(self.hwnd)
        self.rect = (window_rect[0], window_rect[1], window_rect[2] - window_rect[0], window_rect[3] - window_rect[1])

        self.w = window_rect[2] - window_rect[0]
        self.h = window_rect[3] - window_rect[1]

        border_pixels = 0
        titlebar_pixels = 30
        self.w = self.w - border_pixels
        self.h = self.h - titlebar_pixels - border_pixels
        self.cropped_x = border_pixels
        self.cropped_y = titlebar_pixels

        # set the cropped coordinates offset so we can translate screenshot
        # images into actual screen positions
        self.offset_x = window_rect[0] + self.cropped_x
        self.offset_y = window_rect[1] + self.cropped_y

    def get_screenshot(self):
        wDC = win32gui.GetWindowDC(self.hwnd)
        dcObj = win32ui.CreateDCFromHandle(wDC)
        cDC = dcObj.CreateCompatibleDC()
        dataBitMap = win32ui.CreateBitmap()
        dataBitMap.CreateCompatibleBitmap(dcObj, self.w, self.h)
        cDC.SelectObject(dataBitMap)
        cDC.BitBlt((0, 0), (self.w, self.h), dcObj, (self.cropped_x, self.cropped_y), win32con.SRCCOPY)

        signedIntsArray = dataBitMap.GetBitmapBits(True)
        img = np.fromstring(signedIntsArray, dtype='uint8')
        img.shape = (self.h, self.w, 4)

        dcObj.DeleteDC()
        cDC.DeleteDC()
        win32gui.ReleaseDC(self.hwnd, wDC)
        win32gui.DeleteObject(dataBitMap.GetHandle())

        # Drops the alpha channel, or cv.matchTemplate() will throw an error like:
        #   error: (-215:Assertion failed) (depth == CV_8U || depth == CV_32F) && type == _templ.type()
        #   && _img.dims() <= 2 in function 'cv::matchTemplate'
        img = img[..., :3]

        # Makes image C_CONTIGUOUS to avoid errors that look like:
        #   File ... in draw_rectangles
        #   TypeError: an integer is required (got type tuple)
        # see the discussion here:
        # https://github.com/opencv/opencv/issues/14866#issuecomment-580207109
        img = np.ascontiguousarray(img)

        return img

    @staticmethod
    def list_windowname():
        def winEnumHandler(hwnd, ctx):
            if win32gui.IsWindowVisible(hwnd):
                print(hex(hwnd), win32gui.GetWindowText(hwnd))

        win32gui.EnumWindows(winEnumHandler, None)

    def get_screen_position(self, pos):
        return pos[0] + self.offset_x, pos[1] + self.offset_y

    def start(self):
        self.stopped = False
        t = Thread(target=self.run)
        t.start()

    def stop(self):
        self.stopped = True

    def get_hwnd(self):
        return self.hwnd

    def get_rect(self):
        return self.rect

    def run(self):
        while not self.stopped:
            screenshot = self.get_screenshot()
            self.lock.acquire()
            self.screenshot = screenshot
            self.lock.release()

    def show_mouse_coordinates(self):
        # Get the position of the mouse cursor
        x, y = win32api.GetCursorPos()

        # Get the dimensions of the active window
        win_x, win_y, win_width, win_height = self.rect

        # Calculate the mouse position relative to the active window
        rel_x = x - win_x
        rel_y = y - win_y

        # Print the coordinates if the mouse is within the active window
        if 0 <= rel_x <= win_width and 0 <= rel_y <= win_height:
            print(f"Mouse coordinates (Relative to Active Window): ({rel_x}, {rel_y})")
        else:
            print("Mouse is outside the active window.")

        # Sleep to avoid high CPU usage
        win32api.Sleep(100)
