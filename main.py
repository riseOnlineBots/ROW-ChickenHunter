import os
import random
import time
from datetime import datetime, timedelta
from threading import Thread, Event as ThreadEvent

import win32api
import win32con
import win32gui

import keyboard_detector
from admin_privileges import running_as_admin
from device_validation import DeviceValidation
from windowcapture import WindowCapture

os.chdir(os.path.dirname(os.path.abspath(__file__)))

registered_devices = ['D8-BB-C1-17-F1-9E', '04-7C-16-5B-06-1D']

device_registration = DeviceValidation(registered_devices)
running_as_admin()

keyboard_monitor = keyboard_detector.KeyboardDetector()
keyboard_monitor.start()

window_name = 'Rise Online Client'
window = WindowCapture(window_name=window_name)
hwnd = window.get_hwnd()

R = 0x52
Z = 0x5A
ONE = 0x31
current_camp = 1
loading_screen_duration = 22 + 5  # 5 seconds for character flashing
next_camp_change_timedelta = 33
next_camp_change = datetime.now() + timedelta(seconds=next_camp_change_timedelta)


def sleep(min_delay=0.1, max_delay=0.5):
    if min_delay > max_delay:
        min_delay, max_delay = max_delay, min_delay

    time.sleep(random.uniform(min_delay, max_delay))


def press_key(key, min_delay=0.1, max_delay=0.5):
    win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, key, 0)
    sleep(min_delay=min_delay, max_delay=max_delay)
    win32api.SendMessage(hwnd, win32con.WM_KEYUP, key, 0)


def mouse_click(target_x=None, target_y=None, times=1):
    if target_x is None or target_y is None:
        raise ValueError("Both x and y must be provided.")

    # Calling win32gui allows us to re-calculate coordinates dynamically in case the window is moved
    # window.get_rect() calculates the window position once it's initialised.
    rect = win32gui.GetWindowRect(hwnd)
    win_x, win_y, _, _ = rect

    rel_target_x = target_x + win_x
    rel_target_y = target_y + win_y

    # lParam = win32api.MAKELONG(rel_target_x, rel_target_y)
    # win32api.SendMessage(hwnd, win32con.WM_MOUSEMOVE, 0, lParam)
    # button_down_msg = None
    # button_up_msg = None

    # button_down_msg = win32con.WM_LBUTTONDOWN
    # button_up_msg = win32con.WM_LBUTTONUP
    # button_down_msg = win32con.WM_RBUTTONDOWN
    # button_up_msg = win32con.WM_RBUTTONUP

    # if button_down_msg is not None:
    #     win32api.SendMessage(hwnd, button_down_msg, win32con.MK_LBUTTON, lParam)

    # if button_up_msg is not None:
    #     win32api.SendMessage(hwnd, button_up_msg, 0, lParam)

    win32api.SetCursorPos((rel_target_x, rel_target_y))
    sleep(min_delay=0.1, max_delay=0.2)

    for i in range(times):
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        sleep(min_delay=0.1, max_delay=0.2)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)


def get_camp_header():
    return 1802, 53  # Fixed x, y for 1920x1080, because I am tired of integrating OpenCV.


def click_camp_header():
    x, y = get_camp_header()

    mouse_click(target_x=x, target_y=y)


def click_next_camp():
    click_x, click_y = get_camp_header()

    _current_camp = 1 if current_camp == max_camp else current_camp

    next_camp_y = click_y + (20 * _current_camp)

    mouse_click(target_x=click_x, target_y=next_camp_y, times=3)


def run_bot(_max_camp):
    global next_camp_change, current_camp

    if datetime.now() >= next_camp_change:
        click_camp_header()
        sleep(min_delay=0.1, max_delay=0.2)
        click_next_camp()

        time.sleep(loading_screen_duration)

        next_camp_change = datetime.now() + timedelta(seconds=next_camp_change_timedelta)

        if current_camp >= max_camp:
            current_camp = 1
        else:
            current_camp += 1
    else:
        press_key(key=Z)
        press_key(key=R)
        press_key(key=ONE)


def run_thread(_pause_event, _max_camp):
    while True:
        if keyboard_monitor.get_combination_active():
            _pause_event.set()
        else:
            if _pause_event.is_set():
                _pause_event.clear()

            run_bot(_max_camp)


if __name__ == '__main__':
    if device_registration.is_device_legal():
        # print('Set the resolution windowed 1920x1080')
        # print("If you haven't set the resolution yet, close the program and try again.")
        # print('To pause/resume the bot, simultaneously press RIGHT CTRL and RIGHT ALT.')
        # print('The bot begins in 3 seconds')

        print('1920x1080 pencere modunda ayarla oyunu')
        print('Henuz cozunurlugu ayarlamadiysan, programi kapatip tekrar ac')
        print('Duraklatmak/devam ettirmek icin SAG CTRL ve SAG ALT tuslarina aynanda bas')
        print('Bot 3 saniye icinde basliyor')
        time.sleep(3)

        win32gui.SetForegroundWindow(hwnd)
        max_camp = 3
        pause_event = ThreadEvent()
        thread = Thread(target=run_thread, args=(pause_event, max_camp,))
        thread.start()
