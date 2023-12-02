import os
import random
import time
import pyautogui
from time import sleep

import win32api
import win32con
from pywin.framework import app
import pygetwindow as gw

from core import winapi
# from pyWinActivate import win_activate, win_wait_active
from tkinter import *

import winsound

from core.pyWinActive import win_activate

duration = 1000  # milliseconds
freq = 440  # Hz

# Giant dictonary to hold key name and VK value
VK_CODE = {'backspace': 0x08,
           'tab': 0x09,
           'clear': 0x0C,
           'enter': 0x0D,
           'shift': 0x10,
           'ctrl': 0x11,
           'alt': 0x12,
           'pause': 0x13,
           'caps_lock': 0x14,
           'esc': 0x1B,
           'spacebar': 0x20,
           'page_up': 0x21,
           'page_down': 0x22,
           'end': 0x23,
           'home': 0x24,
           'left_arrow': 0x25,
           'up_arrow': 0x26,
           'right_arrow': 0x27,
           'down_arrow': 0x28,
           'select': 0x29,
           'print': 0x2A,
           'execute': 0x2B,
           'print_screen': 0x2C,
           'ins': 0x2D,
           'del': 0x2E,
           'help': 0x2F,
           '0': 0x30,
           '1': 0x31,
           '2': 0x32,
           '3': 0x33,
           '4': 0x34,
           '5': 0x35,
           '6': 0x36,
           '7': 0x37,
           '8': 0x38,
           '9': 0x39,
           'a': 0x41,
           'b': 0x42,
           'c': 0x43,
           'd': 0x44,
           'e': 0x45,
           'f': 0x46,
           'g': 0x47,
           'h': 0x48,
           'i': 0x49,
           'j': 0x4A,
           'k': 0x4B,
           'l': 0x4C,
           'm': 0x4D,
           'n': 0x4E,
           'o': 0x4F,
           'p': 0x50,
           'q': 0x51,
           'r': 0x52,
           's': 0x53,
           't': 0x54,
           'u': 0x55,
           'v': 0x56,
           'w': 0x57,
           'x': 0x58,
           'y': 0x59,
           'z': 0x5A,
           'numpad_0': 0x60,
           'numpad_1': 0x61,
           'numpad_2': 0x62,
           'numpad_3': 0x63,
           'numpad_4': 0x64,
           'numpad_5': 0x65,
           'numpad_6': 0x66,
           'numpad_7': 0x67,
           'numpad_8': 0x68,
           'numpad_9': 0x69,
           'multiply_key': 0x6A,
           'add_key': 0x6B,
           'separator_key': 0x6C,
           'subtract_key': 0x6D,
           'decimal_key': 0x6E,
           'divide_key': 0x6F,
           'F1': 0x70,
           'F2': 0x71,
           'F3': 0x72,
           'F4': 0x73,
           'F5': 0x74,
           'F6': 0x75,
           'F7': 0x76,
           'F8': 0x77,
           'F9': 0x78,
           'F10': 0x79,
           'F11': 0x7A,
           'F12': 0x7B,
           'F13': 0x7C,
           'F14': 0x7D,
           'F15': 0x7E,
           'F16': 0x7F,
           'F17': 0x80,
           'F18': 0x81,
           'F19': 0x82,
           'F20': 0x83,
           'F21': 0x84,
           'F22': 0x85,
           'F23': 0x86,
           'F24': 0x87,
           'num_lock': 0x90,
           'scroll_lock': 0x91,
           'left_shift': 0xA0,
           'right_shift ': 0xA1,
           'left_control': 0xA2,
           'right_control': 0xA3,
           'left_menu': 0xA4,
           'right_menu': 0xA5,
           'browser_back': 0xA6,
           'browser_forward': 0xA7,
           'browser_refresh': 0xA8,
           'browser_stop': 0xA9,
           'browser_search': 0xAA,
           'browser_favorites': 0xAB,
           'browser_start_and_home': 0xAC,
           'volume_mute': 0xAD,
           'volume_Down': 0xAE,
           'volume_up': 0xAF,
           'next_track': 0xB0,
           'previous_track': 0xB1,
           'stop_media': 0xB2,
           'play/pause_media': 0xB3,
           'start_mail': 0xB4,
           'select_media': 0xB5,
           'start_application_1': 0xB6,
           'start_application_2': 0xB7,
           'attn_key': 0xF6,
           'crsel_key': 0xF7,
           'exsel_key': 0xF8,
           'play_key': 0xFA,
           'zoom_key': 0xFB,
           'clear_key': 0xFE,
           '+': 0xBB,
           ',': 0xBC,
           '-': 0xBD,
           '.': 0xBE,
           '/': 0xBF,
           '`': 0xC0,
           ';': 0xBA,
           '[': 0xDB,
           '\\': 0xDC,
           ']': 0xDD,
           "'": 0xDE,
           '`': 0xC0}


def press(*args):
    '''
    one press, one release.
    accepts as many arguments as you want. e.g. press('left_arrow', 'a','b').
    '''
    for i in args:
        win32api.keybd_event(VK_CODE[i], 0, 0, 0)
        time.sleep(.05)
        win32api.keybd_event(VK_CODE[i], 0, win32con.KEYEVENTF_KEYUP, 0)


def pressAndHold(*args):
    '''
    press and hold. Do NOT release.
    accepts as many arguments as you want.
    e.g. pressAndHold('left_arrow', 'a','b').
    '''
    for i in args:
        win32api.keybd_event(VK_CODE[i], 0, 0, 0)
        time.sleep(.05)


def pressHoldRelease(*args):
    '''
    press and hold passed in strings. Once held, release
    accepts as many arguments as you want.
    e.g. pressAndHold('left_arrow', 'a','b').

    this is useful for issuing shortcut command or shift commands.
    e.g. pressHoldRelease('ctrl', 'alt', 'del'), pressHoldRelease('shift','a')
    '''
    for i in args:
        win32api.keybd_event(VK_CODE[i], 0, 0, 0)
        time.sleep(.05)

    for i in args:
        win32api.keybd_event(VK_CODE[i], 0, win32con.KEYEVENTF_KEYUP, 0)
        time.sleep(.1)


w = winapi.WinApi()
w.find_window_wildcard(".*Tibia.*")


class game:
    global w

    def __init__(self):
        self._handle = None

    def eat_food(self):
        try:
            foodlist = ['fish']
            for i in foodlist:
                pos = pyautogui.locateOnScreen(i + ".png")
                w.RightClick((pos.left, pos.top))
        except:
            Exception("Food Eater")

    def useOnLocation(self, x, y, name):
        try:
            loc = pyautogui.locateOnScreen(name + '.png')
            w.UseOn((loc.left, loc.top), (x, y))
        except:
            Exception("Use on Location")

    def getLife(self):
        # health = pyautogui.locateOnScreen("health.png", grayscale=True, confidence=.5)
        #
        # if (pyautogui.pixelMatchesColor(int(health.left + 194), int(health.top + 40), (22, 23, 21), tolerance=10)):#zycie
        #  if (pyautogui.pixelMatchesColor(int(health.left + 194), int(health.top + 60), (22, 23, 21), tolerance=10)):#mana
        #       print('jestesmy zgubieni')
        #       winsound.Beep(freq, duration)
        #       return 70
        #   pyautogui.getWindowsWithTitle("Realesta 7.4 - Efemeryda")[0].maximize()
        try:

            battle = pyautogui.locateOnScreen("battle.png", grayscale=True, confidence=.5)
            # pass
            if not (pyautogui.pixelMatchesColor(int(battle.left + 40), int(battle.top + 122), (29, 30, 27),
                                                tolerance=10)):  # local
                # if not (pyautogui.pixelMatchesColor(int(battle.left + 42), int(battle.top + 118), (29, 30, 27), tolerance=10)):  # remote
                # win32api.SendMessage(self._handle, win32con.LEFT_CTRL_PRESSED, 0)
                pressHoldRelease('ctrl', 'l')
                pressHoldRelease('ctrl', 'l')
                for x in range(5):
                    pass
                    # pyautogui.click(battle)
                    winsound.Beep(freq, duration)
                # time.sleep(900)

            #     # press('enter')

            # self.dancing()
            # time.sleep(1)
            return 0
            # elif (pyautogui.pixelMatchesColor((health.left + 94), (health.top + 7), (219, 79, 79))):
            #     return 903
            # elif (pyautogui.pixelMatchesColor((health.left + 74), (health.top + 7), (219, 79, 79))):
            #     return 70
            # elif (pyautogui.pixelMatchesColor((health.left + 54), (health.top + 7), (219, 79, 79))):
            #     return 50
            # elif (pyautogui.pixelMatchesColor((health.left + 34), (health.top + 7), (219, 79, 79))):
            #     return 30
            # elif (pyautogui.pixelMatchesColor((health.left + 24), (health.top + 7), (219, 79, 79))):
            #     return 20
            # elif (pyautogui.pixelMatchesColor((health.left + 14), (health.top + 7), (219, 79, 79))):
            #     return 15
            # elif (pyautogui.pixelMatchesColor((health.left + 5), (health.top + 7), (219, 79, 79))):
            #     return 10
        except Exception as e:
            # time.sleep(1)
            # win_activate(window_title="Realesta 7.4 - Efemeryda", partial_match=True)
            # time.sleep(1)
            return 0

    def makeStuff(self):
        while True:
            # spij 14 minut + od 1 do 30 sekund
            #sleep(14 * 60 + self.totoLoto(self, 1, 20))
            try:
                #sleep(10)
                # self.dragRune(self)
                # self.dragRune(self)
                #self.eatMeat(self)
                tibiantis = gw.getWindowsWithTitle("Tibiantis")[0]
                tibiantis.activate()
                time.sleep(2)
                myScreenshot = pyautogui.screenshot()
                myScreenshot.save('lol.png')
            except Exception as e:
                time.sleep(2)
                #win_activate(window_title="Tibiantis", partial_match=True)
                tibiantis = gw.getWindowsWithTitle("Tibiantis")[0]
                # tibiantis.restore()
                # tibiantis.minimize()
                # tibiantis.maximize()
                # tibiantis.
                #tibiantis.show()
                tibiantis.activate()
                time.sleep(2)
                myScreenshot = pyautogui.screenshot()
                myScreenshot.save('lol.png')
                return 0

    def totoLoto(self, fro, to):
        return random.randint(fro, to)

    def dragRune(self):
        blank = pyautogui.locateOnScreen("blank.png", grayscale=False, confidence=.99)
        leftHand = pyautogui.locateOnScreen("leftHand.png", grayscale=False, confidence=.9)
        pyautogui.moveTo(blank.left, blank.top)
        sleep(self.totoLoto(self, 1, 4))
        pyautogui.dragTo(leftHand.left, leftHand.top, button='left', duration=0.8)
        sleep(self.totoLoto(self, 1, 4))
        pyautogui.hotkey("f10")
        sleep(self.totoLoto(self, 1, 4))
        pyautogui.dragTo(blank.left, blank.top, button='left', duration=0.8)
        sleep(self.totoLoto(self, 1, 2))

    def eatMeat(self):
        meat = pyautogui.locateOnScreen("ham.png", grayscale=False, confidence=.75)
        pyautogui.moveTo(meat.left, meat.top)
        pyautogui.rightClick()
        pyautogui.rightClick()
        pyautogui.rightClick()
        pyautogui.rightClick()
        sleep(self.totoLoto(self, 1, 4))

    def getMana(self):
        health = pyautogui.locateOnScreen("mana.png")
        if (pyautogui.pixelMatchesColor((health.left + 105), (health.top + 5), (67, 64, 192))):
            return 100
        elif (pyautogui.pixelMatchesColor((health.left + 94), (health.top + 5), (67, 64, 192))):
            return 90
        elif (pyautogui.pixelMatchesColor((health.left + 74), (health.top + 5), (67, 64, 192))):
            return 70
        elif (pyautogui.pixelMatchesColor((health.left + 54), (health.top + 5), (67, 64, 192))):
            return 50
        elif (pyautogui.pixelMatchesColor((health.left + 34), (health.top + 7), (219, 79, 79))):
            return 30
        elif (pyautogui.pixelMatchesColor((health.left + 24), (health.top + 7), (219, 79, 79))):
            return 20
        elif (pyautogui.pixelMatchesColor((health.left + 14), (health.top + 7), (219, 79, 79))):
            return 15
        elif (pyautogui.pixelMatchesColor((health.left + 5), (health.top + 28), (219, 79, 79))):
            return 10
        return 0

    def heal(self, percent, spell, time):
        # life = self.getLife(self)
        while True:
            # w.Write(spell)
            sleep(time)
            life = self.getLife(self)
            # self.makeStuff(self)

    def mp(self, percent, spell, time):
        mana = self.getMana(self)
        while mana > percent:
            w.Write(spell)
            sleep(time)
            mana = self.getMana(self)

    def dancing(self):
        pyautogui.hotkey("ctrl", "up")
        pyautogui.hotkey("ctrl", "left")
        pyautogui.hotkey("ctrl", "right")
        pyautogui.hotkey("ctrl", "down")

    def equipItem(self, name, slot):
        ref = pyautogui.locateOnScreen("slots.png")
        w.find_window_wildcard(".*Tibia.*")
        if (slot == "head"):
            item = pyautogui.locateOnScreen(name + ".png", grayscale=True)
            w.MouseDrag((item.left, item.top), ((ref.left - 22, ref.top + 19)))
        elif (slot == "armor"):
            item = pyautogui.locateOnScreen(name + ".png", grayscale=True)
            w.MouseDrag((item.left, item.top), ((ref.left - 21, ref.top + 50)))
        elif (slot == "legs"):
            item = pyautogui.locateOnScreen(name + ".png", grayscale=True)
            w.MouseDrag((item.left, item.top), ((ref.left - 21, ref.top + 90)))
        elif (slot == "boots"):
            item = pyautogui.locateOnScreen(name + ".png", grayscale=True)
            w.MouseDrag((item.left, item.top), ((ref.left - 21, ref.top + 130)))
        elif (slot == "necklace"):
            item = pyautogui.locateOnScreen(name + ".png", grayscale=True)
            w.MouseDrag((item.left, item.top), ((ref.left - 57, ref.top + 25)))
        elif (slot == "lhand"):
            item = pyautogui.locateOnScreen(name + ".png", grayscale=True)
            w.MouseDrag((item.left, item.top), ((ref.left - 55, ref.top + 70)))
        elif (slot == "rhand"):
            item = pyautogui.locateOnScreen(name + ".png", grayscale=True)
            w.MouseDrag((item.left, item.top), ((ref.left - 20, ref.top + 70)))
        elif (slot == "ring"):
            item = pyautogui.locateOnScreen(name + ".png", grayscale=True)
            w.MouseDrag((item.left, item.top), ((ref.left - 57, ref.top + 100)))
        elif (slot == "backpack"):
            item = pyautogui.locateOnScreen(name + ".png", grayscale=True)
            w.MouseDrag((item.left, item.top), ((ref.left - 20, ref.top + 25)))
        elif (slot == "free"):
            item = pyautogui.locateOnScreen(name + ".png", grayscale=True)
            w.MouseDrag((item.left, item.top), ((ref.left - 20, ref.top + 100)))

    def making_runes(self, rune_spell, rune_time):
        while rune_spell != "":
            try:
                pyautogui.locateOnScreen("mana.png", grayscale=True)
                # Make the rune
                w.Write(self.spell)
                sleep(1)
                self.eat_food()
            except:
                print("Except: making runes")
            print("Making runes: " + rune_spell)
            time.sleep(rune_time)

    def fishing(self):
        try:
            loc = pyautogui.locateOnScreen("water.png", confidence=.8)
            self.useOnLocation(self, loc.left, loc.top, "fishing")
        except:
            Exception("Fishing")
