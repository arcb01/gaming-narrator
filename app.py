import sys, pyautogui
import pygame, win32api, win32con, win32gui
import ctypes
import pyscreenshot as ImageGrab
import os
from pynput import mouse
import sys
from PIL import Image
from io import BytesIO
import win32clipboard
import win32gui, win32com.client
import pygame
import win32api
from pynput import keyboard
import win32con
import win32gui
import time, pyautogui
from pygame import gfxdraw
from pynput.mouse import Listener as MouseListener
from pynput.keyboard import Listener as KeyboardListener
import pynput.mouse as mouse
import pynput.keyboard as keyboard
import time,webbrowser
import sys

from TTS import Narrator
from ocr import OCR


READ_CONTENT = 'º'
REPEAT_KEY = 'r'
QUIT_KEY = 'esc'


class App:
    """
    Class that runs the entire application

    `Attributes:`
        bbox_color: color of the bounding box that is drawn around the text
        rect_width: width of the bounding box
        narrator: An object of the Narrator class that is used to read the text
        OCR: An object of the OCR class that is used to detect text
        clock: Pygame clock object

    `Methods:`
        get_disp_size(): Get the width and height of the display
        get_mouse_pos(): Returns the x and y coordinates of the mouse position
        check_events(): Checks for keyboard events
        load_display(): Loads the display window
        draw_detection(detection): For a given detection, draw a bounding box around the text
        run(): Main loop of the application
    """

    def __init__(self, narrator: object, OCR: object):
        self.bbox_color = (170, 255, 0)
        self.rect_width = 4
        self.narrator = narrator
        self.OCR = OCR
        self.clock = pygame.time.Clock()
        self.transparent = True
        self.mode = win32con.LWA_COLORKEY
        self.alpha = 0
        self.reading_mode = None

    def get_disp_size(self):
        """
        Returns the width and height of the display
        """

        user32 = ctypes.windll.user32
        dwidth, dheight = user32.GetSystemMetrics(0), user32.GetSystemMetrics(1)
        return dwidth, dheight

    def get_mouse_pos(self):
        """
        Returns the x and y coordinates of the mouse position
        """

        x_mouse_pos = pyautogui.position().x
        y_mouse_pos = pyautogui.position().y
        return x_mouse_pos, y_mouse_pos

    def quit(self):
        print("\n ==== Stopping... ====\n")

        self.narrator.say("Quitting Narrator")
        # Delete all images and quit
        self.OCR.delete_imgs()
        sys.exit()

    def interactions(self):
        # Put current window on foreground
        self.minimize()
        # Start OCR
        self.OCR.start()
        # Draw all detections
        for detection in self.OCR.get_all_detections():
            self.draw_detection(detection, all_det=True)

    def load_display(self):
        """
        Loads the display window
        """

        pygame.init()
        info = pygame.display.Info()
        w = info.current_w
        h = info.current_h
        self.screen = pygame.display.set_mode((w, h), pygame.NOFRAME) # For borderless, use pygame.NOFRAME
        done = False
        self.fuchsia = (255, 0, 128)  # Transparency color
        dark_red = (255, 0, 0)
        # Create layered window
        win32gui.SetWindowPos(pygame.display.get_wm_info()['window'], win32con.HWND_TOPMOST, 0,0,0,0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
        hwnd = pygame.display.get_wm_info()["window"]
        win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE,
                            win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE) | win32con.WS_EX_LAYERED)
        # Set window transparency color
        win32gui.SetLayeredWindowAttributes(hwnd, win32api.RGB(*self.fuchsia), self.alpha, self.mode) # NOTE: Transparent
        self.clear_screen = lambda : self.screen.fill(self.fuchsia)

    def check_keyevents(self):

        with keyboard.Events() as keyboardEvents:
            keyboardEvent = keyboardEvents.get(0.1)
            if keyboardEvent != None and "Press" in str(keyboardEvent):
                if keyboardEvent.key == keyboard.Key.f4:
                    self.mode = win32con.LWA_ALPHA
                    self.alpha = 50
                    self.transparent = False
                    self.interactions()
                if keyboardEvent.key == keyboard.Key.esc:
                    self.mode = win32con.LWA_COLORKEY
                    self.alpha = 0
                    self.transparent = True

    def check_mouseevents(self):
        with mouse.Events() as mouseEvents:
            mouseEvent = mouseEvents.get(0.1)
            if mouseEvent != None:
                if "Click" in str(mouseEvent) and not self.transparent:
                    x, y = mouseEvent.x, mouseEvent.y
                    print(x, y)
                else:
                    return


    def draw_detection(self, detection : list, all_det=False):
        """
        For a given detection, draw a bounding box around the text
        :param detection: a list containing the bounding box vertices, text, and confidence
        """

        #if not all_det:
            #self.clear_screen()

        bbox = detection[0]
        self.output_text = detection[1]
        top = bbox[0][0]
        left = bbox[0][1]
        width = bbox[1][0] - bbox[0][0]
        height = bbox[2][1] - bbox[1][1]

        #s = pygame.Surface((width,height), pygame.SRCALPHA, 32)   # per-pixel alpha
        #s.fill((255, 0, 128, 80))                         # notice the alpha value in the color
        #self.screen.blit(s, (top, left))

        pygame.draw.rect(self.screen, self.bbox_color,  pygame.Rect(top, left, 
                                                            width, height), 
                        self.rect_width)
        pygame.display.update()

    def minimize(self):
        """
        toplist = []
        winlist = []

        def enum_callback(hwnd, results):
            winlist.append((hwnd, win32gui.GetWindowText(hwnd)))

        win32gui.EnumWindows(enum_callback, toplist)
        window = [(hwnd, title) for hwnd, title in winlist if 'resident' in title.lower()]
        window_id = window[0]
        win32gui.SetForegroundWindow(window_id[0])
        """
        # use the window handle to set focus
        active_window = win32gui.GetForegroundWindow()
        shell = win32com.client.Dispatch("WScript.Shell")
        shell.SendKeys('%')
        win32gui.SetForegroundWindow(active_window)

    def check_transparency(self):
        if self.transparent:
            self.screen.fill(self.fuchsia)  # Transparent background
        else:
            self.screen.fill((255, 255, 255))

    def run(self):
        """
        Main loop of the application
        """

        print("\n ==== App is running... ====")
        while True:
            self.load_display()
            self.check_keyevents()
            self.check_transparency()
            self.check_mouseevents()

            pygame.display.update()
            self.clock.tick(60)

            
if __name__ == "__main__":

    lang = "es"
    en_voice = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices\Tokens\TTS_MS_EN-US_DAVID_11.0"
    es_voice = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices\Tokens\TTS_MS_ES-ES_HELENA_11.0"

    if lang == "en":
        voice = en_voice
    elif lang == "es":
        voice = es_voice

    tts = Narrator(voice=voice)
    ocr = OCR(lang=lang)
    a = App(tts, ocr)
    a.run()


    # ========== TODO ==========
    # 1. Play the game and see what is needed


    # ========== BUG ==========
    # 4. Pressing m too consistantly causes the program to bug out

    # ========== FIXME ==========
    # 1. Display issues with the calling order

    # ========== FUTURE WORK ==========
    # 1. Add a GUI
    # 2. Only way of speeding up is speeding up EasyOCR or replacing it.
    #   OCR Alternatives:
            # 1. https://github.com/PaddlePaddle/PaddleOCR
            # 2. https://github.com/mindee/doctr

    # 3. Language understanding for bulding sentences

   # ========== References ==========
   # 1. https://github.com/nathanaday/RealTime-OCR
