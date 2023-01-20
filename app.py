import sys, keyboard, pyautogui
import pygame, win32api, win32con, win32gui
import win32gui, win32com.client

from TTS import Narrator
from ocr import OCR
from utils.utils import *

# ======== Key bindings ========
START_READING = 'º'
SWITCH_DET_FORWARD = "flecha derecha"
SWITCH_DET_BACKWARD = "flecha izquierda"
REPEAT_KEY = 'r'
READ_NEAREST = 'h'
READ_OUT_LOUD = '-'
QUIT_KEY = 'esc'


class App:
    """
    Class that runs the entire application

    `Attributes:`
        rect_width: width of the bounding box
        narrator: An object of the Narrator class that is used to read the text
        OCR: An object of the OCR class that is used to detect text
        clock: Pygame clock object
        switch_detection: Boolean that enables swtiching between detections
        det_idx: Index of the current detection
        highlighted_color: Color that highlights the current bounding box
        dimmed_color: Color of the unselected bounding boxes

    `Methods:`
        check_events(): Checks for keyboard events
        load_display(): Loads the display window
        draw_detection(detection): For a given detection, draw a bounding box around the text
        run(): Main loop of the application
    """

    def __init__(self, narrator: object, OCR: object):
        self.app_name = "Gaming Narrator"
        self.app_logo = pygame.image.load('./assets/logo.png')
        self.rect_width = 4
        self.narrator = narrator
        self.OCR = OCR
        self.clock = pygame.time.Clock()
        self.switch_detection = False
        self.engaging = False
        self.dimmed_color = (65, 94, 0)
        self.highlighted_color = (170, 255, 0)
        pygame.init()

    def quit(self):
        try:
            if len(self.OCR.get_all_detections()) > 0:
                self.OCR.empty_results()
                self.narrator.say("Clearing screen.")
                self.clear_screen()
                pygame.display.update()
                # Delete all images 
                self.OCR.delete_imgs()
        except:
            pass

    def end_of_list(self):
        return True if self.det_idx == len(self.OCR.get_all_detections()) - 1 else False

    def check_events(self):
        """
        Captures any keyboard events
        """

        event = keyboard.read_event()

        if event.event_type == keyboard.KEY_DOWN and event.name == QUIT_KEY:
            self.quit()

        if event.event_type == keyboard.KEY_DOWN and event.name == READ_NEAREST:
            # Get mouse position
            x, y = get_mouse_pos()
            # Load display
            self.load_display()
            # Start OCR
            #if self.OCR.check_start():
            self.OCR.start()
            #else:
                #pass
            # Find nearest detection from the mouse position
            nearest_detection = self.OCR.find_nearest_detection(x, y)
            # Draw detection
            self.draw_detection(nearest_detection)
            # Text to speech
            self.narrator.say(self.output_text)

        if event.event_type == keyboard.KEY_DOWN and event.name == START_READING:
            self.det_idx = 0
            self.engaging = True
            self.load_display()
            # Start OCR
            self.OCR.send_screen(self.screen) # For adding a loading screen.
            self.OCR.start()
            if len(self.OCR.get_all_detections()) > 0:
                # Draw all possible OCR detections
                for detection in self.OCR.get_all_detections():
                    self.draw_detection(detection, color=self.dimmed_color)
                # Highlight first detection
                self.draw_detection(self.OCR.get_all_detections()[self.det_idx], color=self.highlighted_color)


        if event.event_type == keyboard.KEY_DOWN and event.name in [SWITCH_DET_FORWARD, SWITCH_DET_BACKWARD]:
            assert len(self.OCR.get_all_detections()) > 0, "No detections found yet. Please start scanning first."

            if event.name == SWITCH_DET_BACKWARD:
                if not self.end_of_list():
                    if self.det_idx > 0:
                        # Apply dimmed color to current detection
                        self.draw_detection(self.OCR.get_all_detections()[self.det_idx], color=self.dimmed_color)
                        # Apply highlighted color to previous detection
                        self.draw_detection(self.OCR.get_all_detections()[self.det_idx - 1], color=self.highlighted_color)
                        self.det_idx  -= 1

            elif event.name == SWITCH_DET_FORWARD:
                if not self.end_of_list():
                    self.det_idx  += 1
                    # Apply highlighted color to current detection
                    self.draw_detection(self.OCR.get_all_detections()[self.det_idx], color=self.highlighted_color)
                    # Apply dimmed color to previous detection
                    self.draw_detection(self.OCR.get_all_detections()[self.det_idx - 1], color=self.dimmed_color)
                else:
                    # Apply dimmed color to previous detection
                    self.draw_detection(self.OCR.get_all_detections()[self.det_idx], color=self.dimmed_color)
                    self.det_idx = 0
                    # Apply highlighted color to current detection
                    self.draw_detection(self.OCR.get_all_detections()[self.det_idx], color=self.highlighted_color)

        if event.event_type == keyboard.KEY_DOWN and event.name == READ_OUT_LOUD:
            assert len(self.OCR.get_all_detections()) > 0, "No detections found yet. Please start scanning first."
            # Read text of current detection
            text = self.OCR.get_all_detections()[self.det_idx][1]
            self.narrator.say(text)

        if event.event_type == keyboard.KEY_DOWN and event.name == REPEAT_KEY:
            assert len(self.OCR.get_all_detections()) > 0, "No detections found yet. Please start scanning first."
            # Repeat text a little bit slower
            text = self.OCR.get_all_detections()[self.det_idx][1]
            self.narrator.slower_saying(text)

    def load_display(self):
        """
        Loads the display window
        """
        # https://stackoverflow.com/questions/550001/fully-transparent-windows-in-pygame

        pygame.init()
        w, h = get_disp_size()
        self.screen = pygame.display.set_mode((w, h)) # For borderless, use pygame.NOFRAME
        pygame.display.set_caption(self.app_name)
        pygame.display.set_icon(self.app_logo)
        fuchsia = (255, 0, 128)  # Transparency bbox_color
        # Lock window on top
        win32gui.SetWindowPos(pygame.display.get_wm_info()['window'], win32con.HWND_TOPMOST, 0,0,0,0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
        # Create layered window
        hwnd = pygame.display.get_wm_info()["window"]
        win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE,
                            win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE) | win32con.WS_EX_LAYERED)
        # Set window transparency bbox_color
        win32gui.SetLayeredWindowAttributes(hwnd, win32api.RGB(*fuchsia), 0, win32con.LWA_COLORKEY)

        self.screen.fill(fuchsia)  # Transparent background
        self.clear_screen = lambda : self.screen.fill(fuchsia)

        pygame.display.update()


    def draw_detection(self, detection : list, color : tuple = (170, 255, 0)):
        """
        For a given detection, draw a bounding box around the text
        :param detection: a list containing the bounding box vertices, text, and confidence
        :param color: the color of the bounding box
        """
    
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

        pygame.draw.rect(self.screen, color,  pygame.Rect(top, left, 
                                                            width, height), 
                        self.rect_width)
        pygame.display.update()


    def run(self):
        """
        Main loop of the application
        """

        print("\n ==== App is running... ====")
        while True:
            self.check_events()
            self.clock.tick(60)

#            pygame.display.flip()
            
            

        
        
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
    # 1. Documentation

    # ========== FIXME ==========
    # 1. Fullscreen issues with some apps
    # 4. NEAREST DETECTION func interaction doesn't work well

    # ========== BUG ==========
    # Sometimes While display loaded, Clicking on the top part of the screen, app crashes

    # ========== FUTURE WORK ==========
    # 1. Add a GUI (key binding selector, etc.)
    # 2. Improve Narrator voices
    # 3. Ways of speeding up.
    #    Using OCR Alternatives:
            # 1. https://github.com/PaddlePaddle/PaddleOCR
            # 2. https://github.com/mindee/doctr
    #    PyTorch 2.0

   # ========== References ==========
   # 1. https://github.com/nathanaday/RealTime-OCR
