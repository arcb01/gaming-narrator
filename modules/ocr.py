import pyautogui
import random
import cv2
import easyocr
import os, glob
from .utils.utils import loading_screen, clear_screen


class OCR:
    """
    Class that captures the screen and performs OCR

    `Attributes:`
        lang: language in which the OCR will be performed
        gpu: use GPU or not (recommended to use GPU)
        imgs: List containing the filenames of the screenshots
        result: List of tuples (bbox, text, prob)
        imgs_dir: Directory where screenshots will be stored
        file_nom: Nomenclature of the screenshots

    `Methods:`
        images_dir(): Makes sure that the imgs directory exists
        take_screenshot(): Take a screenshot and save it in imgs_dir
        start(): Start OCR detection and save results
        delete_imgs(): Deletes all screenshots when the program finishes running
    """

    def __init__(self, lang="en", gpu=True):
        self.lang = lang
        self.gpu = gpu
        self.imgs = []
        self.result = []
        self.imgs_dir = "./imgs/" # NOTE: Directory where images are stored
        self.file_nom = "OCR_pic_"
        self.images_dir()

    def images_dir(self):
        """
        Makes sure that the imgs directory exists
        """
            
        if not os.path.exists(self.imgs_dir):
            os.makedirs(self.imgs_dir)
        
    def take_screenshot(self):
        """
        Function that takes a screenshot and save it in imgs_dir
        """
        myScreenshot = pyautogui.screenshot()
        h = str(random.getrandbits(128))
        filename = self.file_nom + h + ".png"
        myScreenshot.save(self.imgs_dir + filename)
        self.imgs.append(filename)
    
    def send_screen(self, screen):
        """
        Receives the screen from the main app
        """
        self.app_screen = screen

    def start(self):
        """
        Start OCR detection and save results
        """

        # Take screenshot
        self.take_screenshot()
        # Loading screen
        loading_screen(self.app_screen)
        # Get last img
        #self.last_img = self.imgs.pop()
        img = self.imgs[-1]
        # Read img
        imgf = cv2.imread(self.imgs_dir + img)
        # OCR
        reader = easyocr.Reader([self.lang], gpu=self.gpu)
        self.result = reader.readtext(imgf, paragraph=True)
        # Clear loading screen
        clear_screen(self.app_screen)

    def get_all_detections(self):
        """
        Returns the list of all detections
        """

        return self.result

    def empty_results(self):
        """
        Empty the list of results
        """

        self.result = []
    
    def delete_imgs(self):
        """
        Deletes all screenshots taken
        """

        for filename in glob.glob(self.imgs_dir + f"{self.file_nom}*"):
            os.remove(filename) 
