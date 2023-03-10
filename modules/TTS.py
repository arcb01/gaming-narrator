import pyttsx3


class Narrator:
    """
    Class used to read text

    `Attributes:`
        engine: pyttsx3 engine
        voice_speed: Playback speed of the voice

    `Methods:`
        say(): Playsback the text
        stop(): Stop playback
        slower_saying(): Repeats the text at a slower rate
    """

    def __init__(self, voice, voice_speed = 160):
        self._engine = pyttsx3.init()
        self._engine.setProperty('voice', str(voice))
        self.voice_speed = voice_speed

    def say(self, text: str):
        """
        Given a text, read it out loud
        :param text: Text to be read
        """

        self._engine.setProperty('rate', self.voice_speed)
        self._engine.say(text)
        self._engine.runAndWait()

    def stop(self):
        """
        Stop the voice engine
        """

        self._engine.stop()

    def slower_saying(self, text: str):
        """
        Given a text, read it out loud at a slower rate
        :param text: Text to be read
        """

        self._engine.stop()
        self._engine.setProperty('rate', self.voice_speed * 0.5)
        self._engine.say(text)
        self._engine.runAndWait()