# config.py
import asyncio
import os

# Default TTS + OCR configuration
TTS_VOICE = "Microsoft Guy(Natural)"
TTS_VOLUME = "30"
TTS_RATE = "1"
TTS_USE_TCP = True
TTS_SERVER_HOST = "127.0.0.1"
TTS_SERVER_PORT = 5150
TTS_SERVER_AUTO_START = True

CROP_LEFT = 75
CROP_TOP = 110
CROP_RIGHT = 20
CROP_BOTTOM = 50

ocr_lock = asyncio.Lock()

# Paths (you can adjust these)
TESSERACT_PATH = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
TTS_EXE_PATH = (
    r"C:\Programming\CPP\NaturalVoiceSAPIAdapter\ttsapplication\TTSApplicationSample\x64\Debug\TtsApplication.exe"
)

# Store discovered voices in a local directory next to this module
# e.g. c:\Programming\Python\kindleReader\src\voices\voices.json
VOICES_DIR = os.path.join(os.path.dirname(__file__), 'voices')
VOICES_FILE = os.path.join(VOICES_DIR, 'voices.json')