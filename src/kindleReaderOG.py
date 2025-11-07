import asyncio
import re
from pywinauto import Application
import pygetwindow as gw
import win32gui
import win32con
import win32ui
from PIL import Image
import pytesseract
import win32api
from pytesseract import TesseractError
import sys, os

def resource_path(rel_path):
    # When PyInstaller bundles, it extracts to _MEIPASS; otherwise use script dir.
    base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel_path)

# -------------------- Configuration --------------------
# make local paths later
pytesseract.pytesseract.tesseract_cmd = resource_path(r"C:\Program Files\Tesseract-OCR\tesseract.exe")


TTS_EXE = resource_path(r"C:\Programming\CPP\NaturalVoiceSAPIAdapter\ttsapplication\TTSApplicationSample\x64\Debug\TtsApplication.exe")


CROP_LEFT = 75
CROP_TOP = 110
CROP_RIGHT = 20
CROP_BOTTOM = 50

ocr_lock = asyncio.Lock()
TTS_VOICE = "Microsoft Guy(Natural)"
TTS_VOLUME = "30"
TTS_RATE = "1"

TTS_USE_TCP = True
TTS_SERVER_HOST = "127.0.0.1"
TTS_SERVER_PORT = 5150
TTS_SERVER_AUTO_START = True

# need a better way to move forward pages at the right time, dont want to wait for the server to completely 
# finish the text, but dont want to get to far ahead of the server by moving forward independently

# package to exe
#pyinstaller --onefile --console ^
#   --add-binary "C:\\Program Files\\Tesseract-OCR\\tesseract.exe;." ^
#   --add-data   "C:\\Program Files\\Tesseract-OCR\\tessdata;tessdata" ^
#   --add-binary "C:\\Programming\CPP\\NaturalVoiceSAPIAdapter\\ttsapplication\\TTSApplicationSample\\x64\\Debug\\TtsApplication.exe;." ^
#   kindleReader.py

# -------------------- Background Screenshot --------------------
def capture_window_bg(hwnd, crop=None):
    rect = win32gui.GetClientRect(hwnd)
    width = rect[2]
    height = rect[3]

    hwnd_dc = win32gui.GetWindowDC(hwnd)
    mfc_dc = win32ui.CreateDCFromHandle(hwnd_dc)
    save_dc = mfc_dc.CreateCompatibleDC()
    save_bitmap = win32ui.CreateBitmap()
    save_bitmap.CreateCompatibleBitmap(mfc_dc, width, height)
    save_dc.SelectObject(save_bitmap)

    import ctypes
    ctypes.windll.user32.PrintWindow(hwnd, save_dc.GetSafeHdc(), 1)

    bmpinfo = save_bitmap.GetInfo()
    bmpstr = save_bitmap.GetBitmapBits(True)
    img = Image.frombuffer('RGB', (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
                           bmpstr, 'raw', 'BGRX', 0, 1)

    win32gui.DeleteObject(save_bitmap.GetHandle())
    save_dc.DeleteDC()
    mfc_dc.DeleteDC()
    win32gui.ReleaseDC(hwnd, hwnd_dc)

    if crop:
        left, top, right, bottom = crop
        img = img.crop((left, top, width - right, height - bottom))

    return img

# -------------------- TTS --------------------
async def _maybe_start_tts_server():
    if not TTS_SERVER_AUTO_START:
        return None
    try:
        args = [TTS_EXE, "--server", "--voice", TTS_VOICE, "--rate", TTS_RATE, "--volume", TTS_VOLUME]
        proc = await asyncio.create_subprocess_exec(*args)
        await asyncio.sleep(0.2)
        return proc
    except Exception as e:
        print(f"Failed to auto-start TTS server: {e}")
        return None
    
tts_server_proc = None  # global handle

async def start_tts_server_once():
    global tts_server_proc
    if tts_server_proc is None and TTS_USE_TCP and TTS_SERVER_AUTO_START:
        try:
            tts_server_proc = await asyncio.create_subprocess_exec(
                TTS_EXE,
                "--server",
                "--voice", TTS_VOICE,
                "--rate", TTS_RATE,
                "--volume", TTS_VOLUME
            )
            await asyncio.sleep(0.5)  # give server time to start
        except Exception as e:
            print(f"Failed to start TTS server: {e}")
            tts_server_proc = None


async def speak_async(text):
    if TTS_USE_TCP:
        await start_tts_server_once()
        try:
            reader, writer = await asyncio.open_connection(TTS_SERVER_HOST, TTS_SERVER_PORT)
            # Flatten text: remove internal line breaks so it’s one chunk
            safe_text = text.replace("\n", " ").replace("\r", " ")

            # Send as a single message to the server
            writer.write((safe_text + "\n").encode('utf-8'))  # keep newline for server
            await writer.drain()
            try:
                writer.close()
                await writer.wait_closed()
            except Exception:
                pass
            return
        except Exception as e:
            print(f"TTS TCP send error: {e}")
            return
    else:
        print("TTS_USE_TCP is disabled and no fallback is configured.")
        return

def estimate_speech_duration(text, rate=1):
    words = len(text.split())
    base_wpm = 210
    try:
        r = float(rate)
    except Exception:
        try:
            r = float(TTS_RATE)
        except Exception:
            r = 1.0
    wpm = base_wpm * r
    if wpm <= 0:
        wpm = base_wpm
    duration_seconds = (words / wpm) * 60
    print(f"Estimated time: {duration_seconds:.2f} seconds")
    return duration_seconds

# -------------------- Kindle Window --------------------
def find_kindle_window():
    for w in gw.getAllWindows():
        if "Kindle" in w.title:
            return w
    return None

def turn_page_bg(hwnd):
    win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RIGHT, 0)
    win32api.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RIGHT, 0)

# -------------------- Reader --------------------
async def read_kindle_text_async(kindle_window):
    if not kindle_window:
        print("Kindle window not found.")
        return None  # return None if no text

    screenshot = capture_window_bg(kindle_window._hWnd,
                                   crop=(CROP_LEFT, CROP_TOP, CROP_RIGHT, CROP_BOTTOM))
    screenshot.show()
    try:
        async with ocr_lock:
            text = await asyncio.to_thread(pytesseract.image_to_string, screenshot)
    except Exception as e:
        print(f"OCR error: {e}")
        return None

    text = re.sub(r'\s+', ' ', text).strip()
    text = text.replace("- ", "")
    text = text.replace("“", "\"").replace("”", "\"")
    text = text.replace("‘", "\'").replace("’", "\'")

    if not text:
        print("No text detected on this page.")
        return None

    print("Detected text:\n", text)
    await speak_async(text)
    return text  # return the text so main_loop can calculate duration


# -------------------- Main Loop --------------------
async def main_loop():
    kindle_win = find_kindle_window()
    if not kindle_win:
        print("Kindle window not found. Open Kindle and try again.")
        return

    try:
        while True:
            text = await read_kindle_text_async(kindle_win)
            if text:
                # Estimate how long TTS will take
                duration = estimate_speech_duration(text)
                buffer = 0  # seconds to subtract before turning page
                await asyncio.sleep(max(duration - buffer, 0.5))

                # Turn the page AFTER waiting
                turn_page_bg(kindle_win._hWnd)
                await asyncio.sleep(0.5)
            else:
                await asyncio.sleep(0.5)
    except asyncio.CancelledError:
        print("Main loop cancelled. Exiting...")
    except KeyboardInterrupt:
        print("Exiting...")


if __name__ == "__main__":
    print("Starting background Kindle TTS reader. Press Ctrl+C to stop.")
    asyncio.run(main_loop())
