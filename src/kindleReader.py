# kindle_reader.py
import asyncio
import re
# from pywinauto import Application # needed if you dont have ctypes.windll.user32.SetProcessDPIAware()
from pygetwindow import getAllWindows
import win32gui
import win32con
import win32ui
from PIL import Image
import pytesseract
import win32api
import pytesseract
from config import *
from utils import resource_path
import ctypes

pytesseract.pytesseract.tesseract_cmd = resource_path(TESSERACT_PATH)
tts_server_proc = None
DEBUG_SAVE = False
DEBUG_DIR = r"c:\Programming\Python\kindleReader\debug_images"

# make python aware of our dpi (needed for screens that dont have scaling set to 100%, like 4k/2k screens at 125/150% scaling)
try:
    ctypes.windll.user32.SetProcessDPIAware()
except Exception:
    pass


# -------------------- Screenshot --------------------
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

# -------------------- Kindle Controls --------------------
def find_kindle_window():
    for w in getAllWindows():
        if "Kindle for PC" in w.title:
            return w
    return None


def turn_page_bg(hwnd):
    win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RIGHT, 0)
    win32api.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RIGHT, 0)


# -------------------- TTS --------------------
async def start_tts_server_once():
    global tts_server_proc
    if tts_server_proc is None and TTS_USE_TCP and TTS_SERVER_AUTO_START:
        try:
            tts_server_proc = await asyncio.create_subprocess_exec(
                resource_path(TTS_EXE_PATH),
                "--server", "--voice", TTS_VOICE,
                "--rate", TTS_RATE, "--volume", TTS_VOLUME
            )
            await asyncio.sleep(0.5)
        except Exception as e:
            print(f"Failed to start TTS server: {e}")
            tts_server_proc = None


async def speak_async(text):
    if TTS_USE_TCP:
        await start_tts_server_once()
        try:
            reader, writer = await asyncio.open_connection(TTS_SERVER_HOST, TTS_SERVER_PORT)
            safe_text = text.replace("\n", " ").replace("\r", " ")
            writer.write((safe_text + "\n").encode("utf-8"))
            await writer.drain()
            writer.close()
            await writer.wait_closed()
        except Exception as e:
            print(f"TTS TCP send error: {e}")


# -------------------- OCR + Reading --------------------
def estimate_speech_duration(text, rate=1):
    words = len(text.split())
    base_wpm = 210
    wpm = base_wpm * float(rate)
    return (words / wpm) * 60


async def read_kindle_text_async(kindle_window):
    if not kindle_window:
        print("Kindle window not found.")
        return None

    screenshot = capture_window_bg(kindle_window._hWnd,
                                   crop=(CROP_LEFT, CROP_TOP, CROP_RIGHT, CROP_BOTTOM))
    # show the screenshot for debugging (will bring window to foreground)

    try:
        async with ocr_lock:
            text = await asyncio.to_thread(pytesseract.image_to_string, screenshot)
    except Exception as e:
        print(f"OCR error: {e}")
        return None

    text = re.sub(r'\s+', ' ', text).strip()
    text = text.replace("- ", "")
    text = text.replace("“", '"').replace("”", '"')
    text = text.replace("‘", "'").replace("’", "'")

    if not text:
        print("No text detected on this page.")
        return None

    print("Detected text:\n", text)
    await speak_async(text)
    return text


async def main_loop(stop_event):
    kindle_win = find_kindle_window()
    if not kindle_win:
        print("Kindle window not found. Open Kindle and try again.")
        return

    while not stop_event.is_set():
        text = await read_kindle_text_async(kindle_win)
        if text:
            duration = estimate_speech_duration(text, TTS_RATE)
            await asyncio.sleep(duration)
            turn_page_bg(kindle_win._hWnd)
        else:
            await asyncio.sleep(1)
