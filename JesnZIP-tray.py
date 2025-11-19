"""JesnZIP tray agent

Features:
- Monitor Windows clipboard for images and file paths (images/videos).
- On new image/video, upload to s.jesn.zip and copy returned share link to clipboard.
- Show a small toast notification on successful upload.
- Tray menu: Open site, Toggle Auto-Upload, Restart, Exit
"""

import os
import sys
import time
import json
import logging
import traceback
import threading
import tempfile
import hashlib
import webbrowser
from pathlib import Path

import requests
from PIL import Image, ImageGrab
import win32clipboard
import win32con
import pystray
from pystray import MenuItem
from PIL import Image as PILImage
from win32com.client import Dispatch
import tkinter as _tk
from tkinter import simpledialog
try:
    from winrt.windows.ui.notifications import ToastNotificationManager, ToastNotification
    from winrt.windows.data.xml.dom import XmlDocument
    HAVE_WINRT = True
except Exception:
    HAVE_WINRT = False

# Logging
log_file = "JZIP-debug.log"
logging.basicConfig(filename=log_file, filemode="a", level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

BASE_DIR = Path(__file__).parent.resolve()
SETTINGS_PATH = BASE_DIR / "tray_settings.json"
UPLOAD_ENDPOINT = "https://s.jesn.zip/api/upload"
ORIGIN_HEADER = "https://s.jesn.zip"

DEFAULT_SETTINGS = {
    "auto_upload": True,
    "poll_interval": 1.0
}


def load_settings():
    try:
        if SETTINGS_PATH.exists():
            return json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
        else:
            SETTINGS_PATH.write_text(json.dumps(DEFAULT_SETTINGS), encoding="utf-8")
            return DEFAULT_SETTINGS.copy()
    except Exception as e:
        logging.error(f"Failed to load settings: {e}")
        return DEFAULT_SETTINGS.copy()


def save_settings(settings):
    try:
        SETTINGS_PATH.write_text(json.dumps(settings), encoding="utf-8")
    except Exception as e:
        logging.error(f"Failed to save settings: {e}")


settings = load_settings()




def set_clipboard_text(text: str):
    try:
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, text)
        win32clipboard.CloseClipboard()
        logging.debug(f"Set clipboard text: {text}")
    except Exception as e:
        logging.error(f"Failed to set clipboard: {e}")


def show_notification(title: str, message: str, duration: int = 4):
    try:
        # Use native WinRT toasts only
        if HAVE_WINRT:
            try:
                logging.debug("Using WinRT toast")
                tplt = f"<toast><visual><binding template='ToastGeneric'><text>{title}</text><text>{message}</text></binding></visual></toast>"
                xml = XmlDocument()
                xml.load_xml(tplt)
                notifier = ToastNotificationManager.create_toast_notifier('JesnZIP')
                toast = ToastNotification(xml)
                notifier.show(toast)
                logging.debug("WinRT toast shown")
                return
            except Exception as e:
                logging.exception(f"WinRT toast failed: {e}")
                return
        else:
            logging.warning("WinRT not available; skipping notification (winrt package missing)")
            return
    except Exception as e:
        logging.error(f"Failed to show notification: {e}")


def file_hash(path: str):
    try:
        h = hashlib.sha256()
        with open(path, "rb") as f:
            while True:
                chunk = f.read(8192)
                if not chunk:
                    break
                h.update(chunk)
        return h.hexdigest()
    except Exception:
        return None


def image_bytes_hash(img):
    try:
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tf:
            img.save(tf, format="PNG")
            tf.flush()
            h = file_hash(tf.name)
        os.unlink(tf.name)
        return h
    except Exception as e:
        logging.error(f"image_bytes_hash error: {e}")
        return None


def upload_path(path: str, filename: str = None):
    try:
        if not filename:
            filename = os.path.basename(path)
        with open(path, "rb") as f:
            files = {"file": (filename, f)}
            headers = {"origin": ORIGIN_HEADER}
            # include Authorization header when a session_key is set
            sk = settings.get('session_key')
            if sk:
                headers['Authorization'] = sk
            logging.debug(f"Uploading {path} to {UPLOAD_ENDPOINT} with headers keys: {list(headers.keys())}")
            resp = requests.post(UPLOAD_ENDPOINT, files=files, headers=headers, timeout=60)
        if resp.status_code in (200, 201):
            data = resp.json()
            url = data.get("url") or data.get("share_url") or data.get("file_url")
            logging.debug(f"Upload response: {data}")
            return True, url, data
        else:
            logging.error(f"Upload failed: {resp.status_code} {resp.text}")
            return False, None, resp.text
    except Exception as e:
        logging.error(f"Upload exception: {e}")
        return False, None, str(e)


def handle_new_file(path: str):
    logging.info(f"Detected new path to upload: {path}")
    # upload and copy result
    ok, url, data = upload_path(path)
    if ok and url:
        set_clipboard_text(url)
        show_notification("JesnZIP", "Upload completed — link copied to clipboard", duration=4)
    elif ok and not url:
        show_notification("JesnZIP", "Upload completed (no link returned)", duration=4)
    else:
        show_notification("JesnZIP: Upload failed", str(data), duration=6)


def monitor_clipboard_loop():
    last_id = None
    poll = float(settings.get("poll_interval", 1.0))
    while True:
        try:
            grabbed = ImageGrab.grabclipboard()
            if grabbed is None:
                # nothing in clipboard
                time.sleep(poll)
                continue

            # If a list of file paths
            if isinstance(grabbed, list):
                # handle first path that looks like an image or video
                for p in grabbed:
                    if not os.path.exists(p):
                        continue
                    ext = os.path.splitext(p)[1].lower()
                    if ext in ('.png', '.jpg', '.jpeg', '.bmp', '.gif') or ext in ('.mp4', '.mov', '.mkv', '.avi'):
                        identifier = f"file::{os.path.abspath(p)}::{os.path.getsize(p)}::{os.path.getmtime(p)}"
                        if identifier != last_id and settings.get("auto_upload", True):
                            last_id = identifier
                            threading.Thread(target=handle_new_file, args=(p,), daemon=True).start()
                        break
                time.sleep(poll)
                continue

            # If an image object
            if isinstance(grabbed, PILImage.Image) or hasattr(grabbed, 'save'):
                img = grabbed
                # hash image bytes to dedupe
                img_h = image_bytes_hash(img)
                if img_h and img_h != last_id and settings.get("auto_upload", True):
                    last_id = img_h
                    # save to temp file
                    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tf:
                        temp_path = tf.name
                        img.save(temp_path, format='PNG')
                    threading.Thread(target=handle_new_file, args=(temp_path,), daemon=True).start()
                time.sleep(poll)
                continue

        except Exception as e:
            logging.error(f"monitor error: {e}\n{traceback.format_exc()}")
            time.sleep(poll)


def toggle_auto_upload(icon, item):
    settings['auto_upload'] = not settings.get('auto_upload', True)
    save_settings(settings)
    # Update menu text by rebuilding menu
    try:
        icon.menu = make_menu(icon)
    except Exception:
        logging.exception("Failed to update menu after toggling auto_upload")


def prompt_for_session_key(icon, item=None):
    try:
        # Use a simple tkinter dialog to request session key
        root = _tk.Tk()
        root.withdraw()
        key = simpledialog.askstring("JesnZIP Login", "Enter session key (Authorization header):", parent=root)
        root.destroy()
        if key:
            settings['session_key'] = key.strip()
            save_settings(settings)
            show_notification("JesnZIP", "Session key saved", duration=3)
            try:
                icon.menu = make_menu(icon)
            except Exception:
                logging.exception("Failed to update menu after setting session_key")
    except Exception as e:
        logging.error(f"prompt_for_session_key failed: {e}")


def logout(icon, item=None):
    try:
        if 'session_key' in settings:
            del settings['session_key']
            save_settings(settings)
        show_notification("JesnZIP", "Logged out", duration=3)
        try:
            icon.menu = make_menu(icon)
        except Exception:
            logging.exception("Failed to update menu after logout")
    except Exception as e:
        logging.error(f"logout failed: {e}")


def _startup_shortcut_path():
    appdata = os.environ.get('APPDATA')
    if not appdata:
        return None
    startup_dir = os.path.join(appdata, 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup')
    return os.path.join(startup_dir, 'JesnZIP-tray.lnk')


def is_autostart_enabled():
    try:
        path = _startup_shortcut_path()
        if not path or not os.path.exists(path):
            return False
        shell = Dispatch('WScript.Shell')
        lnk = shell.CreateShortcut(path)
        # Check if shortcut points to current script/executable
        target = getattr(lnk, 'TargetPath', '')
        args = getattr(lnk, 'Arguments', '')
        script_path = str(os.path.join(BASE_DIR, os.path.basename(__file__)))
        # If the shortcut's arguments include our script path or target equals our exe, assume enabled
        if script_path in args or os.path.abspath(target) == os.path.abspath(sys.executable):
            return True
        return True
    except Exception:
        return False


def enable_autostart():
    try:
        path = _startup_shortcut_path()
        if not path:
            return False
        shell = Dispatch('WScript.Shell')
        shortcut = shell.CreateShortcut(path)
        # Point to python executable and pass script as argument so it works during development
        shortcut.TargetPath = sys.executable
        shortcut.Arguments = f'"{os.path.join(str(BASE_DIR), os.path.basename(__file__))}"'
        shortcut.WorkingDirectory = str(BASE_DIR)
        icon_path = BASE_DIR / 'ICON.ico'
        if icon_path.exists():
            shortcut.IconLocation = str(icon_path)
        shortcut.Save()
        return True
    except Exception as e:
        logging.error(f"enable_autostart failed: {e}")
        return False


def disable_autostart():
    try:
        path = _startup_shortcut_path()
        if path and os.path.exists(path):
            os.remove(path)
        return True
    except Exception as e:
        logging.error(f"disable_autostart failed: {e}")
        return False


def toggle_autostart(icon, item):
    current = is_autostart_enabled()
    if current:
        ok = disable_autostart()
        msg = "Autostart disabled" if ok else "Failed to disable autostart"
    else:
        ok = enable_autostart()
        msg = "Autostart enabled" if ok else "Failed to enable autostart"
    save_settings(settings)
    try:
        icon.menu = make_menu(icon)
    except Exception:
        logging.exception("Failed to update menu after toggling autostart")
    show_notification("JesnZIP", msg, duration=3)


def open_site(icon, item=None):
    webbrowser.open("https://s.jesn.zip/create")


def restart(icon, item=None):
    icon.stop()
    os.execl(sys.executable, sys.executable, *sys.argv)


def exit_app(icon, item=None):
    icon.stop()


def make_menu(icon):
    auto_label = ("Disable Auto-Upload" if settings.get('auto_upload', True) else "Enable Auto-Upload")
    autostart_label = ("Disable Autostart" if is_autostart_enabled() else "Enable Autostart")
    # Login/Logout menu item
    if settings.get('session_key'):
        auth_item = MenuItem("Logout", logout)
    else:
        auth_item = MenuItem("Login / Set Session Key", prompt_for_session_key)
    return pystray.Menu(
        MenuItem("Open s.jesn.zip", open_site),
        auth_item,
        MenuItem(auto_label, toggle_auto_upload),
        MenuItem(autostart_label, toggle_autostart),
        MenuItem("Restart", restart),
        MenuItem("Exit", exit_app)
    )


def create_icon_and_run():
    # Create a simple icon image
    icon_path = BASE_DIR / "ICON.ico"
    if icon_path.exists():
        try:
            img = PILImage.open(icon_path)
        except Exception:
            img = PILImage.new('RGB', (64, 64), color=(73, 109, 137))
    else:
        img = PILImage.new('RGB', (64, 64), color=(73, 109, 137))

    icon = pystray.Icon("JesnZIP", icon=img)
    # Set initial menu (pass the icon so labels can be computed)
    icon.menu = make_menu(icon)

    # Start clipboard monitor
    t = threading.Thread(target=monitor_clipboard_loop, daemon=True)
    t.start()

    logging.info("Starting JesnZIP tray icon and clipboard monitor")
    icon.run()


if __name__ == '__main__':
    try:
        create_icon_and_run()
    except Exception as e:
        logging.critical(f"Fatal error in tray agent: {e}\n{traceback.format_exc()}")
        raise
