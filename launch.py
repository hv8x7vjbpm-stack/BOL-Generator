"""
launch.py — Main entry point for BOL Generator desktop app.

Features:
- Flask runs in background thread
- pywebview shows the app in a native window
- pystray puts it in the system tray when minimised
- On quit: auto-saves Excel + PDF backup to backups/ folder
- Single instance lock prevents multiple copies running
"""
import os
import sys
import time
import json
import threading
import tempfile
import datetime
import signal
import logging

# ── Logging setup ─────────────────────────────────────────────────────────────
log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
os.makedirs(log_dir, exist_ok=True)
logging.basicConfig(
    filename=os.path.join(log_dir, "bol_generator.log"),
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s"
)
log = logging.getLogger("BOLLauncher")

# ── Resolve base path (works for both .py and PyInstaller .exe) ───────────────
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
    # PyInstaller unpacks to _MEIPASS for read-only resources
    RESOURCE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    RESOURCE_DIR = BASE_DIR

# ── Single instance lock ───────────────────────────────────────────────────────
LOCK_FILE = os.path.join(BASE_DIR, ".bol_lock")

def acquire_lock():
    if os.path.exists(LOCK_FILE):
        try:
            with open(LOCK_FILE) as f:
                pid = int(f.read().strip())
            # Check if that PID is still running
            os.kill(pid, 0)
            return False  # Already running
        except (ValueError, OSError):
            pass  # Stale lock, continue
    with open(LOCK_FILE, "w") as f:
        f.write(str(os.getpid()))
    return True

def release_lock():
    try:
        os.unlink(LOCK_FILE)
    except Exception:
        pass

# ── Backup on quit ────────────────────────────────────────────────────────────
def do_auto_backup():
    """Save Excel + PDF backup when app closes."""
    try:
        backup_dir = os.path.join(BASE_DIR, "backups")
        os.makedirs(backup_dir, exist_ok=True)
        stamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

        # Import app functions
        from app import load_store, build_excel_shortage, generate_bols_pdf

        store = load_store()
        if not store:
            log.info("Auto-backup: no data to save")
            return

        # Save Excel
        xlsx_src = build_excel_shortage(store)
        xlsx_dst = os.path.join(backup_dir, f"BOL_Backup_{stamp}.xlsx")
        import shutil
        shutil.copy2(xlsx_src, xlsx_dst)
        os.unlink(xlsx_src)
        log.info(f"Auto-backup Excel: {xlsx_dst}")

        # Save PDF
        pdf_src = generate_bols_pdf(store)
        pdf_dst = os.path.join(backup_dir, f"BOL_Backup_{stamp}.pdf")
        shutil.copy2(pdf_src, pdf_dst)
        os.unlink(pdf_src)
        log.info(f"Auto-backup PDF: {pdf_dst}")

        log.info(f"Auto-backup complete: {len(store)} BOLs saved")
    except Exception as e:
        log.error(f"Auto-backup failed: {e}")


# ── Flask runner ──────────────────────────────────────────────────────────────
flask_thread = None

def run_flask():
    # Change working directory so Flask finds templates and BOL INPUT.docx
    os.chdir(BASE_DIR)
    from app import app as flask_app
    flask_app.run(host="127.0.0.1", port=5001, debug=False, use_reloader=False, threaded=True)


# ── System tray ───────────────────────────────────────────────────────────────
tray_icon = None
window_ref = None

def make_tray_icon():
    """Create a simple tray icon image using PIL."""
    from PIL import Image, ImageDraw, ImageFont
    img = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    # Dark blue circle background
    draw.ellipse([4, 4, 60, 60], fill=(26, 54, 93, 255))
    # White "B" letter
    try:
        font = ImageFont.truetype("arial.ttf", 32)
    except Exception:
        font = ImageFont.load_default()
    draw.text((18, 12), "B", fill=(255, 255, 255, 255), font=font)
    return img


def show_window(icon=None, item=None):
    """Bring the webview window to front."""
    global window_ref
    if window_ref:
        try:
            window_ref.show()
            window_ref.restore()
        except Exception as e:
            log.error(f"show_window: {e}")


def quit_app(icon=None, item=None):
    """Save backup then exit cleanly."""
    log.info("Quit requested — running auto-backup")
    if tray_icon:
        tray_icon.stop()
    do_auto_backup()
    release_lock()
    os._exit(0)


def setup_tray():
    """Set up and run system tray icon in its own thread."""
    global tray_icon
    try:
        import pystray
        from pystray import MenuItem as Item, Menu

        icon_image = make_tray_icon()
        menu = Menu(
            Item("Open BOL Generator", show_window, default=True),
            Menu.SEPARATOR,
            Item("Save Backup Now", lambda i, it: threading.Thread(target=do_auto_backup, daemon=True).start()),
            Menu.SEPARATOR,
            Item("Quit", quit_app),
        )
        tray_icon = pystray.Icon(
            "BOL Generator",
            icon_image,
            "BOL Generator — Jackson Pottery",
            menu
        )
        tray_icon.run()
    except Exception as e:
        log.error(f"Tray setup failed: {e}")


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    global window_ref, flask_thread

    # Single instance check
    if not acquire_lock():
        import ctypes
        ctypes.windll.user32.MessageBoxW(
            0,
            "BOL Generator is already running.\nCheck your system tray.",
            "BOL Generator",
            0x40  # MB_ICONINFORMATION
        )
        sys.exit(0)

    log.info("BOL Generator starting")

    # Start Flask
    flask_thread = threading.Thread(target=run_flask, daemon=True)
    flask_thread.start()
    log.info("Flask started")

    # Wait for Flask to be ready
    import urllib.request
    for _ in range(20):
        try:
            urllib.request.urlopen("http://127.0.0.1:5001", timeout=1)
            break
        except Exception:
            time.sleep(0.5)

    # Start system tray in background thread
    tray_thread = threading.Thread(target=setup_tray, daemon=True)
    tray_thread.start()

    # Create webview window
    import webview

    def on_closed():
        """Window X button — minimise to tray instead of quitting."""
        log.info("Window closed — minimising to tray")
        # Don't quit, just hide — tray icon stays active

    def on_loaded():
        log.info("Window loaded")

    window_ref = webview.create_window(
        title="BOL Generator — Jackson Pottery Inc",
        url="http://127.0.0.1:5001",
        width=1280,
        height=900,
        min_size=(960, 700),
        resizable=True,
        on_top=False,
    )
    window_ref.events.closed += on_closed
    window_ref.events.loaded += on_loaded

    # Start webview (blocks until all windows closed)
    webview.start(debug=False)

    # When webview exits fully, do backup and quit
    log.info("Webview exited — running auto-backup")
    do_auto_backup()
    release_lock()
    if tray_icon:
        tray_icon.stop()
    sys.exit(0)


if __name__ == "__main__":
    main()
