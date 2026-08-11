"""Tray mandiri untuk memantau penggunaan RAM sistem Windows."""

import ctypes
import logging
import os
import sys
import threading
import time

try:
    import psutil
except ImportError:
    psutil = None


# pythonw.exe tidak memiliki stdout/stderr console.
if sys.stdout is None:
    sys.stdout = open(os.devnull, "w")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w")


# Mutex mandiri agar lifecycle tray lain tidak saling mengikat.
_mutex = ctypes.windll.kernel32.CreateMutexW(
    None, False, "POKJA2026RamMonitorTrayMutex"
)
if ctypes.windll.kernel32.GetLastError() == 183:  # ERROR_ALREADY_EXISTS
    sys.exit(0)


_log_dir = os.path.join(
    os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "POKJA2026"
)
os.makedirs(_log_dir, exist_ok=True)
_log_file = os.path.join(_log_dir, "ram-monitor.log")
logging.basicConfig(
    filename=_log_file,
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    encoding="utf-8",
)
logger = logging.getLogger("pokja-ram-tray")

from infi.systray import SysTrayIcon


_stop = threading.Event()
_tray = None


def _format_gb(value):
    return f"{value / (1024 ** 3):.1f}".replace(".", ",")


def _ram_tooltip():
    if psutil is None:
        return "RAM: psutil tidak tersedia"
    try:
        memory = psutil.virtual_memory()
        return (
            f"RAM: {_format_gb(memory.used)}/{_format_gb(memory.total)} GB "
            f"({memory.percent:.0f}%)"
        )
    except Exception as exc:
        logger.warning("Gagal membaca RAM: %s", exc)
        return "RAM: n/a"


def _monitor_ram():
    while not _stop.is_set():
        tray = _tray
        if tray is None:
            return
        try:
            tray.update(hover_text=_ram_tooltip())
        except Exception as exc:
            logger.warning("Gagal memperbarui tooltip RAM: %s", exc)
        if _stop.wait(5):
            return


def _open_log(systray):
    if os.path.exists(_log_file):
        os.startfile(_log_file)


def _quit(systray):
    global _tray
    _stop.set()
    _tray = None
    logger.info("RAM monitor tray berhenti.")


def main():
    global _tray
    menu = (("Buka Log RAM", None, _open_log),)
    icon_path = os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "ram_monitor.ico"
    )
    icon = icon_path if os.path.exists(icon_path) else None
    _tray = SysTrayIcon(icon, _ram_tooltip(), menu, on_quit=_quit)
    _tray.start()
    logger.info("RAM monitor tray aktif.")
    threading.Thread(target=_monitor_ram, name="ram-monitor", daemon=True).start()

    try:
        while _tray is not None:
            time.sleep(1)
    except (KeyboardInterrupt, SystemExit):
        _stop.set()


if __name__ == "__main__":
    main()
