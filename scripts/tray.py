import os
import sys
import signal
import psutil  # Import the psutil library
from infi.systray import SysTrayIcon
import subprocess


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

icon_path = resource_path('tray.ico')
def start_tray():
    menu_options = (("Restart", None, restart_app),)
    systray = CustomSysTrayIcon(icon_path, "My App", menu_options, on_quit=quit_app)
    systray.start()

class CustomSysTrayIcon(SysTrayIcon):
    def create_menu(self):
        """Create a custom menu without the default 'Quit' option."""
        self._menu = [(text, icon, action) for text, icon, action in self.menu_options]

def quit_app(systray):
    try:
        current_process = psutil.Process(os.getpid())
        children = current_process.children(recursive=True)  # Get all child processes
        for child in children:
            child.terminate()  # Terminate each child process
        current_process.terminate()  # Terminate the current process
    except Exception as e:
        print(f"Failed to terminate processes cleanly: {e}")
        sys.exit(1)  # Force exit if clean termination fails

def restart_app(systray):
    python = sys.executable
    os.execl(python, python, "main.py")


if __name__ == "__main__":
    start_tray()