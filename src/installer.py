import os
import requests
import shutil
import ctypes
import threading
import winshell
import win32com
from win32com.client import Dispatch


def download_file(url: str, save_path: str) -> None:
  response: requests.Response = requests.get(url, stream=True)
  with open(save_path, 'wb') as file:
    shutil.copyfileobj(response.raw, file)
  del response


def create_shortcut(target: str, shortcut_path: str, description: str = "") -> None:
  shell: win32com.client.CDispatch = Dispatch('WScript.Shell')
  shortcut: win32com.client.CDispatch = shell.CreateShortCut(shortcut_path)
  shortcut.TargetPath = target
  shortcut.WorkingDirectory = os.path.dirname(target)
  shortcut.Description = description
  shortcut.save()


def main() -> None:
  home_dir: str = os.path.expanduser("~")
  install_dir: str = os.path.join(home_dir, "keyboardcat")
  exe_path: str = os.path.join(install_dir, "kbdcat.exe")
  shortcut_path: str = os.path.join(winshell.desktop(), "Keyboard Cat.lnk")

  if not os.path.exists(install_dir):
    os.makedirs(install_dir)

  github_url: str = "https://github.com/itzCozi/keyboard-cat/releases/download/1.0/kdbcat.exe"
  download_file(github_url, exe_path)
  create_shortcut(exe_path, shortcut_path, "Keyboard Cat")
  message_thread: threading.Thread = threading.Thread(
    target=lambda: ctypes.windll.user32.MessageBoxW(0,
                                                    "Installation complete. Run Keyboard Cat from your desktop whenever!",
                                                    "Finished",
                                                    0x40))
  message_thread.start()


if __name__ == "__main__":
  main()
