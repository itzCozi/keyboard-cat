import os
import requests
import shutil
import winshell
from win32com.client import Dispatch


def download_file(url, save_path):
  response = requests.get(url, stream=True)
  with open(save_path, 'wb') as file:
    shutil.copyfileobj(response.raw, file)
  del response


def create_shortcut(target, shortcut_path, description=""):
  shell = Dispatch('WScript.Shell')
  shortcut = shell.CreateShortCut(shortcut_path)
  shortcut.TargetPath = target
  shortcut.WorkingDirectory = os.path.dirname(target)
  shortcut.Description = description
  shortcut.save()


def main():
  program_files = os.environ["ProgramFiles"]
  install_dir = os.path.join(program_files, "keyboardcat")
  exe_path = os.path.join(install_dir, "kbdcat.exe")
  shortcut_path = os.path.join(winshell.desktop(), "Keyboard Cat.lnk")

  if not os.path.exists(install_dir):
    os.makedirs(install_dir)

  github_url = "https://github.com/itzcozi/keyboard-cat/releases/download/v1.0.0/kdbcat.exe"
  download_file(github_url, exe_path)
  create_shortcut(exe_path, shortcut_path, "Keyboard Cat")


if __name__ == "__main__":
  main()
