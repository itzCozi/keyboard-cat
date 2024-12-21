import argparse
import ctypes
import os
import sys
import threading
from typing import *

import pystray
from PIL import Image
from pystray import MenuItem as item

from controller import Keyboard


class Program:

  def __init__(this: Self) -> None:
    parser: argparse.ArgumentParser = argparse.ArgumentParser(description="Program to run with a specified key.")
    parser.add_argument('--key', type=str, default='f15',
                        help='Key to press (EX: f12, f15, a, b, c, etc.) (default: f15)')
    parser.add_argument('--interval', type=int, default=300,
                        help='Time between each keystroke in seconds (default: 300)')
    parser.add_argument('--paused', type=bool, default=False,
                        help='Will start the program paused if True (default: False)')
    args: argparse.Namespace = parser.parse_args()
    if args.key.lower() not in Keyboard.vk_codes:
      ctypes.windll.user32.MessageBoxW(0, "Invalid key specified.", "Error", 0x10)
      sys.exit(0)
    elif args.interval < 1:
      ctypes.windll.user32.MessageBoxW(0, "Specified interval is less than one.", "Error", 0x10)
      sys.exit(0)

    this.loop: bool = True
    this.paused: bool = args.paused
    this.rest_time: int = args.interval
    this.key: int = Keyboard.vk_codes[args.key.lower()]
    this.stop_event: threading.Event = threading.Event()
    this.cancel_token: threading.Event = threading.Event()
    # Prevent multiple instances of the program
    this.prevent_multiple_instance()
    message_thread: threading.Thread = threading.Thread(
      target=lambda: ctypes.windll.user32.MessageBoxW(0,
                                                      "Keyboard Cat is now running in your system tray, right click it to learn more.",
                                                      "Meow :3",
                                                      0x40))
    message_thread.start()

  def prevent_multiple_instance(this: Self) -> None:
    # Create a mutex and check if it already exists
    mutex_name: str = "keyboard-cat"
    ctypes.windll.kernel32.CreateMutexW(None, False, mutex_name)
    last_error: int = ctypes.windll.kernel32.GetLastError()
    ERROR_ALREADY_EXISTS: int = 183
    if last_error == ERROR_ALREADY_EXISTS:
      ctypes.windll.user32.MessageBoxW(0, "Another instance is already running.", "Error", 0x10)
      sys.exit(0)

  def get_resource_path(this: Self, relative_path: str) -> str:
    base_path: str = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

  def start(this: Self) -> None:
    this.cancel_token.clear()
    this.proc()

  def stop(this: Self) -> None:
    this.loop: bool = False
    this.stop_event.set()
    this.cancel_token.set()

  def proc(this: Self) -> None:
    pause_event: threading.Event = threading.Event()
    while this.loop:
      if not this.paused:
        pause_event.set()
        if this.cancel_token.wait(this.rest_time):
          break
        Keyboard.pressAndReleaseKey(this.key)
      else:
        pause_event.clear()
        while this.paused and this.loop:
          if this.cancel_token.wait(10):
            break
      if this.cancel_token.is_set():
        break

  def create_menu(this: Self) -> pystray.Menu:
    if this.paused:
      return pystray.Menu(
        item("Resume", lambda: this.resume()),
        item("Quit", lambda: this.on_quit(icon)),
      )
    else:
      return pystray.Menu(
        item("Pause", lambda: this.pause()),
        item("Quit", lambda: this.on_quit(icon)),
      )

  def pause(this: Self) -> None:
    this.paused: bool = True
    icon.menu = this.create_menu()

  def resume(this: Self) -> None:
    this.paused: bool = False
    icon.menu = this.create_menu()

  def on_quit(this: Self, icon: pystray.Icon) -> None:
    icon.stop()
    this.stop()


program: Program = Program()
icon_path: str = program.get_resource_path("icon.ico")
image: Image.Image = Image.open(icon_path)
image: Image.Image = image.resize((64, 64))
icon: pystray.Icon = pystray.Icon("Keyboard Cat", image)
icon.menu = program.create_menu()

if __name__ == "__main__":
  threading.Thread(target=program.start).start()
  icon.run()
