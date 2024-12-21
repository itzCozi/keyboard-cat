import os
import sys
import threading
import ctypes

import pystray
from PIL import Image
from pystray import MenuItem as item

from controller import Keyboard


class Program:
  loop: bool = True
  paused: bool = False
  rest_time: int = 300  # 5 minutes
  key: int = Keyboard.vk_codes["f15"]
  stop_event: threading.Event = threading.Event()
  cancel_token: threading.Event = threading.Event()

  def prevent_multiple_instance(this) -> None:
    # Create a mutex and check if it already exists
    mutex_name = "keyboard-cat"
    ctypes.windll.kernel32.CreateMutexW(None, False, mutex_name)
    last_error = ctypes.windll.kernel32.GetLastError()
    ERROR_ALREADY_EXISTS = 183
    if last_error == ERROR_ALREADY_EXISTS:
      ctypes.windll.user32.MessageBoxW(0, "Another instance is already running.", "Error", 0x10)
      sys.exit(0)

  # This gets the path to the resource folder of our exe
  def get_resource_path(this, relative_path):
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

  def start(this) -> None:
    this.cancel_token.clear()
    this.proc()

  def stop(this) -> None:
    this.loop = False
    this.stop_event.set()
    this.cancel_token.set()

  def proc(this) -> None:
    pause_event = threading.Event()
    while this.loop:
      if not this.paused:
        pause_event.set()
        Keyboard.pressAndReleaseKey(this.key)
        if this.cancel_token.wait(this.rest_time):
          break
      else:
        pause_event.clear()
        while this.paused and this.loop:
          if this.cancel_token.wait(10):
            break
      if this.cancel_token.is_set():
        break

  def create_menu(this) -> pystray.Menu:
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

  def pause(this) -> None:
    this.paused = True
    icon.menu = this.create_menu()

  def resume(this) -> None:
    this.paused = False
    icon.menu = this.create_menu()

  def on_quit(this, icon) -> None:
    icon.stop()
    this.stop()


program = Program()
# This could be an exit point for the program
program.prevent_multiple_instance()

icon_path = program.get_resource_path("icon.ico")
image = Image.open(icon_path)
image = image.resize((64, 64))
icon = pystray.Icon("Keyboard Cat", image)
icon.menu = program.create_menu()

if __name__ == "__main__":
  threading.Thread(target=program.start).start()
  icon.run()
