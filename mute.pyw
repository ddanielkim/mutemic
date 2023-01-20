import win32api
import win32gui
from pynput import keyboard
from pygame import mixer
import configparser
from infi.systray import SysTrayIcon
import time

# Initialize the mixer
mixer.init()

# Loads config and defines structure
config = configparser.ConfigParser()
config.read('resources/config.ini')
volume = float(config['DEFAULT']['volume'])
hotkey = config['DEFAULT']['hotkey']
mute_sound = config['DEFAULT']['mute_sound']
unmute_sound = config['DEFAULT']['unmute_sound']

# Set the volume for the sound files
ms = mixer.Sound(mute_sound)
ums = mixer.Sound(unmute_sound)
ms.set_volume(volume)
ums.set_volume(volume)


# Create the tray icon and set the unmuted icon
icon_muted = 'resources/muted.ico'
icon_unmuted = 'resources/unmuted.ico'

# Defines a global variable 'm' to keep track of mute/unmute status
m = 0

def reset_state():
    global m
    m = 0
    tray.update(icon=icon_unmuted)

def reset_state_wrapper(event):
    reset_state()

def mute():
    global m
    if m == 0:
        ms.play()
        tray.update(icon=icon_muted)
    else:
        ums.play()
        tray.update(icon=icon_unmuted)
    WM_APPCOMMAND = 0x319
    APPCOMMAND_MICROPHONE_VOLUME_MUTE = 0x180000

    hwnd_active = win32gui.GetForegroundWindow()
    win32api.SendMessage(hwnd_active, WM_APPCOMMAND, None, APPCOMMAND_MICROPHONE_VOLUME_MUTE)
    m ^= 1

def mute_wrapper(event):
    mute()

# unmutes before quitting
def quit_program(systray):
    global m
    if m == 1:
        mute()
    time.sleep(1)
    tray.shutdown()

menu_options = (("Mute/Unmute", None, mute_wrapper), ("Reset state to unmuted", None, reset_state_wrapper), ("Quit", None, quit_program),)
tray = SysTrayIcon(icon_unmuted, "Mute/Unmute", menu_options)
tray.start()

h = keyboard.GlobalHotKeys({hotkey: mute})
h.start()