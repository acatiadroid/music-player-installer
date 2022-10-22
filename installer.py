# Installer for Music Player 

import os
import platform
import requests
import threading
import winshell
import pythoncom
import http.client as httpc
from time import sleep
from win32com.client import Dispatch
from subprocess import Popen
from tkinter import *
from tkinter.font import Font
from tkinter.messagebox import showerror

os_used = platform.system()
dest = str(os.environ['LOCALAPPDATA'])
executable = "https://github.com/acatiadroid/music-player/releases/download/installation_source/Music.Player.exe"
repo = "https://github.com/acatiadroid/music-player"
startmenu = str(winshell.start_menu())
desktop = winshell.desktop()
sep = "-----------------------------------------------------\n\n" # just for separating info

app = Tk()
app.wm_title("Install Music Player")

def check_network_connection():
    connection = httpc.HTTPConnection("google.com", timeout=10)

    try:
        connection.request("HEAD", "/")
        connection.close()
        return True
    except:
        return False 

def install():
    """Installs the music player"""
    pythoncom.CoInitialize() # since being used in a thread
    txt = Text(app)
    txt.pack()

    def write(text):
        txt.insert("end", f"{text}\n")

    if os.path.exists(dest + r"\music-player"):
        folder = dest + '\music-player'
        return write(sep + f"Download stopped! The following folder already exists:\n{folder}\nThis likely means the music player is already installed.")    
        
    internet = check_network_connection()

    if not internet:
        write(sep + "Download FAILED! You don't have an internet connection. Please establish a network connection and retry.")
        return showerror("No network connection", "You don't have an internet connection, which is required to run this installer.")

    write("Installation started.")

    write("Downloading files from GitHub... (may take a moment)")
    p = Popen(["git", "clone", repo], cwd=dest)
    p.wait()
    p.kill()

    req = requests.get(executable)

    with open(dest + r"\music-player\Music Player.exe", "wb") as f:
        f.write(req.content)

    write("Download completed. Creating shortcut(s)...")

    sleep(1) # just giving some time before continuing

    path = os.path.join(startmenu, "Music Player.lnk")
    target = dest + r"\music-player\Music Player.exe"
    wdir = dest + r"\music-player"
    
    shortcuts = []

    shortcuts.append("Start Menu")
    shell = Dispatch("WScript.Shell")
    shortcut = shell.CreateShortcut(path)
    shortcut.Targetpath = target
    shortcut.WorkingDirectory = wdir
    shortcut.IconLocation = dest + r"\music-player\player\Assets\musical_note.ico"
    shortcut.save()

    if cvar.get() == 1:
        # create desktop shortcut
        path = os.path.join(desktop, "Music Player.lnk")
        target = dest + r"\music-player\Music Player.exe"
        wdir = dest + r"\music-player"
        
        shell = Dispatch("WScript.Shell")
        shortcut = shell.CreateShortcut(path)
        shortcut.Targetpath = target
        shortcut.WorkingDirectory = wdir
        shortcut.IconLocation = dest + r"\music-player\player\Assets\musical_note.ico"
        shortcut.save()

        shortcuts.append("Desktop")
    
    write(f"Created shortcuts for: {', '.join(shortcuts)}")

    write(sep + "Download completed! You can close this window and delete the installer if you want.")
    
    close = Button(app, text="Close Installer", font=Font(size=18), command=lambda: app.destroy())
    close.pack()
    return

content = """This will install the Music Player to your Windows computer. Python not required!

Git is required."""
Label(app, text=content, font=Font(size=12)).pack()

cvar = IntVar()

check = Checkbutton(app, text="Create Desktop shortcut", variable=cvar)
check.pack(pady=10, padx=10)

btn = Button(app, command=lambda: threading.Thread(target=install).start(), text="   Begin   ", font=Font(size=12))
btn.pack()

app.mainloop()