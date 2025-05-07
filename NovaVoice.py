import os
import sys
import json
import threading
import ctypes
import psutil
import pvporcupine
import pyaudio
import struct
import subprocess
import pyttsx3
import speech_recognition as sr
import webbrowser
import tkinter as tk
from tkinter import filedialog
from dotenv import load_dotenv


# === Helper: Correct path resolution when bundled ===
def resource_path(relative_path):
    if getattr(sys, 'frozen', False):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.abspath(relative_path)


# === Config & Constants ===
load_dotenv()
ACCESS_KEY = "PV_ACESS_KEY"
WAKE_WORD_PPN = resource_path("hey-assistant_en_windows.ppn")
CACHE_FILE = "exe_cache.json"
CUSTOM_CMDS_FILE = "custom_commands.json"
SEARCH_PATHS = [
    r"C:\Program Files",
    r"C:\Program Files (x86)",
    r"C:\Users"
]

# === Load or Initialize Cache & Custom Commands ===
if os.path.exists(CACHE_FILE):
    with open(CACHE_FILE, "r") as f:
        exe_cache = json.load(f)
else:
    exe_cache = {}

if os.path.exists(CUSTOM_CMDS_FILE):
    with open(CUSTOM_CMDS_FILE, "r") as f:
        custom_commands = json.load(f).get("commands", [])
else:
    custom_commands = []


def save_cache():
    with open(CACHE_FILE, "w") as f:
        json.dump(exe_cache, f)


# === TTS & Recognizer ===
tts = pyttsx3.init()
voices = tts.getProperty('voices')
if len(voices) > 1:
    tts.setProperty('voice', voices[1].id)
tts.setProperty('rate', 160)
tts.setProperty('volume', 1.0)

recognizer = sr.Recognizer()
mic = sr.Microphone()


def speak(text):
    tts.say(text)
    tts.runAndWait()


# === System Control Helpers ===
def mute_system():
    VK_MUTE = 0xAD
    KEYEVENTF_EXT = 0x0001
    KEYEVENTF_KEYUP = 0x0002
    ctypes.windll.user32.keybd_event(VK_MUTE, 0, KEYEVENTF_EXT, 0)
    ctypes.windll.user32.keybd_event(VK_MUTE, 0, KEYEVENTF_KEYUP, 0)


def sleep_system():
    # Suspend the system
    ctypes.windll.powrprof.SetSuspendState(False, True, False)


def shutdown_system():
    os.system("shutdown /s /t 0")


def get_battery_status():
    batt = psutil.sensors_battery()
    if batt is None:
        return None
    return batt.percent, batt.power_plugged


# === Finder & Launcher (async helper) ===
def find_executable(exe_name):
    for base in SEARCH_PATHS:
        if not os.path.exists(base):
            continue
        for root, _, files in os.walk(base):
            if exe_name.lower() in (f.lower() for f in files):
                return os.path.join(root, exe_name)
    return None


def launch_executable_async(exe_name, app_name):
    path = exe_cache.get(exe_name)
    if path and not os.path.exists(path):
        exe_cache.pop(exe_name, None)
        path = None
    if path:
        os.startfile(path)
        speak(f"Opening {app_name}.")
        return True

    speak(f"Searching for {app_name}, please wait.")

    def _search():
        found = find_executable(exe_name)
        if found:
            exe_cache[exe_name] = found
            save_cache()
            os.startfile(found)
            speak(f"Found and opening {app_name}.")
        else:
            speak(f"Sorry, I couldn’t find {app_name}.")

    threading.Thread(target=_search, daemon=True).start()
    return False


# === Built-in Web Openers ===
def open_browser():
    try:
        os.system("start microsoft-edge:https://www.google.com")
        speak("Opening Microsoft Edge.")
    except:
        speak("Falling back to default browser.")
        launch_executable_async("chrome.exe", "Google Chrome")


# === GUI Prompt for Manual Path ===
def prompt_for_exe(title="Select executable"):
    root = tk.Tk()
    root.withdraw()
    path = filedialog.askopenfilename(
        title=title,
        filetypes=[("Executable", "*.exe")],
        initialdir=os.path.expanduser("~")
    )
    root.destroy()
    return path


# === Voice Command Handling ===
def handle_command(phrase_time=4.0):
    speak("What would you like me to do?")
    frames, reads = [], int(porcupine.sample_rate / porcupine.frame_length * phrase_time)
    for _ in range(reads):
        pcm = stream.read(porcupine.frame_length, exception_on_overflow=False)
        frames.append(pcm)
    audio = sr.AudioData(b"".join(frames), porcupine.sample_rate, 2)

    try:
        command = recognizer.recognize_google(audio).lower().strip()
        speak(f"You said: {command}")
    except sr.UnknownValueError:
        speak("Sorry, I didn't catch that.");
        return
    except sr.RequestError:
        speak("It seems I'm offline.");
        return

    # Strip wake words
    for trig in ("hey assistant", "assistant", "hey"):
        if command.startswith(trig):
            command = command.replace(trig, "", 1).strip()

    # === Help / List Commands ===
    if "help" in command or "list commands" in command:
        built_ins = [
            "open Google", "open YouTube", "open Edge",
            "open Chrome", "open Firefox", "open Brave",
            "open PyCharm", "open Visual Studio", "open VS Code",
            "open IntelliJ", "mute the system", "put computer to sleep",
            "turn off the computer", "what's my battery level"
        ]
        speak("I can do the following built-in commands:")
        for b in built_ins:
            speak(b)
        if custom_commands:
            speak("And your custom commands are:")
            for cmd in custom_commands:
                speak(cmd["phrase"])
        else:
            speak("You have no custom commands defined.")
        return

    # === System Controls ===
    if "mute" in command:
        mute_system();
        speak("System muted.");
        return
    if "sleep" in command:
        sleep_system();
        speak("Putting the computer to sleep.");
        return
    if "turn off" in command or "shut down" in command:
        shutdown_system();
        return
    if "battery" in command:
        status = get_battery_status()
        if status:
            pct, plugged = status
            plugged_str = "plugged in" if plugged else "not plugged in"
            speak(f"Battery is at {pct} percent and is {plugged_str}.")
        else:
            speak("Battery information unavailable.")
        return

    # === Built-in App & Web Routing ===
    if "google" in command:
        webbrowser.open("https://www.google.com");
        speak("Opening Google.");
        return
    if "youtube" in command:
        webbrowser.open("https://www.youtube.com");
        speak("Opening YouTube.");
        return
    if "edge" in command:
        open_browser();
        return

    if "chrome" in command:
        launch_executable_async("chrome.exe", "Google Chrome");
        return
    if "firefox" in command:
        launch_executable_async("firefox.exe", "Mozilla Firefox");
        return
    if "brave" in command:
        launch_executable_async("brave.exe", "Brave Browser");
        return

    if "pycharm" in command:
        launch_executable_async("pycharm64.exe", "PyCharm");
        return

    if "visual studio" in command:
        key, default = "devenv.exe", r"C:\Program Files\Microsoft Visual Studio\2022\Community\Common7\IDE\devenv.exe"
        if os.path.exists(default):
            os.startfile(default);
            exe_cache[key] = default;
            save_cache();
            speak("Opening Visual Studio.")
        else:
            found = find_executable(key)
            if found:
                exe_cache[key] = found;
                save_cache();
                os.startfile(found);
                speak("Found and opening Visual Studio.")
            else:
                speak("Please select the Visual Studio executable.")
                manual = prompt_for_exe("Locate devenv.exe")
                if manual and os.path.exists(manual):
                    os.startfile(manual);
                    exe_cache[key] = manual;
                    save_cache();
                    speak("Opening Visual Studio.")
                else:
                    speak("That path doesn’t exist.")
        return

    if "code" in command:
        key = "code.exe"
        user_path = os.path.expandvars(r"%LocalAppData%\Programs\Microsoft VS Code\Code.exe")
        system_path = r"C:\Program Files\Microsoft VS Code\Code.exe"
        if os.path.exists(user_path):
            os.startfile(user_path);
            exe_cache[key] = user_path;
            save_cache();
            speak("Opening Visual Studio Code.")
        elif os.path.exists(system_path):
            os.startfile(system_path);
            exe_cache[key] = system_path;
            save_cache();
            speak("Opening Visual Studio Code.")
        else:
            found = find_executable(key)
            if found:
                os.startfile(found);
                exe_cache[key] = found;
                save_cache();
                speak("Found and opening Visual Studio Code.")
            else:
                speak("Please select the VS Code executable.")
                manual = prompt_for_exe("Locate Code.exe")
                if manual and os.path.exists(manual):
                    os.startfile(manual);
                    exe_cache[key] = manual;
                    save_cache();
                    speak("Opening Visual Studio Code.")
                else:
                    speak("That path doesn’t exist.")
        return

    if "intellij" in command:
        launch_executable_async("idea64.exe", "IntelliJ IDEA");
        return

    # === Custom Commands ===
    for cmd in custom_commands:
        phrase = cmd.get("phrase", "").lower()
        if phrase and phrase in command:
            action = cmd.get("action")
            response = cmd.get("response", "")
            if action == "launch_executable":
                launch_executable_async(cmd["exe_name"], cmd["app_name"])
            elif action == "url":
                webbrowser.open(cmd["url"])
                if response: speak(response)
            elif action == "shell":
                subprocess.Popen(cmd["shell_cmd"], shell=True)
                if response: speak(response)
            return

    speak("I’m not sure how to do that.")


# === Main Wake-Word Loop ===
def main():
    global porcupine, pa, stream
    porcupine = pvporcupine.create(
        access_key=ACCESS_KEY,
        keyword_paths=[WAKE_WORD_PPN],
        sensitivities=[0.7]
    )
    pa = pyaudio.PyAudio()
    stream = pa.open(
        rate=porcupine.sample_rate,
        channels=1,
        format=pyaudio.paInt16,
        input=True,
        frames_per_buffer=porcupine.frame_length
    )

    speak("Assistant ready. Say 'hey assistant' to begin.")
    try:
        while True:
            pcm = stream.read(porcupine.frame_length, exception_on_overflow=False)
            frame = struct.unpack_from("h" * porcupine.frame_length, pcm)
            if porcupine.process(frame) >= 0:
                handle_command()
                speak("Listening for wake word again.")
    except KeyboardInterrupt:
        speak("Goodbye.")
    finally:
        stream.stop_stream();
        stream.close()
        pa.terminate();
        porcupine.delete();
        tts.stop()


if __name__ == "__main__":
    main()
