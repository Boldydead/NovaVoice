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
from dotenv import load_dotenv, find_dotenv  # Added find_dotenv for potential path info
import datetime
import time
import logging

# === Basic Logging Setup ===
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


# === Helper: Correct path resolution when bundled ===
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        # Not running in a PyInstaller bundle
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


# === Config & Constants ===

# --- Determine potential locations for the .env file ---
app_name_for_folders = "NovaVoice"  # Used for creating a dedicated config folder
expected_dotenv_locations = []
loaded_dotenv_path_info = {"loaded": False, "path": "Not found or loaded by default search"}

# Location 1: Application-specific config directory in AppData\Roaming (Recommended)
try:
    appdata_roaming_path = os.getenv('APPDATA')
    if appdata_roaming_path:
        app_config_dir = os.path.join(appdata_roaming_path, app_name_for_folders)
        # Note: The script does not create this directory. User should create it if they want to use this location.
        # For auto-creation, you could add: os.makedirs(app_config_dir, exist_ok=True)
        expected_dotenv_locations.append(os.path.join(app_config_dir, ".env"))
    else:
        logging.warning("APPDATA environment variable not found. Cannot check Roaming AppData location for .env.")
except Exception as e_appdata:
    logging.error(f"Error determining APPDATA path for .env: {e_appdata}")

# Location 2: Directory where the .exe or .py script is located (Fallback)
try:
    if getattr(sys, 'frozen', False):  # Running as a bundled executable (PyInstaller)
        script_or_exe_dir = os.path.dirname(sys.executable)  # sys.executable is the path to the .exe
    elif __file__:  # Running as a .py script
        script_or_exe_dir = os.path.dirname(os.path.abspath(__file__))
    else:  # Fallback (e.g., interactive interpreter where __file__ is not defined)
        script_or_exe_dir = os.getcwd()
    expected_dotenv_locations.append(os.path.join(script_or_exe_dir, ".env"))
except Exception as e_exe_dir:
    logging.error(f"Error determining executable/script directory for .env: {e_exe_dir}")
    # As a further fallback if __file__ and sys.executable logic fails, add CWD.
    current_working_dir_env = os.path.join(os.getcwd(), ".env")
    if current_working_dir_env not in expected_dotenv_locations:
        expected_dotenv_locations.append(current_working_dir_env)

# --- Attempt to load .env from the prioritized locations ---
for dotenv_path_to_try in expected_dotenv_locations:
    if dotenv_path_to_try and os.path.exists(dotenv_path_to_try):  # Check if path is valid and exists
        if load_dotenv(dotenv_path_to_try, override=True):  # override=True ensures this takes precedence
            logging.info(f"Successfully loaded .env from: {dotenv_path_to_try}")
            loaded_dotenv_path_info["loaded"] = True
            loaded_dotenv_path_info["path"] = dotenv_path_to_try
            break  # Stop searching once found and successfully loaded
        else:
            logging.warning(
                f"Found .env at {dotenv_path_to_try} but load_dotenv() reported failure (e.g. empty file or permission issue).")
    else:
        if dotenv_path_to_try:  # Only log if it was a valid path attempt
            logging.info(f".env file not found at trial location: {dotenv_path_to_try}")

# If still not loaded after checking specific paths, try a final default search by python-dotenv
# This default search looks in os.getcwd() and its parents by default.
if not loaded_dotenv_path_info["loaded"]:
    logging.info(
        "Attempting final fallback to default load_dotenv() search (e.g., current working directory and parents).")
    # find_dotenv will find the path if default load_dotenv succeeds
    env_path_by_find_dotenv = find_dotenv(usecwd=True, raise_error_if_not_found=False)
    if load_dotenv(override=True):  # override=True
        logging.info(
            f".env file loaded using default search logic (found at: {env_path_by_find_dotenv or 'unknown default path'}).")
        loaded_dotenv_path_info["loaded"] = True
        loaded_dotenv_path_info["path"] = env_path_by_find_dotenv or "Found by default search"
    else:
        logging.error(".env file could not be loaded by any method. PORCUPINE_ACCESS_KEY will be missing.")
        print(
            "CRITICAL: .env file containing PORCUPINE_ACCESS_KEY was not found by any method. Ensure it exists in a checked location.")

PORCUPINE_ACCESS_KEY = os.getenv("PORCUPINE_ACCESS_KEY")
if not PORCUPINE_ACCESS_KEY:
    logging.error(
        "PORCUPINE_ACCESS_KEY is not set in the environment. This usually means the .env file was not found, not loaded correctly, or the key is missing from the .env file. Wake word detection will fail.")
    print(
        "WARNING: PORCUPINE_ACCESS_KEY not loaded. Please ensure your .env file is in a recognized location (e.g., %APPDATA%\\NovaVoice\\.env or next to the executable) and contains the key. Wake word detection will fail.")

WAKE_WORD_PPN_FILENAME = "hey google_windows.ppn"  # Make sure you have this file or update the name
WAKE_WORD_PPN = None  # Initialize to None
try:
    # resource_path is for files BUNDLED with the exe or next to the script for dev
    potential_ppn_path = resource_path(WAKE_WORD_PPN_FILENAME)
    if os.path.exists(potential_ppn_path):
        WAKE_WORD_PPN = potential_ppn_path
        logging.info(f"Wake word file found at: {WAKE_WORD_PPN}")
    else:
        logging.error(f"Wake word file '{WAKE_WORD_PPN_FILENAME}' not found at primary path '{potential_ppn_path}'.")
        # This specific fallback for .ppn might be less relevant if resource_path is robust for bundled files.
        # For development (not bundled), if resource_path defaults to CWD and PPN is not there but in script dir:
        if not getattr(sys, 'frozen', False) and __file__:
            alt_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), WAKE_WORD_PPN_FILENAME)
            if os.path.exists(alt_path):
                WAKE_WORD_PPN = alt_path
                logging.info(f"Found wake word file at alternate dev path: {WAKE_WORD_PPN}")
            else:
                logging.error(f"Also not found at alternate dev path: {alt_path}")
                # WAKE_WORD_PPN remains None
        elif getattr(sys, 'frozen', False):  # Bundled but not found by resource_path
            logging.error(
                f"CRITICAL (Bundled): Wake word file '{WAKE_WORD_PPN_FILENAME}' not found at expected bundled path '{potential_ppn_path}'. Ensure it was added to PyInstaller (--add-data).")
            # WAKE_WORD_PPN remains None

    if not WAKE_WORD_PPN:  # If still None after all checks
        print(f"ERROR: Wake word file '{WAKE_WORD_PPN_FILENAME}' could not be located. Wake word detection will fail.")

except Exception as e:
    logging.error(f"Error resolving resource path for PPN file: {e}")
    WAKE_WORD_PPN = None  # Ensure it's None if path resolution failed
    print(f"ERROR: Could not resolve PPN file path due to an exception. Wake word detection will fail.")

CACHE_FILE = resource_path("exe_cache.json")
CUSTOM_CMDS_FILE = resource_path("custom_commands.json")  # Script handles if this doesn't exist
DAILY_CACHE_FILE = resource_path("daily_text.json")  # Script handles if this doesn't exist

SEARCH_PATHS = [
    r"C:\Program Files",
    r"C:\Program Files (x86)",
    os.path.expanduser("~")
]

APP_LAUNCH_MAP = {
    "chrome": "chrome.exe", "firefox": "firefox.exe", "brave": "brave.exe",
    "edge": "msedge.exe",
    "pycharm": "pycharm64.exe", "intellij": "idea64.exe",
    "visual studio": "devenv.exe", "code": "Code.exe", "vs code": "Code.exe",
    "notepad": "notepad.exe", "calculator": "calc.exe", "explorer": "explorer.exe",
    "word": "winword.exe", "excel": "excel.exe", "powerpoint": "powerpnt.exe",
    "spotify": "spotify.exe", "discord": "discord.exe", "steam": "steam.exe",
    "zoom": "zoom.exe", "teams": "teams.exe",
    "slack": "slack.exe", "obs": "obs64.exe",
    "gimp": "gimp-2.10.exe",
    "vlc": "vlc.exe",
}


# === Load or Initialize Cache & Custom Commands ===
def load_json_data(file_path, default_data=None):
    if default_data is None: default_data = {}
    actual_path = file_path  # Path already resolved by resource_path or similar logic
    if os.path.exists(actual_path):
        try:
            with open(actual_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except json.JSONDecodeError:
            logging.warning(f"Could not decode JSON from {actual_path}. Using default data.")
            return default_data
        except Exception as e:
            logging.error(f"Error loading {actual_path}: {e}. Using default data.")
            return default_data
    else:
        logging.info(f"JSON file not found at {actual_path}. Using default data.")
    return default_data


exe_cache = load_json_data(CACHE_FILE, {})

# For custom_commands, it's better to look for it relative to the .exe/script dir or AppData
# rather than bundling it, as users are meant to edit it.
# Let's use the same logic as .env for custom_commands.json
custom_cmds_path_to_use = None
for loc_base in [os.path.join(os.getenv('APPDATA', ''), app_name_for_folders) if os.getenv('APPDATA') else None,
                 # AppData first
                 script_or_exe_dir]:  # Then next to exe/script
    if loc_base:  # Ensure loc_base is not None (e.g. if APPDATA failed)
        potential_path = os.path.join(loc_base, "custom_commands.json")
        if os.path.exists(potential_path):
            custom_cmds_path_to_use = potential_path
            logging.info(f"Using custom_commands.json from: {custom_cmds_path_to_use}")
            break
if not custom_cmds_path_to_use:
    # If not found in preferred locations, resource_path would try bundled (if added) or CWD.
    # But custom_commands usually shouldn't be bundled if user-editable.
    # We'll let load_json_data use the path from resource_path("custom_commands.json")
    # which will default to CWD or _MEIPASS if bundled (which is not ideal for user editing)
    # So, we explicitly set it to a path that will likely not exist if not in preferred locations,
    # leading to default_data being used by load_json_data if user hasn't created it.
    custom_cmds_path_to_use = os.path.join(script_or_exe_dir, "custom_commands.json")  # Default to looking next to exe
    logging.info(
        f"custom_commands.json not found in primary user locations. Will attempt to load from/create near executable: {custom_cmds_path_to_use} or use default empty commands.")

custom_commands_data = load_json_data(custom_cmds_path_to_use, {"commands": []})
custom_commands = custom_commands_data.get("commands", [])

# === Thread-safe TTS Setup ===
try:
    tts = pyttsx3.init()
    voices = tts.getProperty('voices')
    female_voice_id = None
    for voice in voices:
        if voice.gender == 'female' or 'female' in (voice.name or "").lower() or \
                (len(voices) > 1 and voice.id == voices[1].id):
            female_voice_id = voice.id
            break
    if female_voice_id:
        tts.setProperty('voice', female_voice_id)
    elif voices:
        tts.setProperty('voice', voices[0].id)

    tts.setProperty('rate', 160)
    tts.setProperty('volume', 1.0)
    tts_lock = threading.Lock()
except Exception as e:
    logging.error(f"Failed to initialize TTS: {e}. Voice output will be disabled.")
    tts = None


def speak(text):
    if not tts:
        logging.warning(f"TTS not available. Intended to speak: {text}")
        print(f"ASSISTANT (TTS Disabled): {text}")
        return
    with tts_lock:
        try:
            logging.info(f"SPEAKING: {text}")
            tts.say(text)
            tts.runAndWait()
        except RuntimeError as e:
            logging.error(f"TTS RuntimeError: {e}")
        except Exception as e:
            logging.error(f"General TTS error: {e}")


# === Speech Recognizer ===
recognizer = sr.Recognizer()
try:
    mic = sr.Microphone()
    with mic as source:
        logging.info("Adjusting for ambient noise, please wait...")
        recognizer.adjust_for_ambient_noise(source, duration=1)
        logging.info("Ambient noise adjustment complete.")
except Exception as e:
    logging.error(f"Microphone not found or speech_recognition library issue: {e}")
    if tts:
        speak("I can't access a microphone. Please check your microphone settings.")
    else:
        print("ERROR: Microphone not found or speech_recognition library issue.")
    mic = None


# === System Control Helpers (Windows-specific) ===
def mute_system():
    VK_MUTE = 0xAD
    KEYEVENTF_EXT = 0x0001
    KEYEVENTF_KEYUP = 0x0002
    ctypes.windll.user32.keybd_event(VK_MUTE, 0, KEYEVENTF_EXT, 0)
    ctypes.windll.user32.keybd_event(VK_MUTE, 0, KEYEVENTF_KEYUP, 0)
    logging.info("System mute toggled.")


def sleep_system():
    ctypes.windll.powrprof.SetSuspendState(False, True, False)
    logging.info("System going to sleep.")


def shutdown_system():
    os.system("shutdown /s /t 0")
    logging.info("System shutting down.")


def get_battery_status():
    try:
        batt = psutil.sensors_battery()
        if batt is None: return None
        return batt.percent, batt.power_plugged
    except Exception as e:
        logging.error(f"Could not get battery status: {e}")
        return None


# === Finder & Launcher ===
def find_executable(exe_name):
    logging.debug(f"Searching for executable: {exe_name}")
    for base in SEARCH_PATHS:
        if not os.path.exists(base):
            logging.debug(f"Search path {base} does not exist, skipping.")
            continue
        for root, dirs, files in os.walk(base, topdown=True):
            current_depth = os.path.abspath(root).count(os.sep) - os.path.abspath(base).count(os.sep)
            if current_depth > 5:
                logging.debug(f"Reached max search depth in {root}, pruning walk.")
                dirs[:] = []
                continue

            exe_name_lower_no_ext, exe_ext = os.path.splitext(exe_name.lower())
            if not exe_ext: exe_ext = ".exe"

            for f in files:
                f_lower = f.lower()
                f_lower_no_ext, f_ext = os.path.splitext(f_lower)
                if f_lower_no_ext == exe_name_lower_no_ext and f_ext == exe_ext:
                    found_path = os.path.join(root, f)
                    logging.info(f"Found {exe_name} as {f} at {found_path}")
                    return found_path
    logging.info(f"{exe_name} not found in standard search paths.")
    return None


def launch_executable_async(exe_name, app_name):
    path = exe_cache.get(exe_name.lower())
    if path and not os.path.exists(path):
        logging.info(f"Cached path for {exe_name} ('{path}') no longer exists. Removing from cache.")
        exe_cache.pop(exe_name.lower(), None)
        path = None

    if path:
        try:
            os.startfile(path)
            speak(f"Opening {app_name}.")
            return True
        except Exception as e:
            logging.error(f"Error starting cached {app_name} from {path}: {e}")
            exe_cache.pop(exe_name.lower(), None)

    speak(f"Searching for {app_name}, please wait.")

    def _search_and_launch_thread():
        # For exe_cache and daily_text, determine writable path
        # Prefer AppData, then next to exe/script
        writable_config_dir = script_or_exe_dir  # Default
        try:
            appdata_path_local = os.getenv('LOCALAPPDATA')  # Local AppData often better for caches
            if appdata_path_local:
                potential_dir = os.path.join(appdata_path_local, app_name_for_folders)
                os.makedirs(potential_dir, exist_ok=True)  # Ensure it exists
                writable_config_dir = potential_dir
            elif os.getenv('APPDATA'):  # Fallback to Roaming if LOCALAPPDATA not set
                potential_dir = os.path.join(os.getenv('APPDATA'), app_name_for_folders)
                os.makedirs(potential_dir, exist_ok=True)
                writable_config_dir = potential_dir
        except Exception as e_mkdir:
            logging.warning(
                f"Could not create/use AppData directory for cache, using executable's directory: {e_mkdir}")
            # writable_config_dir remains script_or_exe_dir

        current_cache_file = os.path.join(writable_config_dir, os.path.basename(CACHE_FILE))

        found_path = find_executable(exe_name)
        if found_path:
            exe_cache[exe_name.lower()] = found_path
            try:
                with open(current_cache_file, 'w', encoding="utf-8") as f_cache:
                    json.dump(exe_cache, f_cache)
                logging.info(f"Updated exe_cache at {current_cache_file}")
            except Exception as e_cache_write:
                logging.error(f"Error writing to exe_cache.json at {current_cache_file}: {e_cache_write}")

            try:
                os.startfile(found_path)
                speak(f"Found and opening {app_name}.")
            except Exception as e_start:
                logging.error(f"Error starting found {app_name} from {found_path}: {e_start}")
                speak(
                    f"I found {app_name} but couldn't open it. There might be a permission issue or the file is corrupted.")
        else:
            speak(
                f"Sorry, I couldn't find {app_name} to open. You might need to add it as a custom command if it's in a non-standard location.")

    threading.Thread(target=_search_and_launch_thread, daemon=True).start()
    return False


# === Built-in Browser Opener ===
def open_default_browser(url="https://www.google.com"):
    try:
        webbrowser.open(url)
        logging.info(f"Opened {url} in default browser.")
        return True
    except Exception as e:
        logging.error(f"Could not open URL {url} in browser: {e}")
        return False


# === GUI Prompt for Manual Path ===
def prompt_for_exe(title="Select executable"):
    try:
        root = tk.Tk()
        root.withdraw()
        path = filedialog.askopenfilename(
            title=title,
            filetypes=[("Executable", "*.exe"), ("All files", "*.*")],
            initialdir=os.path.expanduser("~\\Desktop")
        )
        root.destroy()
        return path
    except Exception as e:
        logging.error(f"Error showing Tkinter dialog: {e}")
        return None


# --- Command Handler Functions (for Command Pattern) ---
def cmd_open_google(command_text):
    if open_default_browser("https://www.google.com"):
        speak("Opening Google.")
    else:
        speak("I couldn't open Google in your browser.")


def cmd_open_youtube(command_text):
    if open_default_browser("https://www.youtube.com"):  # Original URL
        speak("Opening YouTube.")
    else:
        speak("I couldn't open YouTube in your browser.")


def cmd_mute_system(command_text):
    mute_system()
    speak("System mute toggled.")


def cmd_sleep_system(command_text):
    speak("Putting the computer to sleep.")
    sleep_system()


def cmd_shutdown_system(command_text):
    speak("Shutting down the computer.")
    shutdown_system()


def cmd_get_battery(command_text):
    status = get_battery_status()
    if status:
        pct, plugged = status
        plugged_str = "plugged in and charging" if plugged else "not plugged in"
        speak(f"Battery is at {pct} percent and is {plugged_str}.")
    else:
        speak("Sorry, I couldn't get the battery information.")


def cmd_get_current_time(command_text):
    now = datetime.datetime.now()
    time_str = now.strftime("%I:%M %p")
    speak(f"The current time is {time_str}.")


def cmd_get_current_date(command_text):
    now = datetime.datetime.now()
    date_str = now.strftime("%A, %B %d, %Y")
    speak(f"Today is {date_str}.")


def cmd_show_help(command_text):
    built_ins = [
        "open Google", "open YouTube",
        "open Edge", "open Chrome",
        "mute the system", "put computer to sleep",
        "turn off the computer", "what's my battery level",
        "what time is it", "what is today's date",
        "tell me today's Bible verse",
        "what can you open"
    ]
    speak("I can understand commands like:")
    for b_cmd_idx, b_cmd in enumerate(built_ins):
        if b_cmd_idx < 6:
            speak(b_cmd)
        else:
            break
    speak("And many application names like 'open Word' or 'launch Spotify'.")
    if custom_commands:
        speak("Your custom commands include:")
        for c_cmd_idx, c_cmd in enumerate(custom_commands):
            if c_cmd_idx < 3:
                speak(c_cmd.get("phrase", "a custom command"))
            else:
                break
    else:
        speak("You haven't defined any custom commands yet.")
    speak("You can also ask 'what can you do' for a broader overview of my capabilities.")


def cmd_list_known_apps(command_text):
    known_apps_friendly_names = set()
    for app_alias in APP_LAUNCH_MAP.keys():
        known_apps_friendly_names.add(app_alias.title())
    if exe_cache:  # exe_cache is loaded from a file determined by writable_config_dir logic implicitly
        for exe_name_lower_key in exe_cache.keys():
            app_name_from_cache = exe_name_lower_key.replace(".exe", "")
            app_name_from_cache = ' '.join(word.capitalize() for word in app_name_from_cache.split())
            known_apps_friendly_names.add(app_name_from_cache)
    for cmd_config in custom_commands:  # custom_commands loaded from preferred path
        if cmd_config.get("action") == "launch_executable":
            app_name = cmd_config.get("app_name")
            if app_name:
                known_apps_friendly_names.add(app_name.title())
            else:
                exe_name = cmd_config.get("exe_name")
                if exe_name:
                    friendly_exe_name = exe_name.lower().replace(".exe", "")
                    friendly_exe_name = ' '.join(word.capitalize() for word in friendly_exe_name.split())
                    known_apps_friendly_names.add(friendly_exe_name)
    if not known_apps_friendly_names:
        speak("I don't have a specific list of pre-identified applications right now. "
              "However, you can ask me to 'open' or 'launch' any application by its name, "
              "like 'open Chrome' or 'launch Notepad', and I'll try my best to find it.")
    else:
        speak_text = "Based on my configuration and past successful launches, I can try to open apps like: "
        apps_to_speak = sorted(list(known_apps_friendly_names))
        if not apps_to_speak:
            speak("I couldn't compile a list of apps right now. Try asking me to open one directly!")
            return
        if len(apps_to_speak) > 7:
            speak_text += ", ".join(apps_to_speak[:6]) + ", and others such as " + apps_to_speak[6] + "."
        elif len(apps_to_speak) > 1:
            speak_text += ", ".join(apps_to_speak[:-1]) + ", and " + apps_to_speak[-1] + "."
        else:
            speak_text += apps_to_speak[0] + "."
        speak(speak_text)
        speak("You can also try asking for any other application not mentioned, and I'll search for it.")


def cmd_tell_capabilities(command_text):
    speak("I can help you with various tasks on your computer after you say my wake word.")
    speak("Here's an overview of what I can do:")
    speak(
        "First, I can open applications. For example, you can ask me to 'open Chrome', 'launch Spotify', or 'open Notepad'. "
        "I have a list of these common applications, and I can also search for others if you ask. "
        "If you'd like a more detailed list of applications I'm already familiar with or have found before, you can say 'what applications can you open'.")
    speak("Second, I can open websites. For example, try saying 'open Google' or 'open YouTube'.")
    speak(
        "Third, I can provide information. You can ask me 'what time is it', 'what is today's date', or 'what's my battery level'.")
    speak(
        "Fourth, I can control some system settings on Windows. For instance, 'mute the system', 'put computer to sleep', or 'turn off the computer'.")
    speak(
        "I can also attempt to fetch a daily Bible verse for you if you ask, though please note the current source for this is a placeholder message.")
    if custom_commands:
        speak(
            "Additionally, I can run custom commands that you've set up. These can include launching specific programs with a unique phrase, opening particular web pages, or running shell scripts.")
    else:
        speak(
            "And if you were to set up custom commands, I could perform those too, allowing you to tailor my actions to your needs!")
    speak("For a shorter list of example command phrases you can use, just say 'help'.")


# === Command Dispatcher ===
COMMAND_DISPATCHER = {
    "what can you do": cmd_tell_capabilities,
    "what are your capabilities": cmd_tell_capabilities,
    "tell me your capabilities": cmd_tell_capabilities,
    "what can you do and what can you open": cmd_tell_capabilities,
    "what can you do and open": cmd_tell_capabilities,
    "what can you open and what can you do": cmd_tell_capabilities,
    "describe your functions": cmd_tell_capabilities,
    "what all can you do": cmd_tell_capabilities,
    "help": cmd_show_help, "list commands": cmd_show_help,
    "google": cmd_open_google, "youtube": cmd_open_youtube,
    "mute": cmd_mute_system, "sleep": cmd_sleep_system,
    "turn off": cmd_shutdown_system, "shut down": cmd_shutdown_system,
    "battery": cmd_get_battery,
    "time": cmd_get_current_time, "what time is it": cmd_get_current_time, "current time": cmd_get_current_time,
    "date": cmd_get_current_date, "today's date": cmd_get_current_date, "what is today's date": cmd_get_current_date,
    "what can you open": cmd_list_known_apps,
    "what apps can you open": cmd_list_known_apps,
    "what applications can you open": cmd_list_known_apps,
    "list apps i can open": cmd_list_known_apps,
    "show known apps": cmd_list_known_apps,
    "tell me what apps you can open": cmd_list_known_apps,
    "list applications": cmd_list_known_apps,
}


# === Voice Command Handling ===
def handle_command(phrase_listen_timeout=5.0, phrase_time_limit=10.0):
    if not mic:
        speak("Microphone is not available, cannot listen for commands.")
        return
    speak("What would you like me to do?")
    command = ""
    try:
        with mic as source:
            logging.info("Listening for command...")
            try:
                audio = recognizer.listen(source, timeout=phrase_listen_timeout, phrase_time_limit=phrase_time_limit)
            except sr.WaitTimeoutError:
                logging.info("No speech detected within timeout.")
                return
        logging.info("Processing command...")
        command = recognizer.recognize_google(audio).lower().strip()
        speak(f"You said: {command}")
    except sr.UnknownValueError:
        speak("Sorry, I didn't catch that clearly.")
        logging.info("Google Speech Recognition could not understand audio.")
        return
    except sr.RequestError as e:
        speak("It seems I'm having trouble reaching the speech service.")
        logging.error(f"Could not request results from Google Speech Recognition service; {e}")
        return
    except Exception as e:
        speak("An unexpected error occurred while trying to understand you.")
        logging.error(f"Error during speech recognition: {e}")
        return
    if not command: return
    common_wake_phrases = ("hey assistant", "assistant", "hey google", "google")
    for trig in common_wake_phrases:
        if command.startswith(trig):
            command = command.replace(trig, "", 1).strip()
            logging.debug(f"Command after stripping '{trig}': {command}")
            break
    if not command:
        speak("I didn't get a specific command after the wake phrase.")
        return

    if command in COMMAND_DISPATCHER:
        COMMAND_DISPATCHER[command](command)
        return
    sorted_keywords = sorted(COMMAND_DISPATCHER.keys(), key=len, reverse=True)
    for keyword in sorted_keywords:
        if keyword and keyword in command:
            func = COMMAND_DISPATCHER[keyword]
            func(command)
            return

    if "edge" in command:
        if open_default_browser("https://www.microsoft.com/edge"):
            speak("Opening Microsoft Edge.")
        else:
            speak("I couldn't open Edge.")
        return
    if command.startswith("open ") or command.startswith("launch "):
        parts = command.split(" ", 1)
        if len(parts) > 1:
            app_to_launch = parts[1].strip()
            if app_to_launch:
                app_exe = APP_LAUNCH_MAP.get(app_to_launch.lower())
                if app_exe:
                    launch_executable_async(app_exe, app_to_launch.title())
                    return
                else:
                    for c_cmd in custom_commands:
                        custom_app_name = c_cmd.get("app_name", "")
                        if custom_app_name and custom_app_name.lower() == app_to_launch.lower() and \
                                c_cmd.get("action") == "launch_executable":
                            exe_name_custom = c_cmd.get("exe_name")
                            if exe_name_custom:
                                launch_executable_async(exe_name_custom, app_to_launch.title())
                                return
                            break
                    else:
                        app_exe_guess = app_to_launch + ".exe" if not app_to_launch.lower().endswith(
                            ".exe") else app_to_launch
                        logging.info(f"Attempting generic launch for: {app_exe_guess} ({app_to_launch.title()})")
                        launch_executable_async(app_exe_guess, app_to_launch.title())
                        return
    if "bible verse" in command or "daily text" in command or "today's text" in command:
        speak("Fetching today's Bible verse, please wait.")
        daily_scripture = fetch_daily_text()
        speak(daily_scripture)
        return
    for cmd_config in custom_commands:
        phrase = cmd_config.get("phrase", "").lower()
        if phrase and phrase == command:
            action = cmd_config.get("action")
            response = cmd_config.get("response", f"Okay, performing your custom action for '{phrase}'.")
            logging.info(f"Executing custom command for exact phrase: '{phrase}', action: {action}")
            if action == "launch_executable":
                exe_name = cmd_config.get("exe_name")
                app_name_custom = cmd_config.get("app_name", exe_name or "the application")
                if exe_name:
                    launch_executable_async(exe_name, app_name_custom)
                else:
                    speak(f"Executable name missing for custom command '{phrase}'.")
            elif action == "url":
                url_to_open = cmd_config.get("url")
                if url_to_open:
                    if open_default_browser(url_to_open):
                        speak(response) if response else speak(f"Opening {url_to_open}")
                    else:
                        speak(f"I couldn't open the URL for '{phrase}'.")
                else:
                    speak(f"URL missing for custom command '{phrase}'.")
            elif action == "shell":
                shell_command_to_run = cmd_config.get("shell_cmd")
                if shell_command_to_run:
                    try:
                        subprocess.Popen(shell_command_to_run, shell=True,
                                         creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0)
                        if response: speak(response)
                    except Exception as e_shell:
                        logging.error(f"Error executing shell command '{shell_command_to_run}': {e_shell}")
                        speak(f"I couldn't run the shell command for '{phrase}'.")
                else:
                    speak(f"Shell command missing for custom command '{phrase}'.")
            else:
                speak(f"Unknown action type '{action}' for custom command '{phrase}'.")
            return
    speak("I'm not sure how to do that yet.")


# === Daily Text & Wisdom Functions ===
def fetch_daily_text():
    today_str = datetime.datetime.now().strftime("%Y-%m-%d")
    # For daily_text_cache, determine writable path
    # Prefer AppData, then next to exe/script
    writable_config_dir_daily = script_or_exe_dir  # Default
    try:
        appdata_path_local = os.getenv('LOCALAPPDATA')
        if appdata_path_local:
            potential_dir = os.path.join(appdata_path_local, app_name_for_folders)
            os.makedirs(potential_dir, exist_ok=True)
            writable_config_dir_daily = potential_dir
        elif os.getenv('APPDATA'):
            potential_dir = os.path.join(os.getenv('APPDATA'), app_name_for_folders)
            os.makedirs(potential_dir, exist_ok=True)
            writable_config_dir_daily = potential_dir
    except Exception as e_mkdir_daily:
        logging.warning(
            f"Could not create/use AppData directory for daily_text_cache, using executable's directory: {e_mkdir_daily}")

    current_daily_cache_file = os.path.join(writable_config_dir_daily, os.path.basename(DAILY_CACHE_FILE))

    cached_data = load_json_data(current_daily_cache_file)  # load_json_data now takes resolved path
    if cached_data.get("date") == today_str and cached_data.get("text"):
        logging.info(f"Using cached daily text from {current_daily_cache_file}.")
        return cached_data["text"]

    logging.warning(
        "Web scraping for jw.org daily text is highly dependent on website structure and may fail or return this message.")
    full_text = "Fetching the daily Bible verse from JW.org is currently unreliable due to website changes. Please check the website directly for today's text."
    try:
        with open(current_daily_cache_file, "w", encoding="utf-8") as f:
            json.dump({"date": today_str, "text": full_text}, f)
        logging.info(f"Updated daily_text_cache at {current_daily_cache_file}")
    except Exception as e:
        logging.error(f"Error writing to daily cache file {current_daily_cache_file}: {e}")
    return full_text


# === Scheduler Thread ===
daily_greeting_done_today = False


def scheduler_thread_func():
    global daily_greeting_done_today
    logging.info("Scheduler thread started.")
    while True:
        now = datetime.datetime.now()
        hour = now.hour
        if 7 <= hour < 8 and not daily_greeting_done_today:
            logging.info("Scheduler: Morning sequence triggered.")
            speak("Good morning! Here's something for your day.")
            time.sleep(0.5)
            text_for_day = fetch_daily_text()
            speak(text_for_day)
            daily_greeting_done_today = True
        if hour < 6:
            if daily_greeting_done_today:
                logging.info("Scheduler: Resetting daily_greeting_done_today flag.")
                daily_greeting_done_today = False
        time.sleep(30)


# === Main Wake-Word Loop ===
porcupine = None
pa = None
audio_stream = None


def main_loop():
    global porcupine, pa, audio_stream
    if not PORCUPINE_ACCESS_KEY or not WAKE_WORD_PPN or not os.path.exists(WAKE_WORD_PPN):
        msg = "Porcupine Access Key or Wake Word PPN file is missing/invalid. Wake word engine cannot start."
        logging.error(msg)
        speak(msg) if tts else print(f"ERROR: {msg}")
        return
    try:
        porcupine = pvporcupine.create(
            access_key=PORCUPINE_ACCESS_KEY,
            keyword_paths=[WAKE_WORD_PPN],
            sensitivities=[0.65]
        )
    except pvporcupine.PorcupineError as e:
        logging.error(f"Porcupine initialization failed: {e}")
        speak("I couldn't start my wake word engine. Please check the access key and keyword file.")
        return
    except Exception as e_porc_general:
        logging.error(f"An unexpected error occurred during Porcupine initialization: {e_porc_general}")
        speak("An unexpected error occurred while starting my wake word engine.")
        return

    pa = pyaudio.PyAudio()
    try:
        audio_stream = pa.open(
            rate=porcupine.sample_rate,
            channels=1,
            format=pyaudio.paInt16,
            input=True,
            frames_per_buffer=porcupine.frame_length,
            input_device_index=None
        )
    except Exception as e:
        logging.error(f"PyAudio stream opening failed: {e}")
        speak("I couldn't open the audio stream. Please check your microphone.")
        if porcupine: porcupine.delete()
        if pa: pa.terminate()
        return

    scheduler = threading.Thread(target=scheduler_thread_func, daemon=True)
    scheduler.start()
    speak("Assistant ready. Say 'Hey Google' or your wake word to begin.")
    logging.info(
        f"Listening for wake word '{os.path.basename(WAKE_WORD_PPN if WAKE_WORD_PPN else WAKE_WORD_PPN_FILENAME)}'...")
    try:
        while True:
            try:
                pcm = audio_stream.read(porcupine.frame_length, exception_on_overflow=False)
                frame = struct.unpack_from("h" * porcupine.frame_length, pcm)
            except IOError as e:
                if hasattr(e, 'errno') and e.errno == pyaudio.paInputOverflowed:
                    logging.warning("Input overflowed. Skipping frame.")
                    continue
                else:
                    logging.error(f"Audio stream read IOError (not overflow): {e}")
                    time.sleep(0.1)
                    continue
            except Exception as e_read_generic:
                logging.error(f"Unexpected error reading or unpacking audio stream: {e_read_generic}")
                time.sleep(0.1)
                continue
            keyword_index = porcupine.process(frame)
            if keyword_index >= 0:
                logging.info("Wake word detected!")
                handle_command()
                logging.info(
                    f"Listening for wake word '{os.path.basename(WAKE_WORD_PPN if WAKE_WORD_PPN else WAKE_WORD_PPN_FILENAME)}' again...")
    except KeyboardInterrupt:
        logging.info("Keyboard interrupt received. Shutting down.")
        if tts: speak("Goodbye.")
    except Exception as e:
        logging.error(f"An unexpected error occurred in the main loop: {e}", exc_info=True)
        if tts: speak("An unexpected error occurred. I might need to restart.")
    finally:
        logging.info("Cleaning up resources...")
        if audio_stream is not None:
            try:
                if audio_stream.is_active(): audio_stream.stop_stream()
                audio_stream.close()
            except Exception as e_stream:
                logging.error(f"Error closing audio stream: {e_stream}")
        if pa is not None:
            try:
                pa.terminate()
            except Exception as e_pa:
                logging.error(f"Error terminating PyAudio: {e_pa}")
        if porcupine is not None:
            try:
                porcupine.delete()
            except Exception as e_porc:
                logging.error(f"Error deleting Porcupine instance: {e_porc}")
        logging.info("Shutdown complete.")


if __name__ == "__main__":
    critical_failure = False
    if not mic:
        logging.error("Main loop not started: Microphone not available.")
        if tts:
            speak("Assistant cannot start because the microphone is not available.")
        else:
            print("ERROR: Assistant cannot start: Microphone not available.")
        critical_failure = True
    if not PORCUPINE_ACCESS_KEY:
        logging.error("Main loop not started: Porcupine Access Key missing.")
        if not tts and not critical_failure:
            print("ERROR: Assistant cannot start: Porcupine Access Key missing.")
        elif tts and not critical_failure:
            speak("Porcupine Access Key is missing. Assistant cannot start.")
        critical_failure = True

    # WAKE_WORD_PPN is now initialized to None, so check its value directly
    if not WAKE_WORD_PPN or not os.path.exists(WAKE_WORD_PPN):  # Check if WAKE_WORD_PPN was successfully set and exists
        logging.error("Main loop not started: Wake word PPN file is missing, invalid, or could not be resolved.")
        if not critical_failure:
            msg = "Wake word file is missing or invalid. Assistant cannot start."
            if tts:
                speak(msg)
            else:
                print(f"ERROR: {msg}")
        critical_failure = True

    if not critical_failure:
        main_loop()
    else:
        logging.info("Assistant did not start due to critical errors listed above.")
