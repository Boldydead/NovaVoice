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
from dotenv import load_dotenv, find_dotenv
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

app_name_for_folders = "NovaVoice"

if getattr(sys, 'frozen', False):
    script_or_exe_dir = os.path.dirname(sys.executable)
elif __file__:
    script_or_exe_dir = os.path.dirname(os.path.abspath(__file__))
else:
    script_or_exe_dir = os.getcwd()

# --- .env File Loading ---
expected_dotenv_locations = []
loaded_dotenv_path_info = {"loaded": False, "path": "Not found"}
try:
    appdata_roaming_path = os.getenv('APPDATA')
    if appdata_roaming_path:
        app_config_dir = os.path.join(appdata_roaming_path, app_name_for_folders)
        os.makedirs(app_config_dir, exist_ok=True) # Ensure dir exists
        expected_dotenv_locations.append(os.path.join(app_config_dir, ".env"))
except Exception as e_appdata:
    logging.error(f"Error determining APPDATA for .env: {e_appdata}")
expected_dotenv_locations.append(os.path.join(script_or_exe_dir, ".env"))

for dotenv_path_to_try in expected_dotenv_locations:
    if dotenv_path_to_try and os.path.exists(dotenv_path_to_try):
        if load_dotenv(dotenv_path_to_try, override=True):
            logging.info(f"Successfully loaded .env from: {dotenv_path_to_try}")
            loaded_dotenv_path_info["loaded"] = True;
            loaded_dotenv_path_info["path"] = dotenv_path_to_try;
            break
        else:
            logging.warning(f"Found .env at {dotenv_path_to_try} but load_dotenv() failed.")
    elif dotenv_path_to_try:
        logging.info(f".env file not found at trial: {dotenv_path_to_try}")

if not loaded_dotenv_path_info["loaded"]:
    logging.info("Attempting fallback default load_dotenv() search.")
    env_path_by_find_dotenv = find_dotenv(usecwd=True, raise_error_if_not_found=False)
    if load_dotenv(override=True):
        logging.info(f".env loaded by default search (path: {env_path_by_find_dotenv or 'unknown'}).")
        loaded_dotenv_path_info["loaded"] = True;
        loaded_dotenv_path_info["path"] = env_path_by_find_dotenv or "Default search path"
    else:
        logging.error(".env could not be loaded. Keys will be missing.")
        print("CRITICAL: .env file not found by any method.")

PORCUPINE_ACCESS_KEY = os.getenv("PORCUPINE_ACCESS_KEY")

if not PORCUPINE_ACCESS_KEY:
    logging.error("PORCUPINE_ACCESS_KEY not set. Wake word detection fails.")
    print("WARNING: PORCUPINE_ACCESS_KEY not loaded. Ensure .env has the key.")

# --- Wake Word PPN File Loading ---
WAKE_WORD_PPN_FILENAME = "hey google_windows.ppn" # Make sure you have this file or update
WAKE_WORD_PPN = None
expected_ppn_locations = []
try:
    appdata_roaming_path = os.getenv('APPDATA')
    if appdata_roaming_path:
        app_config_dir_ppn = os.path.join(appdata_roaming_path, app_name_for_folders)
        os.makedirs(app_config_dir_ppn, exist_ok=True) # Ensure dir exists
        expected_ppn_locations.append(os.path.join(app_config_dir_ppn, WAKE_WORD_PPN_FILENAME))
except Exception:
    pass
expected_ppn_locations.append(os.path.join(script_or_exe_dir, WAKE_WORD_PPN_FILENAME))
try:
    # For bundled files, resource_path is crucial
    expected_ppn_locations.append(resource_path(WAKE_WORD_PPN_FILENAME))
except Exception:
    pass

for potential_path in expected_ppn_locations:
    if isinstance(potential_path, str) and os.path.exists(potential_path):
        WAKE_WORD_PPN = potential_path;
        logging.info(f"Wake word file '{WAKE_WORD_PPN_FILENAME}' found at: {WAKE_WORD_PPN}");
        break
    elif isinstance(potential_path, str):
        logging.info(f"PPN not found at trial: {potential_path}")

if not WAKE_WORD_PPN:
    logging.error(f"CRITICAL: Wake word file '{WAKE_WORD_PPN_FILENAME}' could not be located.")
    print(f"ERROR: Wake word file '{WAKE_WORD_PPN_FILENAME}' could not be located. Wake word detection will fail.")

# --- Writable User Data Directory (for caches, notes, etc.) ---
writable_user_data_dir = script_or_exe_dir # Default if AppData fails
try:
    appdata_path_local = os.getenv('LOCALAPPDATA')
    target_dir_base = None
    if appdata_path_local:
        target_dir_base = appdata_path_local
    elif os.getenv('APPDATA'): # Fallback to Roaming if LOCALAPPDATA is not set
        target_dir_base = os.getenv('APPDATA')

    if target_dir_base:
        potential_dir = os.path.join(target_dir_base, app_name_for_folders)
        os.makedirs(potential_dir, exist_ok=True) # Ensure the directory exists
        writable_user_data_dir = potential_dir
    logging.info(f"Using writable directory for user data (cache, notes, etc.): {writable_user_data_dir}")
except Exception as e_mkdir:
    logging.warning(f"Could not create/use AppData for user data, using script/exe directory as fallback: {e_mkdir}")

CACHE_FILE = os.path.join(writable_user_data_dir, "exe_cache.json")
DAILY_CACHE_FILE = os.path.join(writable_user_data_dir, "daily_text.json")
NOTES_FILE = os.path.join(writable_user_data_dir, "notes.txt")


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
    "slack": "slack.exe", "obs": "obs64.exe", "obs studio": "obs64.exe",
    "gimp": "gimp-2.10.exe", # Example, version might change
    "vlc": "vlc.exe",
    "photoshop": "photoshop.exe", "illustrator": "illustrator.exe",
    "blender": "blender.exe", "audacity": "audacity.exe",
    "putty": "putty.exe", "filezilla": "filezilla.exe",
    "virtualbox": "VirtualBox.exe", "vmware": "vmware.exe",
    # Add more common apps if desired
}


# === Load or Initialize Cache & Custom Commands ===
def load_json_data(file_path, default_data=None):
    if default_data is None: default_data = {}
    if os.path.exists(file_path):
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except json.JSONDecodeError:
            logging.warning(f"Could not decode JSON from {file_path}. Using default data.")
        except Exception as e:
            logging.error(f"Error loading {file_path}: {e}. Using default data.")
    else:
        logging.info(f"JSON file not found at {file_path}. Using default data for it.")
    return default_data


# --- Load Custom Commands ---
custom_cmds_path_to_use = None
custom_cmds_filename = "custom_commands.json"
custom_cmds_locations_to_try = []
try:
    appdata_roaming_path_cmds = os.getenv('APPDATA')
    if appdata_roaming_path_cmds:
        app_config_dir_cmds = os.path.join(appdata_roaming_path_cmds, app_name_for_folders)
        os.makedirs(app_config_dir_cmds, exist_ok=True) # Ensure dir exists
        custom_cmds_locations_to_try.append(os.path.join(app_config_dir_cmds, custom_cmds_filename))
except Exception:
    pass
custom_cmds_locations_to_try.append(os.path.join(script_or_exe_dir, custom_cmds_filename))

for potential_path in custom_cmds_locations_to_try:
    if potential_path and os.path.exists(potential_path):
        custom_cmds_path_to_use = potential_path;
        logging.info(f"Using custom_commands.json from: {custom_cmds_path_to_use}");
        break

if not custom_cmds_path_to_use:
    # If not found in preferred user locations, create a default path next to script/exe for initial creation
    custom_cmds_path_to_use = os.path.join(script_or_exe_dir, custom_cmds_filename)
    logging.info(
        f"{custom_cmds_filename} not found in primary user locations. "
        f"Will attempt to load from/create at: {custom_cmds_path_to_use} or use default empty commands."
    )
    custom_commands_data = {"commands": []} # Default if file doesn't exist at this path either
else:
    custom_commands_data = load_json_data(custom_cmds_path_to_use, {"commands": []})

custom_commands = custom_commands_data.get("commands", [])
exe_cache = load_json_data(CACHE_FILE, {})


# === Thread-safe TTS Setup ===
try:
    tts = pyttsx3.init()
    voices = tts.getProperty('voices')
    female_voice_id = None
    # Prioritize female voices, then a known good one if available (e.g., second voice often Zira on Windows)
    for voice in voices:
        if voice.gender == 'female' or 'female' in (voice.name or "").lower() or \
                (len(voices) > 1 and voice.id == voices[1].id): # voices[1] is often Zira
            female_voice_id = voice.id
            break
    if female_voice_id:
        tts.setProperty('voice', female_voice_id)
    elif voices: # Fallback to the first available voice if no female or specific voice found
        tts.setProperty('voice', voices[0].id)

    tts.setProperty('rate', 160) # Adjust rate as preferred
    tts.setProperty('volume', 1.0) # Full volume
    tts_lock = threading.Lock()
except Exception as e:
    logging.error(f"Failed to initialize TTS: {e}. Voice output will be disabled.")
    tts = None


def speak(text):
    if not tts:
        logging.warning(f"TTS not available. Intended to speak: {text}")
        print(f"ASSISTANT (TTS Disabled): {text}") # Fallback to print if TTS fails
        return
    with tts_lock: # Ensure thread-safe TTS operations
        try:
            logging.info(f"SPEAKING: {text}")
            tts.say(text)
            tts.runAndWait()
        except RuntimeError as e: # Specific error common with pyttsx3 if interrupted
            logging.error(f"TTS RuntimeError: {e}")
        except Exception as e: # Catch any other TTS errors
            logging.error(f"General TTS error: {e}")


# === Speech Recognizer ===
recognizer = sr.Recognizer()
try:
    mic = sr.Microphone()
    with mic as source:
        logging.info("Adjusting for ambient noise, please wait...")
        recognizer.adjust_for_ambient_noise(source, duration=1) # Adjust for 1 second
        logging.info("Ambient noise adjustment complete.")
except Exception as e:
    logging.error(f"Microphone not found or speech_recognition library issue: {e}")
    if tts:
        speak("I can't access a microphone. Please check your microphone settings.")
    else:
        print("ERROR: Microphone not found or speech_recognition library issue.")
    mic = None # Indicate microphone is not available


# === System Control Helpers (Windows-specific) ===
def mute_system():
    VK_MUTE = 0xAD
    KEYEVENTF_EXT = 0x0001 # KEYEVENTF_EXTENDEDKEY
    KEYEVENTF_KEYUP = 0x0002
    ctypes.windll.user32.keybd_event(VK_MUTE, 0, KEYEVENTF_EXT, 0) # Press Mute
    ctypes.windll.user32.keybd_event(VK_MUTE, 0, KEYEVENTF_KEYUP | KEYEVENTF_EXT, 0) # Release Mute
    logging.info("System mute toggled.")


def sleep_system():
    # For Windows: False, True, False = Standby (Sleep)
    # Forcing display off not included here for simplicity
    ctypes.windll.powrprof.SetSuspendState(False, True, False)
    logging.info("System going to sleep.")


def shutdown_system():
    # /s for shutdown, /t 0 for immediate
    os.system("shutdown /s /t 0")
    logging.info("System shutting down.")


def get_battery_status():
    try:
        batt = psutil.sensors_battery()
        if batt is None:
            logging.info("Battery information not available on this system.")
            return None
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
        # Limit recursion depth to avoid excessively long searches
        for root, dirs, files in os.walk(base, topdown=True):
            # Calculate current depth relative to the base search path
            current_depth = os.path.abspath(root).count(os.sep) - os.path.abspath(base).count(os.sep)
            if current_depth > 5: # Max depth (e.g., C:\Program Files\Folder1\Folder2\Folder3\Folder4\Folder5)
                logging.debug(f"Reached max search depth in {root}, pruning walk here.")
                dirs[:] = [] # Don't go deeper in this branch
                continue

            exe_name_lower_no_ext, exe_ext = os.path.splitext(exe_name.lower())
            if not exe_ext: exe_ext = ".exe" # Default to .exe if no extension

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
    if path and not os.path.exists(path): # Check if cached path is still valid
        logging.info(f"Cached path for {exe_name} ('{path}') no longer exists. Removing from cache.")
        exe_cache.pop(exe_name.lower(), None)
        path = None # Force a new search

    if path:
        try:
            os.startfile(path) # Use os.startfile for a non-blocking launch on Windows
            speak(f"Opening {app_name}.")
            return True
        except Exception as e:
            logging.error(f"Error starting cached {app_name} from {path}: {e}")
            exe_cache.pop(exe_name.lower(), None) # Remove bad cache entry

    speak(f"Searching for {app_name}, please wait.")

    def _search_and_launch_thread():
        found_path = find_executable(exe_name)
        if found_path:
            exe_cache[exe_name.lower()] = found_path # Update cache
            try:
                # Ensure CACHE_FILE (defined globally using writable_user_data_dir) is used
                with open(CACHE_FILE, 'w', encoding="utf-8") as f_cache:
                    json.dump(exe_cache, f_cache)
                logging.info(f"Updated exe_cache at {CACHE_FILE}")
            except Exception as e_cache_write:
                logging.error(f"Error writing to exe_cache.json at {CACHE_FILE}: {e_cache_write}")

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
    return False # Indicates that the launch is happening asynchronously


# === Built-in Browser Opener ===
def open_default_browser(url="https://www.google.com"):
    try:
        webbrowser.open(url) # Opens in the default browser
        logging.info(f"Opened {url} in default browser.")
        return True
    except Exception as e:
        logging.error(f"Could not open URL {url} in browser: {e}")
        return False


# === GUI Prompt for Manual Path ===
def prompt_for_exe(title="Select executable"):
    try:
        root = tk.Tk()
        root.withdraw() # Hide the main Tkinter window
        path = filedialog.askopenfilename(
            title=title,
            filetypes=[("Executable", "*.exe"), ("All files", "*.*")],
            initialdir=os.path.expanduser("~\\Desktop") # Start on Desktop
        )
        root.destroy() # Clean up Tkinter window
        return path
    except Exception as e:
        logging.error(f"Error showing Tkinter dialog: {e}")
        return None


# --- Command Handler Functions ---
def cmd_open_google(command_text):
    if open_default_browser("https://www.google.com"):
        speak("Opening Google.")
    else:
        speak("I couldn't open Google in your browser.")


def cmd_open_youtube(command_text):
    # Using a more direct youtube.com URL
    if open_default_browser("https://www.youtube.com"): # Using https://www.youtube.com for a stable redirect
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
        speak("Sorry, I couldn't get the battery information for this device.")


def cmd_get_current_time(command_text):
    now = datetime.datetime.now()
    time_str = now.strftime("%I:%M %p") # e.g., "03:30 PM"
    speak(f"The current time is {time_str}.")


def cmd_get_current_date(command_text):
    now = datetime.datetime.now()
    date_str = now.strftime("%A, %B %d, %Y") # e.g., "Monday, May 27, 2024"
    speak(f"Today is {date_str}.")


def cmd_get_system_info(command_text):
    speak("Getting system information...")
    try:
        # CPU Usage
        cpu_usage = psutil.cpu_percent(interval=1) # Get usage over 1 second
        speak(f"Current CPU usage is {cpu_usage} percent.")

        # Memory (RAM) Usage
        mem = psutil.virtual_memory()
        mem_total_gb = round(mem.total / (1024**3), 1) # Total RAM in GB
        mem_used_gb = round(mem.used / (1024**3), 1)   # Used RAM in GB
        speak(f"RAM usage is {mem.percent} percent. Currently using {mem_used_gb} of {mem_total_gb} gigabytes.")

        # Disk Usage (for C: drive on Windows, / on Linux/Mac)
        disk_path = 'C:\\' if os.name == 'nt' else '/'
        disk = psutil.disk_usage(disk_path)
        disk_total_gb = round(disk.total / (1024**3), 1) # Total disk space in GB
        disk_used_gb = round(disk.used / (1024**3), 1)   # Used disk space in GB
        speak(f"Disk {disk_path} is {disk.percent} percent used. Used {disk_used_gb} of {disk_total_gb} gigabytes.")

    except Exception as e:
        speak("Sorry, I couldn't retrieve all system information at the moment.")
        logging.error(f"Error getting system info: {e}")


def cmd_create_folder(command_text):
    folder_name = ""
    # Define various trigger phrases for creating a folder
    triggers = ["create folder ", "make folder ", "new folder ", "create directory ", "make directory ", "new directory "]
    for trigger in triggers:
        if command_text.lower().startswith(trigger):
            folder_name = command_text[len(trigger):].strip()
            break

    if not folder_name:
        speak("What name would you like for the folder?")
        # Potentially listen for a follow-up, but for now, just prompts.
        return

    # Sanitize folder name (remove invalid characters for Windows filenames)
    invalid_chars = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']
    for char in invalid_chars:
        folder_name = folder_name.replace(char, '')

    if not folder_name: # If name becomes empty after sanitization
        speak("That doesn't seem to be a valid folder name.")
        return

    # Create folder on the user's Desktop
    full_folder_path = os.path.join(os.path.join(os.path.expanduser("~"), "Desktop"), folder_name)
    try:
        if not os.path.exists(full_folder_path):
            os.makedirs(full_folder_path)
            speak(f"Folder '{folder_name}' has been created on your Desktop.")
        else:
            speak(f"The folder '{folder_name}' already exists on your Desktop.")
    except Exception as e:
        speak(f"Sorry, I couldn't create the folder '{folder_name}'.")
        logging.error(f"Error creating folder {full_folder_path}: {e}")


def cmd_empty_recycle_bin(command_text):
    if os.name == 'nt': # Windows-specific
        try:
            # Flags for SHEmptyRecycleBinW:
            SHERB_NOCONFIRMATION = 0x00000001 # No confirmation dialog
            SHERB_NOPROGRESSUI = 0x00000002   # No progress UI
            SHERB_NOSOUND = 0x00000004        # No sound
            flags = SHERB_NOCONFIRMATION | SHERB_NOPROGRESSUI | SHERB_NOSOUND

            # Call the Windows API function
            # The first parameter (hwnd) can be None for no parent window.
            # The second parameter (pszRootPath) can be None to empty all recycle bins.
            result = ctypes.windll.shell32.SHEmptyRecycleBinW(None, None, flags)

            if result == 0: # S_OK
                speak("The Recycle Bin has been emptied.")
            elif result == 1: # S_FALSE (often means already empty or user cancelled if flags were different)
                speak("The Recycle Bin is already empty or the operation was not fully completed.")
            else: # Other HRESULT error codes
                speak("I encountered an issue trying to empty the Recycle Bin.")
                logging.warning(f"SHEmptyRecycleBinW returned an error code: {result}")
        except Exception as e:
            speak("An error occurred while trying to empty the Recycle Bin.")
            logging.error(f"Exception while emptying recycle bin: {e}")
    else:
        speak("I can only empty the Recycle Bin on Windows systems.")


def cmd_take_note(command_text): # Expects the full command text to extract the note
    note_content = command_text # Default to full command if no specific trigger found
    # Define trigger phrases for taking a note
    triggers = ["take a note ", "note to self ", "add note ", "make a note ", "new note "]
    for trigger in triggers:
        if command_text.lower().startswith(trigger):
            note_content = command_text[len(trigger):].strip()
            break

    if not note_content or note_content.lower() == "note": # If only "note" or empty after trigger
        speak("What should the note say?")
        # Could add a listen here for the content, but simple for now
        return

    try:
        # NOTES_FILE is defined globally using writable_user_data_dir
        with open(NOTES_FILE, "a", encoding="utf-8") as f: # Append to notes file
            f.write(f"{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}: {note_content}\n")
        speak("Noted.")
        logging.info(f"Note added to {NOTES_FILE}")
    except Exception as e:
        speak("Sorry, I couldn't save your note right now.")
        logging.error(f"Error writing notes to {NOTES_FILE}: {e}")


def cmd_read_notes(command_text):
    try:
        # NOTES_FILE is defined globally
        if os.path.exists(NOTES_FILE):
            with open(NOTES_FILE, "r", encoding="utf-8") as f:
                notes = f.readlines()
            if notes:
                speak("Here are your latest notes:")
                time.sleep(0.2) # Slight pause for effect
                # Read the last few notes (e.g., last 3)
                for note_line in notes[-min(len(notes), 3):]:
                    # Attempt to extract content after "YYYY-MM-DD HH:MM:SS: "
                    content = note_line.split(": ", 1)[1].strip() if ": " in note_line else note_line.strip()
                    speak(content)
                    time.sleep(0.3) # Pause between notes
            else:
                speak("You don't have any notes saved yet.")
        else:
            speak("You don't have any notes saved yet.")
            logging.info(f"Notes file not found at {NOTES_FILE} when trying to read.")
    except Exception as e:
        speak("Sorry, I couldn't read your notes right now.")
        logging.error(f"Error reading notes from {NOTES_FILE}: {e}")


def cmd_show_help(command_text):
    built_ins = [
        "open Google", "open YouTube",
        "take a note [followed by your note]", "read my notes",
        "create folder [folder name]", "empty recycle bin",
        "system information",
        "mute system", "put computer to sleep", "turn off computer",
        "what's my battery level", "what time is it", "what is today's date",
        "tell me today's Bible verse",
        "what can you open"
    ]
    speak("I can understand commands like:")
    time.sleep(0.2)
    # Speak a few examples
    for i, cmd_example in enumerate(built_ins):
        if i < 7: # Speak first 7 examples
            speak(cmd_example)
            time.sleep(0.1)
        else:
            break
    speak("And I can try to open applications by name, like 'open Word' or 'launch Spotify'.")
    if custom_commands:
        speak("I also know your custom commands, such as:")
        for i, c_cmd in enumerate(custom_commands):
            if i < 2: # Speak first 2 custom command phrases
                speak(c_cmd.get("phrase", "a custom task"))
                time.sleep(0.1)
            else:
                break
        if len(custom_commands) > 2:
            speak("and a few others.")
    speak("For a more detailed overview of my functions, just ask 'what can you do'.")


def cmd_list_known_apps(command_text):
    known_apps_friendly_names = set()
    # From standard APP_LAUNCH_MAP
    for app_alias in APP_LAUNCH_MAP.keys():
        known_apps_friendly_names.add(app_alias.title())

    # From exe_cache (previously found apps)
    if exe_cache:
        for exe_name_lower_key in exe_cache.keys():
            app_name_from_cache = exe_name_lower_key.replace(".exe", "")
            # Simple title casing for display
            app_name_from_cache = ' '.join(word.capitalize() for word in app_name_from_cache.split())
            known_apps_friendly_names.add(app_name_from_cache)

    # From custom commands that launch executables
    for cmd_config in custom_commands:
        if cmd_config.get("action") == "launch_executable":
            app_name = cmd_config.get("app_name")
            if app_name:
                known_apps_friendly_names.add(app_name.title())
            else: # Fallback to deriving from exe_name if app_name isn't specified
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

        if not apps_to_speak: # Should not happen if known_apps_friendly_names was populated
            speak("I couldn't compile a list of apps right now. Try asking me to open one directly!")
            return

        if len(apps_to_speak) > 7: # Speak a selection if the list is long
            speak_text += ", ".join(apps_to_speak[:6]) + ", and others such as " + apps_to_speak[6] + "."
        elif len(apps_to_speak) > 1:
            speak_text += ", ".join(apps_to_speak[:-1]) + ", and " + apps_to_speak[-1] + "."
        else: # Only one app
            speak_text += apps_to_speak[0] + "."

        speak(speak_text)
        speak("You can also try asking for any other application not mentioned, and I'll search for it.")


def cmd_tell_capabilities(command_text):
    speak("I can help you with various tasks on your computer after you say my wake word.")
    time.sleep(0.2)
    speak("Here's an overview of what I can do:")
    time.sleep(0.3)
    speak(
        "First, I can open applications. For example, you can ask me to 'open Chrome', 'launch Spotify', or 'open Notepad'. "
        "I have a list of common applications, and I can also search for others if you ask. "
        "If you'd like a list of applications I'm already familiar with or have found before, you can say 'what applications can you open'.")
    time.sleep(0.5)
    speak("Second, I can open websites. For example, try saying 'open Google' or 'open YouTube'.")
    time.sleep(0.5)
    speak(
        "Third, I can provide information. You can ask me 'what time is it', 'what is today's date', 'what's my battery level', or for 'system information' like CPU and RAM usage.")
    time.sleep(0.5)
    speak(
        "Fourth, I can perform system actions on Windows. For instance, 'mute the system', 'put computer to sleep', 'turn off the computer', or 'empty the recycle bin'.")
    time.sleep(0.5)
    speak(
        "I can also help with simple productivity tasks like 'take a note [your note here]' and 'read my notes', or 'create folder [folder name]' on your desktop.")
    time.sleep(0.5)
    speak(
        "I also try to fetch a daily Bible verse for you if you ask, though please note this is a basic feature.")
    time.sleep(0.3)
    if custom_commands:
        speak(
            "Additionally, I can run custom commands that you've set up in the custom_commands.json file. "
            "These can include launching specific programs with a unique phrase, opening particular web pages, or running shell scripts.")
    else:
        speak(
            "And if you were to set up custom commands in the custom_commands.json file, I could perform those too, allowing you to tailor my actions to your needs!")
    time.sleep(0.5)
    speak("For a shorter list of example command phrases you can use, just say 'help'.")


# === Command Dispatcher ===
COMMAND_DISPATCHER = {
    # Capabilities & Help
    "what can you do": cmd_tell_capabilities, "capabilities": cmd_tell_capabilities,
    "tell me your capabilities": cmd_tell_capabilities, "what can you do and open": cmd_tell_capabilities,
    "describe your functions": cmd_tell_capabilities, "what all can you do": cmd_tell_capabilities,
    "help": cmd_show_help, "list commands": cmd_show_help,

    # Web Opening
    "google": cmd_open_google, "youtube": cmd_open_youtube,

    # System Control
    "mute": cmd_mute_system, "mute system": cmd_mute_system, "toggle mute": cmd_mute_system,
    "sleep": cmd_sleep_system, "pc sleep": cmd_sleep_system, "computer sleep": cmd_sleep_system, "put computer to sleep": cmd_sleep_system,
    "turn off": cmd_shutdown_system, "shut down": cmd_shutdown_system, "shutdown": cmd_shutdown_system, "shutdown computer": cmd_shutdown_system,
    "empty recycle bin": cmd_empty_recycle_bin, "clear recycle bin": cmd_empty_recycle_bin,

    # Information
    "battery": cmd_get_battery, "what's my battery": cmd_get_battery, "battery level": cmd_get_battery, "battery status": cmd_get_battery,
    "time": cmd_get_current_time, "what time is it": cmd_get_current_time, "current time": cmd_get_current_time,
    "date": cmd_get_current_date, "today's date": cmd_get_current_date, "what is today's date": cmd_get_current_date,
    "system information": cmd_get_system_info, "system info": cmd_get_system_info, "pc status": cmd_get_system_info,
    "computer status": cmd_get_system_info, "cpu usage": cmd_get_system_info, "ram usage": cmd_get_system_info,
    "disk space": cmd_get_system_info,

    # Application Listing
    "what can you open": cmd_list_known_apps, "what apps can you open": cmd_list_known_apps,
    "list applications": cmd_list_known_apps, "list apps": cmd_list_known_apps, "show known apps": cmd_list_known_apps,

    # Notes & Folders (These commands often need the full text)
    "take a note": cmd_take_note, "note to self": cmd_take_note, "add note": cmd_take_note,
    "make a note": cmd_take_note, "new note": cmd_take_note, # `handle_command` passes full text
    "read my notes": cmd_read_notes, "show my notes": cmd_read_notes, "what are my notes": cmd_read_notes,
    "display notes": cmd_read_notes, "get notes": cmd_read_notes,
    "create folder": cmd_create_folder, "make folder": cmd_create_folder, "new folder": cmd_create_folder,
    "create directory": cmd_create_folder, "make directory": cmd_create_folder, "new directory": cmd_create_folder, # `handle_command` passes full text
}


# === Voice Command Handling ===
def handle_command(phrase_listen_timeout=5.0, phrase_time_limit=10.0):
    if not mic:
        speak("Microphone is not available, cannot listen for commands.")
        return

    speak("What would you like me to do?")
    command_text = "" # Store the full recognized text
    try:
        with mic as source:
            logging.info("Listening for command...")
            try:
                # Listen for the user's command
                audio = recognizer.listen(source, timeout=phrase_listen_timeout, phrase_time_limit=phrase_time_limit)
            except sr.WaitTimeoutError:
                logging.info("No speech detected within timeout.")
                # speak("I didn't hear anything.") # Optional feedback
                return

        logging.info("Processing command...")
        command_text = recognizer.recognize_google(audio).lower().strip()
        # Don't speak back "You said: nothing" if it's a dismissal.
        # We'll let the specific handling below manage the response.

    except sr.UnknownValueError:
        speak("Sorry, I didn't catch that clearly.")
        logging.info("Google Speech Recognition could not understand audio.")
        return
    except sr.RequestError as e:
        speak("It seems I'm having trouble reaching the speech service.")
        logging.error(f"Could not request results from Google Speech Recognition service; {e}")
        return
    except Exception as e: # Catch-all for other speech recognition issues
        speak("An unexpected error occurred while trying to understand you.")
        logging.error(f"Error during speech recognition: {e}")
        return

    if not command_text:
        return # No command was recognized

    # --- NEW: Handle "nothing" or similar dismissal responses ---
    dismissal_phrases = [
        "nothing", "no thanks", "not now", "i'm good", "that's all",
        "nevermind", "actually nothing", "don't do anything",
        "nothing right now", "nothing for now", "no thank you",
        "it's okay", "it's alright", "no", "nope"
    ]
    if command_text in dismissal_phrases:
        speak("Okay.") # Or "Alright.", "Understood.", "No problem."
        logging.info(f"User indicated no command with: '{command_text}'")
        return # Exit handle_command and go back to listening for wake word

    # If it wasn't a dismissal, then speak back what was heard
    speak(f"You said: {command_text}")


    # --- Process the command (existing logic) ---
    processed_command = command_text # This will be stripped of wake words for matching
    common_wake_phrases = ("hey assistant", "assistant", "hey google", "google") # Add more if needed
    for trig in common_wake_phrases:
        if processed_command.startswith(trig):
            processed_command = processed_command.replace(trig, "", 1).strip()
            logging.debug(f"Command after stripping '{trig}': {processed_command}")
            break # Stop after stripping the first recognized wake phrase

    if not processed_command: # If command was only the wake phrase
        speak("I heard my name, but what would you like me to do?")
        return

    # 1. Exact match on processed_command (for simple commands)
    if processed_command in COMMAND_DISPATCHER:
        func_to_call = COMMAND_DISPATCHER[processed_command]
        # Pass original command_text if the function needs it (e.g., for note content)
        if func_to_call in [cmd_take_note, cmd_create_folder]:
            func_to_call(command_text)
        else:
            func_to_call(processed_command)
        return

    # 2. Substring match (longer keys first for better specificity)
    # Useful for commands embedded in longer phrases, e.g., "assistant, can you tell me the time"
    for keyword in sorted(COMMAND_DISPATCHER.keys(), key=len, reverse=True):
        if keyword and keyword in processed_command: # Check if keyword is non-empty
            func_to_call = COMMAND_DISPATCHER[keyword]
            if func_to_call in [cmd_take_note, cmd_create_folder]:
                func_to_call(command_text) # These need the full text
            elif func_to_call == cmd_get_system_info and any(s_info in command_text for s_info in ["cpu", "ram", "disk", "memory"]):
                 # If asking for specific system info, pass full text so it can be parsed if needed (though current cmd_get_system_info gives all)
                func_to_call(command_text)
            else:
                func_to_call(processed_command) # Others can use the stripped command
            return

    # 3. Handle "open <app>" or "launch <app>"
    if processed_command.startswith(("open ", "launch ")):
        parts = processed_command.split(" ", 1)
        if len(parts) > 1:
            app_to_launch = parts[1].strip()
            if app_to_launch:
                # Check standard map first
                app_exe = APP_LAUNCH_MAP.get(app_to_launch.lower())
                if app_exe:
                    launch_executable_async(app_exe, app_to_launch.title())
                    return

                # Check custom commands for this app name (if action is launch_executable)
                for c_cmd in custom_commands:
                    custom_app_name = c_cmd.get("app_name", "").lower()
                    if custom_app_name == app_to_launch.lower() and c_cmd.get("action") == "launch_executable":
                        exe_name_custom = c_cmd.get("exe_name")
                        if exe_name_custom:
                            launch_executable_async(exe_name_custom, app_to_launch.title())
                            return
                        else:
                            logging.warning(f"Custom command for '{app_to_launch}' is missing 'exe_name'.")
                            break # Stop checking custom commands for this app

                # If not in maps or custom app names, try a generic guess
                # (e.g., "open mygame" -> "mygame.exe")
                app_exe_guess = app_to_launch + ".exe" if not app_to_launch.lower().endswith(".exe") else app_to_launch
                logging.info(f"Attempting generic launch for: {app_exe_guess} (app name: {app_to_launch.title()})")
                launch_executable_async(app_exe_guess, app_to_launch.title())
                return

    # 4. Handle "Edge" specifically if not caught by "open edge"
    if "edge" in processed_command and "open" in processed_command: # "open microsoft edge"
         if open_default_browser("https://www.microsoft.com/edge"): # Or just launch msedge.exe via APP_LAUNCH_MAP
             speak("Opening Microsoft Edge.")
         else:
             speak("I couldn't open Edge.")
         return

    # 5. Daily Bible Verse
    if any(s in processed_command for s in ["bible verse", "daily text", "today's text", "scripture"]):
        speak("Fetching today's Bible verse, please wait.")
        daily_scripture = fetch_daily_text()
        speak(daily_scripture)
        return

    # 6. Custom commands (by exact phrase match on the processed_command)
    for cmd_config in custom_commands:
        phrase = cmd_config.get("phrase", "").lower()
        if phrase and phrase == processed_command: # Exact match on the user-defined phrase
            action = cmd_config.get("action")
            response = cmd_config.get("response", f"Okay, performing your custom action for '{phrase}'.")
            logging.info(f"Executing custom command for exact phrase: '{phrase}', action: {action}")

            if action == "launch_executable":
                exe_name = cmd_config.get("exe_name")
                # Use app_name from config, fallback to exe_name, then generic "application"
                app_name_custom = cmd_config.get("app_name", exe_name or "the application")
                if exe_name:
                    launch_executable_async(exe_name, app_name_custom)
                else:
                    speak(f"Executable name ('exe_name') is missing for the custom command '{phrase}'.")
            elif action == "url":
                url_to_open = cmd_config.get("url")
                if url_to_open:
                    if open_default_browser(url_to_open):
                        speak(response) # Speak custom response or default
                    else:
                        speak(f"I couldn't open the URL for the custom command '{phrase}'.")
                else:
                    speak(f"URL is missing for the custom command '{phrase}'.")
            elif action == "shell":
                shell_command_to_run = cmd_config.get("shell_cmd")
                if shell_command_to_run:
                    try:
                        # CREATE_NO_WINDOW flag for Windows to run silently in background
                        subprocess.Popen(shell_command_to_run, shell=True,
                                         creationflags=0x08000000 if os.name == 'nt' else 0)
                        speak(response)
                    except Exception as e_shell:
                        logging.error(f"Error executing shell command '{shell_command_to_run}': {e_shell}")
                        speak(f"I couldn't run the shell command for '{phrase}'.")
                else:
                    speak(f"Shell command ('shell_cmd') is missing for the custom command '{phrase}'.")
            else:
                speak(f"The action type '{action}' for the custom command '{phrase}' is unknown or not supported.")
            return # Custom command executed

    # Fallback if no command matched
    speak("I'm not sure how to do that yet.")


# === Daily Text Function ===
def fetch_daily_text():
    today_str = datetime.datetime.now().strftime("%Y-%m-%d")
    # DAILY_CACHE_FILE is defined globally using writable_user_data_dir
    cached_data = load_json_data(DAILY_CACHE_FILE)

    if cached_data.get("date") == today_str and cached_data.get("text"):
        logging.info(f"Using cached daily text from {DAILY_CACHE_FILE}.")
        return cached_data["text"]

    # Placeholder: Direct users to the website as scraping is unreliable
    logging.warning(
        "Web scraping for jw.org daily text is highly dependent on website structure and often fails or becomes outdated."
        " Using a placeholder message instead.")
    full_text = (
        "Fetching the daily Bible verse from JW.org directly is currently unreliable due to potential website changes. "
        "For the most accurate and up-to-date daily text, please visit JW.org directly through your web browser."
    )

    try:
        with open(DAILY_CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump({"date": today_str, "text": full_text}, f)
        logging.info(f"Updated daily_text_cache at {DAILY_CACHE_FILE} with placeholder message.")
    except Exception as e:
        logging.error(f"Error writing to daily cache file {DAILY_CACHE_FILE}: {e}")
    return full_text


# === Scheduler Thread ===
daily_greeting_done_today = False

def scheduler_thread_func():
    global daily_greeting_done_today
    logging.info("Scheduler thread started.")
    while True:
        now = datetime.datetime.now()
        hour = now.hour

        # Morning greeting and daily text (e.g., between 7 AM and 8 AM)
        if 7 <= hour < 8 and not daily_greeting_done_today:
            logging.info("Scheduler: Morning sequence triggered.")
            speak("Good morning! Here's something for your day.")
            time.sleep(0.5) # Pause for effect
            text_for_day = fetch_daily_text()
            speak(text_for_day)
            daily_greeting_done_today = True

        # Reset flag after midnight (e.g., if it's before 6 AM and flag is still true)
        if hour < 6: # Resetting period
            if daily_greeting_done_today:
                logging.info("Scheduler: Resetting daily_greeting_done_today flag for the new day.")
                daily_greeting_done_today = False

        time.sleep(30) # Check every 30 seconds


# === Main Wake-Word Loop ===
porcupine = None
pa = None
audio_stream = None


def main_loop():
    global porcupine, pa, audio_stream

    # Critical check for Porcupine setup
    if not PORCUPINE_ACCESS_KEY or not WAKE_WORD_PPN or not os.path.exists(WAKE_WORD_PPN):
        msg = "Porcupine Access Key or Wake Word PPN file is missing, invalid, or not found. Wake word engine cannot start."
        logging.error(msg)
        speak(msg) if tts else print(f"ERROR: {msg}")
        return

    try:
        porcupine = pvporcupine.create(
            access_key=PORCUPINE_ACCESS_KEY,
            keyword_paths=[WAKE_WORD_PPN],
            sensitivities=[0.65] # Adjust sensitivity as needed (0.0 to 1.0)
        )
    except pvporcupine.PorcupineError as pe: # Specific Porcupine errors
        logging.error(f"Porcupine initialization failed: {pe}")
        speak("I couldn't start my wake word engine. Please check the access key and the wake word file path and integrity.")
        return
    except Exception as e_porc_general: # Other unexpected errors
        logging.error(f"An unexpected error occurred during Porcupine initialization: {e_porc_general}")
        speak("An unexpected error occurred while starting my wake word engine. Please check the logs.")
        return

    pa = pyaudio.PyAudio()
    try:
        audio_stream = pa.open(
            rate=porcupine.sample_rate,
            channels=1,
            format=pyaudio.paInt16,
            input=True,
            frames_per_buffer=porcupine.frame_length,
            input_device_index=None # Use default input device
        )
    except Exception as e_audio:
        logging.error(f"PyAudio stream opening failed: {e_audio}")
        speak("I couldn't open the audio stream. Please check your microphone and audio settings.")
        if porcupine: porcupine.delete() # Clean up Porcupine if audio fails
        if pa: pa.terminate() # Clean up PyAudio
        return

    # Start the scheduler thread
    scheduler = threading.Thread(target=scheduler_thread_func, daemon=True)
    scheduler.start()

    speak("Assistant ready. Say 'Hey Google' or your wake word to begin.")
    ppn_base = os.path.basename(WAKE_WORD_PPN if WAKE_WORD_PPN and isinstance(WAKE_WORD_PPN, str) else WAKE_WORD_PPN_FILENAME)
    logging.info(f"Listening for wake word '{ppn_base}'...")

    try:
        while True:
            try:
                pcm = audio_stream.read(porcupine.frame_length, exception_on_overflow=False)
                frame = struct.unpack_from("h" * porcupine.frame_length, pcm)
            except IOError as e_io:
                if hasattr(e_io, 'errno') and e_io.errno == pyaudio.paInputOverflowed:
                    logging.warning("Audio input overflowed. Skipping this frame.")
                    continue # Skip this frame and try again
                # For other IOErrors, log and briefly pause
                logging.error(f"Audio stream read IOError (not overflow): {e_io}")
                time.sleep(0.1)
                continue
            except Exception as e_read_generic: # Catch other potential errors during read/unpack
                logging.error(f"Unexpected error reading or unpacking audio stream: {e_read_generic}")
                time.sleep(0.1) # Brief pause before retrying
                continue

            keyword_index = porcupine.process(frame)
            if keyword_index >= 0: # Wake word detected
                logging.info("Wake word detected!")
                # Potentially add a sound effect here if desired
                handle_command() # Process the command
                logging.info(f"Listening for wake word '{ppn_base}' again...")

    except KeyboardInterrupt:
        logging.info("Keyboard interrupt received. Shutting down assistant.")
        if tts: speak("Goodbye.")
    except Exception as e_main_loop: # Catch-all for unexpected errors in the main loop
        logging.error(f"An unexpected error occurred in the main loop: {e_main_loop}", exc_info=True)
        if tts: speak("An unexpected error occurred. I might need to restart.")
    finally:
        logging.info("Cleaning up resources...")
        if audio_stream is not None:
            try:
                if audio_stream.is_active(): audio_stream.stop_stream()
                audio_stream.close()
            except Exception as e_stream_close:
                logging.error(f"Error closing audio stream: {e_stream_close}")
        if pa is not None:
            try:
                pa.terminate()
            except Exception as e_pa_terminate:
                logging.error(f"Error terminating PyAudio: {e_pa_terminate}")
        if porcupine is not None:
            try:
                porcupine.delete()
            except Exception as e_porc_delete:
                logging.error(f"Error deleting Porcupine instance: {e_porc_delete}")
        logging.info("Shutdown complete.")


if __name__ == "__main__":
    critical_failure = False
    error_messages = []

    if not mic:
        error_messages.append("Microphone not available or not initialized.")
    if not PORCUPINE_ACCESS_KEY:
        error_messages.append("Porcupine Access Key (PORCUPINE_ACCESS_KEY) is missing. Check your .env file.")
    if not WAKE_WORD_PPN or not os.path.exists(WAKE_WORD_PPN):
        error_messages.append(
            f"Wake word PPN file ('{WAKE_WORD_PPN_FILENAME}') is missing, invalid, or could not be located at expected paths."
        )

    if error_messages:
        critical_failure = True
        logging.error("Assistant cannot start due to the following critical errors:")
        for msg in error_messages:
            logging.error(f"- {msg}")
            # Attempt to speak the first critical error if TTS is available, otherwise print
            if tts and msg == error_messages[0]: # Only speak the first one to avoid too much talking
                speak(msg + " The assistant cannot start.")
            elif not tts: # Print all if TTS is not available
                 print(f"STARTUP ERROR: {msg} The assistant cannot start.")
            if tts and msg != error_messages[0]: # Print subsequent errors if TTS spoke the first
                 print(f"ADDITIONAL STARTUP ERROR: {msg}")

    if not critical_failure:
        main_loop()
    else:
        logging.info("Assistant did not start due to critical errors listed above. Please resolve them and try again.")
        if not tts: # Ensure there's some console output if TTS didn't cover it
            print("Assistant startup failed. Check logs for details.")
        # Keep console open for a bit if run directly, so user can see errors
        if not getattr(sys, 'frozen', False): # If not a frozen exe
            try:
                input("Press Enter to exit...")
            except EOFError: # In case input is not available (e.g. piped)
                pass
