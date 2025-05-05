import pvporcupine
import pyaudio
import struct
import subprocess
import os
import pyttsx3
import speech_recognition as sr
import webbrowser
from dotenv import load_dotenv
import os

load_dotenv()
ACCESS_KEY = os.getenv("PV_ACCESS_KEY")

# === Config ===
ACCESS_KEY = "PV_ACCESS_KEY"
WAKE_WORD_PPN = "hey-assistant_en_windows.ppn"

# === Locks & Initialization ===
tts = pyttsx3.init()
tts.setProperty('rate', 175)
tts.setProperty('volume', 1.0)

recognizer = sr.Recognizer()


def speak(text):
    tts.say(text)
    tts.runAndWait()


# === Speech Recognizer ===
recognizer = sr.Recognizer()
mic = sr.Microphone()

# === Porcupine & PyAudio Initialization ===
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


# === Program launcher helper ===
def search_and_launch_program(exe_names, app_name):
    """
    Searches common Program Files directories for any of the given exe_names.
    If found, speaks the full path then launches it.
    Returns True on success, False otherwise.
    """
    for base in (r"C:\Program Files", r"C:\Program Files (x86)"):
        for root, dirs, files in os.walk(base):
            for f in files:
                if f.lower() in exe_names:
                    full_path = os.path.join(root, f)
                    speak(f"Found {app_name} in {root}")
                    subprocess.Popen([full_path])
                    speak(f"Opening {app_name}.")
                    return True
    return False


def open_browser():
    # Try Microsoft Edge first (UWP app)
    try:
        os.system("start microsoft-edge:https://www.google.com")
        speak("Opening Microsoft Edge for you.")
        return
    except:
        pass

    # Fallback to installed browsers
    exe_names = ["chrome.exe", "firefox.exe", "brave.exe"]
    if not search_and_launch_program(exe_names, "browser"):
        speak("Hmm, I couldn't find a browser.")


def open_ide():
    exe_names = ["code.exe", "pycharm64.exe", "idea64.exe", "eclipse.exe"]
    if not search_and_launch_program(exe_names, "code editor"):
        speak("Sorry, I couldn't find any coding apps installed.")


# === Voice Command Logic ===
def handle_command(phrase_time=4.0):
    # Prompt for action
    speak("What would you like me to open?")
    # Capture the next few seconds of audio frames
    frames = []
    reads = int(porcupine.sample_rate / porcupine.frame_length * phrase_time)
    for _ in range(reads):
        pcm = stream.read(porcupine.frame_length, exception_on_overflow=False)
        frames.append(pcm)
    raw_data = b"".join(frames)
    audio_data = sr.AudioData(raw_data, porcupine.sample_rate, 2)

    try:
        command = recognizer.recognize_google(audio_data).lower().strip()
        speak(f"You said: {command}")
    except sr.UnknownValueError:
        speak("Sorry, I couldn't catch that.")
        return
    except sr.RequestError:
        speak("It seems I'm offline.")
        return

    # Remove wake triggers
    for trigger in ("hey assistant", "assistant", "hey"):
        if command.startswith(trigger):
            command = command.replace(trigger, "", 1).strip()

    # Built-in commands
    if "google" in command:
        webbrowser.open("https://www.google.com")
        speak("Opening Google.")
    elif "youtube" in command:
        webbrowser.open("https://www.youtube.com")
        speak("Opening YouTube.")
    elif "edge" in command:
        speak("Launching Microsoft Edge.")
        os.system("start microsoft-edge:https://www.google.com")
    elif "chrome" in command:
        search_and_launch_program(["chrome.exe"], "Google Chrome")
    elif "firefox" in command:
        search_and_launch_program(["firefox.exe"], "Mozilla Firefox")
    elif "brave" in command:
        search_and_launch_program(["brave.exe"], "Brave Browser")
    elif "pycharm" in command:
        # Try auto-detect first
        if not search_and_launch_program(["pycharm64.exe"], "PyCharm"):
            # If not found, ask for directory
            speak("I couldn’t find PyCharm automatically. Please tell me the full path to the PyCharm executable.")
            with mic as source:
                recognizer.adjust_for_ambient_noise(source, duration=1)
                try:
                    audio = recognizer.listen(source, timeout=10, phrase_time_limit=8)
                    custom_path = recognizer.recognize_google(audio).strip()
                    if os.path.exists(custom_path):
                        dir_root = os.path.dirname(custom_path)
                        speak(f"Found PyCharm in {dir_root}")
                        subprocess.Popen([custom_path])
                        speak("Opening PyCharm.")
                    else:
                        speak("That path doesn’t exist. Please check the path and try again.")
                except sr.UnknownValueError:
                    speak("Sorry, I couldn't understand the path.")
                except sr.RequestError:
                    speak("I think I'm offline right now.")
    elif "code" in command or "visual studio" in command:
        search_and_launch_program(["code.exe"], "Visual Studio Code")
    elif "intellij" in command:
        search_and_launch_program(["idea64.exe"], "IntelliJ IDEA")
    else:
        speak("I'm not sure how to open that.")


# === Main Wake-Word Loop ===
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
            handle_command(phrase_time=4.0)
            speak("Listening for wake word again.")
except KeyboardInterrupt:
    pass
finally:
    stream.stop_stream()
    stream.close()
    pa.terminate()
    porcupine.delete()
    tts.stop()
