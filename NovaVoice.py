import pvporcupine
import pyaudio
import struct
import subprocess
import os
import pyttsx3
import speech_recognition as sr
import webbrowser

# === Config ===
ACCESS_KEY = "PV_ACCESS_KEY"
WAKE_WORD = "hey-assistant_en_windows.ppn"

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

# === Porcupine & PyAudio Initialization ===
porcupine = pvporcupine.create(
    access_key=ACCESS_KEY,
    keyword_paths=[WAKE_WORD],
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


# === Search for Program ===
def search_and_launch_program(exe_names, app_name):
    for base in [r"C:\Program Files", r"C:\Program Files (x86)"]:
        for root, dirs, files in os.walk(base):
            for f in files:
                if f.lower() in exe_names:
                    path = os.path.join(root, f)
                    subprocess.Popen([path])
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
        command = recognizer.recognize_google(audio_data).lower()
        speak(f"You said: {command}")
    except sr.UnknownValueError:
        speak("Sorry, I couldn't catch that.")
        return
    except sr.RequestError:
        speak("It seems I'm offline.")
        return

    if "google" in command:
        webbrowser.open("https://www.google.com")
        speak("Opening Google.")
    elif "youtube" in command:
        webbrowser.open("https://www.youtube.com")
        speak("Opening YouTube.")
    elif "edge" in command:
        os.system("start microsoft-edge:https://www.google.com")
        speak("Opening Edge.")
    elif "chrome" in command:
        search_and_launch_program(["chrome.exe"], "Chrome")
    elif "firefox" in command:
        search_and_launch_program(["firefox.exe"], "Firefox")
    elif "brave" in command:
        search_and_launch_program(["brave.exe"], "Brave")
    elif "pycharm" in command:
        search_and_launch_program(["pycharm64.exe"], "PyCharm")
    elif "code" in command or "visual studio" in command:
        search_and_launch_program(["code.exe"], "Visual Studio Code")
    elif "intellij" in command:
        search_and_launch_program(["idea64.exe"], "IntelliJ IDEA")
    else:
        speak("I’m not sure how to open that.")


# === Main Loop ===
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
