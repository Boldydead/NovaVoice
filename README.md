# Python Voice Assistant for Desktop Control

This is a Python-based voice assistant designed to run locally on a Windows machine. It listens for a wake word and can then perform various commands such as opening applications, controlling system settings, providing information, and executing custom user-defined actions.

## Features

* **Wake Word Detection:** Uses Picovoice Porcupine for "Hey Google" (or other custom wake word) detection.
* **Voice Command Recognition:** Utilizes Google Speech Recognition via the `speech_recognition` library.
* **Text-to-Speech Output:** Provides voice feedback using `pyttsx3`.
* **Application Launching:**
    * Opens predefined common applications (e.g., Chrome, Notepad, Spotify).
    * Searches for and launches other applications by name.
    * Caches found application paths for faster future launches.
* **Web Browse:** Opens Google, YouTube, and other URLs.
* **System Control (Windows):**
    * Mute/Unmute system volume.
    * Put the computer to sleep.
    * Shut down the computer.
* **Information Retrieval:**
    * Current time and date.
    * Battery status (percentage and charging state).
    * Daily Bible verse (currently a placeholder, see `fetch_daily_text` function).
* **Customizable Commands:** Supports user-defined commands via a `custom_commands.json` file to:
    * Launch executables.
    * Open specific URLs.
    * Run shell commands.
* **Help & Capabilities Listing:** Can explain its own commands and known applications.
* **Scheduler:** Includes a simple scheduler for tasks like a morning greeting.

## Requirements

* Python 3.7+
* Windows Operating System (for some system-specific commands like sleep, mute, shutdown, and `os.startfile`)
* A working microphone
* Internet connection (for speech recognition and some web features)
* Picovoice Porcupine Access Key
* Porcupine wake word file (`.ppn`)

### Python Libraries:
You'll need to install the following Python libraries. It's recommended to use a virtual environment.
* `pvporcupine`
* `pyaudio` (requires PortAudio; may need separate installation steps depending on your system, often tricky on Windows if build tools are missing. Consider installing from a pre-compiled wheel if you encounter issues: [Unofficial Windows Binaries for Python Extension Packages](https://www.lfd.uci.edu/~gohlke/pythonlibs/#pyaudio))
* `pyttsx3`
* `SpeechRecognition`
* `psutil`
* `python-dotenv`
* `setuptools` (often already present, but good to ensure)
* `wheel` (for installing some packages)

## Setup Instructions

1.  **Clone the Repository (if you haven't already):**
    ```bash
    git clone <your-repository-url>
    cd <repository-name>
    ```

2.  **Create a Virtual Environment (Recommended):**
    ```bash
    python -m venv venv
    ```
    Activate it:
    * Windows: `.\venv\Scripts\activate`
    * macOS/Linux: `source venv/bin/activate`

3.  **Install Dependencies:**
    Create a `requirements.txt` file in your project root with the following content:
    ```
    pvporcupine
    pyaudio
    pyttsx3
    SpeechRecognition
    psutil
    python-dotenv
    # Add any other specific versions if needed, e.g., pvporcupine==1.9.5
    ```
    Then install them:
    ```bash
    pip install -r requirements.txt
    ```
    *If `pyaudio` fails, try installing it using a pre-compiled wheel from the link mentioned in the Requirements section.*

## Configuration

1.  **Porcupine Wake Word Engine:**
    * Obtain a free `PORCUPINE_ACCESS_KEY` from [Picovoice Console](https://console.picovoice.ai/).
    * Create a `.env` file in the root of your project directory and add your access key:
        ```env
        PORCUPINE_ACCESS_KEY="YOUR_ACTUAL_ACCESS_KEY_HERE"
        ```
    * Ensure you have the Porcupine wake word file (e.g., `hey google_windows.ppn`). Place it in the root directory or update the `WAKE_WORD_PPN_FILENAME` variable in the Python script if it's named differently or located elsewhere. The script uses "hey google_windows.ppn" by default.

2.  **Custom Commands (Optional):**
    * To define your own voice commands, create a `custom_commands.json` file in the project root.
    * Structure:
        ```json
        {
          "commands": [
            {
              "phrase": "your voice command phrase",
              "action": "launch_executable",
              "exe_name": "application.exe",
              "app_name": "Friendly App Name",
              "response": "Launching Friendly App Name."
            },
            {
              "phrase": "open my favorite website",
              "action": "url",
              "url": "[https://example.com](https://example.com)",
              "response": "Opening your favorite site!"
            },
            {
              "phrase": "run backup script",
              "action": "shell",
              "shell_cmd": "C:\\path\\to\\your\\backup.bat",
              "response": "Running the backup script."
            }
          ]
        }
        ```
    * If this file doesn't exist, the assistant will operate without custom commands.

3.  **Cache Files (Auto-generated):**
    * `exe_cache.json`: Stores paths to executables found by the assistant to speed up subsequent launches. Automatically created/updated.
    * `daily_text.json`: Caches the daily text to avoid re-fetching. Automatically created/updated.

## Usage

1.  **Starting the Assistant:**
    Run the main Python script from your terminal (ensure your virtual environment is activated):
    ```bash
    python your_script_name.py
    ```
    You should hear "Assistant ready. Say 'Hey Google' or your wake word to begin."

2.  **Interacting with the Assistant:**
    * Say the wake word (e.g., "Hey Google").
    * Wait for the prompt (e.g., "What would you like me to do?").
    * Speak your command.

    **Example Commands:**
    * "Hey Google... open Chrome."
    * "Hey Google... what time is it?"
    * "Hey Google... mute the system."
    * "Hey Google... what can you do?"
    * "Hey Google... what apps can you open?"
    * "Hey Google... launch Spotify."
    * "Hey Google... put computer to sleep."
    * (Any custom commands you've defined)

## Customizing Commands

Modify the `custom_commands.json` file as described in the "Configuration" section to add or change custom voice commands. The assistant loads these commands at startup.

## Key Files in the Project

* `your_script_name.py`: The main Python script for the voice assistant.
* `.env`: Stores your `PORCUPINE_ACCESS_KEY` (you need to create this).
* `hey google_windows.ppn` (or your chosen wake word file): The Porcupine wake word model file.
* `custom_commands.json` (optional): For defining custom commands.
* `exe_cache.json` (auto-generated): Caches paths to found executables.
* `daily_text.json` (auto-generated): Caches the daily text.
* `requirements.txt` (you should create this): Lists Python package dependencies.
* `README.md`: This file.

## Basic Troubleshooting

* **"PORCUPINE_ACCESS_KEY not found"**: Ensure your `.env` file is correctly created in the project root and contains your key.
* **"Wake word file ... not found"**: Make sure the `.ppn` file is in the project root and its name matches `WAKE_WORD_PPN_FILENAME` in the script.
* **Microphone not working / "Microphone not found"**: Check your system's microphone settings and ensure PyAudio can access it. PyAudio installation can sometimes be problematic; refer to its documentation or try installing from a wheel.
* **"I'm not sure how to do that yet"**:
    * The command might not be recognized clearly. Try speaking more clearly.
    * The command phrase may not be defined in `COMMAND_DISPATCHER` or your `custom_commands.json`.
    * For application launching, the app might not be in the predefined list or easily findable in standard search paths.
* **TTS or Speech Recognition Issues**: Ensure `pyttsx3` and `SpeechRecognition` are correctly installed and that you have an active internet connection for Google Speech Recognition.

---

*This README provides a starting point. Feel free to add more details, a "Future Enhancements" section, or a "License" section (e.g., MIT License is common for open-source projects).*
