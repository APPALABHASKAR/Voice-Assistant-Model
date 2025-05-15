import pyttsx3
import speech_recognition as sr
import pywhatkit
import datetime
import wikipedia
import os
import webbrowser
import psutil
import keyboard
import asyncio
import requests



engine = pyttsx3.init()
engine.setProperty('rate', 200)
engine.setProperty('voice', engine.getProperty('voices')[1].id)

installed_apps = {
    "notepad": "notepad.exe",
    "calculator": "calc.exe",
    "vlc": "C:/Program Files/VideoLAN/VLC/vlc.exe",
    "chrome": "C:/Program Files/Google/Chrome/Application/chrome.exe",
    "word": "C:/Program Files/Microsoft Office/root/Office16/WINWORD.EXE",
    "excel": "C:/Program Files/Microsoft Office/root/Office16/EXCEL.EXE",
}

chrome_available = os.path.exists(installed_apps.get("chrome", ""))

if chrome_available:
    webbrowser.register('chrome', None, webbrowser.BackgroundBrowser(installed_apps["chrome"]))

video_paused = False
previous_command = ""

def talk(text):
    engine.say(text)
    engine.runAndWait()

async def take_command():
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        recognizer.adjust_for_ambient_noise(source, duration=0.5)
        try:
            voice = recognizer.listen(source, timeout=5, phrase_time_limit=7)
            return recognizer.recognize_google(voice).lower()
        except sr.UnknownValueError:
            return ""
        except sr.WaitTimeoutError:
            return ""

def get_location():
    try:
        response = requests.get('https://ipinfo.io')
        data = response.json()
        location = data.get('city', 'Unknown City')
        country = data.get('country', 'Unknown Country')
        return f"You are in {location}, {country}."
    except requests.exceptions.RequestException:
        return "Unable to fetch location. Please check your internet connection."

async def open_app(app_name):
    app_name = app_name.lower()
    if app_name in installed_apps:
        try:
            os.startfile(installed_apps[app_name])
            talk(f"Opening {app_name}")
        except Exception as e:
            talk(f"Unable to open {app_name}. Error: {str(e)}")
    elif app_name == "chrome" and chrome_available:
        try:
            webbrowser.get('chrome').open("https://www.google.com")
            talk("Opening Chrome browser")
        except Exception as e:
            talk(f"Unable to open Chrome. Error: {str(e)}")
    else:
        talk(f"I couldn't find {app_name} installed on your system.")

async def close_app(app_name):
    app_name = app_name.lower()
    for process in psutil.process_iter(['name']):
        if app_name in process.info['name'].lower():
            process.terminate()
            talk(f"Closed {app_name}")
            return
    talk(f"I couldn't find {app_name} running.")

async def close_chrome_tab():
    try:
        keyboard.press_and_release('ctrl+w')
        talk("Closed the Chrome tab")
    except Exception as e:
        talk(f"Failed to close the Chrome tab. Error: {str(e)}")

async def stop_or_resume_video(command):
    keyboard.press_and_release('space')
    if "resume" in command:
        talk("Resumed the video.")
    else:
        talk("Paused the video.")

async def perform_task(command):
    global previous_command
    if "play" in command:
        song = command.replace("play", "").strip()
        talk(f"Playing {song} on YouTube")
        pywhatkit.playonyt(song)
    elif "pause" in command or "resume" in command:
        await stop_or_resume_video(command)
    elif "close" in command:
        app_name = command.replace("close", "").strip()
        await close_app(app_name)
    elif "home" in command:
        keyboard.press_and_release('win+d')
        talk("Going to Home screen")
    elif "time" in command:
        current_time = datetime.datetime.now().strftime('%I:%M %p')
        talk(f"The current time is {current_time}")
    elif "present day" in command:
        current_date = datetime.datetime.now().strftime('%A, %B %d, %Y')
        talk(f"The current date is {current_date}")
    elif "open" in command:
        app_name = command.replace("open", "").strip()
        await open_app(app_name)
    elif "close browser tab" in command or "close tab" in command:
        await close_chrome_tab()
    elif "search" in command:
        query = command.replace("search", "").strip()
        search_url = f"https://www.google.com/search?q={query.replace(' ', '+')}"
        if chrome_available:
            webbrowser.get('chrome').open(search_url)
        else:
            webbrowser.open(search_url)
        talk(f"Searching for {query}")
    elif 'tell about' in command:
        topic = command.replace('tell about', '').strip()
        info = wikipedia.summary(topic, sentences=2)
        talk(info)
    elif "give snipping tool" in command:
        keyboard.press_and_release('win+shift+s')
        talk("Snipping tool opened.")
    elif "take a screenshot" in command:
        try:
            keyboard.press_and_release('win+print screen')
            talk("Screenshot taken and saved.")
        except Exception as e:
            talk(f"Unable to take screenshot. Error: {e}")
    elif "previous desktop" in command:
        keyboard.press_and_release('ctrl+win+left')
        talk("Switched to the previous desktop.")
    elif "next desktop" in command:
        keyboard.press_and_release('ctrl+win+right')
        talk("Switched to the next desktop.")
    elif "new tab" in command:
        keyboard.press_and_release('ctrl+t')
        talk("New tab opened.")
    elif "previous tab" in command:
        keyboard.press_and_release('ctrl+shift+tab')
        talk("Switched to the previous tab.")
    elif "next tab" in command:
        keyboard.press_and_release('ctrl+tab')
        talk("Switched to the next tab.")
    elif "close all tabs" in command:
        keyboard.press_and_release('ctrl+shift+w')
        talk("All tabs closed.")
    elif "volume down" in command:
        keyboard.press_and_release('volume down')
        talk("Volume decreased.")
    elif "volume up" in command:
        keyboard.press_and_release('volume up')
        talk("Volume increased.")
    elif "exit" in command or "see you later" in command:
        talk("Goodbye sir! Have a great day.")
        exit()
    elif "recent apps" in command:
        keyboard.press_and_release('win+tab')
        talk("Recent apps opened.")
    elif "control centre" in command:
        keyboard.press_and_release('win+a')
        talk("Control center opened.")
    elif "back" in command:
        keyboard.press_and_release('esc')
        talk("Navigated back.")
    elif "present location" in command:
        location = get_location()
        talk(location)
    elif "file manager" in command:
        keyboard.press_and_release('win+e')
        talk("File manager opened.")
    elif "clipboard" in command:
        keyboard.press_and_release('win+v')
        talk("Clipboard opened.")
    elif "do it again" in command:
        if previous_command:
            talk("Repeating the last command.")
            await perform_task(previous_command)
        else:
            talk("No previous command to repeat.")
    else:
        talk("I didn't understand. Can you repeat?")
    previous_command = command

async def run_assistant():
    talk("Hi sir! I am Mia, Mr.Reddy's personal assistant. Ready to assist you.")
    while True:
        command = await take_command()
        if "do it again" in command:
            talk("Alright doing again.")
            command = previous_command
        if command:
            await perform_task(command)
        else:
            talk("Can u Say again?")

asyncio.run(run_assistant())
