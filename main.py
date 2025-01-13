import win32com.client
import speech_recognition as sr
import webbrowser
import os
import google.generativeai as genai
import subprocess
import time
from threading import Thread

os.environ["GEMINI_API_KEY"] = "AIzaSyARj_UwL4rSnmEOrP7RFRx3yU54NmMsRgY"
speaker = win32com.client.Dispatch("SAPI.SpVoice")

sleeping = False
reminders = []

def gpt(prompt):
    genai.configure(api_key=os.environ["GEMINI_API_KEY"])

    generation_config = {
        "temperature": 1,
        "top_p": 0.95,
        "top_k": 40,
        "max_output_tokens": 8192,
        "response_mime_type": "text/plain",
    }

    model = genai.GenerativeModel(
        model_name="gemini-2.0-flash-exp",
        generation_config=generation_config,
    )

    chat_session = model.start_chat(
        history=[
            {
                "role": "user",
                "parts": [
                    prompt + "\n",
                ],
            },
        ]
    )

    response = chat_session.send_message(prompt)
    print(response.text)
    speaker.Speak(response.text)

def open_website(query):
    sites = [["youtube", "https://www.youtube.com"],
             ["wikipedia", "https://www.wikipedia.com"],
             ["google", "https://www.google.com"],
             ["instagram", "https://www.instagram.com"],
             ["whatsapp","https://www.whatsapp.com"]]
    for site in sites:
        if f"open {site[0]}" in query.lower():
            speaker.Speak(f"Opening {site[0]}...")
            webbrowser.open(site[1])
            return True
    return False

def open_app(query):
    apps = {
        "notepad": "notepad.exe",
        "calculator": "calc.exe",
        "chrome": r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        "explorer": r"C:\Windows\explorer.exe",
        "telegram": r"C:\Program Files\WindowsApps\TelegramMessengerLLP.TelegramDesktop_5.8.3.0_x64__t4vj0pshhgkwm\Telegram.exe",
        "visual code": r"D:\Program Files\Microsoft VS Code\Code.exe"
    }
    for app, path in apps.items():
        if f"open {app}" in query.lower():
            speaker.Speak(f"Opening {app}...")
            subprocess.Popen(path)
            return True
    return False

def takeCom():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        try:
            print("Listening...")
            audio = r.listen(source)
            print("Recognizing...")
            query = r.recognize_google(audio, language="en-in")
            print(f"User said: {query}")
            return query
        except sr.UnknownValueError:
            speaker.Speak("Sorry, I didn't catch that. Could you repeat?")
            return None
        except sr.RequestError as e:
            print(f"Could not request results; {e}")
            speaker.Speak("There was an error connecting to the speech recognition service.")
            return None

def set_reminder(query):
    try:
        speaker.Speak("What should I remind you about?")
        reminder_query = takeCom()
        if reminder_query:
            speaker.Speak("When should I remind you? Please specify in seconds.")
            reminder_time = takeCom()
            if reminder_time and reminder_time.isdigit():
                reminders.append((reminder_query, int(reminder_time)))
                speaker.Speak(f"Reminder set for {reminder_time} seconds.")
    except Exception as e:
        print(f"Error setting reminder: {e}")

def check_reminders():
    while True:
        current_time = time.time()
        for reminder in reminders[:]:
            reminder_text, reminder_time = reminder
            if reminder_time <= 0:
                speaker.Speak(f"Reminder: {reminder_text}")
                reminders.remove(reminder)
            else:
                reminder_time -= 1
        time.sleep(1)

reminder_thread = Thread(target=check_reminders)
reminder_thread.daemon = True
reminder_thread.start()

speaker.Speak("Hello, I'm CYPHER AI")

while True:
    try:
        if not sleeping:
            query = takeCom()
            if query:
                if "go to sleep" in query.lower():
                    speaker.Speak("Alright, I'm going to sleep. Say 'wake up' to activate me again.")
                    sleeping = True
                elif "set a reminder" in query.lower():
                    set_reminder(query)
                elif open_website(query):
                    pass
                elif open_app(query):
                    pass
                else:
                    gpt(query)
        else:
            query = takeCom()
            if query and "wake up" in query.lower():
                speaker.Speak("I'm awake now. How can I assist you?")
                sleeping = False
    except KeyboardInterrupt:
        speaker.Speak("Goodbyee!")
        print("\nExiting the program.")
        break