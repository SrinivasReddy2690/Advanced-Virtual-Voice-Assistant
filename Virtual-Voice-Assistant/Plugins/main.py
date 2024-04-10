try:
    import pyttsx3
    from ttkthemes import ThemedTk
    from time import sleep
    import datetime
    import speech_recognition as sr
    import wikipedia
    import os
    import webbrowser
    import pyjokes
    import pywhatkit as kit
    import time
    from plyer import notification
    import tkinter as tk
    from tkinter import ttk
    from tkinter import LEFT, BOTH, SUNKEN
    from PIL import Image, ImageTk
    from threading import Thread
    # importing prebuilt modules
    import os
    import winshell
    import logging
    import pyttsx3
    logging.disable(logging.WARNING)
    os.environ['TF_CPP_MIN_LOG_LEVEL'] = '3'
    from keras_preprocessing.sequence import pad_sequences
    import numpy as np
    import wmi
    from keras.models import load_model
    from pickle import load
    import speech_recognition as sr
    import sys
    from pydub import AudioSegment
    from pydub.playback import play
    sys.path.insert(0, os.path.expanduser('~')+"/Virtual-Voice-Assistant") # adding voice assistant directory to system path
    # importing modules made for assistant
    from database import *
    from image_generation import generate_image
    from gmail import *
    from wifi import *
    from API_functionalities import *
    from system_operations import *
    from whatsapp_messages import *
    from browsing_functionalities import *
    from brightness import *
    from tell_time import *
    from set_volume import *
    from change_bg import *
    import threading
    import queue

except Exception as e:
    print(e)
    print("ERROR OCCURRED WHILE IMPORTING THE MODULES")
    exit(0)

recognizer = sr.Recognizer()

engine = pyttsx3.init()         
engine.setProperty('rate', 185)

sys_ops = SystemTasks()
tab_ops = TabOpt()
win_ops = WindowOpt()

# load trained model
model = load_model('..\\Data\\chat_model')

# load tokenizer object
with open('..\\Data\\tokenizer.pickle', 'rb') as handle:
    tokenizer = load(handle)

# load label encoder object
with open('..\\Data\\label_encoder.pickle', 'rb') as enc:
    lbl_encoder = load(enc)

stop_flag = False
BG_COLOR = "#D2C6E2"
BUTTON_COLOR = "#F9F4F2"
BUTTON_FONT = ("Arial", 14, "bold")
BUTTON_FOREGROUND = "black"
# HEADING_FONT = ("Arial", 24, "bold")
# INSTRUCTION_FONT = ("Helvetica", 14)

# Global variable to store console messages
console_messages = ["Click to start listen"]
message_queue = queue.Queue()

def on_button_click():
    global stop_flag
    start_assistant()

def update_console(root, label1):
    global console_messages
    
    # Get the latest message
    latest_message = console_messages[-1]

    # Apply different styles based on the message content
    if "ASSISTANT" in latest_message:
        label1.config(text=latest_message, fg="blue", font=("Arial Rounded MT Bold", 12, "bold"), bg="#f0f8ff", relief="solid", padx=10, pady=5)
    elif "LISTENING" in latest_message:
        label1.config(text=latest_message, fg="green", font=("Arial Rounded MT Bold", 12, "bold"), bg="#e6ffe6", relief="solid", padx=10, pady=5)
    elif "COULD NOT UNDERSTAND AUDIO" in latest_message:
        label1.config(text=latest_message, fg="red", font=("Arial Rounded MT Bold", 12, "bold"), bg="#ffe6e6", relief="solid", padx=10, pady=5)
    elif "SAYS" in latest_message:
        label1.config(text=latest_message, fg="black", font=("Arial Rounded MT Bold", 12), bg="#ffffe6", relief="solid", wraplength=1200, padx=10, pady=5)
    elif "Click to start listen" in latest_message:  
        label1.config(text=latest_message, fg="black", font=("Arial Rounded MT Bold", 12, "bold"), bg="#f0f8ff", relief="solid", padx=10, pady=5)
        
    label1.pack(side="top", pady=30)

    # Call update_console again after 1000 milliseconds
    root.after(1000, update_console, root, label1)


def speak(text):
    print("ASSISTANT -> " + text)
    console_messages.append("ASSISTANT -> " + text)

    update_console(root,label1)

    try:
        engine.say(text)
        engine.runAndWait()
    except KeyboardInterrupt or RuntimeError:
        return

def chat(text):
    max_len = 20
    while True:
        result = model.predict(pad_sequences(tokenizer.texts_to_sequences([text]),
                                                                          truncating='post', maxlen=max_len), verbose=False)
        intent = lbl_encoder.inverse_transform([np.argmax(result)])[0]
        return intent
    
def record():
    recognizer = sr.Recognizer()
    recognizer.energy_threshold = 2000  # Adjust energy threshold for ambient noise
    recognizer.dynamic_energy_adjustment_ratio = 1.5  # Adjust dynamic energy adjustment ratio
    recognizer.pause_threshold = 0.5  # Adjust pause threshold
    recognizer.operation_timeout = 4  # Adjust operation timeout
    with sr.Microphone() as mic:
        recognizer.adjust_for_ambient_noise(mic, duration=1)  # Adjust for ambient noise
        print("Listening...")
        console_messages.append("LISTENING")
        update_console(root,label1)
        audio = recognizer.listen(mic, timeout=None)  # Listen indefinitely until speech is detected
    try:
        text = recognizer.recognize_google(audio, language='en-US').lower()
        print("USER -> " + text)
        console_messages.append(f"SAYS {text}")
        root.after(0, lambda: update_console(root, label1))  # Call update_console in the main thread

        # Check if the user said "exit" and close the GUI if true
        if "exit" in text:
            root.destroy()

        return text
    except sr.UnknownValueError:
        print("Could not understand audio.")
        console_messages.append("COULD NOT UNDERSTAND AUDIO")
        root.after(0, lambda: update_console(root, label1))  # Call update_console in the main thread
        return None
    except sr.RequestError as e:
        print("Error occurred during recognition: {0}".format(e))
        return None

def start_assistant():
    listen_thread = threading.Thread(target=listen_audio)
    listen_thread.start()

def listen_audio():
    while True:
        response = record()
        if response is None:
            continue
        else:
            perform_task(response)
            message_queue.put(response)  # Put the response into the queue
        sleep(3)

def perform_task(query):
    # add_data(query)
    intent = chat(query)
    done = False
    if ("google" in query and "search" in query) or ("google" in query and "how to" in query) or "google" in query:
        googleSearch(query)
        return
    elif ("youtube" in query and "search" in query) or "play" in query or ("how to" in query and "youtube" in query):
        youtube(query)
        return
    elif "distance" in query or "map" in query:
        get_map(query)
        return
    elif 'time now' in query or 'what is the time now' in query or 'tell me the time' in query:
        tim = tell_time()
        if time:
            done = True
    elif intent == "joke" and "joke" in query:
        joke = get_joke()
        if joke:
            speak(joke)
            done = True
    elif intent == "news" and "news" in query:
        news = get_news()
        if news:
            speak(news)
            done = True
    elif "what is" in query or "who is" in query or "who" in query:
        res = get_general_response(query)
        if  res:
            speak(res)
            done = True
    elif "change brightness to " in query:
        brightness = brightness_level(query)
        if brightness:
            done = True
    elif "change volume to" in query:
        volume_level = float(re.search(r"\d+", query).group()) / 100  
        set_volume(volume_level)  
        speak(f"Volume changed to {int(volume_level*100)} percent") 
        done = True
    elif "send whatsapp message" in query or "whatsapp message" in query:
        whatsapp = sendWhatsappMessage()
        if whatsapp:
            done = True
    elif "change background" in query:
        change_bg(r"C:\Users\banka\Documents\MAJOR PROJECT\Virtual-Voice-Assistant\Plugins\background_images")
        speak("System background changed successfully")
        done  =True
    elif 'empty recycle bin' in query:
        winshell.recycle_bin().empty(confirm = False, show_progress = False, sound = True)
        speak("Recycle Bin Recycled")
    elif 'shutdown system' in query:
        speak("Hold On a Sec! Your system is on its way to shut down")
        subprocess.call('shutdown /p /f')
    elif 'lock window' in query or "system ko lock Karen" in query:
        speak("locking the device")
        ctypes.windll.user32.LockWorkStation()
    elif "hibernate" in query or "sleep" in query:
        speak("Hibernating")
        subprocess.call("shutdown /i /h")
    elif "restart" in query:
        speak("System Restarting")
        subprocess.call(["shutdown", "/r"])
    elif "where is" in query:
        query=query.replace("where is","")
        location = query
        speak("User asked to Locate")
        speak(location)
        webbrowser.open("https://www.google.nl/maps/place/" + location + "")
    elif intent == "ip" and "ip" in query:
        ip = get_ip()
        if ip:
            speak(ip)
            done = True
    elif intent == "movies" and "movies" in query:
        speak("Some of the latest popular movies are as follows :")
        get_popular_movies()
        done = True
    elif intent == "tv_series" and "tv series" in query:
        speak("Some of the latest popular tv series are as follows :")
        get_popular_tvseries()
        done = True
    elif intent == "weather" and "weather" in query:
        city = re.search(r"(in|of|for) ([a-zA-Z]*)", query)
        if city:
            city = city[2]
            weather = get_weather(city)
            speak(weather)
        else:
            weather = get_weather()
            speak(weather)
        done = True
    elif intent == "internet_speedtest" and "internet" in query:
        speak("Getting your internet speed, this may take some time")
        speed = get_speedtest()
        if speed:
            speak(speed)
            done = True
    elif intent == "system_stats" and "stats" in query:
        stats = system_stats()
        speak(stats)
        done = True
    elif intent == "image_generation" and "image" in query:
        speak("what kind of image you want to generate?")
        text = record()
        speak("Generating image please wait..")
        generate_image(text)
        done = True
    elif intent == "system_info" and ("info" in query or "specs" in query or "information" in query):
        info = systemInfo()
        speak(info)
        done = True
    elif intent == "email" and "email" in query:
        receiver_id = None
        subject = None
        body = None

        # Function to handle sending email
        def send_email_gui():
            nonlocal receiver_id, subject, body
            receiver_id = receiver_entry.get()
            subject = subject_entry.get()
            body = body_text.get("1.0", tk.END)
            if check_email(receiver_id):
                if send_email(receiver_id, subject, body):
                    speak('Email sent successfully')
                else:
                    speak("Error occurred while sending email")
            else:
                speak("Invalid email address")
        # GUI to input email details
        root = tk.Tk()
        root.title("Email")

        receiver_label = tk.Label(root, text="Receiver Email:")
        receiver_label.pack()
        receiver_entry = tk.Entry(root)
        receiver_entry.pack()

        subject_label = tk.Label(root, text="Subject:")
        subject_label.pack()
        subject_entry = tk.Entry(root)
        subject_entry.pack()

        body_label = tk.Label(root, text="Body:")
        body_label.pack()
        body_text = tk.Text(root, height=10, width=30)
        body_text.pack()

        send_button = tk.Button(root, text="Send Email", command=send_email_gui)
        send_button.pack()

        root.mainloop()

    elif intent == "select_text" and "select" in query:
        sys_ops.select()
        done = True
    elif intent == "copy_text" and "copy" in query:
        sys_ops.copy()
        done = True
    elif intent == "paste_text" and "paste" in query:
        sys_ops.paste()
        done = True
    elif intent == "delete_text" and "delete" in query:
        sys_ops.delete()
        done = True
    elif intent == "pause" and "pause" in query:
        speak("video is paused")
        sys_ops.pause()
        done = True
    elif intent == "resume" and "resume" in query:
        speak("Video is resumed")
        sys_ops.resume()
        done = True
    elif intent == "new_file" and "new file" in query:
        sys_ops.new_file()
        done = True
    elif intent == "switch_tab" and "switch" in query and "tab" in query:
        tab_ops.switchTab()
        done = True
    elif intent == "close_tab" and "close" in query and "tab" in query:
        tab_ops.closeTab()
        done = True
    elif intent == "new_tab" and "new" in query and "tab" in query:
        tab_ops.newTab()
        done = True
    elif intent == "close_window" and "close" in query:
        win_ops.closeWindow()
        done = True
    elif intent == "switch_window" and "switch" in query:
        win_ops.switchWindow()
        done = True
    elif intent == "minimize_window" and "minimize" in query:
        win_ops.minimizeWindow()
        done = True
    elif intent == "maximize_window" and "maximize" in query:
        win_ops.maximizeWindow()
        done = True
    elif intent == "screenshot" and "screenshot" in query:
        screenshot()
        if screenshot:
            speak("Screenshot saved successfully")
        done = True
    elif intent == "wikipedia" and ("tell" in query or "about" in query):
        description = tell_me_about(query)
        if description:
            speak(description)
        else:
            googleSearch(query)
        done = True
    elif intent == "open_website":
        completed = open_specified_website(query)
        if completed:
            done = True
    elif intent == "open_app":
        completed = open_app(query)
        if completed:
            done = True
    elif intent == "note" and "note" in query:
        speak("what would you like to take down?")
        note = record()
        take_note(note)
        done = True
    # elif intent == "get_data" and "history" in query:
    #     get_data()
    #     done = True
    elif intent == "exit" and ("exit" in query or "terminate" in query or "quit" in query):
        sys.exit(0)


if __name__ == "__main__":
    try:
        print("Started")
    except Exception as e:
        print(f"An error occurred: {e}")


root = tk.Tk()
root.title("Voice Assistant")
root.geometry("500x700")
root.configure(bg=BG_COLOR)


# Create a frame for status bar and console output
status_frame = ttk.Frame(root)
status_frame.pack(pady=10)

# Load and set the background image
background_image = Image.open(r"C:\Users\banka\Documents\MAJOR PROJECT\Virtual-Voice-Assistant\Plugins\background_images\1.jpeg")
background_photo = ImageTk.PhotoImage(background_image)
background_label = ttk.Label(root, image=background_photo)
background_label.place(x=0, y=0, relwidth=1, relheight=1)

f1 = ttk.Frame(root)
f1.pack(pady=100)

image2 = Image.open(r"C:\Users\banka\Documents\MAJOR PROJECT\Virtual-Voice-Assistant\Plugins\background_images\2.jpg")
resized_image = image2.resize((120, 120))
p2 = ImageTk.PhotoImage(resized_image)
l2 = ttk.Label(f1, image=p2, relief="solid")
l2.pack(side="top", fill="both")

# heading_label = ttk.Label(root, text="Voice Assistant", font=HEADING_FONT, background=BG_COLOR)
# heading_label.pack(pady=20)

HEADING_FONT = ("Arial Rounded MT Bold", 24, "bold") 

# Create the heading label with styles
heading_label = ttk.Label(root, text="Virtual Voice Assistant", font=HEADING_FONT)
heading_label.config(foreground="black", relief="raised", borderwidth=2, padding=5)
heading_label.pack(pady=(20, 40))


INSTRUCTION_FONT = ("Arial Rounded MT Bold", 13, "bold") 

instruction_label = ttk.Label(root, text="Click the button below to start the Voice Assistant.",
                            font=INSTRUCTION_FONT, relief="raised", borderwidth=2, padding=3)
instruction_label.pack(pady=(10,10))

button_frame = ttk.Frame(root)
button_frame.pack()

BUTTON_FONT = ("Arial Rounded MT Bold", 14, "bold")

# Configure the style for the buttons
style = ttk.Style()
style.configure("TButton", font=BUTTON_FONT, borderwidth=2, relief="raised", padding=5)
style.map("TButton", foreground=[("pressed", "red"), ("active", "blue")])

start_button = ttk.Button(button_frame, text="Start Voice Assistant", command=on_button_click,
                          style="TButton", width=20)
start_button.pack(side="left", padx=10, pady=10)

stop_button = ttk.Button(button_frame, text="Stop Assistant", command=exit,
                         style="TButton", width=20)
stop_button.pack(side="left", padx=10, pady=10)


style = ttk.Style(root)
style.configure("TButton", font=BUTTON_FONT, background=BUTTON_COLOR, foreground=BUTTON_FOREGROUND)
style.configure("CButton", font=BUTTON_FONT, background="#FFF2F2", foreground=BUTTON_FOREGROUND)

# Create a text widget for console messages
# console_text = tk.Text(root, height=10, width=60, bg="#FFFFFF", fg="#000000", font=("Arial", 12))
# console_text.pack(pady=10)

# Run the GUI main loop
# update_status_message()
label1 = tk.Label(root,text=console_messages[-1])
label1.pack()
root.after(1, update_console, root, label1)         
root.mainloop()



