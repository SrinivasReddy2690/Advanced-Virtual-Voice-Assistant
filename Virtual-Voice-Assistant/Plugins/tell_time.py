import datetime
import pyttsx3
import time

def speak(text):
    engine = pyttsx3.init()
    engine.setProperty('rate', -5)
    engine.say(text)
    engine.runAndWait()

def tell_time():
    current_time = datetime.datetime.now()
    hour = current_time.hour
    minute = current_time.minute

    # Convert hour to 12-hour format
    if hour >= 12:
        meridiem = "PM"
        hour -= 12
    else:
        meridiem = "AM"
    
    # Special case: 12 AM is midnight, 12 PM is noon
    if hour == 0:
        hour = 12
        
    if 5 <= current_time.hour < 12:
        time_of_day = "morning"
    elif 12 <= current_time.hour < 17:
        time_of_day = "afternoon"
    elif 17 <= current_time.hour < 20:
        time_of_day = "evening"
    else:
        time_of_day = "night"

    # Construct the time string
    str_time = f"{hour}:{minute:02d} {meridiem}"
    print(f"Sir, the time is {str_time}. Good {time_of_day}!")
    speak(f"Sir, the time is {str_time}" + "Good " + time_of_day)
    
# tell_time()