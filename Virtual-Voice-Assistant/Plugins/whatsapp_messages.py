import time
import speech_recognition as sr
import pyttsx3
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.webdriver import WebDriver

def speak(text):
    engine = pyttsx3.init()
    engine.say(text)
    engine.runAndWait()

def takeCommandUser():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Name of User or Group")
        r.pause_threshold = 1
        audio = r.listen(source)
        try:
            query = r.recognize_google(audio, language='en-in')
            print(f'Client to whom the message is to be sent is : {query}\n')
        except Exception as e:
            print(e)
            print("Unable to recognize Client name")
            speak("Unable to recognize Client Name")
            print("Check your Internet Connectivity")
        return query

def takeCommandMessage():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Enter Your Message")
        r.pause_threshold = 1
        audio = r.listen(source)
        try:
            query = r.recognize_google(audio, language='en-in')
            print(f'Message to be sent is : {query}\n')
        except Exception as e:
            print(e)
            print("Unable to recognize your message")
            print("Check your Internet Connectivity")
        return query

def sendWhatsappMessage():
    chromedriver_path = 'C:\\Users\\banka\\Documents\\MAJOR PROJECT\\chromedriver-win64\\chromedriver.exe'

    service = Service(chromedriver_path)
    driver = WebDriver(service=service)

    driver.get('https://web.whatsapp.com/')
    speak("Scan QR code before proceeding")
    time.sleep(25)
    speak("Enter Name of Group or User")
    name = takeCommandUser()
    user = driver.find_element(By.XPATH, f'//span[@title="{name}"]')
    user.click()
    speak("Enter Your Message")
    msg = takeCommandMessage()
    msg_box = driver.find_element(By.CLASS_NAME, '_ak1l')
    msg_box.send_keys(msg)
    button = driver.find_element(By.CLASS_NAME, '_ak1u')
    button.click()
    if button.click:
        speak("Message sent successfully")
    time.sleep(5)
    driver.quit()

# sendWhatsappMessage()
