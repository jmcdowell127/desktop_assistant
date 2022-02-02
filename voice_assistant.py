import wolframalpha as wolframalpha
import subprocess
import pyttsx3
# import tkinter
import json
# import random
# import operator
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import winshell
import pyjokes
# import feedparser
import smtplib
import ctypes
import time
import requests
import shutil
from clint.textui import progress
# from ecapture import ecapture as ec
# from bs4 import BeautifulSoup
# import win32com.client as wincl
from urllib.request import urlopen
from config import email, pword, news_api, weather_api, wolfram_api

# set engine to Pyttsx3 which is used for text to speech and sapi5 which is microsoft speech application
#     platform interface being used for text to speech function
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)


# voice ID 1 is for female, 0 is for male

def speak(audio):
    engine.say(audio)
    engine.runAndWait()


def wishMe():
    hour = int(datetime.datetime.now().hour)
    if 0 <= hour < 12:
        speak('Good Morning Sir !')

    elif 12 <= hour < 18:
        speak('Good Afternoon Sir !')

    else:
        speak('Good Evening Sir !')

    assname = ('Jarvis')
    speak('I am your Assistant')
    speak(assname)


def username():
    speak('What should I call you sir?')
    uname = takeCommand()
    speak('Welcome Mister')
    speak(uname)
    columns = shutil.get_terminal_size().columns

    print('#####'.center(columns))
    print('Welcome Mr.', uname.center(columns))
    print('#####'.center(columns))

    speak('How can I help you, Sir')


def takeCommand():
    r = sr.Recognizer()

    with sr.Microphone() as source:

        print('Listening...')
        r.pause_threshold = 0.8
        r.energy_threshold = 4000
        audio = r.listen(source)

    try:
        print('Recognizing...')
        query = r.recognize_google(audio, language='en-in')
        print(f'User said: {query}\n')

    except Exception as e:
        print(e)
        print('Unable to recognize your voice.')
        return 'None'

    return query


def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()

    # enable low security in gmail
    server.login(email, pword)
    server.sendmail(email, to, content)
    server.close()


# main function starts here. will now call all above functions in main function
if __name__ == '__main__':
    clear = lambda: os.system('cls')

    # this function will clean any command before execution of this python file
    clear()
    wishMe()
    username()

    while True:

        query = takeCommand().lower()

        # all the commands said by me will be stored here in 'query' and will be converted to lower
        # case for easily recognition of command
        if 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace('wikipedia', '')
            results = wikipedia.summary(query, sentences=3)
            speak('According to Wikipedia')
            print(results)
            speak(results)

        elif 'open youtube' in query:
            speak('Here you go to Youtube\n')
            webbrowser.open('youtube.com')

        elif 'open google' in query:
            speak('Here you go to Google\n')
            webbrowser.open('google.com')

        elif 'open stackoverflow' in query:
            speak('Here you go to Stack Over Flow.Happy coding')
            webbrowser.open('stackoverflow.com')

        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime('% H:% M: %S')
            speak(f'John, the time is {strTime}')

        elif 'email to john' in query:
            try:
                speak('What should I say?')
                content = takeCommand()
                to = 'Receiver email address'
                sendEmail(to, content)
                speak('Email has been sent !')
            except Exception as e:
                print(e)
                speak('I am not able to send this email')

        elif 'send a mail' in query:
            try:
                speak('What should I say?')
                content = takeCommand()
                to = 'Receiver email address'
                sendEmail(to, content)
                speak('Email has been sent !')
            except Exception as e:
                print(e)
                speak('I am not able to sent this email')

        elif 'how are you' in query:
            speak('I am fine, Thank you')
            speak('How are you, Sir')

        elif 'fine' in query or 'good' in query:
            speak("It's good to know that you're fine")

        elif 'change my name to' in query:
            query = query.replace('change my name to', '')
            assname = query

        elif 'change name' in query:
            speak('What would you like to call me, Sir ')
            assname = takeCommand()
            speak('I like the new name')

        elif "what's your name" in query or "What is your name" in query:
            speak('My friends call me')
            speak(assname)
            print('My friends call me', assname)

        elif 'exit' in query:
            speak('Good bye sir')
            exit()

        elif 'who made you' in query or 'who created you' in query:
            speak('I have been created by John.')

        elif 'joke' in query:
            speak(pyjokes.get_joke())

        elif 'calculate' in query:

            app_id = wolfram_api
            client = wolframalpha.Client(app_id)
            indx = query.lower().split().index('calculate')
            query = query.split()[indx + 1:]
            res = client.query(' '.join(query))
            answer = next(res.results).text
            print('The answer is ' + answer)
            speak('The answer is ' + answer)

        elif 'search' in query or 'play' in query:

            query = query.replace('search', '')
            query = query.replace('play', '')
            webbrowser.open(query)

        elif 'who I am' in query:
            speak('If you talk then you could be human')

        elif 'go to Discord' in query:
            speak('opening Discord..')
            discord = r'C:\Users\John\AppData\Roaming\Microsoft\Start Menu\Programs\Discord Inc'
            os.startfile(discord)

        elif 'news' in query:

            try:
                jsonObj = urlopen(
                    '''https://newsapi.org/v2/top-headlines?country=us&category=business&apiKey=73b1cc55a1984c89b6730d03b428bcd4''')
                data = json.load(jsonObj)
                i = 1

                speak("here are some of today's business stories")
                print('''=============BUSINESS NEWS==============''' + '\n')

                for item in data['articles']:
                    print(str(i) + '. ' + item['title'] + '\n')
                    print(item['description'] + '\n')
                    speak(str(i) + '. ' + item['title'] + '\n')
                    i += 1
            except Exception as e:
                print(str(e))

        elif 'lock window' in query:
            speak('locking the device')
            ctypes.windll.user32.LockWorkStation()

        elif 'shutdown system' in query:
            speak('Hold on a sec ! Your system is on its way to shutting down')
            subprocess.call('shutdown / p /f')

        elif 'empty recycle bin' in query:
            winshell.recycle_bin().empty(confirm=False, show_progress=False, sound=True)
            speak('Trash has been thrown out')

        elif 'change background' in query:
            ctypes.windll.user32.SystemParametersInfoW(20,
                                                       0,
                                                       "Location of wallpaper",
                                                       0)
            speak("Background changed successfully")

        elif "don't listen" in query or "stop listening" in query:
            speak("how long would you like for me to zone out")
            a = int(takeCommand())
            time.sleep(a)
            print(a)

        elif "where is" in query:
            query = query.replace("where is", "")
            location = query
            speak("User asked to Locate")
            speak(location)
            webbrowser.open("https://www.google.nl / maps / place/" + location + "")

        # elif 'camera' in query or 'take a photo' in query:
        #     ec.capture(0, 'Jarvis Camera ', 'img.jpg')

        elif 'restart' in query:
            subprocess.call(["shutdown", "/r"])

        elif 'hibernate' in query or 'sleep' in query:
            speak('Hibernating')
            subprocess.call("shutdown / h")

        elif "log off" in query or "sign out" in query:
            speak("Make sure all applications are closed before sign-out")
            time.sleep(5)
            subprocess.call(["shutdown", "/l"])

        elif "update assistant" in query:
            speak("After downloading file please replace this file with the downloaded one")
            url = '# url after uploading file'
            r = requests.get(url, stream=True)

            with open("Voice.py", "wb") as Pypdf:
                total_length = int(r.headers.get('content-length'))
                for ch in progress.bar(r.iter_content(chunk_size=2391975),
                                       expected_size=(total_length/1024) + 1):
                    if ch:
                        Pypdf.write(ch)

        elif "jarvis" in query:
            wishMe()
            speak("Jarvis 1 point o in your service Mister")
            speak(assname)

        elif 'weather' in query:
            api_key = weather_api
            base_url = "http://api.openweathermap.org / data / 2.5 / weather?"
            speak(" City name ")
            print("City name : ")
            city_name = takeCommand()
            complete_url = base_url + "appid =" + api_key + "&q =" + city_name
            response = requests.get(complete_url)
            x = response.json()

            if x["cod"] != "404":
                y = x["main"]
                current_temperature = y["temp"]
                current_pressure = y["pressure"]
                current_humidity = y["humidity"]
                z = x["weather"]
                weather_description = z[0]["description"]
                print(" Temperature (in kelvin unit) = " +str(current_temperature)+"\n atmospheric pressure (in hPa unit) ="+str(current_pressure) +"\n humidity (in percentage) = " +str(current_humidity) +"\n description = " +str(weather_description))

            else:
                speak(" City Not Found ")

        elif "wikipedia" in query:
            webbrowser.open("wikipedia.com")

        elif "Good Morning" in query:
            speak("A warm" +query)
            speak("How are you Mister")
            speak(assname)

        elif "what is" in query or "who is" in query:
            client = wolframalpha.Client(app_id)
            res = client.query(query)

            try:
                print(next(res.results).text)
                speak(next(res.results).text)
            except StopIteration:
                print("No results")


