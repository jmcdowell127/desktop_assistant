import wolframalpha as wolframalpha
import subprocess
import pyttsx3
import tkinter
import json
import random
import operator
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import winshell
import pyjokes
import feedparser
import smtplib
import ctypes
import time
import requests
import shutil
from twilio.rest import Client
from clint.textui import progress
from ecapture import ecapture as ec
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen
from config import email, pword


# set engine to Pyttsx3 which is used for test to speech and sapi5 which is microsoft speech application
#     platform interface being used for text to speech function
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)
# voice ID 1 is for female, 0 is for male

def speak(aduio):
    engine.say(audio)
    engine.runAndWait()

def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speak('Good Morning Sir !')

    elif hour >= 12 and hour < 18:
        speak('Good Afternoon Sir !')

    else:
        speak('Good Evening Sir !')

    assname = ('Wendy')
    speak('I am your Assistant')
    speak(assname)

def username():
    speak('What should I call you sir?')
    uname = takeCommand()
    speak('Welcome Mister')
    speak(uname)
    columns = shutil.get_terminal_size().columns

    print('#######################'.center(columns))
    print('Welcome Mr.', uname.center(columns))
    print('#######################'.center(columns))

    speak('How can I help you, Sir')

def takeCommand():

    r = sr.Recognizer()

    with sr.Microphone() as source:

        print('Listening...')
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print('Recognizing...')
        query = r.recognize_google(audio, language = 'en-in')
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
            results = wikipedia.summary(query, sentences = 3)
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

            app_id = 'G64WUQ-AT7V4TV6P4'
            client = wolframalpha.Client(app_id)
            indx = query.lower().split().index('calculate')
            query = query.split()[indx + 1:]
            res = client.query(' '.join(query))
            answer = next(res.results).text
            print('The answer is ' + answer)
            speak('The answer is ' + answer)

