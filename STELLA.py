# Importing the modules.
import subprocess
import wolframalpha
import pyttsx3
import json
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
from os import startfile
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
import win32com.client as wincl
from urllib.request import urlopen
import pywhatkit as pwt
import pyautogui


# Initializing the engine for text to speech using pyttsx3.
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)
chrome_path = "C:/Program Files/Google/Chrome/Application/chrome.exe %s"

def speak(audio):
    engine.say(audio)
    engine.runAndWait()


def wishMe():
    global assname
    assname = ("Stella")
    greet = ("I am your Assistant"+assname)
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speak("Good Morning !"+greet)
    elif hour >= 12 and hour < 18:
        speak("Good Afternoon !"+greet)

    else:
        speak("Good Evening !"+greet)







def takeCommand():

    r = sr.Recognizer()

    with sr.Microphone() as source:

        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        
        print(f"User said: {query}\n")

    except Exception as e:
        print(e)
        print("Voice not clear...Try saying that again")
        return "None"

    return query


def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()

    # Enable low security in gmail
    server.login('your email id', 'your email password')
    server.sendmail('your email id', to, content)
    server.close()

def open_chrome():
    url="https://www.google.co.in/"
    webbrowser.get(chrome_path).open(url)

if __name__ == '__main__':
    def clear(): return os.system('cls')

    # This Function will clean any
    # command before execution of this python file
    clear()
    wishMe()
    speak("How can i Help you.")

    while True:

        query = takeCommand().lower()

        # All the commands said by user will be
        # stored here in 'query' and will be
        # converted to lower case for easily
        # recognition of command
        if 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=3)
            speak("According to Wikipedia")
            print(results)
            speak(results)

        elif 'open youtube' in query:
            speak("Here you go to Youtube\n")
            webbrowser.open("youtube.com")

        elif 'open google' in query:
            speak("Here you go to Google\n")
            webbrowser.open("google.com")

        elif 'open stackoverflow' in query:
            speak("Here you go to Stack Over flow.Happy coding")
            webbrowser.open("stackoverflow.com")

        elif 'play music' in query or "play song" in query:
            speak("which song do you want to listen")
            song_name = takeCommand()
            os.system("spotify")
            time.sleep(5)
            pyautogui.hotkey('ctrl','l')
            pyautogui.write(song_name, interval = 0.2)
            for key in ['enter','pagedown','tab','enter','enter']:
                time.sleep(3)
                pyautogui.press(key)
            speak("enjoy sire")

        elif 'stop music' in query or 'pause song' in query or 'pause music' in query or 'stop song' in query:
            os.system("spotify")
            pyautogui.press('enter')

        elif 'resume song' in query or 'start song' in query or 'resume music' in query or 'start music' in query:
            os.system("spotify")
            pyautogui.press('enter')

        elif 'favourite music' in query or "favourite songs" in query:
            os.system("spotify")
            time.sleep(3)
            pyautogui.hotkey('ctrl','l')
            pyautogui.write('liked songs', interval = 0.2)
            for key in ['enter','pagedown','tab','enter','tab','tab','enter']:
                time.sleep(2)
                pyautogui.press(key)
            speak("your favourite music coming right up sire")


        elif 'text' in query or 'message' in query or 'whatsapp' in query:
            os.system("WhatsApp")
            time.sleep(3)
            speak("to whom you want to text")
            reciepent= takeCommand()
            pyautogui.write(reciepent,interval=0.2)
            for key in ['pagedown','enter','enter']:
                time.sleep(2)
                pyautogui.press(key)
            speak("what do you want your message to be")
            text = takeCommand()
            pyautogui.write(text,interval=0.2)
            speak("message sent successfully")

        
        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"The time is {strTime}")

        elif 'open opera' in query:
            codePath = r"C:\\Users\\aryan\\AppData\\Local\\Programs\\Opera\\launcher.exe"
            os.startfile(codePath)

        

        elif 'send a mail' in query:
            try:
                speak("What should I say?")
                content = takeCommand()
                speak("whome should i send")
                to = input()
                sendEmail(to, content)
                speak("Email has been sent !")
            except Exception as e:
                print(e)
                speak("I am not able to send this email")

        elif 'how are you' in query:
            speak("I am fine, Thank you")
            speak("How are you")

        elif 'fine' in query or "good" in query:
            speak("It's good to know that your fine")



        elif "change my name to" in query:
            query = query.replace("change my name to", "")
            uname = query

        elif "change your name" in query:
            speak("What would you like to call me")
            assname = takeCommand()
            speak("Thanks for naming me")

        elif "what's your name" in query or "What is your name" in query:
            speak("My friends call me"+assname)
            print("My friends call me", assname)

        elif 'exit' in query or 'you can rest' in query or 'terminate' in query:
            speak("Thanks for giving me your time")
            exit()

        elif 'joke' in query:
            speak(pyjokes.get_joke())

        elif 'thanks' in query or 'thank you' in query:
            speak("my pleasure sire")

        elif "calculate" in query:

            app_id = "H9222P-7L4T8PWU7H"
            client = wolframalpha.Client(app_id)
            indx = query.lower().split().index('calculate')
            query = query.split()[indx + 1:]
            res = client.query(' '.join(query))
            answer = next(res.results).text
            print("The answer is " + answer)
            speak("The answer is " + answer)

        elif 'search' in query:

            query = query.replace("search", "")
            webbrowser.get(chrome_path).open_new_tab(query)

        elif 'play'in query:
            query = query.replace("play","")
            pwt.playonyt(query)
            speak("enjoy thw video")

        elif "who am i" in query:
            speak("If you talk then definitely your human.")

        elif "why you came to world" in query or 'how you came into the world' in query:
            speak("Thanks to my master. further It's a secret")

        # elif 'power point presentation' or 'presentation" in query:
        #     speak("opening Power Point presentation")
        #     power = r"C:\Users\ACER\Desktop\Stella.pptx"  # ppt path
        #     os.startfile(power)

        elif 'is love' in query:
            speak("love is the most twisted curse of all")

        elif "who are you" in query:
            speak("I am your virtual assistant"+assname)

        elif 'reason for you' in query or 'why were you built' in query or 'your purpose' in query:
            speak("I was created as a Minor project for first year.")

        elif 'change background' in query:
            ctypes.windll.user32.SystemParametersInfoW(20,0,"Location of wallpaper",0)
            speak("Background changed successfully")

        elif 'open bluestack' in query:
            appli = r"C:\\ProgramData\\BlueStacks\\Client\\Bluestacks.exe"
            os.startfile(appli)

        elif 'news' in query:

            try:
                jsonObj = urlopen("http://timesofindia.indiatimes.com/rssfeeds/-2128838597.cms?feedtype=sjson")
                data = json.load(jsonObj)
                i = 1

                speak('here are some top news from the times of india')
                print('''=============== TIMES OF INDIA ============''' + '\n')

                for item in data['articles']:

                    print(str(i) + '. ' + item['title'] + '\n')
                    print(item['description'] + '\n')
                    speak(str(i) + '. ' + item['title'] + '\n')
                    i += 1
            except Exception as e:

                print(str(e))

        elif 'lock window' in query:
            speak("locking the device")
            ctypes.windll.user32.LockWorkStation()

        elif 'shutdown' in query:
            speak("Hold On a Sec ! Your system is on its way to shut down")
            subprocess.call('shutdown /p /f')

        elif 'empty recycle bin' in query:
            winshell.recycle_bin().empty(confirm=False, show_progress=False, sound=True)
            speak("Recycle Bin Recycled")

        elif "don't listen" in query or "stop listening" in query:
            speak("for how much time you want to stop Sylvie from listening commands")
            a = int(takeCommand())
            time.sleep(a)
            print(a)
            print("I am back sire")

        elif "where is" in query:
            query = query.replace("where is", "")
            location = query
            speak("User asked to Locate")
            speak(location)
            webbrowser.open("https://www.google.com/maps/place/" + location)

        elif "take a photo" in query:
            ec.capture(0, "Sylvie Camera ", "img.jpg")

        elif "camera" in query or "open camera" in query:
            pyautogui.press('win')
            time.sleep(1)
            pyautogui.write("camera",interval = 0.1)
            pyautogui.press('enter')
            pyautogui.press('enter')

        elif "restart" in query:
            subprocess.call(["shutdown", "/r"])

        elif "hibernate" in query or "sleep" in query:
            speak("Hibernating")
            subprocess.call("shutdown","/h")

        elif "log off" in query or "sign out" in query:
            speak("Make sure all the application are closed before sign-out")
            time.sleep(5)
            subprocess.call(["shutdown", "/l"])

        elif "write a note" in query:
            speak("What should i write")
            note = takeCommand()
            file = open('Sylvie.txt', 'w+')
            speak("Should i include date and time as well")
            snfm = takeCommand()
            if 'yes' in snfm or 'sure' in snfm:
                strTime = datetime.datetime.now().strftime("%H:%M:%S")
                file.write(strTime)
                file.write(" :- ")
                file.write(note)
            else:
                file.write(note)

        elif "show note" in query:
            speak("Showing Notes")
            file = open("Sylvie.txt", "r")
            print(file.read())
            speak(file.read())

        
        elif 'who created you'in query:
            speak("I was created by four students from     D     J     S     C    E      aryan sharma     aryan singh      abhinav niar      and    amit upadhyay.\n")

        elif "update assistant" in query:
            speak(
                "After downloading file please replace this file with the downloaded one")
            url = '# url after uploading file'
            r = requests.get(url, stream=True)

            with open("Voice.py", "wb") as Pypdf:

                total_length = int(r.headers.get('content-length'))

                for ch in progress.bar(r.iter_content(chunk_size=2391975),
                                    expected_size=(total_length / 1024) + 1):
                    if ch:
                        Pypdf.write(ch)

        # NPPR9-FWDCX-D2C8J-H872K-2YT43
        elif "assistant" in query:

            hour = int(datetime.datetime.now().hour)
            if hour >= 0 and hour < 12:
                speak("Good Morning !")
            elif hour >= 12 and hour < 18:
                speak("Good Afternoon !")

            else:
                speak("Good Evening !")

            speak("at your service sire")
            

        elif "weather" in query:

            # Google Open weather website
            # to get API of Open weather
            api_key = "Api key"
            base_url = "http://api.openweathermap.org/data/2.5/weather?"
            speak("which city sire ")
            print("which city sire: ")
            city_name = takeCommand()
            complete_url = base_url + "appid =" + api_key + "&q =" + city_name
            response = requests.get(complete_url)
            x = response.json()

            if x["code"] != "404":
                y = x["main"]
                current_temperature = y["temp"]
                current_pressure = y["pressure"]
                current_humidiy = y["humidity"]
                z = x["weather"]
                weather_description = z[0]["description"]
                print(" Temperature (in kelvin unit) = " + str(current_temperature)+"\n atmospheric pressure (in hPa unit) ="+str(
                current_pressure) + "\n humidity (in percentage) = " + str(current_humidiy) + "\n description = " + str(weather_description))

            else:
                speak(" City Not Found ")

        elif "send message " in query:
            # You need to create an account on Twilio to use this service
            account_sid = 'Account Sid key'
            auth_token = 'Auth token'
            client = Client(account_sid, auth_token)

            message = client.messages \
                .create(
                    body=takeCommand(),
                    from_='Sender No',
                    to='Receiver No'
                )

            print(message.sid)

        elif "morning" in query:
            speak("A warm good morning to you too")
            speak("How are you ")
            speak(uname)

        # most asked question from google Assistant
        elif "will you be my gf" in query or "will you be my bf" in query:
            speak("I'm not sure about, may be you should give me some time")


        elif "i love you" in query:
            speak("I Love myself too")

        elif "what is" in query or "who is" in query:

            # Use the same API key
            # that we have generated earlier
            client = wolframalpha.Client("4211fe82c01f049ce4adf91324bd8cd1")
            res = client.query(query)

            try:
                print(next(res.results).text)
                speak(next(res.results).text)
            except StopIteration:
                print("No results")

        # elif "" in query:
            # Command here
            # For adding more commands

        else:
            speak("sorry it seems that is out of my current scope right now")
            speak("im sorry for the inconvenience ")
