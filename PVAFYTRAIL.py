from tkinter.tix import InputOnly
import pyttsx3 #pip install pyttsx3
import speech_recognition as sr #pip install speechRecognition
import datetime
import wikipedia #pip install wikipedia
import webbrowser
import os
import smtplib
import subprocess
import wolframalpha
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
import os
from twilio.rest import Client
from clint.textui import progress
#from ecapture import ecapture as ec
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen


engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
print(voices[1].id)
engine.setProperty('voice', voices[1].id)


def speak(audio):
    engine.say(audio)
    engine.runAndWait()


def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speak("Good Morning!")

    elif hour>=12 and hour<18:
        speak("Good Afternoon!")   

    else:
        speak("Good Evening!")  

    speak("Hello. Please tell me how may I help you")
  

def takeCommand():
    #It takes microphone input from the user and returns string output

    r = sr.Recognizer()
    audio = ''
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source, phrase_time_limit = 5)

    try:
        print("Recognizing...")    
        query = r.recognize_google(audio, language='en-in')
        print(f"User said: {query}\n")

    except Exception as e:
        # print(e)    
        print("Say that again please...")  
        return "None"
    return query

def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login('youremail@gmail.com', 'your-password')
    server.sendmail('youremail@gmail.com', to, content)
    server.close()

if __name__ == "__main__":
    clear = lambda: os.system('cls')
    clear()
    wishMe()
    while True:
    # if 1:
        query = takeCommand().lower()

        # Logic for executing tasks based on query
        if 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            speak("According to Wikipedia")
            print(results)
            speak(results)

        elif 'open youtube' in query:
            speak("Opening Youtube\n")
            webbrowser.open("youtube.com")

        elif 'open google' in query:
            speak("Opening Google\n")
            webbrowser.open("google.com")

        elif 'open stackoverflow' in query:
            speak("Opening Stackoverflow\n")
            webbrowser.open("stackoverflow.com")   


        elif 'play music' in query:
            music_dir = 'D:\\IMPORTANT\\music'
            songs = os.listdir(music_dir)
            print(songs)    
            os.startfile(os.path.join(music_dir, songs[0]))

        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")    
            speak(f"Sir, the time is {strTime}")

        elif 'open code' in query:
            codePath = "C:\\Users\\Haris\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
            os.startfile(codePath)

        elif 'email to harry' in query:
            try:
                speak("What should I say?")
                content = takeCommand()
                to = "harryyourEmail@gmail.com"    
                sendEmail(to, content)
                speak("Email has been sent!")
            except Exception as e:
                print(e)
                speak("Sorry I am not able to send this email")

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
                    speak("How are you, Sir")

        elif 'fine' in query or "good" in query:
                    speak("It's good to know that your fine")

        elif "what's your name" in query or "What is your name" in query:
                    speak("My friends call me")
                    speak(assname)
                    print("My friends call me Paffy")

        elif 'exit' in query:
                    speak("Thanks for giving me your time")
                    exit()

        elif 'joke' in query:
                    speak(pyjokes.get_joke())

        elif "calculate" in query:
                    
                    app_id = "UU46TX-JWQ958E4GG"
                    client = wolframalpha.Client(app_id)
                    indx = query.lower().split().index('calculate')
                    query = query.split()[indx + 1:]
                    res = client.query(' '.join(query))
                    answer = next(res.results).text
                    print("The answer is " + answer)
                    speak("The answer is " + answer)

        elif 'search' in query or 'play' in query:
                    
                    query = query.replace("search", "")
                    query = query.replace("play", "")		
                    webbrowser.open(query)

        elif "who made you" in query or "who created you" in query:
                    speak("I have been created by GOD.")

        elif "who am i" in query:
                    speak("If you talk then definitely you're a human.")

        # elif 'news' in query:
                    
        #             try:
        #                 jsonObj = urlopen('''https://newsapi.org / v1 / articles?source = the-times-of-india&sortBy = top&apiKey =\\times of India Api key\\''')
        #                 data = json.load(jsonObj)
        #                 i = 1
                        
        #                 speak('here are some top news from the times of india')
        #                 print('''=============== TIMES OF INDIA ============'''+ '\n')
                        
        #                 for item in data['articles']:
                            
        #                     print(str(i) + '. ' + item['title'] + '\n')
        #                     print(item['description'] + '\n')
        #                     speak(str(i) + '. ' + item['title'] + '\n')
        #                     i += 1
        #             except Exception as e:
                        
        #                 print(str(e))

        elif 'lock window' in query:
                        speak("locking the device")
                        ctypes.windll.user32.LockWorkStation()

        elif 'shutdown system' in query:
                        speak("Hold On a Sec ! Your system is on its way to shut down")
                        subprocess.call('shutdown / p /f')

        elif "restart" in query:
                    subprocess.call(["shutdown", "/r"])
                    
        elif "hibernate" in query or "sleep" in query:
                    speak("Hibernating")
                    subprocess.call("shutdown / h")

        elif "log off" in query or "sign out" in query:
                    speak("Make sure all the application are closed before sign-out")
                    time.sleep(5)
                    subprocess.call(["shutdown", "/l"])

        elif "write a note" in query:
                    speak("What should i write, sir")
                    note = takeCommand()
                    file = open('PVAFY.txt', 'w')
                    speak("Sir, Should i include date and time")
                    snfm = takeCommand()
                    if 'yes' in snfm or 'sure' in snfm:
                        strTime = datetime.datetime.now().strftime("% H:% M:% S")
                        file.write(strTime)
                        file.write(" :- ")
                        file.write(note)
                    else:
                        file.write(note)
                
        elif "show note" in query:
                    speak("Showing Notes")
                    file = open("PVAFY.txt", "r")
                    print(file.read())
                    speak(file.read(6))

        elif "where is" in query:
                            query = query.replace("where is", "")
                            location = query
                            speak("User asked to Locate")
                            speak(location)
                            webbrowser.open("https://www.google.nl / maps / place/" + location + "")

        elif "don't listen" in query or "stop listening" in query:
                    speak("for how much time you want to stop PVAFY from listening commands")
                    a = int(takeCommand())
                    time.sleep(a)
                    print(a)

        elif "PVAFY" in query:
                    
                    wishMe()
                    speak("PVAFY is at your service now")
                    speak(assname)

        elif "weather" in query:	
                    # Google Open weather website
                    # to get API of Open weather
                    api_key = "f79f3ceb2d0574b4478962809dbe6e54"
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
                        current_humidiy = y["humidity"]
                        z = x["weather"]
                        weather_description = z[0]["description"]
                        print(" Temperature (in kelvin unit) = " +str(current_temperature)+"\n atmospheric pressure (in hPa unit) ="+str(current_pressure) +"\n humidity (in percentage) = " +str(current_humidiy) +"\n description = " +str(weather_description))
                    
                    else:
                        speak(" City Not Found ")
        

        elif "open wikipedia" in query:
                    webbrowser.open("wikipedia.com")

        elif "Good Morning" in query:
                    speak("A warm" +query)
                    speak("How are you Mister")
                    speak(assname)

                # most asked question from google Assistant
        elif "will you be my gf" in query or "will you be my bf" in query:
                    speak("I'm not sure about, may be you should give me some time")

        elif "how are you" in query:
                    speak("I'm fine, glad you asked me that")

        elif "i love you" in query:
                    speak("It's hard to understand")

        elif "what is" in query or "who is" in query:
                    
                    # Use the same API key
                    # that we have generated earlier
                    client = wolframalpha.Client("UU46TX-JWQ958E4GG")
                    res = client.query(query)
                    
                    try:
                        print (next(res.results).text)
                        speak (next(res.results).text)
                    except StopIteration:
                        print ("No results")
        
        elif "send message " in query or "send sms" in query:
                print("sir what should i send")
                speak("sir what should i send")
            
                from twilio.rest import Client
                account_sid = 'AC525e13c2bd33cd0eeecd06518056c1db'
                auth_token = '479078fd8510f265467208b5c211d3fa'
                client = Client(account_sid, auth_token)

                message = client.messages \
                        .create(
                            body = takeCommand(),
                            from_='+16814994704',
                            to =input("ENTER A NUM  ")
                            )
                print(message.sid)  
        
        elif "call" in query:
            print("Whom should i call")
            speak("whom shoudl i call")
            
            from twilio.rest import Client
            account_sid = 'AC525e13c2bd33cd0eeecd06518056c1db'
            auth_token = '479078fd8510f265467208b5c211d3fa'
            client = Client(account_sid, auth_token)

            message = client.calls \
                .create(
                    twiml='<Response><say>This is the second testing message from pvafy..</say></Response>',
                    from_='+16814994704',
                    to =input("ENTER A NUM  ")
                    )
            print(message.sid)    
            
        elif "open chrome" in query:
            speak("Google Chrome")
            os.startfile('C:\Program Files\Google\Chrome\Application\chrome.exe')
           
  
       # elif "open mozilla firefox" in query:
        #    speak("Opening Mozilla Firefox")
         #   os.startfile('C:\Users\HP\AppData\Local\Mozilla Firefox\Mozilla Firefox.exe')
            
  
       # elif "open Office" in query:
       #      speak("Opening  office")
        #     os.startfile('https://www.office.com/?auth=1')
        
        elif "set alarm" in query:
            speak("sir please tell me the time to set alarm.for example, set alarm to 5:30 pm")
            tt = takeCommand()
            tt = tt.replace(" set alarm to ", "")
            tt = tt.replace(".","")
            tt = tt.upper()
            import alarm 
            alarm.alarm(tt)
            
        elif 'volume up' in query:
            pyautogui.press("volumeup")    
            
            
            
            

    
  
             
                      
                                