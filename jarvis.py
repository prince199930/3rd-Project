import pyttsx3
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import random
import requests
import json
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
# print(voices[0].id)
engine.setProperty('voice', voices[0].id)
def speak(audio):
    engine.say(audio)
    engine.runAndWait()
def Wishme():
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speak("Good Morning")

    elif hour>=12 and hour<18:
        speak("Good Afternoon")

    else:
        speak("good evening")
          

    speak(" I am Jarvis sir.How may I help you")           
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
        # print(e)    
        print("Say that again please...")  
        return "None"
    return query
def speak_1(str):
    from win32com.client import Dispatch
    speak = Dispatch("SAPI.SpVoice")
    speak.Speak(str)    
 


if __name__ == "__main__":
    Wishme()
    while True:
        query = takeCommand().lower()
        if 'bye' in query:
            print('bye sir')
            speak("bye sir")
            exit()
        elif 'hello jarvis' in query or 'hi' in query:
            print('hello sir')
            speak("hello sir")    
        elif 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            speak("According to Wikipedia")
            print(results)
            speak(results)
        elif 'open youtube' in query:
            webbrowser.open("youtube.com")   
        elif 'open google' in query:
            webbrowser.open("google.com")   
        elif 'open stackoverflow' in query:
            webbrowser.open("stackoverflow.com")
        elif 'play music' in query :
            music_dir = "C:\\Users\\hp\\Music\\my music"
            songs =os.listdir(music_dir)
            print(songs)
            randomfile = random.choice(songs)
            os.startfile(os.path.join(music_dir, randomfile))
        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"Sir, the time is {strTime}")
        elif 'open code' in query:
            codepath = "C:\\Users\\hp\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
            os.startfile(codepath)
        elif 'open command prompt' in query:
            print('opening command prompt')
            speak('opening command prompt')
            codepath = "C:\\windows\\system32\cmd.exe"
            os.startfile(codepath)       
        elif 'open notepad' in query:
            speak('opening notepad')
            print('opening notepad')
            notepad_dir ="C:\\windows\\system32\\cmd.exe"
            os.startfile(notepad_dir)
        elif 'force to close notepad' in query:
            speak('closing notepad')
            print('closing notepad')
            os.system('TASKKILL /F /IM notepad.exe')  
        elif 'open shareit' in query:
            print('opening shareit')
            speak('opening shareit')
            shareit_dir ="C:\\Program Files (x86)\\SHAREit Technologies\\SHAREit\\SHAREit.exe"
            os.startfile(shareit_dir)
        elif 'force to close shareit' in query:
            speak('closing shareit')
            print('closing shareit')
            os.system('TASKKILL /F /IM SHAREit.exe')      
        elif 'how are you' in query:
            print('I am fine sir and you')
            speak('I am fine sir and you')
        elif 'i am fine jarvis' in query:
            print('its good')
            speak('its good')    
        elif 'who is your owner' in query:
            print('Prince sir')
            speak('Prince sir')
            print('Next Command')
            speak('Next Command')   
        elif 'open my images' in query or 'open images' in query:
            images_dir ="C:\\Users\\hp\\Downloads\\SHAREit\\MAR-LX2\\photo"   
            images =os.listdir(images_dir)
            print(images)
            os.startfile(os.path.join(images_dir))   
        elif 'open microsoft office word' in query:
            microsoft_dir ="C:\\Users\\hp\\Desktop"
            microsoft =os.listdir(microsoft_dir)
            print(microsoft)
            os.startfile(os.path.join(microsoft_dir))  
        elif 'news' in query:
            speak_1("News for today.. Lets begin")
            url = "https://newsapi.org/v2/top-headlines?country=in&apiKey=a68d16adea2e4fdb9d1c49982c72cbe7"
            news = requests.get(url).text
            news_dict = json.loads(news)
            print(news_dict["articles"])
            arts = news_dict['articles']
            for article in arts:
                while True:
                    speak_1(article['title'])
                    query_1 = takeCommand()
                    print('Next>>')
                    if 'next' in query_1:
                        speak_1("..Moving on the next news...")
                        break
                if 'stop' in query_1:
                    break
            speak_1("Thanks for listening")    
        else:
            print("sir, I don't understand, Please say it again")    
            speak("sir, I don't understand, Please say it again")
    
    
            
    

            
    

              

    
    

        