from win32com.client import Dispatch
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import smtplib
import operator
from word2number import w2n
from googletrans import Translator
from googletrans import LANGCODES
from gtts import gTTS
from pygame import mixer
import random
import re
from PyDictionary import PyDictionary
from nltk.corpus import wordnet


synonyms = []
antonyms = []

dict = PyDictionary()

trans = Translator()

speak = Dispatch("SAPI.SpVoice")
dic={
  "Afrikaans": [
    ["South Africa", "af-ZA"]
  ],
  "Arabic" : [
    ["Algeria","ar-DZ"],
    ["Bahrain","ar-BH"],
    ["Egypt","ar-EG"],
    ["Israel","ar-IL"],
    ["Iraq","ar-IQ"],
    ["Jordan","ar-JO"],
    ["Kuwait","ar-KW"],
    ["Lebanon","ar-LB"],
    ["Morocco","ar-MA"],
    ["Oman","ar-OM"],
    ["Palestinian Territory","ar-PS"],
    ["Qatar","ar-QA"],
    ["Saudi Arabia","ar-SA"],
    ["Tunisia","ar-TN"],
    ["UAE","ar-AE"]
  ],
  "Basque": [
    ["Spain", "eu-ES"]
  ],
  "Bulgarian": [
    ["Bulgaria", "bg-BG"]
  ],
  "Catalan": [
    ["Spain", "ca-ES"]
  ],
  "Chinese Mandarin": [
    ["China (Simp.)", "cmn-Hans-CN"],
    ["Hong Kong SAR (Trad.)", "cmn-Hans-HK"],
    ["Taiwan (Trad.)", "cmn-Hant-TW"]
  ],
  "Chinese Cantonese": [
    ["Hong Kong", "yue-Hant-HK"]
  ],
  "Croatian": [
    ["Croatia", "hr_HR"]
  ],
  "Czech": [
    ["Czech Republic", "cs-CZ"]
  ],
  "Danish": [
    ["Denmark", "da-DK"]
  ],
  "English": [
    ["Australia", "en-AU"],
    ["Canada", "en-CA"],
    ["India", "en-IN"],
    ["Ireland", "en-IE"],
    ["New Zealand", "en-NZ"],
    ["Philippines", "en-PH"],
    ["South Africa", "en-ZA"],
    ["United Kingdom", "en-GB"],
    ["United States", "en-US"]
  ],
  "Farsi": [
    ["Iran", "fa-IR"]
  ],
  "French": [
    ["France", "fr-FR"]
  ],
  "Filipino": [
    ["Philippines", "fil-PH"]
  ],
  "Galician": [
    ["Spain", "gl-ES"]
  ],
  "German": [
    ["Germany", "de-DE"]
  ],
  "Greek": [
    ["Greece", "el-GR"]
  ],
  "Finnish": [
    ["Finland", "fi-FI"]
  ],
  "Hebrew" :[
    ["Israel", "he-IL"]
  ],
  "Hindi": [
    ["India", "hi-IN"]
  ],
  "Hungarian": [
    ["Hungary", "hu-HU"]
  ],
  "Indonesian": [
    ["Indonesia", "id-ID"]
  ],
  "Icelandic": [
    ["Iceland", "is-IS"]
  ],
  "Italian": [
    ["Italy", "it-IT"],
    ["Switzerland", "it-CH"]
  ],
  "Japanese": [
    ["Japan", "ja-JP"]
  ],
  "Korean": [
    ["Korea", "ko-KR"]
  ],
  "Lithuanian": [
    ["Lithuania", "lt-LT"]
  ],
  "Malaysian": [
    ["Malaysia", "ms-MY"]
  ],
  "Dutch": [
    ["Netherlands", "nl-NL"]
  ],
  "Norwegian": [
    ["Norway", "nb-NO"]
  ],
  "Polish": [
    ["Poland", "pl-PL"]
  ],
  "Portuguese": [
    ["Brazil", "pt-BR"],
    ["Portugal", "pt-PT"]
  ],
  "Romanian": [
    ["Romania", "ro-RO"]
  ],
  "Russian": [
    ["Russia", "ru-RU"]
  ],
  "Serbian": [
    ["Serbia", "sr-RS"]
  ],
  "Slovak": [
    ["Slovakia", "sk-SK"]
  ],
  "Slovenian": [
    ["Slovenia", "sl-SI"]
  ],
  "Spanish": [
    ["Argentina", "es-AR"],
    ["Bolivia", "es-BO"],
    ["Chile", "es-CL"],
    ["Colombia", "es-CO"],
    ["Costa Rica", "es-CR"],
    ["Dominican Republic", "es-DO"],
    ["Ecuador", "es-EC"],
    ["El Salvador", "es-SV"],
    ["Guatemala", "es-GT"],
    ["Honduras", "es-HN"],
    ["México", "es-MX"],
    ["Nicaragua", "es-NI"],
    ["Panamá", "es-PA"],
    ["Paraguay", "es-PY"],
    ["Perú", "es-PE"],
    ["Puerto Rico", "es-PR"],
    ["Spain", "es-ES"],
    ["Uruguay", "es-UY"],
    ["United States", "es-US"],
    ["Venezuela", "es-VE"]
  ],
  "Swedish": [
    ["Sweden", "sv-SE"]
  ],
  "Thai": [
    ["Thailand", "th-TH"]
  ],
  "Turkish": [
    ["Turkey", "tr-TR"]
  ],
  "Ukrainian": [
    ["Ukraine", "uk-UA"]
  ],
  "Vietnamese": [
    ["Viet Nam", "vi-VN"]
  ],
  "Zulu": [
    ["South Africa", "zu-ZA"]
  ]
}
def synant(word, ch):
    for syn in wordnet.synsets(word):
        for l in syn.lemmas():
            synonyms.append(l.name())
            if l.antonyms():
                antonyms.append(l.antonyms()[0].name())
    if ch is 's':
        print(set(synonyms))
        speak.Speak("the synonyms of the given word is")
        speak.Speak(set(synonyms))

    elif ch is 'e':
        speak.Speak("the antonyms of the given word is")
        print(set(antonyms))
        speak.Speak(set(antonyms))


def translate(line, lan):
    lang = LANGCODES[lan]

    t = trans.translate(
        line, src='en', dest=lang
    )
    print(f'Source: {t.src}')
    print(f'Destination: {t.dest}')
    print(f'{t.origin} -> {t.text}')
    ob = str(t.text)
    obj =gTTS(text = ob, slow = False, lang = lang)
    id = random.randint(0,1000000)
    id =str(id)
    obj.save(id+'.mp3')
    mixer.init()
    mixer.music.load(id+'.mp3')
    mixer.music.play()

    print()

def get_operator_fn(op):
    return {
        '+' : operator.add,
        '-' : operator.sub,
        'x' : operator.mul,
        'divided' :operator.__truediv__,
        'Mod' : operator.mod,
        'mod' : operator.mod,
        '^' : operator.xor,
        }[op]
def eval_binary_expr(op1, oper, op2):
    op1,op2 = w2n.word_to_num(op1), w2n.word_to_num(op2)
    return get_operator_fn(oper)(op1, op2)

def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour >=0 and hour<12:
        speak.Speak("Good Morning!")

    elif hour>12 and hour <=18:
        speak.Speak("Good Afternoon!")

    elif hour>18 and hour <=24:
        speak.Speak("Good Evening!")

    speak.Speak("I am King 01 bot Sir. How may I help you,Sir?")

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print("Recognizing....")
        query = r.recognize_google(audio, language = 'en-in')
        print(f"user said: {query}\n")
    except:
        print("Say that again please..")
        return "None"
    return query

def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login('anupampoddar1997@gmail.com','bestintheworld')
    server.sendmail('anupampoddar1997@gmail.com', to, content)
    server.close()

if __name__=="__main__":
    wishMe()
    while True:
        query = takeCommand().lower()

        if'wikipedia' in query:
            speak.Speak('Searching Wikipedia...')
            query = query.replace("wikipedia", " ")
            results = wikipedia.summary(query, sentences=2)
            speak("According to wikipedia")
            print(results)
            speak(results)

        elif 'open youtube' in query:
            speak.Speak("What would you like to search on youtube, sir?")
            sear =str(takeCommand())
            sear = sear.replace(" ","+")

            webbrowser.get("C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s").open(
                f'https://www.youtube.com/results?search_query={sear}')


        elif 'open google' in query:
            webbrowser.get("C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s").open(
                'https://www.google.com')
        elif 'open stackoverflow' in query:
            webbrowser.get("C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s").open(
                'https://www.stackoverflow.com')
        elif 'play music' in query:
            music_dir = "C:\\Users\\Acer\\Music"
            songs = os.listdir(music_dir)
            speak.Speak("Index of movie you want to open?")
            print(songs)
            index = int(input())
            os.startfile(os.path.join(music_dir, songs[index]))
        elif 'play movie' in query:
            movie_dir ="E:\\"
            movies = os.listdir(movie_dir)
            print(movies)
            speak.Speak("Index of movie you want to open?")
            index =int(input())

            os.startfile(os.path.join(movie_dir,movies[index]))
        elif 'the time' in query:
            time= datetime.datetime.now().strftime("%H:%M:%S")
            speak.Speak(f"Sir, the time is {time}")
        elif 'open code' in query:
            path = r"C:\Program Files\JetBrains\PyCharm Community Edition 2019.3.3\bin\pycharm64.exe"
            os.startfile(path)
        elif 'calculator' in query:
            speak.Speak("what do you want to calculate,sir?")
            ques = takeCommand()
            print(ques)
            ans = str(eval_binary_expr(*(ques.split())))
            speak.Speak(f"Sir, the answer is {ans}")
        elif 'language translator' in query:
            speak.Speak("what would you like to translate?")
            text = takeCommand().lower()
            speak.Speak("sir,What language you would like to translate the text into?")
            lan =  takeCommand().lower()
            translate(text, lan)
        elif re.search("open.*site", query):
            speak.Speak("What website would you like to open, sir?")
            site = takeCommand()
            print(site)
            webbrowser.get("C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s").open(
               f'https://{site}')
        elif re.search("word", query):
            if 'meaning' in query:
                speak.Speak("what is the word sir?")
                word = str(takeCommand())
                print(dict.meaning(word))
                speak.Speak(dict.meaning(word))
            elif 'synonym' in query:
                speak.Speak("what is the word sir?")
                word = takeCommand()
                synant(word,'s')

            elif 'antonym' in query:
                speak.Speak("what is the word sir?")
                word = takeCommand()
                synant(word, 'a')

        elif 'email to' in query:
            try:
                speak.Speak("What is the email id of recepient?")
                to= takeCommand().lower()
                to =to.replace(" ","")
                print(to)

                speak.Speak("what should i say?")
                content = takeCommand()
                sendEmail(to,content)
                speak("Email has been sent")

            except Exception as e:
                print(e)
                speak.Speak("Sorry my lord Anupam. I am not able send this mail")

        elif 'quit' in query:
            speak.Speak("See you soon, Thankyou sir!")

            exit()

