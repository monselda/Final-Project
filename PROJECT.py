import datetime
import os
import time
import pyttsx3  #pip install pyttsx3==2.71
import speech_recognition as sr


#setting text to speech engine and voice rate property
engine = pyttsx3.init("sapi5")
engine.setProperty('rate', 150)


#function for synthesizer
def say(text):
    print("Assistant: " + text)
    engine.say(text)
    engine.runAndWait()


#getting input from microphone
def get_audio():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.adjust_for_ambient_noise(source)
        audio = r.listen(source)
        query = ""
    try:
        query = r.recognize_google(audio, language='en')
        print("User: " + query)
    except sr.UnknownValueError:
        print("Sorry, I did not get that.")
    return query.lower()

#function for greeting the user once the program starts
def greet():
    hour = int(datetime.datetime.now().hour)

    if 0 <= hour < 12:
        say("good morning")
    elif 12 <= hour < 18:
        say("good afternoon")
    else:
        say("good evening")


#function to paste and read something in notepad
def notepad():
    if "paste in notepad" in said:
        say("i am ready")
        text = get_audio()
        file = open(r"C:/Users/Computer-1\Documents/3rd Term - SY 2019-2020/CPE106L_B3/FINAL PROJECT/FP_2/notes/note.txt", "w+") 
        file.write(text)
        file.close()
    
    if "show text" in said:
        say("opening notepad")
        os.startfile('C:/Users/Computer-1\Documents/3rd Term - SY 2019-2020/CPE106L_B3/FINAL PROJECT/FP_2/notes/note.txt')
        with open("C:/Users/Computer-1\Documents/3rd Term - SY 2019-2020/CPE106L_B3/FINAL PROJECT/FP_2/notes/note.txt","r") as f:
            data = f.readlines()
            say(str(data))


#function that opens an application
def open_application():
    if "chrome" in said:
        say("opening Google Chrome")
        os.startfile('C:/Program Files (x86)/Google/Chrome/Application/chrome.exe')
            
    elif "brave" in said:
        say("opening brave browser")
        os.startfile("C:/Program Files (x86)/BraveSoftware/Brave-Browser/Application/brave.exe")
            
    elif "garena" in said:
        say("opening garena")
        os.startfile("C:/Program Files (x86)/Garena/Garena/Garena.exe")
            
    elif "word" in said:
        say("opening M S word")
        os.startfile("C:/Program Files (x86)/Microsoft Office/root/Office16/WINWORD.EXE")
            
    elif "powerpoint" in said:
        say("opening powerpoint")
        os.startfile("C:/Program Files (x86)/Microsoft Office/root/Office16/POWERPNT.EXE")
            
    else:
        say("sorry, that application is not installed in the system. or I wasn't programmed to open "
                "that application")


#function for closing the program
def quit():
    if "goodbye" in said:
        say("okay, goodbye.")
        exit()
            


#main function that contains the commands of other functions
def main(said):
    if "open" in said:
        open_application()
    quit()


if __name__ == "__main__":
    greet()
    while True:
        print("Assistant: Listening...")
        said = get_audio()
        main(said)
        notepad()

        if "write mode" in said:
            say("write mode activated.")
            while True:
                data = input(">> ")
                said = data.lower()
                if data == "deactivate":
                    break
                else:
                    main(said)