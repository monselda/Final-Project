#modules for reminder google calendar
from __future__ import print_function   
import datetime
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
#end of modules for reminder google calendar

import pytz
import pyttsx3  #pip install pyttsx3==2.71
import speech_recognition as sr
import datetime
import os
import time
import mysql.connector



#db connector
mydb = mysql.connector.connect(
	host='localhost', 
	user='root', 
	password='GOODLUCK!@#', 
    database='commands',
	# Important!!!
	auth_plugin='mysql_native_password'
	# Important!!!
    )
mycursor = mydb.cursor()


#setting text to speech engine and voice rate property
engine = pyttsx3.init("sapi5")
engine.setProperty('rate', 150)


# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']

MONTHS = ["january", "february", "march", "april", "may", "june","july", "august", "september","october", "november", "december"]
DAYS = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]
DAY_EXTENTIONS = ["rd", "th", "st", "nd"]



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



#function for closing the program
def quit():
    if "goodbye" in said:
        say("okay, goodbye.")
        exit()
            


#main function that contains the commands of other functions
def main(said):
    quit()




if __name__ == "__main__":
    while True:
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
                    