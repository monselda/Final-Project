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
import pyttsx3  #pip install pyttsx3==2.71
import speech_recognition as sr


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


#function for controlling laptop's brightness, monitor, volume.
def nircmd():
    '''if 'monitor off' in said:
        os.system("C:/nircmd.exe monitor off")
        get_audio()

    if 'monitor on' in said:
        os.system("C:/nircmd.exe monitor on")'''

    if 'increase volume' in said or 'volume up' in said:
        os.system("C:/nircmd.exe changesysvolume 5000")
        say("volume increased")

    if 'decrease volume' in said or 'volume down' in said:
        os.system("C:/nircmd.exe changesysvolume -5000")
        say("volume decreased")

    if 'max volume' in said:
        os.system("C:/nircmd.exe setsysvolume 65535")
        say("max volume") 

    if 'mute' in said:
        os.system("C:/nircmd.exe mutesysvolume 1")

    if 'unmute' in said:
        os.system("C:/nircmd.exe mutesysvolume 0")
        say("volume unmuted")

    if 'stand by' in said:
        os.system("C:/nircmd.exe standby")


#function to shutdown and restart computer
def snr():
    if "shutdown computer" in said:
        say("computer will shutdown in 5 seconds")
        os.system("shutdown /s /t 5")

    if "restart computer" in said:
        say("computer will restart in 5 seconds")
        os.system("shutdown /r /t 5")




#function to get google calendar account
def authenticate_user():
    """Shows basic usage of the Google Calendar API.
    Prints the start and name of the next 10 events on the user's calendar.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('calendar', 'v3', credentials=creds)

    return service

#fuction to get reminders
def get_reminder(day, service):
    if "what do i have" in said or "reminders" in said or "schedule" in said:
        # Call the Calendar API
        date = datetime.datetime.combine(day, datetime.datetime.min.time())
        end_date = datetime.datetime.combine(day, datetime.datetime.max.time())
        utc = pytz.UTC
        date = date.astimezone(utc)
        end_date = end_date.astimezone(utc)

        events_result = service.events().list(calendarId='primary', timeMin=date.isoformat(), timeMax=end_date.isoformat(),
                                            singleEvents=True,
                                            orderBy='startTime').execute()
        events = events_result.get('items', [])

        if not events:
            say('No upcoming events found.')
        else:
            say(f"You have {len(events)} events on this day.")
            for event in events:
                start = event['start'].get('dateTime', event['start'].get('date'))
                print(start, event['summary'])
                start_time = str(start.split("T")[1].split("+")[0])
                if int(start_time.split(":")[0]) < 12:
                    start_time = start_time + "am"
                else:
                    start_time = str(int(start_time.split(":")[0])-12)
                    start_time = start_time + "pm"

                say(event["summary"] + " at " + start_time)

#function to get date from reminders
def get_date(said):
    said = said.lower()
    today = datetime.date.today()

    if said.count("today") > 0:
        return today

    day = -1
    day_of_week = -1
    month = -1
    year = today.year

    for word in said.split():
        if word in MONTHS:
            month = MONTHS.index(word) + 1
        elif word in DAYS:
            day_of_week = DAYS.index(word)
        elif word.isdigit():
            day = int(word)
        else:
            for ext in DAY_EXTENTIONS:
                found = word.find(ext)
                if found > 0:
                    try:
                        day = int(word[:found])
                    except:
                        pass

    if month < today.month and month != -1:  # if the month mentioned is before the current month set the year to the next
        year = year+1

    # This is slighlty different from the video but the correct version
    if month == -1 and day != -1:  # if we didn't find a month, but we have a day
        if day < today.day:
            month = today.month + 1
        else:
            month = today.month

    # if we only found a dta of the week
    if month == -1 and day == -1 and day_of_week != -1:
        current_day_of_week = today.weekday()
        dif = day_of_week - current_day_of_week

        if dif < 0:
            dif += 7
            if said.count("next") >= 1:
                dif += 7

        return today + datetime.timedelta(dif)

    if day != -1: 
        return datetime.date(month=month, day=day, year=year)



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
    nircmd()



if __name__ == "__main__":
    greet()
    while True:
        service = authenticate_user()
        print("Assistant: Listening...")
        said = get_audio()
        main(said)
        notepad()
        date = get_date(said)
        get_reminder(date, service)

        if "write mode" in said:
            say("write mode activated.")
            while True:
                data = input(">> ")
                said = data.lower()
                if data == "deactivate":
                    break
                else:
                    main(said)
                    date = get_date(said)
                    get_reminder(date, service)