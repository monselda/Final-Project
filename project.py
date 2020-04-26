import pyttsx3  #pip install pyttsx3==2.71
import speech_recognition as sr


#setting text to speech engine and voice rate property
engine = pyttsx3.init("sapi5")
engine.setProperty('rate', 150)


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
def main():
    quit()


if __name__ == "__main__":
    
    while True:
        print("Assistant: Listening...")
        said = get_audio()
        main()

        if "write mode" in said:
            say("write mode activated.")
            while True:
                data = input(">> ")
                said = data.lower()
                if data == "deactivate":
                    break
                else:
                    main()
