#print("this this Taylor personal assistant")

import speech_recognition as sr
import webbrowser
import pyttsx3 
from fuzzywuzzy import fuzz
import requests
import openai
import openpyxl
import os
from datetime import datetime

# Initialize recognizer class
recognizer = sr.Recognizer()

# Initialize pyttsx3 class

engine = pyttsx3.init()

api_key = '30c65596d93143ee83466cee73320a2c'
url = f'https://newsapi.org/v2/top-headlines?country=us&apiKey={api_key}'
response = requests.get(url)
data = response.json()


# weather setup 

weather_api_key = '4a89db6b4da96e6662ef58e25bc4cc67'

def get_weather(city="Dallas"):  # Default to a city like Denton
    weather_url = f"http://api.openweathermap.org/data/2.5/weather?q={city}&appid={weather_api_key}&units=imperial"
    response = requests.get(weather_url)
    if response.status_code == 200:
        weather_data = response.json()
        temp = weather_data['main']['temp']
        description = weather_data['weather'][0]['description']
        speak(f"The current temperature in {city} is {temp} degrees Fahrenheit with {description}.")
    else:
        speak("Sorry, I couldn't fetch the weather right now.")

client = openai.OpenAI(api_key="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")
def ai_process(command):
    try:
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": "You are a virtual assistant named Taylor, skilled in general tasks like Alexa and Google Assistant."},
                {"role": "user", "content": command}
            ]
        )
        reply = response.choices[0].message.content
        return reply
    except Exception as e:
        return f"An error occurred: {e}"
    
    
def log_conversation_to_excel(user_input, jarvis_reply, filename="conversation_log.xlsx"):
    if not os.path.exists(filename):
        # Create a new workbook and sheet
        wb = openpyxl.Workbook()
        sheet = wb.active
        sheet.title = "Chat History"
        # Create headers
        sheet.append(["Timestamp", "User", "Jarvis"])
    else:
        wb = openpyxl.load_workbook(filename)
        sheet = wb["Chat History"]

    # Append the new row
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    sheet.append([timestamp, user_input, jarvis_reply])
    wb.save(filename)    

def speak(text):
    engine.say(text)
    engine.runAndWait()

def is_match(command, keyword, threshold=80):
    return fuzz.partial_ratio(command.lower(), keyword.lower()) >= threshold

def processcommand(c):
    c = c.lower().strip()
    
    if c.startswith("open"):
        website = c.replace("open", "").strip().replace(" ", "")
        url = f"https://{website}.com"
        webbrowser.open(url)
    elif is_match(c, "news") or is_match(c, "read news"):
        speak("Here are the top headlines.")
        articles = data.get("articles", [])[:3]  # Read top 3 headlines
        if not articles:
            speak("Sorry, I couldn't find any news right now.")
        else:
            for i, article in enumerate(articles, start=1):
                title = article.get("title", "No title")
                speak(f"Headline {i}: {title}")
    elif is_match(c, "weather") or is_match(c, "what's the weather"):
        get_weather()
    else:
        output = ai_process(c)
        speak(output)
        log_conversation_to_excel(c, output)
        

    


if __name__ == "__main__":
    speak("Initialing tayler")
    while True:
        
            r = sr.Recognizer()

            print("recognizing.....")
            try:
                with sr.Microphone() as source:
                    print("Listening...")
                    audio = r.listen(source,timeout=3, phrase_time_limit=2 )

                command1 = r.recognize_google(audio)
                if (command1.lower()== "hey tayler!"):
                    speak("Ya")
                    
                with sr.Microphone() as source:
                    speak("Ya")
                    print("tayler Active")
                    audio = r.listen(source, timeout=3, phrase_time_limit=2)
                    command = r.recognize_google(audio)    
                    
                    processcommand(command)
                
            except Exception as e:
                print("Error; {0}".format(e))
        