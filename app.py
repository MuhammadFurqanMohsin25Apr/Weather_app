import requests
import json
import win32com.client as wincom
import os
city=input("Enter the name of the city: ")
api = os.getenv("my_api")
url=f"http://api.weatherapi.com/v1/current.json?key={api}&q={city}&aqi=no"

r=requests.get(url)
wdic=json.loads(r.text)
w_c=wdic['current']["temp_c"]
w_f=wdic['current']["temp_f"]
w_s=wdic['current']["wind_kph"]
w_m=wdic['current']["wind_mph"]
w_d=wdic['current']["wind_dir"]
f_c=wdic['current']["feelslike_c"]
f_f=wdic['current']["feelslike_f"]
speak = wincom.Dispatch("SAPI.SpVoice")
text = f"The current temperature of {city} is {w_c} degrees in celcius. and in fahrenheit it is {w_f} degrees. the wind speed is {w_s} kilometers per hour.and in miles per hour it is {w_m} miles per hour. the wind direction is {w_d}. the feels like temperature is {f_c} degrees in celcius and {f_f} degrees in fahrenheit."
speak.Speak(text)