import speech_recognition as sr
import pyttsx3 as pytx
import datetime
import wikipedia
import pywhatkit
import subprocess
import requests
from datetime import datetime



API_KEY = '82794bcc5aaf9c93301da72f07edd116'


audio = sr.Recognizer() # Realizar reconhecimento do audio
machine = pytx.init()   # Inicializar a maquina para o text to speech

def execute():
    try:
        with sr.Microphone() as mic: # Abertura e fechamento do Mic
            audio.adjust_for_ambient_noise(mic)
            print('listening...')
            voice = audio.listen(mic) # Reconhece o que foi ouvido no Mic
            command = audio.recognize_google(voice, language="pt-BR") # Utiliza o recognize_google pra traduzir o que ouviu em texto
            command = command.lower() 
            if 'jarvis' in command:
                command = command.replace('jarvis', '')
                #machine.say(command)
                machine.runAndWait()

    except:
        print('Mic is not working...')

    return command


def command_voice():
    command = execute()
    if 'horas' in command:
        horas = datetime.datetime.now().strftime('%H:%M')
        machine.say('Agora são ' + horas)
        machine.runAndWait()
        

    elif 'procure por' in command:
        search = command.replace('procure por', '')
        wikipedia.set_lang('pt')
        result = wikipedia.summary(search, 2)
        print(result)
        machine.say(result)
        machine.runAndWait() 

    elif 'toque' in command:
        musica = command.replace('toque','')
        youtube_result = pywhatkit.playonyt(musica)
        machine.say('Sim senhor chefe')
        machine.runAndWait()

    elif 'abrir excel' in command:
        path = '/Applications/Microsoft Excel.app'
        subprocess.Popen(["open", path])
        
    elif 'abrir notas' in command:
        path = '/System/Applications/Notes.app'
        subprocess.Popen(["open", path])
    
    elif 'abrir word' in command:
        path = '/Applications/Microsoft Word.app'
        subprocess.Popen(["open", path])

    elif 'abrir powerpoint' in command:
        path = '/Applications/Microsoft PowerPoint.app'
        subprocess.Popen(["open", path])

    elif 'temperatura' in command:
        city_raw = command.replace('temperatura','')
        city_strip = city_raw.strip()
        if ' ' in city_strip:
            city = city_strip.replace(' ', '%20')
            print(city)
        else:
            city = city_raw
        
        link = f'https://api.openweathermap.org/data/2.5/weather?q={city}&appid=fee510c86e406b8d39d0be50f9b0727a'
        data_request = requests.get(link)
        dictionary = data_request.json()
        weather = str(round(dictionary['main']['temp'] - 273.15))
        machine.say(f'A temperatura em {city_raw} é ' + weather + 'Graus Celsius')
        machine.runAndWait() 

    
    elif 'Que dia é hoje' in command:
        today = datetime.today()
        d1 = str(today.strftime("%d/%m/%Y"))
        machine.say(d1)
        machine.runAndWait()

command_voice()
