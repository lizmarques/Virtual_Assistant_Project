# Importando pacotes
import speech_recognition as sr
import sounddevice as sd
import wavio as wv
import webbrowser
import datetime
from datetime import date
import pandas as pd
import requests
import random
from tkinter import *
from PIL import Image, ImageTk, ImageSequence
from time import sleep
import subprocess
from pynput.keyboard import Key, Controller as Key_controller
import playsound
from gtts import gTTS
from O365 import Connection, FileSystemTokenBackend, Account
import time
import os
from bs4 import BeautifulSoup
import config

user = "Liz"


# Função para tratamento de erros
def error_treat():
    assist_speak(f"{user}, sua fala não foi processada. Estou desligando o sistema.")
    close_window()


# Função de fala da assistente
def assist_speak(text):
    text = str(text)
    tts = gTTS(text=text, lang='pt')        # gTTS = Google Text-to-Speech API
    audio_tts = "audio.mp3"
    tts.save(audio_tts)
    playsound.playsound(audio_tts)
    # print(text)
    os.remove(audio_tts)


# Função de gravação do áudio
def recording_audio():
    freq = 48000                                                    # Frequência
    duration = 5                                                    # Duração de cada gravação
    assist_speak('Ouvindo...')
    recording = sd.rec(int(duration * freq),                        # Gravar matrizes NumPy contendo sinais de áudio
                       samplerate=freq, channels=2)
    sd.wait()                                                       # Verifica se a gravação já terminou
    wv.write("audio_vox.wav", recording, freq, sampwidth=2)         # Grava uma matriz Numpy em um arquivo WAV


# Função para converter o áudio em texto
def audio_to_text():
    # Iniciando o reconhecimento de fala
    r = sr.Recognizer()
    filename = "audio_vox.wav"

    # Abrindo o arquivo
    with sr.AudioFile(filename) as source:
        # "Escutando" o arquivo
        audio_data = r.listen(source)
        # Convertendo o audio em texto
        try:
            text = r.recognize_google(audio_data, language='pt-BR')         # Google Speech Recognition API
            # Escrevendo o que foi dito
            print('{}: '.format(user) + text.lower())
            return text.lower()

        except sr.UnknownValueError:
            return error_treat()

        except sr.RequestError:
            assist_speak("Desculpa, mas não existe conexão com a internet."
                         "Tente novamente quando estiver conectado.")
            close_window()


# Função de Saudação
def greet():
    now = datetime.datetime.now()
    six_am = now.replace(hour=6, minute=00, second=0, microsecond=0)
    mid_day = now.replace(hour=12, minute=00)
    six_pm = now.replace(hour=18, minute=00)
    mid_night = now.replace(hour=00, minute=00)

    if mid_day > now >= six_am:
        d = str("Bom dia {}! Como eu posso te ajudar?".format(user))
        assist_speak(d)

    elif six_pm > now >= mid_day:
        t = str("Boa tarde {}! Como eu posso te ajudar?".format(user))
        assist_speak(t)

    elif now >= six_pm:
        n = str("Boa noite {}! Como eu posso te ajudar?".format(user))
        assist_speak(n)

    elif six_am > now >= mid_night:
        m = str("Boa madrugada, {}! Como eu posso te ajudar?".format(user))
        assist_speak(m)

    return


# Função de Interação
def assist_interaction():
    interaction = ["Deseja mais alguma coisa?", "Posso te ajudar em algo mais?",
                   "Mais alguma informação, {}?".format(user)]
    return assist_speak(random.choice(interaction))

# Abrindo a interface da Ártemis
janela = Tk()


# Função para rodar o GIF no Tkinter
def play_gif():
    global img
    img = Image.open("videos/voicebot_sem_logo.gif")

    lbl = Label(janela)
    lbl.place(x=0, y=0)
    for img in ImageSequence.Iterator(img):
        img = img.resize((1362, 764))
        img = ImageTk.PhotoImage(img)
        lbl.config(image=img)
        janela.update()
        sleep(0.01)
    janela.after(0, play_gif)


# Função para o fechamento da janela
def close_window():
    janela.destroy()


# Funções de atalho no teclado
def cursor_to_end():
    keyboard = Key_controller()
    keyboard.press(Key.ctrl.value)
    keyboard.press(Key.end.value)
    keyboard.release(Key.ctrl.value)
    keyboard.release(Key.end.value)


def save_file():
    keyboard = Key_controller()
    keyboard.press(Key.ctrl.value)
    keyboard.press("s")
    keyboard.release(Key.ctrl.value)
    keyboard.release("s")


def close_file():
    keyboard = Key_controller()
    keyboard.press(Key.ctrl.value)
    keyboard.press("w")
    keyboard.release(Key.ctrl.value)
    keyboard.release("w")


# Funcionalidades da Assistente Virtual
# Função 1: Notícias do Dia - site: Folha de São Paulo
def news_scraping():
    folha_de_sp = "https://www1.folha.uol.com.br/tec/"
    res = requests.get(folha_de_sp)
    soup = BeautifulSoup(res.content, 'html.parser')

    headlines_1 = soup.find_all('h2', {'class': 'c-main-headline__title'})
    headlines_2 = soup.find_all('h2', {'class': 'c-headline__title'})
    headlines_2_final = random.sample(headlines_2, 2)

    headlines = headlines_1 + headlines_2_final
    all_headlines = [h.text + "." for h in headlines]
    # print(all_headlines)

    data_folha_de_sp = pd.DataFrame(headlines)
    data_folha_de_sp.to_csv("folha_de_sp.csv", index=True)

    return all_headlines


# Função 2: Abrir o Microsoft Teams
def open_teams():
    path = r"C:\Users\liznm\AppData\Local\Microsoft\Teams\current\Teams.exe"
    subprocess.Popen(path)


# Função 3: Agenda do Dia - vinculada ao Outlook
def authenticate_outlook():
    # Autenticação das credenciais da API do Microsoft Graph
    # Microsoft Graph é o gateway para dados e inteligência no Microsoft 365

    credentials = (config.outlook_client_id, config.outlook_client_secret)
    token_backend = FileSystemTokenBackend(
        token_path=config.outlook_token_path, token_filename=config.outlook_token_filename
    )
    account = Account(credentials, token_backend=token_backend)
    if not account.is_authenticated:
        # Caso não seja autenticado, mostrar erro
        account.authenticate(scopes=config.outlook_scopes)

    connection = Connection(credentials, token_backend=token_backend, scopes=config.outlook_scopes)
    connection.refresh_token()

    # print("Authenticated Outlook.")
    return account


def get_outlook_events(calendar):

    now = datetime.datetime.now()
    mid_night = now.replace(hour=00, minute=00)
    before_mid_night = now.replace(hour=23, minute=59)

    query = calendar.new_query('start').greater_equal(mid_night)
    query.chain('and').on_attribute('end').less_equal(before_mid_night)

    events = calendar.get_events(query=query, limit=None, include_recurring=True)
    events = list(events)

    complete_date = str(date.today())

    final_schedule = [str(event).replace("Subject: ", "").replace("from:", "de").replace(":00 to:", " a")
                          .replace("(on:", "").replace(":00)", "").replace(complete_date, "") for event in events]
    # print(final_schedule)

    return final_schedule


def daily_meetings():
    try:
        # Autenticação do Outlook
        outlook_acct = authenticate_outlook()

        # Acessa e recupera todos os eventos do calendário no outlook
        outlook_calendar = outlook_acct.schedule().get_default_calendar()
        outlook_events = get_outlook_events(outlook_calendar)

        # outlook_events_reminder = outlook_calendar.remind_before_minutes = 15

        weekday_number = date.today().weekday()
        current_day = date.today().day
        current_month = date.today().month

        days = {0: "segunda feira", 1: "terça feira", 2: "quarta feira", 3: "quinta feira", 4: "sexta feira",
                5: "sábado", 6: "domingo"}
        weekday_written = [value for key, value in days.items() if key == weekday_number]

        months = {1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril", 5: "maio", 6: "junho",
                  7: "julho", 8: "agosto", 9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro"}
        current_month_written = [value for key, value in months.items() if key == current_month]

        assist_speak("Agenda de {}, dia {} de {}:".format(weekday_written, current_day, current_month_written))

        assist_speak(outlook_events)

    except AssertionError:
        assist_speak("{}, sua agenda está livre hoje".format(user))


# Função 5: Bloco de Notas por Voz
def voice_note():
    file_name = "voice_notes.txt"
    if os.path.exists("voice_notes.txt") is False:
        open("voice_notes.txt", "w").close()
    subprocess.Popen(["notepad.exe", file_name])
    keyboard = Key_controller()
    r = sr.Recognizer()
    subprocess.run('', shell=True)
    while True:
        with sr.Microphone() as source:
            assist_speak("Ouvindo")
            audio = r.listen(source)
            try:
                final_audio = r.recognize_google(audio, language='pt-BR')
                if "sair" in final_audio:
                    break
                else:
                    cursor_to_end()
                    keyboard.type("- ")
                    for c in final_audio.capitalize():
                        keyboard.type(c)
                        time.sleep(0.1)
                    keyboard.press(Key.enter)
            except sr.UnknownValueError:
                assist_speak(f"Desculpa,{user} eu não entendi o que você falou"
                             f"Você pode repetir?")

    save_file()
    close_file()


# Função 6: Pesquisa no Google
def internet_search():
    speech = audio_to_text()
    search_term = speech.replace("pesquisar", "")
    webbrowser.open("http://google.com/search?q=" + search_term)
