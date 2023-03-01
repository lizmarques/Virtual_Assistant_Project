# FUNCIONALIDADES DA ASSISTENTE VIRTUAL

# Função 0: Quem é a Ártemis?
# Função 1: Notícias do Dia - site Folha de São Paulo
# Função 2: Abrir o Microsoft Teams
# Função 3: Agenda do Dia - vinculada ao Outlook
# Função 5: Bloco de Notas por Voz
# Função 6: Pesquisa no Google

# Importações
from voicebot_functions import *

# Modificando o logo do Tkinter
janela.iconbitmap(r"C:\Users\liznm\PycharmProjects\voicebot_project\VoiceBotProject\images\1.ico")
# janela.attributes('-fullscreen',True)
janela.geometry("1920x1080")
play_gif()

# Saudação inicial
greet()

# Loop do sistema
while True:

    # Função de gravação do áudio
    recording_audio()

    # Variável "speech": contém o áudio transformado em texto
    speech = audio_to_text()

    # Funcionamento da Assistente Virtual
    if "nome" in speech:
        assist_speak("Meu nome é Ártemis e eu sou uma assistente virtual."
                     "E pra você que não sabe, o nome Ártemis é em homenagem"
                     " a deusa grega da caça e da natureza.")
        assist_interaction()

    if "notícias" in speech:
        info = news_scraping()
        assist_speak("Principais notícias sobre tecnologia do site Folha de São Paulo.{}".format(info))
        assist_interaction()

    if "abrir" in speech:
        open_teams()
        assist_interaction()

    if "bloco de notas" in speech:
        assist_speak("O que você quer escrever?")
        voice_note()
        assist_speak("Anotado!")
        assist_interaction()

    if "pesquisar" in speech:
        internet_search()
        assist_interaction()

    if "agenda" in speech:
        daily_meetings()
        assist_interaction()

    goodbyes_list = ["isso é tudo", "tchau", "obrigada"]
    if speech in goodbyes_list:
        assist_speak("Precisando, é só chamar. Até breve!")
        close_window()
        break
