import sys

import pymongo

from get_audio import get_audio_eng
from speak import speak_eng
from voice_assistant_ENG import commands_eng
from Games_eng.tic_tac_toe_ENG import tic_tac_toe_eng

sys.path.append("../get_audio")
sys.path.append("../speak")
import pyodbc

login = ''


def introductory_talk(conn):
    speak_eng(f"{login} Do you need some more my help?")
    text = get_audio_eng()
    agreement = ('yes', 'yeah', 'that\'s right', 'Absolutely correct', 'correct')
    refusal = ('no', 'bye', 'goodbye', 'doesn\'t need', 'incorrect')
    if text in agreement:
        main_eng_va("", login)
    elif text in refusal:
        conn.commit()
        speak_eng("Have a good day")
    else:
        speak_eng("Sorry, could you say yes or no")
        introductory_talk(conn)


def main_eng_va(check_command="", name="ERROR"):
    global login
    login = name
    myclient = pymongo.MongoClient("mongodb://localhost:27017/")
    mydb = myclient["voice_assistant"]
    mongo_collect = mydb["actions"]
    if check_command == "":
        conn = pyodbc.connect('Driver={SQL Server};'
                              'Server=USER-PC\SQLEXP;'
                              'Database=Voice_assistant;'
                              'Trusted_Connection=yes;')
        cursor = conn.cursor()
        speak_eng(f'Nice to meet you {login}. How can I help you?')
        text = get_audio_eng()
    else:
        text = check_command
    if "timer" in text:
        cursor.execute(f"Insert into dbo.History values('{login}', 'establish timer')")
        commands_eng.timer(mongo_collect, login)
        cursor.commit()
    elif "time" in text:
        cursor.execute(f"INSERT INTO dbo.HISTORY values('{login}', 'checking current time')")
        commands_eng.time(mongo_collect, login)
    elif "weather" in text:
        cursor.execute(f"Insert into dbo.History values('{login}', 'checking the current weather')")
        commands_eng.weather(mongo_collect, login)
        conn.commit()
    elif "word file" in text:
        cursor.execute(f"Insert into dbo.History values('{login}', 'creating a word file')")
        commands_eng.creating_word_file(mongo_collect, login)
        cursor.commit()
    elif "computer" in text:
        cursor.execute(f"Insert into dbo.History values('{login}', 'opening my computer')")
        commands_eng.open_my_computer(mongo_collect, login)
        cursor.commit()
    elif "recycle bin" in text:
        cursor.execute(f"Insert into dbo.History values('{login}', 'cleaning the recycle bin')")
        commands_eng.cleaning_recycle_bin(mongo_collect, login)
        cursor.commit()
    elif "browser" in text:
        cursor.execute(f"Insert into dbo.History values('{login}', 'opening the browser')")
        commands_eng.opening_the_browser(mongo_collect, login)
        cursor.commit()
    elif "calculate" in text:
        cursor.execute(f"Insert into dbo.History values('{login}', 'using the calculator')")
        commands_eng.calculator(mongo_collect, login)
        cursor.commit()
    elif "what can you do" in text:
        cursor.execute(f"Insert into dbo.History values('{login}', 'checking possibility of voice assistant')")
        commands_eng.capability(mongo_collect, login)
        cursor.commit()
    elif "search" in text:
        cursor.execute(f"Insert into dbo.History values('{login}', 'using Windows searching')")
        commands_eng.search(mongo_collect, login)
        cursor.commit()
    elif "language" in text:
        cursor.execute(f"Insert into dbo.History values('{login}', 'changing the language')")
        commands_eng.change_keyboard_lang(mongo_collect, login)
        cursor.commit()
    elif "volume" in text:
        if "up" in text or "high" in text or "more" in text:
            cursor.execute(f"Insert into dbo.History values('{login}', 'volume up')")
            commands_eng.volume("up", mongo_collect, login)
            cursor.commit()
        else:
            cursor.execute(f"Insert into dbo.History values('{login}', 'volume down')")
            commands_eng.volume("down", mongo_collect, login)
            cursor.commit()
    elif "task manager" in text:
        cursor.execute(f"Insert into dbo.History values('{login}', 'opening task manager')")
        commands_eng.task_manager(mongo_collect, login)
        cursor.commit()
    elif "command" in text:
        cursor.execute(f"Insert into dbo.History values('{login}', 'checking commands')")
        commands_eng.check_command(mongo_collect, login)
        cursor.commit()
    elif "game" in text:
        cursor.execute(f"Insert into dbo.History values('{login}', 'opening the game')")
        tic_tac_toe_eng(mongo_collect, login)
        cursor.commit()
    elif "history" in text:
        cursor.execute(f"Insert into dbo.History values('{login}', 'checking history')")
        commands_eng.history(login, cursor, mongo_collect)
        cursor.commit()
    elif "money" in text:
        cursor.execute(f"Insert into dbo.History values('{login}', 'checking money')")
        commands_eng.money(mongo_collect, login)
        cursor.commit()
    else:
        speak_eng(f"Sorry {login},I can't do this or I couldn't hear you properly, please say something else or repeat")
    if check_command == "":
        introductory_talk(conn)

