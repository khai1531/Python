import tkinter as tk
from customtkinter import CTkTextbox, CTkEntry, CTkButton, CTkFrame
from tkinter import filedialog, Menu
import pyttsx3
import datetime
import speech_recognition as sr
import webbrowser as wb
import time
import threading
import wikipedia
import win32api
import win32con
import os
import sys

# Initialize the speech engine
bot = pyttsx3.init()
voices = bot.getProperty('voices')
bot.setProperty('voice', voices[0].id)

# Speak function to convert text to speech
def speak(audio):
    chat_log("B.BOT: " + audio)
    bot.say(audio)
    bot.runAndWait()

# Get the current time
def timenow():
    Time = datetime.datetime.now().strftime('%I:%M %p')
    speak(Time)

# Welcome user
def welcome():
    hour = datetime.datetime.now().hour
    if 6 <= hour < 12:
        speak('Good morning Sir')
    elif 12 <= hour < 18:
        speak('Good afternoon Sir')
    else:
        speak('Good evening Sir')
    speak('How can I assist you?')
    speak('Press listen to talk or type your command')

# Listen for commands
def command():
    c = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        c.pause_threshold = 2
        try:
            audio = c.listen(source, timeout=5.0)
        except sr.WaitTimeoutError:
            speak("Listening timed out, please try again.")
            return ""
    try:
        query = c.recognize_google(audio, language='en')
        chat_log('Boss: ' + query)
    except sr.UnknownValueError:
        speak('Please repeat or type the command')
        query = str(input('Your request is: '))
    return query

# Process user query
def process_query(query):
    query = query.lower()
    response = "This is a response to your query: " + query
    
    # Display the response in the text area
    chat_log(response)

    if 'google' in query:
        speak('What should I search for, sir?')
        search = command().lower()
        url = f'https://www.google.com/search?q={search}'
        wb.open(url)
        speak(f'Here is your search on Google for {search}.')
    elif 'youtube' in query:
        speak('What should I search for, sir?')
        search = command().lower()
        url = f'https://www.youtube.com/results?search_query={search}'
        wb.open(url)
        speak(f'Here is your search on YouTube for {search}.')
    elif 'music' in query:
        url = f'spotify:search:{query}'
        wb.open(url)
        speak(f'Playing music on Spotify.')
        time.sleep(3)
        win32api.keybd_event(0x20, 0, 0, 0)  # Press space bar
        time.sleep(0.05)
        win32api.keybd_event(0x20, 0, win32con.KEYEVENTF_KEYUP, 0)  # Release space bar
    elif 'time' in query:
        speak('It is:')
        timenow()
    elif 'quit' in query or 'goodbye' in query:
        speak('Goodbye Sir.')
        window.quit()
    else:
        try:
            wikipedia.set_lang('en')
            robot = wikipedia.summary(query, sentences=1)
            speak(robot)
        except wikipedia.exceptions.PageError:
            speak('Sorry, I can not find the result.')
        except wikipedia.exceptions.DisambiguationError:
            speak('Sorry, I can not find the result.')
        else:
            speak('I did not understand your command.')

# Function for listen button
def listen_command():
    threading.Thread(target=lambda: process_query(command())).start()

# Handle user input in the GUI
def send_message():
    user_input = entry_command.get()
    chat_log("You: " + user_input)
    entry_command.delete(0, tk.END)
    process_query(user_input)

# Function to log chat history
def chat_log(message):
    with open("chat_history.txt", "a", encoding="utf-8") as f:
        f.write(message + "\n")
    chat_history.insert(tk.END, message + "\n")
    chat_history.yview(tk.END)
    window.update_idletasks()

# Function to save chat history
def save_chat():
    file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
    if file_path:
        with open(file_path, "w", encoding="utf-8") as f:
            f.write(chat_history.get(1.0, tk.END))

# Function to handle resource path
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# Create main window
window = tk.Tk()
window.title('ChatBot')
window.geometry('750x700')

# Create menu
main_menu = Menu(window)
file_menu = Menu(main_menu, tearoff=0)
file_menu.add_command(label="Save Chat", command=save_chat)
file_menu.add_separator()
file_menu.add_command(label="Exit", command=window.quit)
main_menu.add_cascade(label="File", menu=file_menu)
window.config(menu=main_menu)

# Main frame
frame = CTkFrame(window)
frame.pack(pady=10)

# Chat history area
chat_history = CTkTextbox(frame, fg_color='#E0E0E0', height=400, width=700, corner_radius=10)
chat_history.pack(pady=10)

# Command entry
entry_command = CTkEntry(frame, fg_color='#E0E0E0', height=40, width=500, corner_radius=10)
entry_command.pack(side=tk.LEFT, padx=(10, 5), pady=10)

# Send button
button_send = CTkButton(frame, text='SEND', command=send_message, height=40, width=100)
button_send.pack(side=tk.LEFT, padx=5, pady=10)

# Listen button
button_listen = CTkButton(frame, text='LISTEN', command=listen_command, height=40, width=100)
button_listen.pack(side=tk.LEFT, padx=5, pady=10)

# Start welcome in a separate thread
threading.Thread(target=welcome).start()

# Start main loop
window.mainloop()
