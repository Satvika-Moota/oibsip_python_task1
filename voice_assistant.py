import speech_recognition as sr
import datetime
import webbrowser
import threading
import tkinter as tk
from tkinter import scrolledtext
import time
import os

# Speech Engine Setup
try:
    import win32com.client
    USE_WIN32 = True
except ImportError:
    import pyttsx3
    USE_WIN32 = False

class VoiceAssistantGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("NOVA | AI Voice Assistant")
        self.root.geometry("900x750")
        
        # --- PROFESSIONAL COLOR PALETTE ---
        self.colors = {
            "bg": "#0F172A",          
            "card": "#1E293B",        
            "accent": "#38BDF8",      
            "secondary": "#94A3B8",   
            "user_text": "#F0F9FF",   
            "bot_text": "#38BDF8",    
            "status_ok": "#10B981",   
            "status_err": "#EF4444"    
        }

        self.root.configure(bg=self.colors["bg"])
        
        # Center window
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
        
        self.recognizer = sr.Recognizer()
        self.is_listening = False
        self.use_win32 = USE_WIN32
        
        self.setup_ui()
        
        # Welcome message
        self.root.after(500, lambda: self.display_and_speak("System online. Hello! I'm NOVA. How can I assist you today?", "assistant"))
    
    def speak(self, text):
        def speak_thread():
            try:
                if self.use_win32:
                    import pythoncom
                    pythoncom.CoInitialize()
                    try:
                        speaker = win32com.client.Dispatch("SAPI.SpVoice")
                        speaker.Rate = 1
                        speaker.Volume = 100
                        speaker.Speak(text)
                    finally:
                        pythoncom.CoUninitialize()
                else:
                    engine = pyttsx3.init()
                    engine.setProperty('rate', 170)
                    engine.say(text)
                    engine.runAndWait()
            except Exception as e:
                print(f"Speech error: {e}")
        
        threading.Thread(target=speak_thread, daemon=True).start()
    
    def setup_ui(self):
        # --- HEADER ---
        header = tk.Frame(self.root, bg=self.colors["bg"], height=100)
        header.pack(fill=tk.X, pady=20)

        title = tk.Label(
            header,
            text="NOVA",
            font=("Helvetica", 36, "bold"),
            fg=self.colors["accent"],
            bg=self.colors["bg"]
        )
        title.pack()

        subtitle = tk.Label(
            header,
            text="VIRTUAL ASSISTANT INTERFACE",
            font=("Helvetica", 10, "bold"),
            fg=self.colors["secondary"],
            bg=self.colors["bg"]
        )
        subtitle.pack()

        # --- STATUS BAR ---
        self.status_label = tk.Label(
            self.root,
            text="● SYSTEM READY",
            font=("Consolas", 11, "bold"),
            fg=self.colors["status_ok"],
            bg=self.colors["bg"]
        )
        self.status_label.pack(pady=5)

        # --- CHAT INTERFACE ---
        chat_frame = tk.Frame(self.root, bg=self.colors["card"], highlightthickness=1, highlightbackground="#334155")
        chat_frame.pack(padx=40, pady=10, fill=tk.BOTH, expand=True)

        self.chat_display = scrolledtext.ScrolledText(
            chat_frame,
            wrap=tk.WORD,
            font=("Segoe UI", 11),
            bg=self.colors["card"],
            fg=self.colors["user_text"],
            insertbackground=self.colors["accent"],
            relief=tk.FLAT,
            padx=20,
            pady=20,
            borderwidth=0
        )
        self.chat_display.pack(fill=tk.BOTH, expand=True)
        self.chat_display.config(state=tk.DISABLED)

        self.chat_display.tag_config("user", foreground=self.colors["user_text"], font=("Segoe UI", 11, "bold"))
        self.chat_display.tag_config("assistant", foreground=self.colors["bot_text"], font=("Segoe UI", 11))
        self.chat_display.tag_config("timestamp", foreground="#475569", font=("Consolas", 9))

        # --- BUTTON PANEL ---
        btn_panel = tk.Frame(self.root, bg=self.colors["bg"])
        btn_panel.pack(pady=30)

        button_style = {
            "font": ("Helvetica", 12, "bold"),
            "bg": self.colors["accent"],
            "fg": self.colors["bg"],
            "activebackground": "#7DD3FC",
            "relief": tk.FLAT,
            "padx": 30,
            "pady": 12,
            "cursor": "hand2"
        }

        self.listen_btn = tk.Button(
            btn_panel,
            text="START LISTENING",
            command=self.toggle_listening,
            **button_style
        )
        self.listen_btn.pack(side=tk.LEFT, padx=15)

        clear_btn = tk.Button(
            btn_panel,
            text="CLEAR LOGS",
            command=self.clear_chat,
            font=("Helvetica", 10, "bold"),
            bg="#334155",
            fg="white",
            relief=tk.FLAT,
            padx=20,
            pady=12,
            cursor="hand2"
        )
        clear_btn.pack(side=tk.LEFT, padx=15)

    def display_message(self, message, sender="assistant"):
        self.chat_display.config(state=tk.NORMAL)
        ts = datetime.datetime.now().strftime("%H:%M")
        
        if sender == "user":
            self.chat_display.insert(tk.END, f"\n[{ts}] ", "timestamp")
            self.chat_display.insert(tk.END, "USER: ", "user")
            self.chat_display.insert(tk.END, f"{message.upper()}\n")
        else:
            self.chat_display.insert(tk.END, f"\n[{ts}] ", "timestamp")
            self.chat_display.insert(tk.END, "NOVA: ", "assistant")
            self.chat_display.insert(tk.END, f"{message}\n")
        
        self.chat_display.see(tk.END)
        self.chat_display.config(state=tk.DISABLED)

    def display_and_speak(self, message, sender="assistant"):
        self.display_message(message, sender)
        if sender == "assistant":
            self.speak(message)

    def update_status(self, status, color):
        self.status_label.config(text=f"● {status}", fg=color)

    def toggle_listening(self):
        if not self.is_listening:
            self.is_listening = True
            self.listen_btn.config(text="STOP LISTENING", bg=self.colors["status_err"])
            self.update_status("MIC ACTIVE / LISTENING", self.colors["accent"])
            threading.Thread(target=self.listen_continuous, daemon=True).start()
        else:
            self.is_listening = False
            self.listen_btn.config(text="START LISTENING", bg=self.colors["accent"])
            self.update_status("SYSTEM READY", self.colors["status_ok"])

    def listen_continuous(self):
        while self.is_listening:
            command = self.listen()
            if command:
                self.process_command(command)
            else:
                time.sleep(0.5)

    def listen(self):
        try:
            with sr.Microphone() as source:
                self.recognizer.adjust_for_ambient_noise(source, duration=0.3)
                audio = self.recognizer.listen(source, timeout=5, phrase_time_limit=10)
                self.update_status("DECODING SPEECH...", "#A855F7")
                command = self.recognizer.recognize_google(audio).lower()
                self.display_message(command, "user")
                return command
        except:
            self.update_status("MIC ACTIVE / LISTENING", self.colors["accent"])
            return None

    def process_command(self, command):
        if "your name" in command or "who are you" in command:
            self.display_and_speak("I am a voice assistant and my boss who created me is Satvika.")
        
        elif "favorite color" in command or "favourite color" in command:
            self.display_and_speak("I don't have eyes to see color, but based on my code, I find electric blue quite pleasing!")

        elif any(word in command for word in ["hello", "hi", "hey"]):
            self.display_and_speak("Greetings. System is ready. How can I help?")
            
        elif any(word in command for word in ["exit", "quit", "bye", "stop"]):
            self.display_and_speak("Powering down. Goodbye!")
            if self.is_listening:
                self.toggle_listening()
        
        elif "time" in command:
            current_time = datetime.datetime.now().strftime("%I:%M %p")
            self.display_and_speak(f"The current time is {current_time}")
        
        elif "date" in command or "today" in command:
            current_date = datetime.datetime.now().strftime("%B %d, %Y")
            self.display_and_speak(f"Today's date is {current_date}")
        
        elif "search" in command:
            query = command.replace("search for", "").replace("search", "").strip()
            url = f"https://www.google.com/search?q={query}"
            webbrowser.open(url)
            self.display_and_speak(f"Accessing web for {query}")
        
        else:
            self.display_and_speak("Command not recognized. Please try again.")
        
        if self.is_listening:
            self.update_status("MIC ACTIVE / LISTENING", self.colors["accent"])

    def clear_chat(self):
        self.chat_display.config(state=tk.NORMAL)
        self.chat_display.delete(1.0, tk.END)
        self.chat_display.config(state=tk.DISABLED)
        self.display_message("Communication logs cleared.", "assistant")

def main():
    root = tk.Tk()
    app = VoiceAssistantGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()