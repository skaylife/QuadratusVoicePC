import tkinter as tk
from tkinter import ttk, messagebox
import threading
import queue
import time
import json
import os
import sys
import zipfile
import requests
import numpy as np
import pyperclip
from datetime import datetime

# ОПРЕДЕЛЕНИЕ ПУТЕЙ (Самый важный блок для EXE)
if getattr(sys, 'frozen', False):
    # Если запущено из EXE
    BASE_DIR = os.path.dirname(sys.executable)
else:
    # Если запущен просто .py файл
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Попытки импорта зависимостей
try:
    import pyaudio
except ImportError:
    pyaudio = None

try:
    import win32gui, win32process, psutil, keyboard
except ImportError:
    keyboard = None

# ИСПРАВЛЕННЫЙ ИМПОРТ VOSK
try:
    from vosk import Model, KaldiRecognizer
    VOSK_AVAILABLE = True
except Exception as e:
    VOSK_AVAILABLE = False
    VOSK_ERROR_MSG = str(e)

# Константы (используем BASE_DIR)
SETTINGS_FILE = os.path.join(BASE_DIR, "voice_settings.json")
LOGS_DIR = os.path.join(BASE_DIR, "voice_logs")
MODELS_DIR = os.path.join(BASE_DIR, "models")
CHUNK = 1024
RATE = 16000
FORMAT = pyaudio.paInt16 if pyaudio else None
CHANNELS = 1

MODEL_LINKS = {
    "Маленькая (50 МБ, быстрая)": "https://alphacephei.com/vosk/models/vosk-model-small-ru-0.22.zip",
    "Большая (1.8 ГБ, точная)": "https://alphacephei.com/vosk/models/vosk-model-ru-0.42.zip"
}

audio_queue = queue.Queue()
recognizing = False
running = False
vosk_recognizer = None
current_audio_level = 0.0
current_session_text = ""
session_start_time = None

default_settings = {
    "vosk_model_path": "",
    "auto_insert_mode": True,
    "insert_mode": "paste",
    "show_overlay_when_no_focus": True,
    "save_session_logs": True,
    "auto_copy_to_clipboard": True,
    "selected_microphone": 0,
    "sensitivity": 0.5,
    "wave_visualization": True,
    "language": "ru",
    "input_mode": "paste"
}

settings = default_settings.copy()

def load_settings():
    global settings
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                settings.update(json.load(f))
        except: pass
    else:
        save_settings()

def save_settings():
    try:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(settings, f, indent=2, ensure_ascii=False)
    except: pass

def save_session_log():
    global current_session_text
    if not settings.get("save_session_logs") or not current_session_text.strip(): return
    if not os.path.exists(LOGS_DIR): os.makedirs(LOGS_DIR)
    log_file = os.path.join(LOGS_DIR, f"session_{datetime.now().strftime('%Y-%m-%d')}.txt")
    with open(log_file, "a", encoding="utf-8") as f:
        f.write(f"\n=== {datetime.now().strftime('%H:%M:%S')} ===\n" + current_session_text + "\n")
    current_session_text = ""

def is_text_input_active():
    try:
        hwnd = win32gui.GetForegroundWindow()
        if hwnd == 0: return False
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        p_name = psutil.Process(pid).name().lower()
        return any(app in p_name for app in ["notepad.exe", "code.exe", "chrome.exe", "telegram.exe", "word.exe"])
    except: return False

def calculate_audio_level(data):
    try:
        audio_data = np.frombuffer(data, dtype=np.int16)
        if len(audio_data) == 0: return 0.0
        m_sq = np.mean(np.square(audio_data.astype(np.float64)))
        return np.sqrt(m_sq) / 5000.0 if m_sq > 0 else 0.0
    except: return 0.0

def audio_capture_thread():
    global running, current_audio_level
    if not pyaudio: return
    try:
        pa = pyaudio.PyAudio()
        stream = pa.open(format=FORMAT, channels=CHANNELS, rate=RATE, input=True,
                        input_device_index=settings.get("selected_microphone", 0), frames_per_buffer=CHUNK)
        while running:
            try:
                data = stream.read(CHUNK, exception_on_overflow=False)
                current_audio_level = calculate_audio_level(data)
                audio_queue.put(data)
            except: break
        stream.stop_stream(); stream.close(); pa.terminate()
    except: pass

def speech_processing_thread(text_callback):
    global recognizing, running, vosk_recognizer
    if not VOSK_AVAILABLE:
        messagebox.showerror("Ошибка", f"Библиотека Vosk не найдена в EXE!\n{VOSK_ERROR_MSG}")
        return

    path = settings.get("vosk_model_path", "")
    if not path or not os.path.exists(path): return
    try:
        model = Model(path)
        vosk_recognizer = KaldiRecognizer(model, RATE)
        while running:
            try:
                data = audio_queue.get(timeout=1)
            except queue.Empty: continue
            if not recognizing: continue
            if vosk_recognizer.AcceptWaveform(data):
                res = json.loads(vosk_recognizer.Result())
                txt = res.get("text", "").strip()
                if txt: text_callback(txt, final=True)
            else:
                part = json.loads(vosk_recognizer.PartialResult())
                txt = part.get("partial", "").strip()
                if txt: text_callback(txt, final=False)
    except Exception as e:
        messagebox.showerror("Ошибка модели", f"Путь: {path}\nОшибка: {e}")

class VoiceAssistantApp:
    def __init__(self):
        load_settings()
        self.root = tk.Tk()
        self.setup_window()
        self.create_widgets()
        self.setup_drag()
        self.overlay_window = None
        self.drag_data = {"x": 0, "y": 0}

    def setup_window(self):
        self.root.title("🎤 Voice Assistant")
        self.root.geometry("100x100")
        self.root.overrideredirect(True)
        self.root.attributes('-topmost', True)
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - 50
        y = (self.root.winfo_screenheight() // 2) - 50
        self.root.geometry(f"+{x}+{y}")
        self.main_frame = tk.Frame(self.root, bg="#202020", relief="flat", bd=0)
        self.main_frame.pack(fill="both", expand=True)

    def create_widgets(self):
        self.mic_button = tk.Button(self.main_frame, text="🎤", font=("Segoe UI Emoji", 24),
            bg="#333333", fg="#ff4444", relief="flat", bd=0, command=self.toggle_recording)
        self.mic_button.place(x=25, y=25, width=50, height=50)
        self.settings_button = tk.Button(self.main_frame, text="⚙", font=("Segoe UI Emoji", 8),
            bg="#202020", fg="#888888", relief="flat", bd=0, command=self.open_settings)
        self.settings_button.place(x=5, y=5, width=15, height=15)
        self.close_button = tk.Button(self.main_frame, text="✕", font=("Segoe UI", 8),
            bg="#202020", fg="#ff4444", relief="flat", bd=0, command=self.close_app)
        self.close_button.place(x=80, y=5, width=15, height=15)
        self.audio_indicator = tk.Frame(self.main_frame, bg="#333333", height=3, width=60)
        self.audio_indicator.place(x=20, y=85)

    def setup_drag(self):
        def start(e): self.drag_data = {"x": e.x, "y": e.y}
        def move(e):
            nx, ny = self.root.winfo_x() + (e.x - self.drag_data["x"]), self.root.winfo_y() + (e.y - self.drag_data["y"])
            self.root.geometry(f"+{nx}+{ny}")
        self.main_frame.bind("<Button-1>", start); self.main_frame.bind("<B1-Motion>", move)

    def toggle_recording(self):
        global recognizing, running, audio_thread, processing_thread
        path = settings.get("vosk_model_path", "")
        if not path or not os.path.exists(path):
            self.show_model_downloader()
            return
        if not running:
            running = True; recognizing = True
            self.mic_button.config(text="🎙", bg="#ff4444", fg="#ffffff")
            audio_thread = threading.Thread(target=audio_capture_thread, daemon=True)
            processing_thread = threading.Thread(target=speech_processing_thread, args=(self.on_recognized_text,), daemon=True)
            audio_thread.start(); processing_thread.start(); self.start_visualization()
        else:
            if recognizing:
                save_session_log(); recognizing = False
                self.mic_button.config(text="🎤", bg="#333333", fg="#ff4444"); self.close_overlay()
            else:
                recognizing = True; self.mic_button.config(text="🎙", bg="#ff4444", fg="#ffffff")

    def show_model_downloader(self):
        win = tk.Toplevel(self.root); win.title("Модели"); win.geometry("380x350"); win.configure(bg="#212121"); win.attributes('-topmost', True)
        tk.Label(win, text="Модель не найдена!", fg="#ff4444", bg="#212121", font=("Segoe UI", 12, "bold")).pack(pady=10)
        if not os.path.exists(MODELS_DIR): os.makedirs(MODELS_DIR)
        local = [d for d in os.listdir(MODELS_DIR) if os.path.isdir(os.path.join(MODELS_DIR, d))]
        if local:
            tk.Label(win, text="Уже скачанные:", fg="white", bg="#212121").pack()
            box = ttk.Combobox(win, values=local, state="readonly"); box.pack(pady=5)
            def pick():
                settings["vosk_model_path"] = os.path.abspath(os.path.join(MODELS_DIR, box.get()))
                save_settings(); win.destroy(); messagebox.showinfo("Успех", "Выбрано!")
            tk.Button(win, text="Использовать эту", command=pick).pack(pady=5)
        for name, url in MODEL_LINKS.items():
            btn = tk.Button(win, text=name, bg="#444444", fg="white", relief="flat", command=lambda u=url: self.download_task(u, win))
            btn.pack(fill="x", padx=40, pady=3)

    def download_task(self, url, parent):
        lbl = tk.Label(parent, text="Скачивание... ждите", fg="yellow", bg="#212121"); lbl.pack(pady=10)
        def run():
            try:
                r = requests.get(url, stream=True); z_path = os.path.join(MODELS_DIR, "temp.zip")
                with open(z_path, "wb") as f:
                    for ch in r.iter_content(chunk_size=8192): f.write(ch)
                with zipfile.ZipFile(z_path, 'r') as z:
                    folder = z.namelist()[0].split('/')[0]; z.extractall(MODELS_DIR)
                os.remove(z_path); settings["vosk_model_path"] = os.path.abspath(os.path.join(MODELS_DIR, folder))
                save_settings(); self.root.after(0, lambda: [parent.destroy(), messagebox.showinfo("Успех", "Готово!")])
            except Exception as e: self.root.after(0, lambda: messagebox.showerror("Ошибка", str(e)))
        threading.Thread(target=run, daemon=True).start()

    def on_recognized_text(self, text, final):
        global current_session_text
        if final:
            current_session_text += text + " "
            if settings.get("auto_copy_to_clipboard"): pyperclip.copy(current_session_text.strip())
            if is_text_input_active() and settings.get("auto_insert_mode"):
                if settings.get("insert_mode") == "paste":
                    pyperclip.copy(text + " "); time.sleep(0.05); keyboard.press_and_release('ctrl+v')
                else: keyboard.write(text + " ")
            else:
                if settings.get("show_overlay_when_no_focus"): self.show_overlay(current_session_text)
        else:
            if not is_text_input_active() and settings.get("show_overlay_when_no_focus"):
                self.show_overlay(current_session_text + f"🎤 {text}")

    def show_overlay(self, text):
        if not self.overlay_window or not self.overlay_window.winfo_exists():
            self.overlay_window = tk.Toplevel(self.root); self.overlay_window.geometry("400x200"); self.overlay_window.attributes("-topmost", True); self.overlay_window.configure(bg="#212121")
            self.ov_text = tk.Text(self.overlay_window, bg="#1a1a1a", fg="white", font=("Segoe UI", 10), wrap="word", bd=0); self.ov_text.pack(fill="both", expand=True, padx=10, pady=10)
            tk.Button(self.overlay_window, text="✕", bg="#ff4444", fg="white", command=self.close_overlay).pack(pady=5)
        self.ov_text.delete(1.0, tk.END); self.ov_text.insert(tk.END, text); self.overlay_window.deiconify()

    def close_overlay(self):
        if self.overlay_window and self.overlay_window.winfo_exists(): self.overlay_window.withdraw()

    def start_visualization(self):
        def update():
            if recognizing and settings.get("wave_visualization"):
                color = "#00ff44" if current_audio_level > 0.1 else "#333333"
                self.audio_indicator.config(bg=color, width=int(20 + (current_audio_level * 40)))
            if running: self.root.after(100, update)
        update()

    def open_settings(self):
        win = tk.Toplevel(self.root); win.title("Настройки"); win.geometry("350x400"); win.configure(bg="#212121"); win.attributes("-topmost", True)
        tk.Label(win, text="Настройки", fg="white", bg="#212121", font=("Segoe UI", 12, "bold")).pack(pady=10)
        vars = {}
        for t, k in [("Авто-вставка", "auto_insert_mode"), ("Оверлей", "show_overlay_when_no_focus"), ("Логи", "save_session_logs"), ("Волна", "wave_visualization")]:
            v = tk.BooleanVar(value=settings.get(k, True))
            tk.Checkbutton(win, text=t, variable=v, bg="#212121", fg="white", selectcolor="#444444").pack(anchor="w", padx=20)
            vars[k] = v
        iv = tk.StringVar(value=settings.get("insert_mode", "paste"))
        tk.Radiobutton(win, text="Ctrl+V", variable=iv, value="paste", bg="#212121", fg="white").pack(anchor="w", padx=20)
        tk.Radiobutton(win, text="Печать", variable=iv, value="keyboard", bg="#212121", fg="white").pack(anchor="w", padx=20)
        def save():
            for k, v in vars.items(): settings[k] = v.get()
            settings["insert_mode"] = iv.get(); save_settings(); win.destroy()
        tk.Button(win, text="Сменить модель", command=self.show_model_downloader).pack(pady=10)
        tk.Button(win, text="💾 Сохранить", command=save, bg="#0078d4", fg="white").pack(fill="x", padx=20)

    def close_app(self):
        global running; running = False; self.root.destroy()

if __name__ == "__main__":
    app = VoiceAssistantApp()
    app.root.mainloop()