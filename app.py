import queue
import threading
import sounddevice as sd
import vosk
import json
from docx import Document
from rus2num import Rus2Num
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import sys

# ================== ПУТИ ==================

BASE_DIR = os.path.dirname(
    sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(__file__)
)
MODEL_PATH = os.path.join(BASE_DIR, "vosk-model-small-ru-0.22")

# ================== ГЛОБАЛЬНЫЕ ==================

q = queue.Queue()
running = False
doc_path = None

# ================== VOSK ==================

model = vosk.Model(MODEL_PATH)
rec = vosk.KaldiRecognizer(model, 16000)

# ================== АУДИО ==================

def callback(indata, frames, time, status):
    if rec.AcceptWaveform(bytes(indata)):
        q.put(json.loads(rec.Result()))

# ================== ЛОГИКА ==================

def normalize_value(text):
    if not text:
        return text
    return text[0].upper() + text[1:]

def insert_value(spoken_text):
    if not doc_path:
        return

    doc = Document(doc_path)

    parts = spoken_text.split(maxsplit=1)
    if len(parts) < 2:
        return

    field_word, value_text = parts

    r2n = Rus2Num()
    try:
        value = r2n(value_text)
    except Exception:
        value = value_text

    value = normalize_value(str(value))


    for para in doc.paragraphs:
        if field_word in para.text.lower():
            if ":" in para.text:
                left, _ = para.text.split(":", 1)
                para.text = f"{left}: {value}"
            else:
                para.text += f" {value}"

            doc.save(doc_path)
            print(f"Обновлено: {field_word} → {value}")
            return

    print(f"Поле '{field_word}' не найдено")

# ================== ОБРАБОТКА РЕЧИ ==================

def process_spoken(text):
    text = text.lower().strip()

    if text in ["стоп", "выход", "закончить"]:
        stop_listening()
        return

    if text == "диктант включить":
        dictation_mode.set(True)
        print("Медицинский диктант ВКЛ")
        return

    if text == "диктант выключить":
        dictation_mode.set(False)
        print("Медицинский диктант ВЫКЛ")
        return

    if dictation_mode.get():
        if not text.startswith("запись"):
            return
        text = text.replace("запись", "", 1).strip()

    insert_value(text)

# ================== ЦИКЛ ПРОСЛУШИВАНИЯ ==================

def listen_loop():
    global running

    with sd.RawInputStream(
        samplerate=16000,
        blocksize=8000,
        dtype='int16',
        channels=1,
        callback=callback
    ):
        print("Прослушивание начато")

        while running:
            result = q.get()
            spoken = result.get("text", "")
            if spoken:
                print("Вы сказали:", spoken)
                process_spoken(spoken)

        print("Прослушивание остановлено")

# ================== УПРАВЛЕНИЕ ==================

def start_listening():
    global running

    if not doc_path:
        messagebox.showwarning("Ошибка", "Сначала выберите документ")
        return

    if not running:
        running = True
        threading.Thread(target=listen_loop, daemon=True).start()

def stop_listening():
    global running
    running = False

def choose_file():
    global doc_path
    path = filedialog.askopenfilename(
        filetypes=[("Word documents", "*.docx")]
    )
    if path:
        doc_path = path
        file_label.config(text=os.path.basename(path))

def toggle_dictation(event=None):
    dictation_mode.set(not dictation_mode.get())
    print("Диктант:", "ВКЛ" if dictation_mode.get() else "ВЫКЛ")

# ================== GUI ==================

root = tk.Tk()
root.title("Голосовое заполнение документа")
root.geometry("430x260")

tk.Button(root, text="Выбрать документ", command=choose_file).pack(pady=10)

file_label = tk.Label(root, text="Файл не выбран")
file_label.pack()

dictation_mode = tk.BooleanVar(value=False)

tk.Checkbutton(
    root,
    text="Медицинский диктант (через слово «запись»)",
    variable=dictation_mode
).pack(pady=5)

tk.Button(root, text="▶ Запуск (F9)", command=start_listening).pack(pady=10)
tk.Button(root, text="■ Стоп (F10)", command=stop_listening).pack()

# ================== ГОРЯЧИЕ КЛАВИШИ ==================

root.bind("<F9>", lambda e: start_listening())
root.bind("<F10>", lambda e: stop_listening())
root.bind("<F8>", toggle_dictation)
root.bind("<Escape>", lambda e: stop_listening())

root.mainloop()
