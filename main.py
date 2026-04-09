import tkinter as tk
import threading
import re
import pythoncom
import win32com.client
import speech_recognition as sr
from custom_classes import Calculator

calculator = Calculator()
expression = ""
speech_lock = threading.Lock()


def value_to_word(value):
    mapping = {
        "0": "Zero",
        "1": "One",
        "2": "Two",
        "3": "Three",
        "4": "Four",
        "5": "Five",
        "6": "Six",
        "7": "Seven",
        "8": "Eight",
        "9": "Nine",
        ".": "Point",
        "÷": "Over",
        "x": "Times",
        "+": "Plus",
        "-": "Minus"
    }
    return mapping.get(value, value)


def _speak_worker(text):
    with speech_lock:
        pythoncom.CoInitialize()
        try:
            speaker = win32com.client.Dispatch("SAPI.SpVoice")
            speaker.Speak(text)
        except Exception:
            pass
        finally:
            pythoncom.CoUninitialize()


def speak(text):
    threading.Thread(target=_speak_worker, args=(text,), daemon=True).start()


def press(value, speak_out=True):
    global expression
    expression += str(value)
    entry_var.set(expression)

    if speak_out:
        speak(value_to_word(str(value)))


def clear():
    global expression
    expression = ""
    entry_var.set("")
    status_var.set("Ready")


def calculate_result(speak_result=True):
    global expression

    try:
        expr = expression.replace("÷", "/").replace("x", "*")
        result = eval(expr)

        if isinstance(result, float) and result.is_integer():
            result = int(result)

        expression = str(result)
        entry_var.set(expression)

        if speak_result:
            speak(f"is {result}")

    except ZeroDivisionError:
        expression = ""
        entry_var.set("Error: division by zero")
        if speak_result:
            speak("Error")
    except Exception:
        expression = ""
        entry_var.set("Error")
        if speak_result:
            speak("Error")


def normalize_voice_text(text):
    text = text.lower().strip()

    replacements = [
        ("multiplied by", " x "),
        ("multiply by", " x "),
        ("times", " x "),
        ("divided by", " ÷ "),
        ("divide by", " ÷ "),
        ("over", " ÷ "),
        ("slash", " ÷ "),
        ("plus", " + "),
        ("minus", " - "),
        ("point", " . "),
        ("equals", " = "),
        ("equal", " = "),
        ("is", " = "),
        ("zero", " 0 "),
        ("one", " 1 "),
        ("two", " 2 "),
        ("three", " 3 "),
        ("four", " 4 "),
        ("five", " 5 "),
        ("six", " 6 "),
        ("seven", " 7 "),
        ("eight", " 8 "),
        ("nine", " 9 "),
    ]

    for old, new in replacements:
        text = text.replace(old, new)

    # bazen recognizer bunları verir
    text = text.replace("*", " x ")
    text = text.replace("/", " ÷ ")
    text = text.replace("×", " x ")

    text = re.sub(r"\s+", " ", text).strip()
    return text


def extract_expression(text):
    text = normalize_voice_text(text)

    allowed = set("0123456789+-x÷.= ")
    filtered = "".join(ch for ch in text if ch in allowed)

    filtered = re.sub(r"\s+", " ", filtered).strip()
    tokens = filtered.split()

    expr_tokens = []
    for token in tokens:
        if token in {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9", ".", "+", "-", "x", "÷", "="}:
            expr_tokens.append(token)

    return expr_tokens


def process_voice_command(text):
    global expression

    tokens = extract_expression(text)

    expression = ""
    entry_var.set("")

    if not tokens:
        status_var.set("Could not parse expression")
        speak("I did not understand")
        return

    for token in tokens:
        if token == "=":
            calculate_result(speak_result=False)
            break
        else:
            press(token, speak_out=False)

    result_text = entry_var.get()
    if result_text.startswith("Error"):
        speak("Error")
    elif result_text:
        speak(f"is {result_text}")


def listen_voice():
    recognizer = sr.Recognizer()
    recognizer.dynamic_energy_threshold = True
    recognizer.pause_threshold = 0.6

    try:
        with sr.Microphone() as source:
            status_var.set("Listening...")
            root.update()

            recognizer.adjust_for_ambient_noise(source, duration=0.5)
            audio = recognizer.listen(source, timeout=5, phrase_time_limit=4)

        text = recognizer.recognize_google(audio, language="en-US")
        status_var.set(f"You said: {text}")
        process_voice_command(text)

    except sr.WaitTimeoutError:
        status_var.set("Listening timed out")
        speak("No speech detected")
    except sr.UnknownValueError:
        status_var.set("Could not understand audio")
        speak("I did not understand")
    except sr.RequestError:
        status_var.set("Speech service error")
        speak("Speech service error")
    except Exception:
        status_var.set("Microphone error")
        speak("Microphone error")


def start_listening():
    threading.Thread(target=listen_voice, daemon=True).start()


root = tk.Tk()
root.title("Calculator")
root.geometry("350x560")
root.resizable(False, False)

entry_var = tk.StringVar()
status_var = tk.StringVar()
status_var.set("Ready")

entry = tk.Entry(
    root,
    textvariable=entry_var,
    font=("Arial", 24),
    bd=10,
    relief="sunken",
    justify="right"
)
entry.pack(fill="both", ipadx=8, ipady=15, padx=10, pady=10)

mic_button = tk.Button(
    root,
    text="🎤 Microphone",
    font=("Arial", 16),
    command=start_listening
)
mic_button.pack(fill="x", padx=10, pady=5)

status_label = tk.Label(
    root,
    textvariable=status_var,
    font=("Arial", 12)
)
status_label.pack(pady=5)

button_frame = tk.Frame(root)
button_frame.pack(expand=True, fill="both")

buttons = [
    ["7", "8", "9", "÷"],
    ["4", "5", "6", "x"],
    ["1", "2", "3", "-"],
    ["0", ".", "C", "+"],
    ["="]
]

for r, row in enumerate(buttons):
    for c, button_text in enumerate(row):
        if button_text == "=":
            btn = tk.Button(
                button_frame,
                text=button_text,
                font=("Arial", 20),
                command=calculate_result
            )
            btn.grid(row=r, column=0, columnspan=4, sticky="nsew", padx=3, pady=3)
        elif button_text == "C":
            btn = tk.Button(
                button_frame,
                text=button_text,
                font=("Arial", 20),
                command=clear
            )
            btn.grid(row=r, column=c, sticky="nsew", padx=3, pady=3)
        else:
            btn = tk.Button(
                button_frame,
                text=button_text,
                font=("Arial", 20),
                command=lambda val=button_text: press(val)
            )
            btn.grid(row=r, column=c, sticky="nsew", padx=3, pady=3)

for i in range(5):
    button_frame.rowconfigure(i, weight=1)

for i in range(4):
    button_frame.columnconfigure(i, weight=1)

root.mainloop()