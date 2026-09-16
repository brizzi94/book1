"""
RESISTOMAT 2316 - manuelle Messung mit UI
------------------------------------------
Ablauf:
1. Beim Start: Anzahl Messungen eingeben, dann Speicherort fuer die
   Excel-Datei per Dialog waehlen.
2. Das Fenster zeigt fortlaufend den aktuellen Messwert vom Geraet an
   (wird im Hintergrund alle 0.3s abgefragt).
3. Du misst manuell am Geraet. Sobald der angezeigte Wert passt,
   druecke ENTER (oder klicke "Bestaetigen") - der aktuell angezeigte
   Wert wird mit Zeitstempel in die Excel-Datei geschrieben.
4. Nach N Messungen ("Anzahl") ist Schluss, das Fenster meldet "Fertig".

Vorher pruefen:
- pip install pyserial openpyxl
- Am Geraet, Menue 150 (RS232): BLOCKCHECK muss auf OFF stehen
- SERIAL_PORT und BAUDRATE unten anpassen
"""

import json
import re
import time
import datetime
import threading
import queue
import tkinter as tk
from tkinter import simpledialog, filedialog, messagebox

import serial
from pathlib import Path
from openpyxl import Workbook, load_workbook

SERIAL_PORT = "COM4"   # <-- anpassen
BAUDRATE = 9600        # <-- muss mit Geraete-Einstellung uebereinstimmen
POLL_INTERVAL = 0.3

# Merkt sich den zuletzt gewaehlten Speicherort, damit der Dialog beim
# naechsten Start denselben Ordner/Dateinamen vorschlaegt.
CONFIG_PATH = Path.home() / ".resistomat_2316_config.json"


def load_last_excel_path():
    try:
        data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        return data.get("last_excel_path")
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        return None


def save_last_excel_path(path: str):
    try:
        CONFIG_PATH.write_text(json.dumps({"last_excel_path": path}), encoding="utf-8")
    except OSError:
        pass

STX = b"\x02"
ETX = b"\x03"
ENQ = b"\x05"
ACK = b"\x06"
NAK = b"\x15"
LF = b"\x0A"
EOT = b"\x04"
ADDR = b"0000"

UNIT_FACTORS = {
    "OHM": 1.0,
    "KOHM": 1_000.0,
    "MOHM": 0.001,
    "UOHM": 0.000_001,
}


# ---------------------------------------------------------
# RESISTOMAT-Kommunikation (wie im bisherigen Skript)
# ---------------------------------------------------------
def send(ser, command: str) -> bool:
    frame = EOT + ADDR + b"sr" + STX + command.encode("ascii") + LF + ETX
    ser.write(frame)
    return ser.read(1) == ACK


def poll(ser):
    ser.write(EOT + ADDR + b"po" + ENQ)
    raw = ser.read_until(ETX)
    if STX not in raw:
        return None
    ser.write(ACK)
    ser.read(1)  # trailing EOT
    return raw.split(STX, 1)[1].split(ETX, 1)[0].decode("ascii", errors="replace").strip()


def fetch_value(ser):
    if not send(ser, "fetc?"):
        return None
    return poll(ser)


def parse_ohm(raw_value: str):
    match = re.match(r"([-+]?\d*\.?\d+)\s*([A-Za-z]*)", raw_value)
    if not match:
        return None
    number = float(match.group(1))
    unit = match.group(2).upper()
    factor = UNIT_FACTORS.get(unit, 1.0)
    return number * factor


def ensure_excel_file(path: str):
    file = Path(path)
    if file.exists():
        wb = load_workbook(path)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Messwerte"
        ws.append(["Datum/Zeit", "Messwert (Ohm)"])
    return wb, ws


def append_measurement(wb, ws, path: str, value_ohm: float):
    ws.append([datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), value_ohm])
    try:
        wb.save(path)
        return True
    except PermissionError:
        return False


# ---------------------------------------------------------
# Hintergrund-Thread: fragt laufend den aktuellen Wert ab
# ---------------------------------------------------------
def polling_worker(ser, value_queue, stop_event):
    while not stop_event.is_set():
        raw_value = fetch_value(ser)
        if raw_value is not None:
            value_queue.put(raw_value)
        time.sleep(POLL_INTERVAL)


# ---------------------------------------------------------
# UI
# ---------------------------------------------------------
class ResistomatUI:
    def __init__(self, root, ser, excel_path, target_count):
        self.root = root
        self.ser = ser
        self.excel_path = excel_path
        self.target_count = target_count
        self.count = 0
        self.current_raw_value = None

        self.wb, self.ws = ensure_excel_file(excel_path)

        self.value_queue = queue.Queue()
        self.stop_event = threading.Event()
        self.worker_thread = threading.Thread(
            target=polling_worker,
            args=(ser, self.value_queue, self.stop_event),
            daemon=True,
        )
        self.worker_thread.start()

        self.build_ui()
        self.poll_queue()

    def build_ui(self):
        self.root.title("RESISTOMAT 2316 - Messung")
        self.root.geometry("420x320")

        tk.Label(self.root, text="Aktueller Messwert:", font=("Segoe UI", 12)).pack(pady=(20, 0))
        self.value_label = tk.Label(self.root, text="--", font=("Segoe UI", 28, "bold"))
        self.value_label.pack(pady=10)

        self.progress_label = tk.Label(self.root, text=f"Messung 0 / {self.target_count}", font=("Segoe UI", 11))
        self.progress_label.pack(pady=(0, 10))

        self.confirm_button = tk.Button(
            self.root, text="Bestaetigen (Enter)", font=("Segoe UI", 12),
            command=self.confirm_measurement
        )
        self.confirm_button.pack(pady=10)

        tk.Label(self.root, text="Letzte Eintraege:", font=("Segoe UI", 10)).pack(pady=(10, 0))
        self.history_listbox = tk.Listbox(self.root, height=8, font=("Consolas", 10))
        self.history_listbox.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        self.root.bind("<Return>", lambda event: self.confirm_measurement())

    def poll_queue(self):
        try:
            while True:
                raw_value = self.value_queue.get_nowait()
                self.current_raw_value = raw_value
                self.value_label.config(text=raw_value)
        except queue.Empty:
            pass

        if self.count < self.target_count:
            self.root.after(150, self.poll_queue)

    def confirm_measurement(self):
        if self.count >= self.target_count:
            return
        if self.current_raw_value is None:
            messagebox.showwarning("Kein Wert", "Es liegt noch kein Messwert vor.")
            return

        ohm_value = parse_ohm(self.current_raw_value)
        if ohm_value is None:
            messagebox.showwarning("Fehler", f"Konnte Wert nicht lesen: {self.current_raw_value}")
            return

        saved = append_measurement(self.wb, self.ws, self.excel_path, ohm_value)
        if not saved:
            messagebox.showerror(
                "Excel gesperrt",
                "Die Excel-Datei ist gerade geoeffnet (z.B. in Excel).\nBitte schliessen und erneut bestaetigen."
            )
            return

        self.count += 1
        self.history_listbox.insert(0, f"#{self.count}: {ohm_value} Ohm")
        self.progress_label.config(text=f"Messung {self.count} / {self.target_count}")

        if self.count >= self.target_count:
            self.confirm_button.config(state="disabled", text="Fertig!")
            self.value_label.config(text="Fertig")
            messagebox.showinfo("Fertig", f"Alle {self.target_count} Messungen erfasst.")
            self.stop_event.set()

    def on_close(self):
        self.stop_event.set()
        self.ser.close()
        self.root.destroy()


def main():
    root = tk.Tk()
    root.withdraw()

    target_count = simpledialog.askinteger(
        "Anzahl Messungen", "Wie viele Messungen sollen erfasst werden?",
        minvalue=1, initialvalue=30
    )
    if not target_count:
        return

    last_path = load_last_excel_path()
    if last_path:
        initialdir = str(Path(last_path).parent)
        initialfile = Path(last_path).name
    else:
        initialdir = str(Path.home())
        initialfile = "resistomat_messwerte.xlsx"

    excel_path = filedialog.asksaveasfilename(
        title="Excel-Datei waehlen/erstellen",
        defaultextension=".xlsx",
        filetypes=[("Excel-Datei", "*.xlsx")],
        initialdir=initialdir,
        initialfile=initialfile,
    )
    if not excel_path:
        return
    save_last_excel_path(excel_path)

    try:
        ser = serial.Serial(
            port=SERIAL_PORT,
            baudrate=BAUDRATE,
            bytesize=serial.EIGHTBITS,
            parity=serial.PARITY_NONE,
            stopbits=serial.STOPBITS_ONE,
            timeout=5,
        )
    except serial.SerialException as e:
        messagebox.showerror("Verbindungsfehler", f"Konnte {SERIAL_PORT} nicht oeffnen:\n{e}")
        return

    if not send(ser, "*idn?") or poll(ser) is None:
        messagebox.showerror("Geraetefehler", "Geraet antwortet nicht auf *idn? - Verbindung pruefen.")
        ser.close()
        return

    root.deiconify()
    app = ResistomatUI(root, ser, excel_path, target_count)
    root.protocol("WM_DELETE_WINDOW", app.on_close)
    root.mainloop()


if __name__ == "__main__":
    main()
