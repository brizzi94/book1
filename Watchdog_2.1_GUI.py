import sys
import time
import queue
from pathlib import Path
from datetime import datetime

import tkinter as tk
from tkinter import ttk

# Wichtig: PollingObserver statt Observer -> funktioniert zuverlässig auch
# auf Netzlaufwerken (SMB-Freigaben wie S:\...), wo die native Windows-
# Benachrichtigung (ReadDirectoryChangesW) Events nach dem ersten oft verschluckt.
from watchdog.observers.polling import PollingObserver
from watchdog.events import FileSystemEventHandler


# ============================================================
# Hier den zu überwachenden Ordner eintragen:
ZIELORDNER = r"S:\SAG\Projects\20_E-Projects\E-0783_USE 2410\5_Qualification\1_Type Testing (internal)\2026-09-07 HTOL\Measurements"

# Nur diese Dateiendungen melden, z.B. [".csv", ".xlsx"]
# Leere Liste [] = alle Dateien melden
DATEIENDUNGEN = []

# Auch Unterordner überwachen?
REKURSIV = False

# Wie oft der Ordner abgefragt wird (Sekunden). Bei Netzlaufwerken reichen
# 1-2 Sekunden; kleinere Werte erzeugen mehr Last auf dem Server.
POLL_INTERVALL = 1.0
# ============================================================

try:
    from plyer import notification
    HAS_PLYER = True
except ImportError:
    HAS_PLYER = False


class NewFileHandler(FileSystemEventHandler):
    def __init__(self, event_queue, extensions=None):
        super().__init__()
        self.event_queue = event_queue
        # z.B. {'.csv', '.xlsx'} oder None = alle Dateien
        self.extensions = {e.lower() for e in extensions} if extensions else None

    def on_created(self, event):
        if event.is_directory:
            return

        path = Path(event.src_path)

        if self.extensions and path.suffix.lower() not in self.extensions:
            return

        self._notify(path)
        # Observer läuft in eigenem Thread -> Event nur über Queue an die GUI geben
        self.event_queue.put((datetime.now(), path))

    def _notify(self, path: Path):
        message = f"Neue Datei: {path.name}"
        timestamp = time.strftime("%H:%M:%S")
        print(f"[{timestamp}] {message}  ({path})")

        if HAS_PLYER:
            try:
                notification.notify(
                    title="Neues Dokument erkannt",
                    message=path.name,
                    timeout=5,
                )
            except Exception as e:
                print(f"  (Desktop-Benachrichtigung fehlgeschlagen: {e})")


class WatchdogGUI:
    def __init__(self, root, folder, event_queue, observer):
        self.root = root
        self.event_queue = event_queue
        self.observer = observer
        self.total_count = 0

        root.title("Watchdog 2.1 – Ordnerüberwachung")
        root.geometry("640x480")
        root.minsize(480, 360)
        root.protocol("WM_DELETE_WINDOW", self.on_close)

        ttk.Label(
            root, text=f"Überwacht: {folder}",
            wraplength=620, font=("Segoe UI", 9, "bold"),
        ).pack(fill="x", padx=10, pady=(10, 0))

        self.status_var = tk.StringVar(value="0 Dateien erkannt")
        ttk.Label(root, textvariable=self.status_var, font=("Segoe UI", 10)).pack(
            fill="x", padx=10, pady=(2, 8)
        )

        list_frame = ttk.Frame(root)
        list_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side="right", fill="y")

        self.listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, font=("Consolas", 9))
        self.listbox.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=self.listbox.yview)

        self.poll_queue()

    def poll_queue(self):
        try:
            while True:
                ts, path = self.event_queue.get_nowait()
                self._add_event(ts, path)
        except queue.Empty:
            pass
        self.root.after(200, self.poll_queue)

    def _add_event(self, ts: datetime, path: Path):
        self.total_count += 1
        self.status_var.set(f"{self.total_count} Dateien erkannt")
        self.listbox.insert(0, f"[{ts.strftime('%H:%M:%S')}] {path.name}")

    def on_close(self):
        self.observer.stop()
        self.observer.join(timeout=2)
        self.root.destroy()


def main():
    folder = Path(ZIELORDNER).expanduser().resolve()
    if not folder.is_dir():
        print(f"Fehler: '{folder}' ist kein gültiger Ordner.")
        print("Bitte ZIELORDNER oben im Script anpassen.")
        sys.exit(1)

    if not HAS_PLYER:
        print("Hinweis: 'plyer' nicht installiert -> nur Konsolenausgabe, keine Desktop-Meldung.")
        print("  Installieren mit: pip install plyer --break-system-packages\n")

    event_queue = queue.Queue()
    handler = NewFileHandler(event_queue, extensions=DATEIENDUNGEN or None)
    observer = PollingObserver(timeout=POLL_INTERVALL)
    observer.schedule(handler, str(folder), recursive=REKURSIV)
    observer.start()

    print(f"Überwache Ordner (Polling, alle {POLL_INTERVALL}s): {folder}")
    if DATEIENDUNGEN:
        print(f"Gefilterte Endungen: {', '.join(DATEIENDUNGEN)}")
    print("Fenster schließen zum Beenden.\n")

    root = tk.Tk()
    WatchdogGUI(root, folder, event_queue, observer)
    root.mainloop()


if __name__ == "__main__":
    main()
