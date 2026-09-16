import sys
import time
from pathlib import Path

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
    def __init__(self, extensions=None):
        super().__init__()
        # z.B. {'.csv', '.xlsx'} oder None = alle Dateien
        self.extensions = {e.lower() for e in extensions} if extensions else None

    def on_created(self, event):
        if event.is_directory:
            return

        path = Path(event.src_path)

        if self.extensions and path.suffix.lower() not in self.extensions:
            return

        self._notify(path)

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


def main():
    folder = Path(ZIELORDNER).expanduser().resolve()
    if not folder.is_dir():
        print(f"Fehler: '{folder}' ist kein gültiger Ordner.")
        print("Bitte ZIELORDNER oben im Script anpassen.")
        sys.exit(1)

    if not HAS_PLYER:
        print("Hinweis: 'plyer' nicht installiert -> nur Konsolenausgabe, keine Desktop-Meldung.")
        print("  Installieren mit: pip install plyer --break-system-packages\n")

    handler = NewFileHandler(extensions=DATEIENDUNGEN or None)
    observer = PollingObserver(timeout=POLL_INTERVALL)
    observer.schedule(handler, str(folder), recursive=REKURSIV)
    observer.start()

    print(f"Überwache Ordner (Polling, alle {POLL_INTERVALL}s): {folder}")
    if DATEIENDUNGEN:
        print(f"Gefilterte Endungen: {', '.join(DATEIENDUNGEN)}")
    print("Zum Beenden: Strg+C\n")

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
        print("\nBeendet.")
    observer.join()


if __name__ == "__main__":
    main()
