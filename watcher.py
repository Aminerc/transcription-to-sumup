import os
import time
import sys
from datetime import datetime
from dotenv import load_dotenv
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

load_dotenv(dotenv_path=os.path.join(os.path.dirname(__file__), ".env"))

WATCH_FOLDER  = os.getenv("WATCH_FOLDER", ".")
EXTENSIONS    = {".docx", ".txt", ".vtt"}
PROCESSED     = set()

# ==================== HANDLER ====================
class TranscriptionHandler(FileSystemEventHandler):

    def on_created(self, event):
        if event.is_directory:
            return
        self._handle(event.src_path)

    def on_moved(self, event):
        if event.is_directory:
            return
        self._handle(event.dest_path)

    def _handle(self, path):
        ext = os.path.splitext(path)[1].lower()
        if ext not in EXTENSIONS:
            return
        if path in PROCESSED:
            return

        # Attendre que le fichier soit complet (ecriture terminee)
        time.sleep(3)
        if not os.path.exists(path):
            return

        PROCESSED.add(path)
        print(f"\n[{datetime.now().strftime('%H:%M:%S')}] Nouveau fichier detecte : {path}")

        try:
            # Recharger process.py a chaque run pour prendre les modifs sans restart du watcher
            import importlib
            import process as process_module
            importlib.reload(process_module)
            process_module.process(path)
        except Exception as e:
            print(f"[ERREUR] {e}")

# ==================== MAIN ====================
def main():
    os.makedirs(WATCH_FOLDER, exist_ok=True)
    print(f"Surveillance active : {WATCH_FOLDER}")
    print(f"Extensions surveillees : {', '.join(EXTENSIONS)}")
    print("En attente de nouveaux fichiers... (Ctrl+C pour arreter)\n")

    handler  = TranscriptionHandler()
    observer = Observer()
    observer.schedule(handler, WATCH_FOLDER, recursive=False)
    observer.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
        print("\nSurveillance arretee.")

    observer.join()

if __name__ == "__main__":
    main()
