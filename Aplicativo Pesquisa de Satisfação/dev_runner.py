import subprocess
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import time
import os

class ReloadHandler(FileSystemEventHandler):
    def __init__(self, file_to_run):
        self.file_to_run = file_to_run
        self.process = None
        self.run_app()

    def run_app(self):
        if self.process:
            self.process.terminate()
        print("\n🔄 Reiniciando o app...\n")
        self.process = subprocess.Popen(["python", self.file_to_run])

    def on_modified(self, event):
        if event.src_path.endswith(".py"):
            self.run_app()

if __name__ == "__main__":
    app_file = "main.py"  # 👈 coloca aqui o arquivo que roda teu app
    event_handler = ReloadHandler(app_file)
    observer = Observer()
    observer.schedule(event_handler, path=".", recursive=True)
    observer.start()

    print("👀 Observando alterações... (CTRL+C para parar)")
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
        if event_handler.process:
            event_handler.process.terminate()
    observer.join()
