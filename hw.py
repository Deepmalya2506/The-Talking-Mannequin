"""
Raspberry Pi 4 Voice Agent + PowerPoint Control
Uses Vosk for offline speech recognition (works better on RPi)
"""

import os
import time
import queue
import threading
import traceback
from dotenv import load_dotenv
from langchain_groq import ChatGroq
from langgraph.graph import StateGraph, END
from typing import TypedDict
from datetime import datetime
from pathlib import Path
import subprocess
import signal
import sys
import json
import pyaudio
from vosk import Model, KaldiRecognizer

# ---------------- CONFIG ----------------
load_dotenv()
DOCUMENTS_FOLDER = Path.home() / "Documents" / "Presentations"
DOCUMENTS_FOLDER.mkdir(parents=True, exist_ok=True)
TIMEZONE = 'Asia/Kolkata'
WAKE_WORDS = ['hello', 'wake up', 'hey', 'jarvis']

# Vosk model path - will be downloaded if not exists
VOSK_MODEL_PATH = Path.home() / ".vosk_models" / "vosk-model-small-en-us-0.15"

# Presentation state
current_document = None
current_page = 0

# ---------------- SHARED MEMORY ----------------
conversation_history = []
user_input_queue = queue.Queue()
shutdown_flag = threading.Event()
is_agent_awake = threading.Event()
has_greeted = threading.Event()
is_speaking = threading.Event()
speech_interrupted = threading.Event()

# ---------------- LLM MODEL ----------------
llm = ChatGroq(model="llama-3.3-70b-versatile", temperature=0.7)

def model_call(prompt: str) -> str:
    """Centralized LLM call with error handling"""
    try:
        response = llm.invoke(prompt)
        return response.content.strip()
    except Exception as e:
        print(f"[Model Error] {e}")
        return "I'm having trouble processing that right now."

# ---------------- POWERPOINT CONTROLLER (LibreOffice) ----------------
class LibreOfficeController:
    """LibreOffice Impress controller for Raspberry Pi"""
    
    def _init_(self):
        self.process = None
        self.presentation_path = None
        self.current_slide = 1
        self.total_slides = 0
        self.presentation_open = False
    
    def open_presentation(self, filename: str) -> bool:
        """Open a presentation file with LibreOffice Impress"""
        try:
            filepath = DOCUMENTS_FOLDER / filename
            
            if not filepath.exists():
                print(f"❌ File not found: {filepath}")
                return False
            
            # Close existing presentation
            if self.process:
                self.close_presentation()
            
            # Open with LibreOffice Impress in presentation mode
            try:
                self.process = subprocess.Popen([
                    'libreoffice',
                    '--impress',
                    '--show',
                    str(filepath.absolute())
                ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                
                self.presentation_path = filepath
                self.current_slide = 1
                self.presentation_open = True
                time.sleep(2)
                
                self.total_slides = self._count_slides(filepath)
                print(f"✅ Opened: {filename} (~{self.total_slides} slides)")
                return True
                
            except FileNotFoundError:
                print("❌ LibreOffice not found. Install: sudo apt install libreoffice-impress")
                return False
                
        except Exception as e:
            print(f"❌ Error opening presentation: {e}")
            traceback.print_exc()
            return False
    
    def _count_slides(self, filepath: Path) -> int:
        """Try to count slides using python-pptx"""
        try:
            from pptx import Presentation
            prs = Presentation(str(filepath))
            return len(prs.slides)
        except:
            return 10
    
    def next_slide(self) -> bool:
        """Go to next slide"""
        if not self.presentation_open:
            return False
        
        try:
            subprocess.run(['xdotool', 'key', 'Right'], check=False)
            if self.current_slide < self.total_slides:
                self.current_slide += 1
                print(f"➡ Slide {self.current_slide}")
                return True
            else:
                print("⚠ Last slide")
                return False
        except Exception as e:
            print(f"❌ Error: {e}")
            return False
    
    def previous_slide(self) -> bool:
        """Go to previous slide"""
        if not self.presentation_open:
            return False
        
        try:
            subprocess.run(['xdotool', 'key', 'Left'], check=False)
            if self.current_slide > 1:
                self.current_slide -= 1
                print(f"⬅ Slide {self.current_slide}")
                return True
            else:
                print("⚠ First slide")
                return False
        except Exception as e:
            print(f"❌ Error: {e}")
            return False
    
    def goto_slide(self, slide_number: int) -> bool:
        """Go to specific slide"""
        if not self.presentation_open:
            return False
        
        try:
            if 1 <= slide_number <= self.total_slides:
                diff = slide_number - self.current_slide
                
                if diff > 0:
                    for _ in range(diff):
                        subprocess.run(['xdotool', 'key', 'Right'], check=False)
                        time.sleep(0.2)
                elif diff < 0:
                    for _ in range(abs(diff)):
                        subprocess.run(['xdotool', 'key', 'Left'], check=False)
                        time.sleep(0.2)
                
                self.current_slide = slide_number
                print(f"🎯 Slide {slide_number}")
                return True
            else:
                print(f"⚠ Invalid slide: {slide_number}")
                return False
        except Exception as e:
            print(f"❌ Error: {e}")
            return False
    
    def close_presentation(self) -> bool:
        """Close presentation"""
        try:
            if self.process:
                try:
                    subprocess.run(['xdotool', 'key', 'Escape'], check=False)
                    time.sleep(0.5)
                except:
                    pass
                
                self.process.terminate()
                try:
                    self.process.wait(timeout=3)
                except subprocess.TimeoutExpired:
                    self.process.kill()
                
                self.process = None
                self.presentation_open = False
                print("✅ Presentation closed")
                return True
            return False
        except Exception as e:
            print(f"❌ Error: {e}")
            return False
    
    def get_current_slide(self) -> int:
        """Get current slide number"""
        return self.current_slide if self.presentation_open else 0

# Global controller instance
ppt_controller = LibreOfficeController()

# ---------------- SPEECH MANAGER ----------------
class SpeechManager:
    """Text output manager with espeak TTS"""
    def _init_(self):
        self.lock = threading.Lock()
        self.use_espeak = self._check_espeak()
        
    def _check_espeak(self) -> bool:
        """Check if espeak is available"""
        try:
            subprocess.run(['espeak', '--version'], 
                         stdout=subprocess.DEVNULL, 
                         stderr=subprocess.DEVNULL, 
                         check=True)
            print("✅ espeak TTS available")
            return True
        except:
            print("⚠ espeak not found. Install: sudo apt install espeak")
            return False

    def speak(self, text):
        """Output text and speak using espeak"""
        print(f"\n🔊 Assistant: {text}")

        with self.lock:
            is_speaking.set()
            speech_interrupted.clear()

            if self.use_espeak:
                try:
                    process = subprocess.Popen(
                        ['espeak', '-s', '160', text],
                        stdout=subprocess.DEVNULL,
                        stderr=subprocess.DEVNULL
                    )
                    
                    while process.poll() is None:
                        if speech_interrupted.is_set():
                            process.terminate()
                            print("⏸ Speech interrupted")
                            break
                        time.sleep(0.1)
                    
                except Exception as e:
                    print(f"[TTS Error] {e}")

            is_speaking.clear()
            speech_interrupted.clear()
    
    def interrupt(self):
        """Interrupt ongoing speech"""
        speech_interrupted.set()

speech_manager = SpeechManager()

# ---------------- VOSK SPEECH RECOGNITION ----------------
def download_vosk_model():
    """Download Vosk model if not exists"""
    if VOSK_MODEL_PATH.exists():
        return True
    
    print("📦 Vosk model not found. Downloading small English model...")
    print("   This is a one-time download (~40MB)")
    
    try:
        import urllib.request
        import zipfile
        
        VOSK_MODEL_PATH.parent.mkdir(parents=True, exist_ok=True)
        
        url = "https://alphacephei.com/vosk/models/vosk-model-small-en-us-0.15.zip"
        zip_path = VOSK_MODEL_PATH.parent / "model.zip"
        
        print("   Downloading...")
        urllib.request.urlretrieve(url, zip_path)
        
        print("   Extracting...")
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(VOSK_MODEL_PATH.parent)
        
        zip_path.unlink()
        print("✅ Model downloaded successfully")
        return True
        
    except Exception as e:
        print(f"❌ Failed to download model: {e}")
        print("   Manual download: https://alphacephei.com/vosk/models")
        print(f"   Extract to: {VOSK_MODEL_PATH}")
        return False

def start_vosk_listener():
    """Initialize Vosk speech recognition in background"""
    
    if not download_vosk_model():
        print("❌ Cannot start without Vosk model")
        return
    
    try:
        model = Model(str(VOSK_MODEL_PATH))
        print("✅ Vosk model loaded")
    except Exception as e:
        print(f"❌ Failed to load Vosk model: {e}")
        return
    
    # Audio configuration
    RATE = 16000
    CHUNK = 8000
    
    try:
        audio = pyaudio.PyAudio()
        stream = audio.open(
            format=pyaudio.paInt16,
            channels=1,
            rate=RATE,
            input=True,
            frames_per_buffer=CHUNK
        )
        stream.start_stream()
        print("✅ Audio stream started")
    except Exception as e:
        print(f"❌ Failed to open audio stream: {e}")
        print("   Check microphone connection")
        return
    
    recognizer = KaldiRecognizer(model, RATE)
    recognizer.SetWords(True)
    
    print("🎤 Vosk listener active - speak now!\n")
    
    def listen_loop():
        while not shutdown_flag.is_set():
            try:
                data = stream.read(CHUNK, exception_on_overflow=False)
                
                if recognizer.AcceptWaveform(data):
                    result = json.loads(recognizer.Result())
                    text = result.get('text', '').strip().lower()
                    
                    if text:
                        print(f"👂 Heard: {text}")
                        
                        if is_speaking.is_set():
                            print("🛑 Interrupting speech")
                            speech_manager.interrupt()
                            user_input_queue.put(text)
                        else:
                            if not is_agent_awake.is_set():
                                if any(wake_word in text for wake_word in WAKE_WORDS):
                                    print("⏰ Wake word detected!")
                                    user_input_queue.put(text)
                                    is_agent_awake.set()
                            else:
                                user_input_queue.put(text)
                
            except Exception as e:
                if not shutdown_flag.is_set():
                    print(f"[Listener Error] {e}")
