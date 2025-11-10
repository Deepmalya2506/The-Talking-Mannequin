import os
import time
import queue
import threading
import traceback
import asyncio
import json
import speech_recognition as sr
import pyttsx3
from dotenv import load_dotenv
from langchain_groq import ChatGroq
from langgraph.graph import StateGraph, END
from typing import TypedDict, Set
from datetime import datetime
from zoneinfo import ZoneInfo
from fastapi import FastAPI, WebSocket
from fastapi.middleware.cors import CORSMiddleware
import uvicorn

# ---------------- CONFIG ----------------
load_dotenv()
TIMEZONE = 'Asia/Kolkata'
WAKE_WORDS = ['hello', 'wake up', 'hey', 'jarvis']
WEBSOCKET_PORT = 8000

# Presentation state
current_document = None
current_page = 0
total_pages = 0

# ---------------- SHARED MEMORY ----------------
conversation_history = []
user_input_queue = queue.Queue()
shutdown_flag = threading.Event()
is_agent_awake = threading.Event()
has_greeted = threading.Event()
is_speaking = threading.Event()
speech_interrupted = threading.Event()

# WebSocket clients
connected_clients: Set[WebSocket] = set()

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

# ---------------- WEBSOCKET COMMAND SENDER ----------------
async def send_to_laptop(command: dict):
    """Send command to all connected laptop clients"""
    if not connected_clients:
        print("⚠️ No laptop connected")
        return
    
    disconnected = set()
    for client in connected_clients:
        try:
            await client.send_json(command)
            print(f"📤 Sent to laptop: {command}")
        except Exception as e:
            print(f"❌ Failed to send to client: {e}")
            disconnected.add(client)
    
    # Remove disconnected clients
    for client in disconnected:
        connected_clients.discard(client)

def send_command_sync(command: dict):
    """Synchronous wrapper for sending commands"""
    try:
        loop = asyncio.get_event_loop()
        if loop.is_running():
            asyncio.create_task(send_to_laptop(command))
        else:
            loop.run_until_complete(send_to_laptop(command))
    except Exception as e:
        print(f"❌ Error sending command: {e}")

# ---------------- FASTAPI WEBSOCKET SERVER ----------------
app = FastAPI(title="Pi Voice Agent")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.websocket("/ws")
async def websocket_endpoint(websocket: WebSocket):
    """WebSocket connection for laptop client"""
    await websocket.accept()
    connected_clients.add(websocket)
    print(f"🔗 Laptop connected! Total clients: {len(connected_clients)}")
    
    try:
        while True:
            # Keep connection alive and listen for responses
            data = await websocket.receive_text()
            print(f"📥 Received from laptop: {data}")
    except Exception as e:
        print(f"❌ WebSocket error: {e}")
    finally:
        connected_clients.discard(websocket)
        print(f"🔌 Laptop disconnected. Total clients: {len(connected_clients)}")

def start_websocket_server():
    """Start FastAPI WebSocket server in background thread"""
    config = uvicorn.Config(app, host="0.0.0.0", port=WEBSOCKET_PORT, log_level="warning")
    server = uvicorn.Server(config)
    
    def run():
        asyncio.run(server.serve())
    
    thread = threading.Thread(target=run, daemon=True)
    thread.start()
    print(f"🌐 WebSocket server started on port {WEBSOCKET_PORT}")

# ---------------- SPEECH MANAGER ----------------
class SpeechManager:
    """Thread-safe text-to-speech manager with interruption support"""
    def __init__(self):
        self.lock = threading.Lock()
        self.current_engine = None

    def speak(self, text):
        """Convert text to speech with interruption capability"""
        print(f"\n🔊 Assistant: {text}")
        
        with self.lock:
            is_speaking.set()
            speech_interrupted.clear()
            
            try:
                self.current_engine = pyttsx3.init('sapi5')
                self.current_engine.setProperty('rate', 190)
                self.current_engine.setProperty('volume', 0.9)
                
                chunks = []
                current_chunk = []
                
                for part in text.replace('!', '! ').replace('?', '? ').replace('.', '. ').split():
                    current_chunk.append(part)
                    if part.endswith(('.', '!', '?')):
                        chunks.append(' '.join(current_chunk))
                        current_chunk = []
                
                if current_chunk:
                    chunks.append(' '.join(current_chunk))
                
                for chunk in chunks:
                    if speech_interrupted.is_set():
                        print("⏸️ Speech interrupted by user input")
                        break
                    
                    if chunk.strip():
                        self.current_engine.say(chunk)
                        self.current_engine.runAndWait()
                
                del self.current_engine
                self.current_engine = None
                
            except Exception as e:
                print(f"[TTS Error: {e}]")
            finally:
                is_speaking.clear()
                speech_interrupted.clear()
    
    def interrupt(self):
        """Interrupt ongoing speech"""
        speech_interrupted.set()
        if self.current_engine:
            try:
                self.current_engine.stop()
            except:
                pass

speech_manager = SpeechManager()

# ---------------- BACKGROUND LISTENER ----------------
def start_listener():
    """Initialize continuous speech recognition in background"""
    recognizer = sr.Recognizer()
    recognizer.energy_threshold = 4000
    recognizer.dynamic_energy_threshold = True
    recognizer.pause_threshold = 1.0
    
    mic = sr.Microphone()
    
    with mic as source:
        print("🎤 Calibrating microphone...")
        recognizer.adjust_for_ambient_noise(source, duration=2)
        print("✅ Calibration complete")

    def callback(recognizer, audio):
        try:
            text = recognizer.recognize_google(audio).strip().lower()
            if not text:
                return
            
            print(f"👂 Heard: {text}")
            
            if is_speaking.is_set():
                print("🛑 User spoke while I was talking - interrupting myself")
                speech_manager.interrupt()
            
            if not is_agent_awake.is_set():
                if any(wake_word in text for wake_word in WAKE_WORDS):
                    print("⏰ Wake word detected!")
                    user_input_queue.put(text)
                    is_agent_awake.set()
            else:
                user_input_queue.put(text)
                
        except sr.UnknownValueError:
            pass
        except Exception as e:
            print(f"[Listener Error] {e}")

    recognizer.listen_in_background(mic, callback, phrase_time_limit=15)
    print("🎧 Background listener active\n")

# ---------------- STATE DEFINITION ----------------
class AgentState(TypedDict):
    user_input: str
    intent: str
    response: str
    should_continue: bool
    waiting_for_input: bool

# ---------------- GRAPH NODES ----------------
def wake_up_node(state: AgentState) -> AgentState:
    """Initial greeting when agent activates (only once)"""
    
    if not has_greeted.is_set():
        current_time = datetime.now(ZoneInfo(TIMEZONE))
        hour = current_time.hour
        
        if 5 <= hour < 12:
            greeting = "Good morning"
        elif 12 <= hour < 17:
            greeting = "Good afternoon"
        else:
            greeting = "Good evening"
        
        prompt = f"""You're a friendly AI assistant being activated for the first time. 
        Current time: {current_time.strftime('%I:%M %p')}
        Give a brief, warm {greeting} greeting (1-2 sentences) and ask how you can help."""
        
        response = model_call(prompt)
        state['response'] = response
        
        conversation_history.append({"role": "assistant", "content": response})
        speech_manager.speak(response)
        
        has_greeted.set()
    
    state['should_continue'] = True
    state['waiting_for_input'] = True
    return state

def listen_node(state: AgentState) -> AgentState:
    """Wait for user input before proceeding"""
    print("\n⏳ Waiting for your input...")
    
    while not user_input_queue.empty():
        try:
            user_input_queue.get_nowait()
        except:
            break
    
    timeout_counter = 0
    max_timeout = 120
    
    while True:
        try:
            user_input = user_input_queue.get(timeout=1)
            state['user_input'] = user_input
            state['waiting_for_input'] = False
            print(f"✅ Got input: {user_input}")
            break
            
        except queue.Empty:
            timeout_counter += 1
            
            if timeout_counter >= max_timeout:
                prompt_msg = "Are you still there? I'm listening whenever you're ready."
                speech_manager.speak(prompt_msg)
                timeout_counter = 0
            
            if shutdown_flag.is_set():
                state['should_continue'] = False
                state['waiting_for_input'] = False
                break
            
            continue
    
    return state

def classify_intent_node(state: AgentState) -> AgentState:
    """Classify user intent: chat, presentation, or sleep"""
    user_input = state['user_input']
    
    context = "\n".join([
        f"{msg['role'].title()}: {msg['content']}" 
        for msg in conversation_history[-4:]
    ]) if conversation_history else "No prior conversation"
    
    prompt = f"""Classify the following input into ONE category based on user intent:

Categories:
- 'chat': General conversation, questions, or requests
- 'presentation': Requests to present/show slides, navigate presentation (next, previous, open presentation)
- 'sleep': Wants to end session (goodbye, that's all, sleep, stop, exit, etc.)

Recent context:
{context}

Current input: "{user_input}"

Respond with ONLY ONE WORD: chat, presentation, or sleep"""
    
    intent = model_call(prompt).lower().strip()
    
    if intent not in ['chat', 'presentation', 'sleep']:
        intent = 'chat'
    
    state['intent'] = intent
    print(f"🎯 Intent: {intent}")
    
    return state

def chat_node(state: AgentState) -> AgentState:
    """Handle general conversation"""
    user_input = state['user_input']
    
    context = "\n".join([
        f"{msg['role'].title()}: {msg['content']}" 
        for msg in conversation_history[-8:]
    ]) if conversation_history else "No prior conversation"
    
    prompt = f"""You are a helpful, conversational AI assistant.

Recent conversation:
{context}

User: {user_input}

Provide a natural, concise response (2-3 sentences max). Be helpful and friendly."""
    
    response = model_call(prompt)
    state['response'] = response
    
    conversation_history.append({"role": "user", "content": user_input})
    conversation_history.append({"role": "assistant", "content": response})
    
    speech_manager.speak(response)
    
    state['should_continue'] = True
    state['waiting_for_input'] = True
    return state

def presentation_node(state: AgentState) -> AgentState:
    """Handle presentation-related queries - sends commands to laptop"""
    global current_document, current_page, total_pages
    
    user_input = state['user_input']
    user_input_lower = user_input.lower()
    
    response = ""
    
    try:
        # Check if laptop is connected
        if not connected_clients:
            response = "No laptop is connected. Please start the laptop client first."
            speech_manager.speak(response)
            state['response'] = response
            state['should_continue'] = True
            state['waiting_for_input'] = True
            return state
        
        # Command: Open presentation
        if any(word in user_input_lower for word in ['open', 'start', 'show', 'present', 'load']):
            # For now, we'll use a default filename - you can expand this
            filename = "NeuralTwin.pptx"  # This should match the file on laptop
            
            command = {
                "action": "open",
                "filename": filename
            }
            send_command_sync(command)
            
            current_document = filename
            current_page = 1
            response = f"Opening presentation: {filename}"
        
        # Command: Next slide
        elif any(word in user_input_lower for word in ['next', 'forward', 'next slide']):
            if current_document:
                command = {"action": "next"}
                send_command_sync(command)
                current_page += 1
                response = f"Moving to slide {current_page}"
            else:
                response = "No presentation is currently open. Please ask me to open a presentation first."
        
        # Command: Previous slide
        elif any(word in user_input_lower for word in ['previous', 'back', 'previous slide', 'go back']):
            if current_document:
                if current_page > 1:
                    command = {"action": "previous"}
                    send_command_sync(command)
                    current_page -= 1
                    response = f"Moving to slide {current_page}"
                else:
                    response = "Already on the first slide"
            else:
                response = "No presentation is currently open."
        
        # Command: Go to specific slide
        elif 'slide' in user_input_lower and any(char.isdigit() for char in user_input):
            if current_document:
                # Extract slide number
                import re
                numbers = re.findall(r'\d+', user_input)
                if numbers:
                    slide_num = int(numbers[0])
                    command = {
                        "action": "goto",
                        "slide_number": slide_num
                    }
                    send_command_sync(command)
                    current_page = slide_num
                    response = f"Going to slide {slide_num}"
                else:
                    response = "I couldn't understand which slide number you want."
            else:
                response = "No presentation is currently open."
        
        # Command: Close presentation
        elif any(word in user_input_lower for word in ['close', 'exit presentation', 'stop presenting']):
            if current_document:
                command = {"action": "close"}
                send_command_sync(command)
                response = f"Closing presentation"
                current_document = None
                current_page = 0
            else:
                response = "No presentation is currently open."
        
        # Default: General explanation about presentation
        else:
            response = "I can help you control presentations. Try saying 'open presentation', 'next slide', 'previous slide', or 'go to slide 5'."
        
    except Exception as e:
        print(f"[Presentation Error] {e}")
        traceback.print_exc()
        response = "I encountered an error with the presentation control."
    
    state['response'] = response
    
    conversation_history.append({"role": "user", "content": user_input})
    conversation_history.append({"role": "assistant", "content": response})
    
    speech_manager.speak(response)
    
    state['should_continue'] = True
    state['waiting_for_input'] = True
    return state

def sleep_node(state: AgentState) -> AgentState:
    """Handle session end"""
    global current_document, current_page
    
    current_time = datetime.now(ZoneInfo(TIMEZONE))
    hour = current_time.hour
    
    if 5 <= hour < 17:
        farewell = "Have a great day"
    else:
        farewell = "Have a wonderful evening"
    
    prompt = f"""The user is ending the session. 
    Give a brief, warm farewell (1 sentence). 
    Use this context: {farewell}. 
    Add a caring note."""
    
    response = model_call(prompt)
    state['response'] = response
    
    conversation_history.append({"role": "user", "content": state['user_input']})
    conversation_history.append({"role": "assistant", "content": response})
    speech_manager.speak(response)
    
    # Close any open presentation
    if current_document:
        send_command_sync({"action": "close"})
        current_document = None
        current_page = 0
    
    is_agent_awake.clear()
    has_greeted.clear()
    
    state['should_continue'] = False
    state['waiting_for_input'] = False
    print("💤 Agent going to sleep. Say a wake word to reactivate.\n")
    
    return state

# ---------------- ROUTING LOGIC ----------------
def route_by_intent(state: AgentState) -> str:
    """Route to appropriate node based on intent"""
    intent = state.get('intent', 'chat')
    
    route_map = {
        'chat': 'chat',
        'presentation': 'presentation',
        'sleep': 'sleep'
    }
    
    return route_map.get(intent, 'chat')

def route_after_processing(state: AgentState) -> str:
    """Determine if we should continue or end"""
    if state.get('should_continue', False):
        return 'continue'
    else:
        return 'end'

# ---------------- BUILD GRAPH ----------------
def build_graph():
    """Construct the agent workflow graph"""
    workflow = StateGraph(AgentState)
    
    workflow.add_node("wake_up", wake_up_node)
    workflow.add_node("listen", listen_node)
    workflow.add_node("classify", classify_intent_node)
    workflow.add_node("chat", chat_node)
    workflow.add_node("presentation", presentation_node)
    workflow.add_node("sleep", sleep_node)
    
    workflow.set_entry_point("wake_up")
    workflow.add_edge("wake_up", "listen")
    workflow.add_edge("listen", "classify")
    
    workflow.add_conditional_edges(
        "classify",
        route_by_intent,
        {
            "chat": "chat",
            "presentation": "presentation",
            "sleep": "sleep"
        }
    )
    
    workflow.add_conditional_edges(
        "chat",
        route_after_processing,
        {
            "continue": "listen",
            "end": END
        }
    )
    
    workflow.add_conditional_edges(
        "presentation",
        route_after_processing,
        {
            "continue": "listen",
            "end": END
        }
    )
    
    workflow.add_edge("sleep", END)
    
    return workflow.compile()

# ---------------- MAIN LOOP ----------------
def main():
    print("=" * 60)
    print("✨ VOICE AGENT SYSTEM (Raspberry Pi)")
    print("=" * 60)
    print(f"💤 Agent is sleeping. Say one of these to wake: {', '.join(WAKE_WORDS)}")
    print("🎤 Speak clearly - I'll wait for you to finish")
    print("🛑 You can interrupt me anytime while I'm speaking")
    print("💬 Say 'goodbye' or 'sleep' to put me to sleep")
    print("⌨️  Press Ctrl+C to exit\n")
    
    # Start WebSocket server
    start_websocket_server()
    time.sleep(2)  # Give server time to start
    
    # Start voice listener
    start_listener()
    
    # Build agent graph
    agent_graph = build_graph()
    
    while not shutdown_flag.is_set():
        try:
            if not is_agent_awake.is_set():
                try:
                    time.sleep(0.5)
                    continue
                except KeyboardInterrupt:
                    break
            
            state = AgentState(
                user_input="",
                intent="",
                response="",
                should_continue=True,
                waiting_for_input=False
            )
            
            try:
                result = agent_graph.invoke(state)
                
                if not result.get('should_continue', False):
                    print("\n" + "=" * 60)
                    print("💤 Conversation ended. Listening for wake word...")
                    print("=" * 60 + "\n")
                    
            except Exception as e:
                print(f"[Graph Error] {e}")
                traceback.print_exc()
                time.sleep(1)
                
        except KeyboardInterrupt:
            print("\n\n👋 Shutting down...")
            if is_agent_awake.is_set():
                speech_manager.speak("Goodbye!")
            break
        except Exception as e:
            print(f"[Main Error] {e}")
            traceback.print_exc()
            time.sleep(1)
    
    shutdown_flag.set()
    print("✨ Agent system terminated.")

if __name__ == "__main__":
    main()