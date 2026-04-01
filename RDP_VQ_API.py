import os
import subprocess
import datetime
import time
import threading
import psutil
import paramiko
import logging
import sys
import re
import pyttsx3
import gc
import json
import pyautogui
import speech_recognition as sr
import win32gui
import win32con
import requests
from flask import Flask, request, jsonify
from flask_cors import CORS
from VQ_Analysis import analyze_audio
from ai_conversation import AIConversationAgent

# Disable PyAutoGUI fail-safe
pyautogui.FAILSAFE = False

# === CONFIGURATION ===
VB_CABLE_DEVICE = "CABLE Output (2- VB-Audio Virtual Cable)"
RECORD_DEVICE = VB_CABLE_DEVICE
OUTPUT_DIR = r"C:\Users\eommhoh\Desktop\VOICERecordings"
LOG_FILE = os.path.join(OUTPUT_DIR, "automation_log.txt")
FFMPEG_PATH = r"C:\Users\eommhoh\AppData\Local\Microsoft\WinGet\Packages\Gyan.FFmpeg_Microsoft.Winget.Source_8wekyb3d8bbwe\ffmpeg-8.0-full_build\bin\ffmpeg.exe"
ANSWER_BUTTON_IMAGE = r"C:\Users\eommhoh\Desktop\answer.png"
ANSWER_BUTTON_TIMEOUT = 30
ANSWER_BUTTON_CONFIDENCE = 0.8
POST_GOODBYE_GRACE_SECONDS = 2

# SSH Configuration
HOST = "10.73.220.21"
PORT = 22
USERNAME = "srv00273"
PASSWORD = "AErPY#146Etuhevq_25"

SECOND_HOST = "10.6.6.222"
SECOND_USERNAME = "k.saboor.erc"
SECOND_PASSWORD = "Jkm.nmj.722811"

PATTERN_ANSWER = r"B-ANSWER RECEIVED"
PATTERN_DISCONNECT = r"B-FORCED DISCONNECT"
FORCED_RELEASE_PATTERN = r"FORCED\s*(DISCONNECT|RELEASE)"

# SignalR Configuration
SIGNALR_HUB_URL = "http://localhost:8089/logHub"
signalr_connection = None

# API Server Configuration
API_PORT = 5055
app = Flask(__name__)
CORS(app)

# Parameters from API
TARGET_BNUMBER = None
ANB = None
BNB = None
VOICE_ID = None
CALL_TYPE = None
RECORD_DURATION = 30
USE_AI_CONVERSATION = True
OPENAI_API_KEY = os.getenv('OPENAI_API_KEY')
ORIGINATING_SERVER = None  # Server that initiated the request

# === LOGGING ===
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger()

# === GLOBAL STATE ===
tts_active = False
auto_answer_active = False
auto_answer_thread = None
call_in_progress = False

def find_jabber_window():
    def callback(hwnd, windows):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if "Cisco Jabber" in title or "Jabber" in title:
                windows.append(hwnd)
        return True

    windows = []
    win32gui.EnumWindows(callback, windows)
    return windows[0] if windows else None

def minimize_jabber_window():
    try:
        jabber_hwnd = find_jabber_window()
        if jabber_hwnd:
            win32gui.ShowWindow(jabber_hwnd, win32con.SW_MINIMIZE)
            logger.info("Jabber window minimized")
            return True
    except Exception as e:
        logger.warning(f"Failed to minimize Jabber window: {e}")
    return False

def disconnect_jabber_call():
    try:
        jabber_hwnd = find_jabber_window()
        if jabber_hwnd:
            win32gui.ShowWindow(jabber_hwnd, win32con.SW_RESTORE)
            time.sleep(0.2)
            win32gui.SetForegroundWindow(jabber_hwnd)
            time.sleep(0.3)
        pyautogui.hotkey('ctrl', 'k')
        logger.info("Sent Jabber disconnect shortcut (Ctrl+K)")
        return True
    except Exception as e:
        logger.error(f"Failed to disconnect Jabber call: {e}")
        return False

# === MILESTONE LOGGING ===
def send_milestone_log(milestone, status="INFO", details=None):
    """Send milestone logs via SignalR"""
    try:
        log_data = {
            "timestamp": str(datetime.datetime.now()),
            "milestone": milestone,
            "status": status,
            "voiceId": VOICE_ID,
            "details": details or {}
        }
        
        print(f"📊 Milestone: {milestone} [{status}]")
        logger.info(f"Milestone: {milestone} [{status}] - {details}")
        
        if signalr_connection:
            try:
                signalr_connection.send("SendMilestone", [log_data])
            except:
                pass  # SignalR not available yet
                
    except Exception as e:
        logger.error(f"Milestone logging error: {e}")

# === AUTO-ANSWER FUNCTIONS ===
def find_and_click_answer_button(timeout=ANSWER_BUTTON_TIMEOUT, confidence=ANSWER_BUTTON_CONFIDENCE):
    start_time = time.time()
    
    while time.time() - start_time < timeout:
        try:
            # Try without confidence (no OpenCV required)
            button_location = pyautogui.locateOnScreen(ANSWER_BUTTON_IMAGE)
            
            if button_location:
                button_x, button_y = pyautogui.center(button_location)
                print(f"✅ Answer button found at ({button_x}, {button_y})")
                
                # Click the button multiple times to ensure it registers
                for i in range(3):
                    pyautogui.click(button_x, button_y)
                    time.sleep(0.2)
                
                logger.info("Answer button clicked successfully")
                print("📞 Answer button clicked!")
                send_milestone_log("CALL_ANSWERED", "SUCCESS", {"method": "auto_answer"})
                return True
                
        except pyautogui.ImageNotFoundException:
            pass  # Button not found yet
        except Exception:
            pass  # Silently ignore errors
        
        time.sleep(0.3)
    
    return False

def start_auto_answer_monitoring():
    global auto_answer_active, auto_answer_thread
    
    if auto_answer_active:
        print("⚠️ Auto-answer monitoring already active")
        return
    
    def monitor_and_answer():
        global auto_answer_active
        auto_answer_active = True
        print("🔊 Auto-answer monitoring started (running silently in background)")
        logger.info("Auto-answer monitoring started")
        
        while auto_answer_active:
            try:
                if find_and_click_answer_button(timeout=10, confidence=ANSWER_BUTTON_CONFIDENCE):
                    print("✅ Call answered successfully")
                    logger.info("Call answered successfully")
                    time.sleep(5)
                else:
                    time.sleep(1)
                    
            except Exception as e:
                logger.error(f"Auto-answer error: {e}")
                time.sleep(1)
    
    auto_answer_thread = threading.Thread(target=monitor_and_answer, daemon=True)
    auto_answer_thread.start()
    print("✅ Auto-answer monitoring thread started")

def stop_auto_answer_monitoring():
    global auto_answer_active
    if auto_answer_active:
        auto_answer_active = False
        print("🛑 Auto-answer monitoring stopped")
        logger.info("Auto-answer monitoring stopped")

# === SIGNALR FUNCTIONS ===
def setup_signalr():
    global signalr_connection
    try:
        from signalrcore.hub_connection_builder import HubConnectionBuilder
        
        print(f"🔗 SignalR Hub - Attempting connection to: {SIGNALR_HUB_URL}")
        logger.info(f"SignalR setup - Connecting to: {SIGNALR_HUB_URL}")
        
        signalr_connection = HubConnectionBuilder().with_url(SIGNALR_HUB_URL).with_automatic_reconnect({
            "type": "raw",
            "keep_alive_interval": 15,
            "reconnect_interval": 5,
            "max_attempts": 5
        }).build()
        
        def on_signalr_open():
            print("✅ SignalR Connected")
            logger.info("SignalR connection opened")
            
        def on_signalr_close():
            print("❌ SignalR Disconnected")
            logger.warning("SignalR connection closed")
            
        def on_signalr_error(msg):
            if "CompletionMessage" not in str(msg):
                print(f"❌ SignalR Error: {msg}")
                logger.error(f"SignalR error: {msg}")
            
        signalr_connection.on_open(on_signalr_open)
        signalr_connection.on_close(on_signalr_close)
        signalr_connection.on_error(on_signalr_error)
        
        signalr_connection.start()
        time.sleep(2)
        print("✅ SignalR connection started")
        logger.info("SignalR connection started successfully")
        return True
    except Exception as e:
        print(f"⚠️ SignalR setup failed: {e} (will continue without SignalR)")
        logger.warning(f"SignalR setup failed: {e}")
        return False

def send_analysis_result(analysis_data):
    try:
        if signalr_connection:
            if "callId" not in analysis_data:
                analysis_data["callId"] = VOICE_ID
            
            print(f"📤 Sending analysis result: {analysis_data}")
            signalr_connection.send("SendAnalysisResult", [analysis_data])
            logger.info(f"Analysis result sent: {analysis_data}")
    except Exception as e:
        print(f"⚠️ Analysis result send error: {e}")
        logger.error(f"Analysis result send error: {e}")

def send_results_to_originating_server(audio_file_path, analysis_results):
    """Send audio file and evaluation results back to originating server"""
    if not ORIGINATING_SERVER:
        logger.warning("No originating server specified")
        return False
    
    try:
        url = f"http://{ORIGINATING_SERVER}/api/receive_results"
        
        # Prepare the multipart form data
        with open(audio_file_path, 'rb') as audio_file:
            files = {'audio': (os.path.basename(audio_file_path), audio_file, 'audio/wav')}
            data = {
                'voiceId': VOICE_ID,
                'results': json.dumps(analysis_results)
            }
            
            print(f"📤 Sending results to {url}")
            logger.info(f"Sending audio and results to {url}")
            
            response = requests.post(url, files=files, data=data, timeout=30)
            
            if response.status_code == 200:
                print(f"✅ Results sent successfully to {ORIGINATING_SERVER}")
                logger.info(f"Results sent successfully: {response.json()}")
                return True
            else:
                print(f"❌ Failed to send results: {response.status_code}")
                logger.error(f"Failed to send results: {response.text}")
                return False
                
    except Exception as e:
        print(f"❌ Error sending results to server: {e}")
        logger.error(f"Error sending results to server: {e}")
        return False

# === API ENDPOINTS ===
@app.route('/api/test_answer_button', methods=['GET'])
def test_answer_button():
    """Test if answer button can be detected on screen"""
    try:
        print("\n🔍 Testing answer button detection...")
        print(f"   Image path: {ANSWER_BUTTON_IMAGE}")
        print(f"   Image exists: {os.path.isfile(ANSWER_BUTTON_IMAGE)}")
        
        if not os.path.isfile(ANSWER_BUTTON_IMAGE):
            return jsonify({
                "status": "error",
                "message": f"Answer button image not found: {ANSWER_BUTTON_IMAGE}"
            }), 404
        
        # Try to find button with 5 second timeout
        result = find_and_click_answer_button(timeout=5, confidence=ANSWER_BUTTON_CONFIDENCE)
        
        if result:
            return jsonify({
                "status": "success",
                "message": "Answer button found and clicked!"
            }), 200
        else:
            return jsonify({
                "status": "not_found",
                "message": "Answer button not found on screen",
                "tips": [
                    "Make sure Jabber has an incoming call",
                    "Ensure answer button is visible on screen",
                    "Check if answer button image matches current Jabber UI",
                    "Try taking a new screenshot of the answer button"
                ]
            }), 200
            
    except Exception as e:
        return jsonify({
            "status": "error",
            "message": str(e)
        }), 500

@app.route('/api/health', methods=['GET'])
def health_check():
    return jsonify({
        "status": "healthy",
        "timestamp": str(datetime.datetime.now()),
        "call_in_progress": call_in_progress
    })

@app.route('/api/start_call', methods=['POST'])
def start_call():
    global TARGET_BNUMBER, ANB, BNB, VOICE_ID, CALL_TYPE, call_in_progress, ORIGINATING_SERVER
    
    try:
        data = request.json
        
        # Validate required parameters
        required_params = ["TARGET_BNUMBER", "ANB", "BNB", "VoiceID", "CALL_TYPE"]
        missing_params = [p for p in required_params if not data.get(p)]
        
        if missing_params:
            return jsonify({
                "status": "error",
                "message": f"Missing required parameters: {', '.join(missing_params)}"
            }), 400
        
        if call_in_progress:
            return jsonify({
                "status": "error",
                "message": "Another call is already in progress"
            }), 409
        
        # Set parameters
        TARGET_BNUMBER = data["TARGET_BNUMBER"]
        ANB = data["ANB"]
        BNB = data["BNB"]
        VOICE_ID = data["VoiceID"]
        CALL_TYPE = data["CALL_TYPE"].upper()
        ORIGINATING_SERVER = None
        
        if CALL_TYPE not in ["IVR", "MANUAL"]:
            return jsonify({
                "status": "error",
                "message": "CALL_TYPE must be 'IVR' or 'MANUAL'"
            }), 400
        
        print(f"\n📞 New call request received:")
        print(f"   VoiceID: {VOICE_ID}")
        print(f"   TARGET_BNUMBER: {TARGET_BNUMBER}")
        print(f"   ANB: {ANB}")
        print(f"   BNB: {BNB}")
        print(f"   CALL_TYPE: {CALL_TYPE}")
        send_milestone_log("CALL_REQUEST_RECEIVED", "SUCCESS", {
            "voiceId": VOICE_ID,
            "targetNumber": TARGET_BNUMBER,
            "callType": CALL_TYPE
        })
        
        # Start call execution in separate thread
        call_thread = threading.Thread(target=execute_call, daemon=True)
        call_thread.start()
        
        return jsonify({
            "status": "success",
            "message": "Call initiated successfully",
            "voiceId": VOICE_ID
        }), 200
        
    except Exception as e:
        logger.error(f"API start_call error: {e}")
        return jsonify({
            "status": "error",
            "message": str(e)
        }), 500

@app.route('/api/call_status/<voice_id>', methods=['GET'])
def get_call_status(voice_id):
    return jsonify({
        "voiceId": voice_id,
        "status": "in_progress" if call_in_progress else "completed",
        "timestamp": str(datetime.datetime.now())
    })

# === RECORDING ===
def start_recording():
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    outfile = os.path.join(OUTPUT_DIR, f"JabberCall_{timestamp}.wav")
 
    cmd = [
        FFMPEG_PATH,
        "-y", "-f", "dshow",
        "-i", f"audio={RECORD_DEVICE}",
        "-ac", "2", "-ar", "44100", "-t", "180",
        outfile
    ]
 
    proc = subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.PIPE)
    logger.info(f"🎙️ Started recording: {outfile} (PID={proc.pid})")
    print(f"🎙️ Recording started: {outfile}")
    send_milestone_log("RECORDING_STARTED", "SUCCESS", {"file": outfile})
    return proc, outfile
 
def stop_recording(proc):
    if proc and proc.poll() is None:
        try:
            proc.terminate()
            proc.wait(timeout=3)
        except Exception:
            proc.kill()
        print("🛑 Recording stopped.")
        logger.info("Recording stopped.")
        send_milestone_log("RECORDING_STOPPED", "SUCCESS")

# === SSH MONITOR ===
def monitor_remote_logs(answered_event, disconnect_event, recording_ref, a_answered_event):
    try:
        send_milestone_log("SSH_CONNECTION_STARTED", "INFO")
        
        client = paramiko.SSHClient()
        client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        client.connect(HOST, PORT, USERNAME, PASSWORD, look_for_keys=False, allow_agent=False)
        chan = client.invoke_shell()
        chan.settimeout(0.5)
        
        send_milestone_log("SSH_CONNECTED", "SUCCESS", {"host": HOST})
 
        # SSH into second node
        print(f"🔗 Connecting to second node {SECOND_HOST}...")
        chan.send(f"ssh {SECOND_USERNAME}@{SECOND_HOST}\n")
        time.sleep(2)
        
        buffer = ""
        start = time.time()
        while time.time() - start < 10:
            if chan.recv_ready():
                data = chan.recv(4096).decode(errors="replace")
                buffer += data
                sys.stdout.write(data)
                sys.stdout.flush()
                if "password:" in buffer.lower():
                    chan.send(f"{SECOND_PASSWORD}\n")
                    print(f"🔑 Password sent")
                    time.sleep(2)
                    break
            time.sleep(0.2)
        
        send_milestone_log("SECOND_NODE_CONNECTED", "SUCCESS", {"host": SECOND_HOST})
        
        print("📡 Sending mml; command...")
        chan.send("mml\n")
        time.sleep(3)
        
        buffer = ""
        start = time.time()
        while time.time() - start < 5:
            if chan.recv_ready():
                data = chan.recv(4096).decode(errors="replace")
                buffer += data
                sys.stdout.write(data)
                sys.stdout.flush()
            time.sleep(0.2)
        
        send_milestone_log("MML_MODE_ACTIVATED", "SUCCESS")
        
        commands = [
            f"EXTPE: BNB={BNB};",
            f"EXTPI: BNB={BNB};",
            f"tctdi:bnb=#11 {TARGET_BNUMBER},cl=1,anb={ANB};"
        ]
 
        print("🚀 Sending commands to remote server...")
        send_milestone_log("COMMANDS_EXECUTION_STARTED", "INFO")
        
        ready_for_connection = False
        
        for cmd in commands:
            chan.send(cmd + "\n")
            print(f"🖥️ Sent: {cmd}")
            logger.info(f"Sent: {cmd}")
            
            wait_time = 120 if "tctdi" in cmd.lower() else 2
            start = time.time()
 
            buffer = ""
            while time.time() - start < wait_time:
                if chan.recv_ready():
                    data = chan.recv(4096).decode(errors="replace")
                    buffer += data
                    sys.stdout.write(data)
                    sys.stdout.flush()
                    
                    if "tctdi" in cmd.lower() and "RINGING OPERATOR" in buffer:
                        print(f"📞 Calling A number ({ANB})")
                        send_milestone_log("A_NUMBER_RINGING", "INFO", {"anb": ANB})
                    
                    if ("READY FOR CONNECTION" in buffer or "ANSWER RECEIVED" in buffer):
                        a_answered_event.set()
                        ready_for_connection = True
                        print("✅ A number answered - READY FOR CONNECTION received")
                        send_milestone_log("A_NUMBER_ANSWERED", "SUCCESS", {"anb": ANB})
                        # Break out of wait loop to send CON; immediately
                        break
                    
                    if re.search(FORCED_RELEASE_PATTERN, buffer, re.IGNORECASE):
                        print("🚨 FORCED RELEASE DETECTED")
                        send_milestone_log("FORCED_RELEASE", "ERROR")
                        disconnect_event.set()
                        if recording_ref[0]:
                            stop_recording(recording_ref[0])
                        client.close()
                        return
                else:
                    time.sleep(0.2)
            
            # If READY FOR CONNECTION received, break out of command loop to send CON;
            if ready_for_connection:
                break
        
        # Send CON; only after READY FOR CONNECTION
        if ready_for_connection:
            print("🔗 READY FOR CONNECTION received - Sending CON; command...")
            time.sleep(1)  # Small delay before sending CON;
            chan.send("CON;\n")
            logger.info("Sent: CON;")
            send_milestone_log("CON_COMMAND_SENT", "INFO")
            
            wait_time = 120
            start = time.time()
            buffer = ""
            
            while time.time() - start < wait_time:
                if chan.recv_ready():
                    data = chan.recv(4096).decode(errors="replace")
                    buffer += data
                    sys.stdout.write(data)
                    sys.stdout.flush()
                    
                    if "FREE SUBSCRIBER" in buffer:
                        print(f"📞 Calling B number ({TARGET_BNUMBER})")
                        send_milestone_log("B_NUMBER_CALLING", "INFO", {"bnumber": TARGET_BNUMBER})
                    
                    if re.search(PATTERN_ANSWER, buffer, re.IGNORECASE):
                        answered_event.set()
                        print("✅ B number answered")
                        send_milestone_log("B_NUMBER_ANSWERED", "SUCCESS", {"bnumber": TARGET_BNUMBER})
                        break  # Exit loop once B answers
                    
                    if re.search(FORCED_RELEASE_PATTERN, buffer, re.IGNORECASE):
                        print("🚨 FORCED RELEASE DETECTED")
                        send_milestone_log("FORCED_RELEASE", "ERROR")
                        disconnect_event.set()
                        if recording_ref[0]:
                            stop_recording(recording_ref[0])
                        client.close()
                        return
                else:
                    time.sleep(0.2)
        else:
            print(f"❌ READY FOR CONNECTION not received - A number {ANB} did not answer")
            send_milestone_log("A_NUMBER_NO_ANSWER", "ERROR", {"anb": ANB})
            disconnect_event.set()
            client.close()
            return
 
        client.close()
    except Exception as e:
        print(f"SSH Error: {e}")
        logger.error(f"SSH Error: {e}")
        send_milestone_log("SSH_ERROR", "ERROR", {"error": str(e)})

# === EXECUTE CALL ===
def execute_call():
    global call_in_progress
    call_in_progress = True
    recording_proc = None
    ai_agent = None
    answered_event = threading.Event()
    disconnect_event = threading.Event()
    a_answered_event = threading.Event()
    recording_ref = [None]
 
    try:
        send_milestone_log("CALL_EXECUTION_STARTED", "INFO")
        
        ssh_thread = threading.Thread(
            target=monitor_remote_logs,
            args=(answered_event, disconnect_event, recording_ref, a_answered_event),
            daemon=True
        )
        ssh_thread.start()
        
        # Wait for A number
        print("⏳ Waiting for A number to answer...")
        if not a_answered_event.wait(timeout=120):
            print("⚠️ A number did not answer")
            send_milestone_log("A_NUMBER_TIMEOUT", "ERROR")
            return

        if minimize_jabber_window():
            send_milestone_log("JABBER_MINIMIZED", "SUCCESS")

        if CALL_TYPE == "MANUAL" and USE_AI_CONVERSATION:
            try:
                print("AI preload: preparing conversation engine before B answers...")
                send_milestone_log("AI_PRELOAD_STARTED", "INFO")
                ai_agent = AIConversationAgent(
                    api_key=OPENAI_API_KEY,
                    input_device_name=VB_CABLE_DEVICE
                )
                ai_agent.ensure_playback_routing()
                ai_agent.warmup()
                send_milestone_log("AI_PRELOAD_READY", "SUCCESS")
            except Exception as e:
                ai_agent = None
                logger.error(f"AI preload error: {e}")
                send_milestone_log("AI_PRELOAD_ERROR", "ERROR", {"error": str(e)})

        # Wait for B number
        if not answered_event.wait(timeout=60):
            print("⚠️ B number did not answer")
            send_milestone_log("B_NUMBER_TIMEOUT", "ERROR")
            return
        
        # Start recording
        recording_proc, wav_path = start_recording()
        recording_ref[0] = recording_proc
        
        # Handle call based on type
        if CALL_TYPE == "IVR":
            print("📞 IVR mode: Sending DTMF tones...")
            send_milestone_log("IVR_MODE_STARTED", "INFO")
            time.sleep(3)
            
            def find_jabber_window():
                def callback(hwnd, windows):
                    if win32gui.IsWindowVisible(hwnd):
                        title = win32gui.GetWindowText(hwnd)
                        if "Cisco Jabber" in title or "Jabber" in title:
                            windows.append(hwnd)
                    return True
                windows = []
                win32gui.EnumWindows(callback, windows)
                return windows[0] if windows else None
            
            jabber_hwnd = find_jabber_window()
            if jabber_hwnd:
                win32gui.ShowWindow(jabber_hwnd, win32con.SW_RESTORE)
                win32gui.SetForegroundWindow(jabber_hwnd)
                time.sleep(0.5)
            
            pyautogui.press('1')
            time.sleep(2)
            pyautogui.press('2')
            time.sleep(2)
            pyautogui.press('1')
            send_milestone_log("DTMF_SENT", "SUCCESS", {"sequence": "1-2-1"})
            
        elif CALL_TYPE == "MANUAL":
            print("📞 MANUAL mode")
            send_milestone_log("MANUAL_MODE_STARTED", "INFO")
            
            # Check if AI conversation is enabled
            if USE_AI_CONVERSATION:
                print("🤖 Using AI-powered conversation")
                send_milestone_log("AI_CONVERSATION_STARTED", "INFO")
                
                try:
                    if ai_agent is None:
                        ai_agent = AIConversationAgent(
                            api_key=OPENAI_API_KEY,
                            input_device_name=VB_CABLE_DEVICE
                        )
                    
                    # Start 3-minute timeout timer
                    timeout_event = threading.Event()
                    
                    def check_timeout():
                        time.sleep(180)  # 3 minutes
                        if not disconnect_event.is_set():
                            timeout_event.set()
                            print("⏰ 3-minute timeout reached")
                            send_milestone_log("CALL_TIMEOUT", "WARNING")
                            try:
                                disconnect_jabber_call()
                                print("📞 Timeout - disconnecting call")
                            except:
                                pass
                            disconnect_event.set()
                    
                    timeout_thread = threading.Thread(target=check_timeout, daemon=True)
                    timeout_thread.start()
                    
                    # Run AI conversation
                    conversation_result = ai_agent.run_conversation(
                        expected_cli=ANB,
                        max_duration=120  # 2 minutes max
                    )
                    
                    # Extract results
                    end_reason = conversation_result.get('end_reason', 'completed')
                    recording_consent = conversation_result.get('recording_consent', 'Unknown')
                    otp_value = conversation_result.get('otp_value', 'Not generated')
                    otp_confirmed = conversation_result.get('otp_confirmed', False)
                    caller_status = conversation_result.get('caller_status', 'Not captured')
                    audio_quality = conversation_result['audio_quality']
                    spoken_cli = conversation_result['cli_reported']
                    cli_match = conversation_result['cli_match']
                    
                    print(f"Caller Status: {caller_status}")
                    print(f"???? Audio Quality: {audio_quality}")
                    print(f"???? CLI: {spoken_cli} ({cli_match})")
                    print(f"???? Recording Consent: {recording_consent}")
                    print(f"???? OTP: {otp_value} (Confirmed={otp_confirmed})")
                    
                    send_milestone_log("CALLER_STATUS_CAPTURED", "SUCCESS", {"status": caller_status})
                    send_milestone_log("AUDIO_QUALITY_CAPTURED", "SUCCESS", {"quality": audio_quality})
                    send_milestone_log("CLI_CAPTURED", "SUCCESS", {"cli": spoken_cli, "match": cli_match})
                    send_milestone_log("RECORDING_CONSENT_CAPTURED", "SUCCESS", {"consent": recording_consent})
                    send_milestone_log("OTP_CAPTURED", "SUCCESS", {"otp": otp_value, "confirmed": otp_confirmed})
                    
                    # Send results to database
                    analysis_data = {
                        "callId": VOICE_ID,
                        "CallerStatus": caller_status,
                        "RecordingConsent": recording_consent,
                        "OTPValue": otp_value,
                        "OTPConfirmed": otp_confirmed,
                        "AudioQualityReported": audio_quality,
                        "CLIInput": ANB,
                        "CLIReported": spoken_cli,
                        "CLIMatch": cli_match,
                        "ConversationEndReason": end_reason
                    }
                    send_analysis_result(analysis_data)
                    
                    # Cleanup AI agent
                    ai_agent.cleanup()
                    time.sleep(POST_GOODBYE_GRACE_SECONDS)
                    
                    # Disconnect call
                    print("✅ AI conversation completed - disconnecting call")
                    try:
                        if disconnect_jabber_call():
                            print("📞 Sent Jabber disconnect command")
                        else:
                            print("❌ Failed to send Jabber disconnect command")
                    except Exception as e:
                        print(f"❌ Failed to disconnect: {e}")
                    disconnect_event.set()
                    if end_reason in {"recording_denied", "otp_not_confirmed", "cli_not_confirmed"}:
                        if recording_proc:
                            stop_recording(recording_proc)
                        send_milestone_log("CALL_ENDED_EARLY", "INFO", {"reason": end_reason})
                        return

                except Exception as e:
                    print(f"AI conversation error: {e}")
                    logger.error(f"AI conversation error: {e}")
                    send_milestone_log("AI_CONVERSATION_ERROR", "ERROR", {"error": str(e)})
                    print("AI failed - microphone capture or call audio routing may not be configured")
                    print(f"Tip: Ensure FFmpeg/Jabber use '{VB_CABLE_DEVICE}' for the VB-Cable recording path")
                    # Disconnect on error
                    try:
                        disconnect_jabber_call()
                    except:
                        pass
                    disconnect_event.set()
            else:
                print("AI conversation is disabled")
                send_milestone_log("AI_DISABLED", "WARNING")
                # Just wait and disconnect
                time.sleep(30)
                try:
                    disconnect_jabber_call()
                except:
                    pass
                disconnect_event.set()
        
        # Wait for recording to complete
        time.sleep(60)
        
        # Disconnect call
        print("📞 Disconnecting call...")
        disconnect_jabber_call()
        send_milestone_log("CALL_DISCONNECTED", "SUCCESS")
        disconnect_event.set()
        
        # Stop recording
        if recording_proc:
            stop_recording(recording_proc)
        
        send_milestone_log("CALL_COMPLETED", "SUCCESS", {"voiceId": VOICE_ID})
        print(f"✅ Call {VOICE_ID} completed successfully")
        
        # Analyze audio
        print(f"🔬 Starting audio analysis...")
        send_milestone_log("ANALYSIS_STARTED", "INFO")
        analysis_result = analyze_audio(wav_path, VOICE_ID, signalr_connection)
        if analysis_result:
            send_milestone_log("ANALYSIS_COMPLETED", "SUCCESS")
            
            # Send results back to originating server
            if ORIGINATING_SERVER:
                print(f"📤 Sending results back to originating server...")
                send_milestone_log("SENDING_RESULTS_TO_SERVER", "INFO", {"server": ORIGINATING_SERVER})
                
                if send_results_to_originating_server(wav_path, analysis_result):
                    send_milestone_log("RESULTS_SENT_TO_SERVER", "SUCCESS", {"server": ORIGINATING_SERVER})
                else:
                    send_milestone_log("RESULTS_SEND_FAILED", "ERROR", {"server": ORIGINATING_SERVER})
        else:
            send_milestone_log("ANALYSIS_FAILED", "ERROR")
            
    except Exception as e:
        print(f"Execution error: {e}")
        logger.error(f"Execution error: {e}")
        send_milestone_log("EXECUTION_ERROR", "ERROR", {"error": str(e)})
    finally:
        if ai_agent:
            ai_agent.cleanup()
        if recording_proc:
            stop_recording(recording_proc)
        call_in_progress = False

# === MAIN ===
if __name__ == "__main__":
    print("🔍 Validating configuration...")
    
    if not os.path.isfile(FFMPEG_PATH):
        print(f"❌ FFmpeg not found: {FFMPEG_PATH}")
        sys.exit(1)
    print(f"✅ FFmpeg found")
    
    if not os.path.isfile(ANSWER_BUTTON_IMAGE):
        print(f"⚠️ Answer button image not found: {ANSWER_BUTTON_IMAGE}")
    else:
        print(f"✅ Answer button image found")
    
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    print(f"✅ Recordings directory ready")
    
    print("\n🚀 Starting Voice Quality Automation Service...")
    
    # Start auto-answer monitoring
    start_auto_answer_monitoring()
    
    # Setup SignalR (optional)
    setup_signalr()
    
    # Start Flask API server
    print(f"\n🌐 Starting API server on port {API_PORT}...")
    print(f"📡 API Endpoints:")
    print(f"   GET  http://localhost:{API_PORT}/api/health")
    print(f"   POST http://localhost:{API_PORT}/api/start_call")
    print(f"   GET  http://localhost:{API_PORT}/api/call_status/<voice_id>")
    print(f"\n✅ Service ready - waiting for API requests...\n")
    
    try:
        app.run(host='0.0.0.0', port=API_PORT, debug=False, threaded=True)
    except KeyboardInterrupt:
        print("\n🛑 Service stopped")
        stop_auto_answer_monitoring()
    except Exception as e:
        print(f"❌ Fatal error: {e}")
        logger.error(f"Fatal error: {e}")
        sys.exit(1)
