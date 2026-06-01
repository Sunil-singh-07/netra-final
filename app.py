import cv2
import numpy as np
import tensorflow as tf
import json
import time
import base64
import threading
import queue
import os
import speech_recognition as sr
import win32com.client
import requests
from flask import Flask, render_template, Response, request, jsonify
from google import genai
from ultralytics import YOLO
from tensorflow.keras.applications.efficientnet import preprocess_input
from dotenv import load_dotenv
from datetime import datetime

load_dotenv()

# ══════════════════════════════════════════
#   CONFIG
# ══════════════════════════════════════════
MODEL_PATH     = r"C:\netra_final\netra-final\currency_model_v5.tflite"
LABELS_PATH    = r"C:\netra_final\netra-final\currency_model_v5_class_map.json"
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
GROQ_API_KEY   = os.getenv("GROQ_API_KEY")

CENTER_ZONE       = 0.35
ANNOUNCE_COOLDOWN = 45   # 45s between scene scans — stays under Gemini free quota
CURRENCY_CONF     = 0.80

GROQ_MODEL = "llama-3.3-70b-versatile"
GROQ_URL   = "https://api.groq.com/openai/v1/chat/completions"

# Gemini model selection
# Currency: best vision accuracy — only called ~20-50 times/day
GEMINI_CURRENCY_MODEL = "gemini-2.5-flash"
# Scene: fast + cheap — called every 45s continuously
GEMINI_SCENE_MODEL    = "gemini-2.0-flash"

# All YOLO COCO objects that matter for navigation safety
# YOLO can detect these — walls/floors handled by Gemini scene scan
HAZARD_OBJECTS = {
    # People
    'person':        {'en': 'Person',        'hi': 'Vyakti'},
    # Furniture
    'chair':         {'en': 'Chair',         'hi': 'Kursi'},
    'couch':         {'en': 'Sofa',          'hi': 'Sofa'},
    'bed':           {'en': 'Bed',           'hi': 'Palang'},
    'dining table':  {'en': 'Table',         'hi': 'Mez'},
    'desk':          {'en': 'Desk',          'hi': 'Mez'},
    # Electronics
    'tv':            {'en': 'TV',            'hi': 'TV'},
    'laptop':        {'en': 'Laptop',        'hi': 'Laptop'},
    # Doors / stairs
    'door':          {'en': 'Door',          'hi': 'Darwaza'},
    'stairs':        {'en': 'Stairs',        'hi': 'Seedhiyan'},
    # Vehicles (outdoor)
    'bicycle':       {'en': 'Bicycle',       'hi': 'Cycle'},
    'motorcycle':    {'en': 'Motorcycle',    'hi': 'Motorcycle'},
    'car':           {'en': 'Car',           'hi': 'Gaadi'},
    'bus':           {'en': 'Bus',           'hi': 'Bus'},
    'truck':         {'en': 'Truck',         'hi': 'Truck'},
    # Animals
    'dog':           {'en': 'Dog',           'hi': 'Kutta'},
    'cat':           {'en': 'Cat',           'hi': 'Billi'},
    # Common indoor objects that cause trips/falls
    'bottle':        {'en': 'Bottle',        'hi': 'Bottle'},
    'cup':           {'en': 'Cup',           'hi': 'Cup'},
    'bowl':          {'en': 'Bowl',          'hi': 'Bowl'},
    'backpack':      {'en': 'Bag',           'hi': 'Bag'},
    'suitcase':      {'en': 'Suitcase',      'hi': 'Suitcase'},
    'umbrella':      {'en': 'Umbrella',      'hi': 'Chhata'},
    'handbag':       {'en': 'Handbag',       'hi': 'Bag'},
    'sports ball':   {'en': 'Ball',          'hi': 'Ball'},
    'potted plant':  {'en': 'Plant',         'hi': 'Paudha'},
    'sink':          {'en': 'Sink',          'hi': 'Sink'},
    'toilet':        {'en': 'Toilet',        'hi': 'Toilet'},
    'refrigerator':  {'en': 'Fridge',        'hi': 'Fridge'},
    'oven':          {'en': 'Oven',          'hi': 'Oven'},
    'microwave':     {'en': 'Microwave',     'hi': 'Microwave'},
    'bench':         {'en': 'Bench',         'hi': 'Bench'},
    'book':          {'en': 'Book',          'hi': 'Kitaab'},
    'clock':         {'en': 'Clock',         'hi': 'Ghadi'},
    'vase':          {'en': 'Vase',          'hi': 'Vase'},
    'scissors':      {'en': 'Scissors',      'hi': 'Kainchi'},
    'cell phone':    {'en': 'Phone',         'hi': 'Phone'},
    'remote':        {'en': 'Remote',        'hi': 'Remote'},
    'keyboard':      {'en': 'Keyboard',      'hi': 'Keyboard'},
    'mouse':         {'en': 'Mouse',         'hi': 'Mouse'},
}

CURRENCY_MAP_EN = {
    '10': 'This is a ten rupee note',
    '20': 'This is a twenty rupee note',
    '50': 'This is a fifty rupee note',
    '100': 'This is a one hundred rupee note',
    '200': 'This is a two hundred rupee note',
    '500': 'This is a five hundred rupee note',
    'Background': ''
}

CURRENCY_MAP_HI = {
    '10': 'Das rupaye ka note hai',
    '20': 'Bees rupaye ka note hai',
    '50': 'Pachaas rupaye ka note hai',
    '100': 'Ek sau rupaye ka note hai',
    '200': 'Do sau rupaye ka note hai',
    '500': 'Paanch sau rupaye ka note hai',
    'Background': ''
}

DISTANCE_EN = {
    'very_close': '1 to 2 steps ahead',
    'close':      '3 to 4 steps ahead',
    'far':        '5 or more steps ahead',
}

DISTANCE_HI = {
    'very_close': '1 se 2 kadam aage',
    'close':      '3 se 4 kadam aage',
    'far':        '5 ya zyada kadam aage',
}

CURRENCY_TRIGGERS = [
    'check', 'scan', 'note', 'currency', 'rupee', 'rupees',
    'money', 'paisa', 'identify note', 'what note', 'which note',
    'how much', 'kitna', 'kitne', 'kaun sa note', 'check note'
]

# ── Wake word — Netra only responds when you say her name first ──
# e.g. "Netra how is the weather" or "Netra scan this note"
WAKE_WORD = "netra"   # lowercase — we compare against lowercased text

# ══════════════════════════════════════════
#   LOAD MODELS
# ══════════════════════════════════════════
print("Loading currency model...")
interpreter = tf.lite.Interpreter(model_path=MODEL_PATH)
interpreter.allocate_tensors()
input_details  = interpreter.get_input_details()
output_details = interpreter.get_output_details()
with open(LABELS_PATH) as f:
    currency_labels = json.load(f)
print("Currency model loaded.")

print("Loading YOLO model...")
yolo = YOLO('yolov8n.pt')
print("YOLO model loaded.")

gemini_client = genai.Client(api_key=GEMINI_API_KEY)

# ══════════════════════════════════════════
#   TTS — win32com SAPI (runs in own thread)
#
#   WHY NOT pyttsx3:
#   pyttsx3 has a known Windows bug where
#   engine.runAndWait() silently does nothing
#   when called from a non-main thread.
#   Since our TTS worker IS a background
#   thread, pyttsx3 produces text output
#   but zero audio — exactly what you saw.
#
#   win32com SAPI works from any thread
#   as long as CoInitialize is called first,
#   which we do at the top of tts_worker.
#
#   BUSY EVENT FLOW (correct):
#   1. voice_listener captures audio
#      → sets busy_event
#   2. process_audio runs in side thread
#      → calls speak(reply)
#      → puts text in tts_queue
#      → returns WITHOUT clearing busy
#   3. tts_worker picks text from queue
#      → speaks it (blocking)
#      → clears busy_event after done
#   4. mic reopens only after speech done
# ══════════════════════════════════════════
tts_queue   = queue.Queue()
is_speaking = False
busy_event  = threading.Event()


def tts_worker():
    """
    win32com SAPI must call CoInitialize in
    the thread where it will be used.
    """
    global is_speaking

    # CoInitialize required for COM in a non-main thread
    import pythoncom
    pythoncom.CoInitialize()

    speaker = win32com.client.Dispatch("SAPI.SpVoice")

    # Set speaking rate: -10 (slow) to 10 (fast), 0 is default
    # -1 is slightly slower than default — clearer for visually impaired
    speaker.Rate = -1
    speaker.Volume = 100

    print("[TTS] Worker ready — using Windows SAPI")

    while True:
        text = tts_queue.get()
        if text is None:
            break
        try:
            is_speaking = True
            print(f"[TTS] Speaking: {text[:50]}")
            speaker.Speak(text)          # blocking — waits until done
            print(f"[TTS] Done")
            time.sleep(0.5)              # 500ms echo buffer before mic reopens
        except Exception as e:
            print(f"[TTS] Error: {e}")
        finally:
            is_speaking = False
            busy_event.clear()           # mic reopens here, nowhere else
            print(f"[TTS] Mic unblocked")


tts_thread = threading.Thread(target=tts_worker, daemon=True)
tts_thread.start()
time.sleep(0.3)   # give TTS worker time to init before first speak()


def speak(text, priority=False):
    """
    Queue text for TTS. Does NOT touch busy_event.
    tts_worker is the sole owner of busy_event.
    """
    if not text or not text.strip():
        busy_event.clear()   # nothing to say — unblock mic immediately
        return
    print(f"[Netra] {text}")
    if priority:
        while not tts_queue.empty():
            try:
                tts_queue.get_nowait()
            except Exception:
                pass
    tts_queue.put(text)


# ══════════════════════════════════════════
#   CAMERA CONFIG
#   Set PHONE_CAMERA_URL to your IP Webcam
#   address, or leave as None to use laptop
#   webcam automatically.
#
#   IP Webcam setup (Android):
#   1. Install "IP Webcam" by Pavel Khlebovich
#   2. Open app → Video prefs → 1280x720, 60fps
#   3. Scroll down → Start Server
#   4. Note the IP shown (e.g. 192.168.1.5:8080)
#   5. Set PHONE_CAMERA_URL below
#   6. Test in browser: http://IP:8080/video
# ══════════════════════════════════════════
PHONE_CAMERA_URL = None
PHONE_CAMERA_URL = "http://192.168.210.196:8080/video"  # ← uncomment and set your IP


def find_best_camera():
    # ── Try phone camera first if URL is set ──
    if PHONE_CAMERA_URL:
        print(f"[Camera] Connecting to phone camera: {PHONE_CAMERA_URL}")
        try:
            cam = cv2.VideoCapture(PHONE_CAMERA_URL)
            cam.set(cv2.CAP_PROP_BUFFERSIZE, 1)   # minimize latency
            deadline = time.time() + 5
            while time.time() < deadline:
                if cam.isOpened():
                    ret, frame = cam.read()
                    if ret and frame is not None:
                        h, w = frame.shape[:2]
                        print(f"[Camera] Phone camera connected — {w}x{h}")
                        print("[Camera] Using phone camera — high quality mode")
                        return cam
                time.sleep(0.2)
            cam.release()
            print("[Camera] Phone camera not reachable — falling back to laptop webcam")
            print(f"[Camera] Check: is IP Webcam running? Can you open {PHONE_CAMERA_URL} in browser?")
        except Exception as e:
            print(f"[Camera] Phone camera error: {e} — falling back to laptop webcam")

    # ── Laptop webcam fallback ──
    print("[Camera] Looking for laptop webcam...")
    for idx in range(3):
        try:
            cam = cv2.VideoCapture(idx, cv2.CAP_MSMF)
            if cam.isOpened():
                ret, frame = cam.read()
                if ret and frame is not None:
                    cam.set(cv2.CAP_PROP_FRAME_WIDTH,  1280)
                    cam.set(cv2.CAP_PROP_FRAME_HEIGHT, 720)
                    cam.set(cv2.CAP_PROP_FPS,          30)
                    cam.set(cv2.CAP_PROP_AUTOFOCUS,    1)
                    cam.set(cv2.CAP_PROP_AUTO_EXPOSURE, 1)
                    w   = int(cam.get(cv2.CAP_PROP_FRAME_WIDTH))
                    h   = int(cam.get(cv2.CAP_PROP_FRAME_HEIGHT))
                    fps = int(cam.get(cv2.CAP_PROP_FPS))
                    print(f"[Camera] Laptop webcam at index {idx} — {w}x{h} @ {fps}fps")
                    return cam
            cam.release()
        except Exception as e:
            print(f"[Camera] Index {idx} error: {e}")
    return None


cap = find_best_camera()
if cap is None:
    raise RuntimeError("No camera found. Check webcam is not in use by another app.")


current_frame = None


def capture_frames():
    global current_frame
    while True:
        ret, frame = cap.read()
        if ret:
            # Rotate 90 degrees clockwise → portrait mode
            # Change cv2.ROTATE_90_CLOCKWISE to cv2.ROTATE_90_COUNTERCLOCKWISE
            # if your camera is upside down
            frame = cv2.rotate(frame, cv2.ROTATE_90_CLOCKWISE)
            current_frame = frame.copy()
        time.sleep(0.033)


threading.Thread(target=capture_frames, daemon=True).start()
time.sleep(0.5)

# ══════════════════════════════════════════
#   APP STATE
# ══════════════════════════════════════════
# ══════════════════════════════════════════
#   APP STATE — 3 mode system
#
#   SILENT  → default on startup
#             camera + YOLO run, nothing spoken
#             waits for "Netra analyse"
#
#   ANALYSE → triggered by "Netra analyse"
#             speaks scene every 45s
#             waits for "Netra stop" / "Netra scan note"
#
#   SCAN    → triggered by "Netra scan note"
#             YOLO announcements stop
#             waits silently
#             each "Netra scan" → scans and speaks result
#             waits for "Netra analyse surroundings"
# ══════════════════════════════════════════
APP_MODE_SILENT  = 'silent'
APP_MODE_ANALYSE = 'analyse'
APP_MODE_SCAN    = 'scan'

state = {
    'app_mode':       APP_MODE_SILENT,   # current app mode
    'language':       'en',
    'last_announced': {},
    'last_dist_key':  {},
    'conversation':   [],
    'mic_active':     True,
    'scanning':       False,
}

# ── Scene change tracking ──
scene_state = {
    'last_labels':    set(),
    'last_counts':    {},
    'announce_count': {},
}


# ══════════════════════════════════════════
#   OBJECT DETECTION — YOLO (full scene)
#
#   Detects ALL objects in frame, not just
#   centre zone. Each object gets:
#   - distance (very_close/close/far)
#   - position (left/centre/right)
#   Sorted by size (closest first).
# ══════════════════════════════════════════
def get_distance(box_h, frame_h):
    ratio = box_h / frame_h
    if ratio > 0.55:
        return 'very_close', '1 to 2 steps ahead'
    elif ratio > 0.30:
        return 'close', '3 to 4 steps ahead'
    else:
        return 'far', '5 or more steps ahead'


def get_position(obj_cx, frame_w):
    """Return left / centre / right based on x position."""
    third = frame_w / 3
    if obj_cx < third:
        return 'left'
    elif obj_cx > 2 * third:
        return 'right'
    else:
        return 'centre'


def detect_objects(frame, lang):
    h, w = frame.shape[:2]
    results = []

    yolo_results = yolo(frame, verbose=False, conf=0.45)[0]  # lower threshold = more detections

    for box in yolo_results.boxes:
        conf  = float(box.conf[0])
        cls   = int(box.cls[0])
        label = yolo.names[cls]

        if label not in HAZARD_OBJECTS:
            continue

        x1, y1, x2, y2 = map(int, box.xyxy[0])
        obj_cx   = (x1 + x2) / 2
        box_h    = y2 - y1

        dist_key, dist_text = get_distance(box_h, h)
        position            = get_position(obj_cx, w)

        name = HAZARD_OBJECTS[label]['en'] if lang == 'en' else HAZARD_OBJECTS[label]['hi']

        results.append({
            'label':    label,
            'name':     name,
            'dist_key': dist_key,
            'dist_text': dist_text,
            'position': position,
            'box':      [x1, y1, x2, y2],
            'conf':     conf,
        })

    # Sort by box height descending (largest = closest)
    results.sort(key=lambda x: x['box'][3] - x['box'][1], reverse=True)
    return results


# ══════════════════════════════════════════
#   CURRENCY — STEP 1: CAPTURE CLEAR FRAME
#   Uses Laplacian variance to measure blur.
#   Waits up to 4s for a sharp frame.
# ══════════════════════════════════════════
def capture_clear_frame(timeout=4.0, blur_threshold=80.0):
    print("[Scan] Waiting for clear frame...")
    best_frame = None
    best_score = -1
    deadline   = time.time() + timeout

    while time.time() < deadline:
        frame = current_frame
        if frame is None:
            time.sleep(0.05)
            continue
        frame = frame.copy()
        gray  = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        score = cv2.Laplacian(gray, cv2.CV_64F).var()

        if score > best_score:
            best_score = score
            best_frame = frame.copy()

        if score >= blur_threshold:
            print(f"[Scan] Sharp frame ready (score={score:.1f})")
            return best_frame

        time.sleep(0.08)

    print(f"[Scan] Using best available frame (score={best_score:.1f})")
    return best_frame


# ══════════════════════════════════════════
#   CURRENCY — STEP 2: YOUR ML MODEL
# ══════════════════════════════════════════
currency_lock = threading.Lock()


def predict_currency(frame):
    img = cv2.resize(frame, (224, 224))
    img = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
    img = preprocess_input(img.astype(np.float32))
    img = np.expand_dims(img, axis=0)
    with currency_lock:
        interpreter.set_tensor(input_details[0]['index'], img)
        interpreter.invoke()
        preds = interpreter.get_tensor(output_details[0]['index'])[0]
    confidence = float(np.max(preds))
    class_idx  = int(np.argmax(preds))
    label      = currency_labels[str(class_idx)]
    print(f"[ML Model] {label} ({confidence*100:.1f}%)")
    return label, confidence


# ══════════════════════════════════════════
#   CURRENCY — STEP 3: GEMINI VERIFICATION
#   Same frame sent to Gemini to confirm
#   or correct your ML model prediction.
# ══════════════════════════════════════════
def gemini_currency(frame, tflite_label, tflite_conf):
    try:
        _, buffer = cv2.imencode('.jpg', frame, [cv2.IMWRITE_JPEG_QUALITY, 92])
        img_b64   = base64.b64encode(buffer).decode('utf-8')

        prompt = f"""You are verifying an Indian currency note.
My local ML model predicted: Rs.{tflite_label} with {tflite_conf*100:.1f}% confidence.
Look carefully at: denomination number, note color, Gandhi portrait, RBI text.
Return ONLY this JSON — no markdown, no extra text:
{{
    "denomination": "10" or "20" or "50" or "100" or "200" or "500" or "unknown",
    "agrees_with_model": true or false,
    "confidence": "high" or "medium" or "low",
    "is_authentic": true or false,
    "side": "front" or "back" or "unknown"
}}"""

        response = gemini_client.models.generate_content(
            model=GEMINI_CURRENCY_MODEL,
            contents=[{"parts": [
                {"text": prompt},
                {"inline_data": {"mime_type": "image/jpeg", "data": img_b64}}
            ]}]
        )
        raw  = response.text.strip().replace("```json", "").replace("```", "").strip()
        data = json.loads(raw)
        print(f"[Gemini] denomination={data.get('denomination')} "
              f"agrees={data.get('agrees_with_model')} "
              f"confidence={data.get('confidence')}")
        return data
    except Exception as e:
        print(f"[Gemini] Error: {e}")
        return None


# ══════════════════════════════════════════
#   GROQ — FAST CONVERSATION (English only)
# ══════════════════════════════════════════
def groq_conversation(user_text):
    try:
        now_time = datetime.now().strftime("%I:%M %p")
        now_date = datetime.now().strftime("%A, %d %B %Y")

        system_prompt = f"""You are Netra, a warm and caring voice assistant
for visually impaired people in India.

STRICT RULES — follow every one:
1. Reply in 1 to 2 short sentences ONLY. Never longer. This is spoken aloud.
2. No bullet points, lists, markdown, asterisks or emojis.
3. Be warm, direct and helpful — like a trusted friend.
4. Always reply in English only.
5. User is in Kanpur, Uttar Pradesh, India.
   Weather questions: give a confident real seasonal answer for Kanpur.
   March in Kanpur = warm and sunny, 28 to 33 degrees Celsius.
   Never say "check your local forecast" — give the actual answer.
6. Current time: {now_time}. Today: {now_date}.
7. Never say "I cannot help" or "I don't have real-time data".
   Always give a confident helpful answer."""

        messages = [{"role": "system", "content": system_prompt}]
        for entry in state['conversation'][-6:]:
            role = "user" if entry['role'] == "user" else "assistant"
            messages.append({"role": role, "content": entry['text']})
        messages.append({"role": "user", "content": user_text})

        print(f"[Groq] Sending: {user_text[:60]}")

        resp = requests.post(
            GROQ_URL,
            headers={
                "Authorization": f"Bearer {GROQ_API_KEY}",
                "Content-Type":  "application/json"
            },
            json={
                "model":       GROQ_MODEL,
                "messages":    messages,
                "max_tokens":  120,
                "temperature": 0.7,
                "stream":      False
            },
            timeout=8
        )

        data  = resp.json()
        reply = data["choices"][0]["message"]["content"].strip()
        print(f"[Groq] Reply: {reply[:80]}")
        return reply

    except requests.Timeout:
        print("Groq timeout")
        return "Sorry, I could not get a response. Please ask again."
    except Exception as e:
        print(f"Groq error: {e}")
        return "I am having a little trouble. Please try again."


# ══════════════════════════════════════════
#   SCENE DESCRIPTION — every 15 seconds
# ══════════════════════════════════════════
last_scene_time = 0


def build_yolo_scene_text(detections):
    if not detections:
        return ""
    centre, left, right = [], [], []
    for d in detections[:6]:
        entry = d['name'] + ' ' + d['dist_text']
        if d['position'] == 'centre':
            centre.append(entry)
        elif d['position'] == 'left':
            left.append(d['name'] + ' on left')
        else:
            right.append(d['name'] + ' on right')
    parts = centre + left + right
    return ', '.join(parts)


def gemini_scene_description(frame, yolo_text):
    try:
        _, buffer = cv2.imencode('.jpg', frame, [cv2.IMWRITE_JPEG_QUALITY, 75])
        img_b64   = base64.b64encode(buffer).decode('utf-8')
        yolo_context = ("YOLO already detected: " + yolo_text + ".") if yolo_text else "YOLO detected nothing."
        prompt = (
            "Describe this scene for a blind person in ONE short sentence only.\n"
            + yolo_context + "\n"
            "Only mention what YOLO missed: walls, floor, open space, doors.\n"
            "ONE sentence. Under 12 words. No lists. No explanation.\n"
            "Example: Wall ahead, path clear on right."
        )
        response = gemini_client.models.generate_content(
            model=GEMINI_CURRENCY_MODEL,
            contents=[{"parts": [
                {"text": prompt},
                {"inline_data": {"mime_type": "image/jpeg", "data": img_b64}}
            ]}]
        )
        return response.text.strip()
    except Exception as e:
        print("[Scene] Gemini error: " + str(e))
        return ""


def describe_scene_now(frame, lang):
    global last_scene_time
    last_scene_time = time.time()
    detections = detect_objects(frame, lang)
    yolo_text  = build_yolo_scene_text(detections)

    def _run():
        gemini_text = gemini_scene_description(frame, yolo_text)
        if yolo_text and gemini_text:
            full = yolo_text + ". " + gemini_text
        elif yolo_text:
            full = yolo_text
        elif gemini_text:
            full = gemini_text
        else:
            full = "The path ahead looks clear."
        print("[Scene] " + full)
        speak(full)

    threading.Thread(target=_run, daemon=True).start()
    return detections


# ══════════════════════════════════════════
#   CURRENCY — FULL SCAN PIPELINE
#
#   Flow:
#   1. Wait for a clear sharp frame
#   2. Run YOUR TFLite ML model on it
#   3. Send same frame to Gemini
#   4. Combine both results → speak reply
#
#   Decision logic:
#   • Both agree          → trust result
#   • Only ML confident   → trust ML
#   • Only Gemini knows   → trust Gemini
#   • Neither confident   → ask to retry
# ══════════════════════════════════════════
def deep_scan(lang):
    state['scanning'] = True
    currency_map = CURRENCY_MAP_EN if lang == 'en' else CURRENCY_MAP_HI

    # ── Step 1: Get a clear frame ──
    frame = capture_clear_frame(timeout=4.0, blur_threshold=80.0)
    if frame is None:
        speak("I cannot access the camera. Please try again.")
        state['scanning'] = False
        return {'denomination': 'unknown', 'confidence': 'low',
                'isAuthentic': True, 'agreed': False, 'side': 'unknown'}

    # ── Step 2: Your ML model ──
    ml_label, ml_conf = predict_currency(frame)

    # ── Step 3: Gemini on same frame ──
    gemini_result = gemini_currency(frame, ml_label, ml_conf)

    # ── Step 4: Decide final answer ──
    if gemini_result is None:
        # Gemini failed — rely on ML model alone
        if ml_label == 'Background' or ml_conf < CURRENCY_CONF:
            speak("I cannot see the note clearly. Please hold it flat and closer.")
            state['scanning'] = False
            return {'denomination': 'unknown', 'confidence': 'low',
                    'isAuthentic': True, 'agreed': False, 'side': 'unknown'}
        final     = ml_label
        authentic = True
        conf_pct  = f"{ml_conf*100:.0f}%"
        side      = 'unknown'
        agreed    = False
        print(f"[Scan] ML only — {final}")

    else:
        gemini_denom = gemini_result.get('denomination', 'unknown')
        agreed       = gemini_result.get('agrees_with_model', False)
        authentic    = gemini_result.get('is_authentic', True)
        side         = gemini_result.get('side', 'unknown')
        gem_conf     = gemini_result.get('confidence', 'low')
        conf_pct     = {'high': '95%', 'medium': '80%', 'low': '60%'}.get(gem_conf, gem_conf)

        if agreed and ml_label != 'Background':
            # Both ML and Gemini say the same thing — most reliable
            final = ml_label
            print(f"[Scan] Both agree — {final}")
        elif gemini_denom != 'unknown':
            # Gemini has a confident answer (may differ from ML)
            final = gemini_denom
            print(f"[Scan] Gemini answer — {final} (ML said {ml_label})")
        elif ml_conf >= CURRENCY_CONF and ml_label != 'Background':
            # Only ML is confident
            final = ml_label
            print(f"[Scan] ML confident — {final}")
        else:
            final = 'unknown'
            print(f"[Scan] Neither model confident")

    # ── Step 5: Speak the result ──
    if final == 'unknown' or final == 'Background':
        speak("I could not identify the note. Please hold it steady and try again.")
    else:
        if not authentic:
            speak("Warning — this note may not be genuine. Please verify with someone.")
        else:
            reply = currency_map.get(final, f"This is a {final} rupee note.")
            speak(reply)

    state['scanning'] = False
    return {'denomination': final if final != 'Background' else 'unknown',
            'confidence': conf_pct if gemini_result else f"{ml_conf*100:.0f}%",
            'isAuthentic': authentic,
            'agreed': agreed,
            'side': side}


# ══════════════════════════════════════════
#   VOICE HELPERS
# ══════════════════════════════════════════
def is_currency_query(text):
    return any(t in text.lower() for t in CURRENCY_TRIGGERS)


def add_to_convo(role, text):
    state['conversation'].append({
        'role': role,
        'text': text,
        'time': datetime.now().strftime("%I:%M %p")
    })
    if len(state['conversation']) > 8:
        state['conversation'] = state['conversation'][-8:]


# ══════════════════════════════════════════
#   PROCESS ONE UTTERANCE (side thread)
#
#   busy_event is set by voice_listener
#   before this function runs.
#   This function must EITHER:
#     a) call speak() — tts_worker will
#        clear busy after speech finishes
#     b) call busy_event.clear() directly
#        if no speech is needed
#   Never both, never neither.
# ══════════════════════════════════════════
def process_audio(audio, recognizer):
    try:
        text = recognizer.recognize_google(
            audio,
            language='en-US'
        ).strip()

        if not text or len(text) < 2:
            busy_event.clear()
            return

        print(f"[Voice] {text}")

        # ── Wake word check ──
        # Fuzzy match for "Netra" — Google STT often mishears it as:
        # nitra, nethra, naetra, neutral, nature, letra, needa, nether
        text_lower = text.lower()

        WAKE_VARIANTS = [
            'netra', 'nitra', 'nethra', 'naetra', 'neutral',
            'nature', 'letra', 'needa', 'nether', 'need a',
            'neetra', 'netra,', 'hetra', 'metra', 'petra',
        ]

        wake_found    = None
        wake_detected = False
        for variant in WAKE_VARIANTS:
            if variant in text_lower:
                wake_found    = variant
                wake_detected = True
                break

        if not wake_detected:
            print(f"[Voice] No wake word — ignoring: {text}")
            busy_event.clear()
            return

        # Strip the matched wake variant from text
        clean_text = text_lower.replace(wake_found, '').strip()
        # Also strip common punctuation/filler
        clean_text = clean_text.lstrip(',. ').strip()

        print(f"[Voice] Wake word [{wake_found}] → command: [{clean_text}]")
        add_to_convo('user', text)

        # ── Just wake word alone ──
        if not clean_text:
            speak("Yes, I am listening.")
            add_to_convo('netra', "Yes, I am listening.")
            return

        # ══════════════════════════════════
        #   COMMAND ROUTING — 3 MODE SYSTEM
        # ══════════════════════════════════

        # ── "Netra analyse surroundings" ──
        # Works from any mode → switches to ANALYSE
        analyse_triggers = ['analyse', 'analyze', 'start', 'describe',
                            'surrounding', 'look around', 'what is around',
                            'tell me surroundings']
        if any(t in clean_text for t in analyse_triggers):
            state['app_mode'] = APP_MODE_ANALYSE
            scene_state['last_labels']    = set()
            scene_state['last_counts']    = {}
            scene_state['announce_count'] = {}
            speak("Starting surroundings analysis.")
            add_to_convo('netra', "Starting surroundings analysis.")
            return

        # ── "Netra stop" ──
        # From ANALYSE → back to SILENT
        stop_triggers = ['stop', 'quiet', 'silence', 'pause', 'enough']
        if any(t in clean_text for t in stop_triggers):
            state['app_mode'] = APP_MODE_SILENT
            speak("Stopped. I am watching silently.")
            add_to_convo('netra', "Stopped.")
            return

        # ── "Netra scan note" ──
        # From any mode → switches to SCAN, waits silently
        if is_currency_query(clean_text):
            if state['app_mode'] == APP_MODE_SCAN:
                # Already in scan mode — do the actual scan now
                speak("Scanning.")
                while busy_event.is_set():
                    time.sleep(0.05)
                busy_event.set()
                state['scanning'] = True
                result = deep_scan(state['language'])
                state['scanning'] = False
                denom = result.get('denomination', 'unknown')
                if denom != 'unknown':
                    reply = f"{denom} rupee note."
                else:
                    reply = "Could not identify. Please show clearly."
                if not result.get('isAuthentic', True):
                    reply = "Warning — note may be fake. " + reply
                add_to_convo('netra', reply)
            else:
                # Enter scan mode
                state['app_mode'] = APP_MODE_SCAN
                speak("Scan mode on. Show a note and say Netra scan.")
                add_to_convo('netra', "Scan mode on.")
            return

        # ── General question (any mode) ──
        reply = groq_conversation(clean_text)
        if reply and reply.strip():
            add_to_convo('netra', reply)
            speak(reply, priority=True)
        else:
            busy_event.clear()

    except sr.UnknownValueError:
        busy_event.clear()
    except sr.RequestError as e:
        print(f"STT error: {e}")
        busy_event.clear()
    except Exception as e:
        print(f"process_audio error: {e}")
        busy_event.clear()


# ══════════════════════════════════════════
#   VOICE LISTENER — PERSISTENT MIC LOOP
# ══════════════════════════════════════════
def voice_listener():
    recognizer = sr.Recognizer()
    recognizer.pause_threshold          = 0.6
    recognizer.phrase_threshold         = 0.3
    recognizer.non_speaking_duration    = 0.4
    recognizer.dynamic_energy_threshold = False
    recognizer.energy_threshold         = 300

    print("Voice listener started.")
    speak("Hello! I am Netra. I am watching silently. Say Netra analyse surroundings to start, or Netra scan note to check currency.")

    # ── Find the correct microphone device ──
    # Find the real microphone, skip virtual audio devices
    def find_real_mic_index():
        import pyaudio
        pa = pyaudio.PyAudio()
        real_mic = None
        print("[Mic] Available audio input devices:")
        for i in range(pa.get_device_count()):
            info = pa.get_device_info_by_index(i)
            if info['maxInputChannels'] > 0:
                name = info['name']
                print(f"[Mic]   [{i}] {name}")
                # Skip virtual audio devices
                skip_keywords = ['droidcam', 'virtual', 'cable', 'voicemeeter',
                                 'stereo mix', 'what u hear', 'loopback']
                if any(k in name.lower() for k in skip_keywords):
                    print(f"[Mic]       ^ skipping (virtual device)")
                    continue
                # Prefer real microphone keywords
                prefer_keywords = ['microphone', 'mic', 'realtek', 'audio input',
                                   'headset', 'headphone', 'built-in', 'internal']
                if any(k in name.lower() for k in prefer_keywords):
                    real_mic = i
                    print(f"[Mic]       ^ selected as real mic")
                    break
                # If nothing preferred yet, use first valid input
                if real_mic is None:
                    real_mic = i
        pa.terminate()
        return real_mic

    mic_index = find_real_mic_index()
    print(f"[Mic] Using microphone index: {mic_index}")

    # ── Open mic with explicit device index ──
    mic_source = sr.Microphone(device_index=mic_index)

    with mic_source as source:
        print("Calibrating for ambient noise, please wait...")
        try:
            recognizer.adjust_for_ambient_noise(source, duration=1.5)
            print(f"[Mic] Energy threshold: {recognizer.energy_threshold:.0f}")
        except Exception as e:
            print(f"[Mic] Calibration warning: {e}")
            recognizer.energy_threshold = 300

        while True:
            if (not state['mic_active']
                    or state['scanning']
                    or busy_event.is_set()):
                time.sleep(0.05)
                continue

            try:
                audio = recognizer.listen(source, timeout=2, phrase_time_limit=10)
            except sr.WaitTimeoutError:
                continue
            except Exception as e:
                print(f"Listen error: {e}")
                time.sleep(0.5)
                # If stream is broken, try to reconnect
                if '-9988' in str(e) or '-9999' in str(e):
                    print("[Mic] Stream broken — reconnecting in 2s...")
                    time.sleep(2)
                    try:
                        source.__exit__(None, None, None)
                        source.__enter__()
                        recognizer.adjust_for_ambient_noise(source, duration=0.5)
                        print("[Mic] Reconnected")
                    except Exception as e2:
                        print(f"[Mic] Reconnect failed: {e2}")
                continue

            busy_event.set()
            threading.Thread(
                target=process_audio,
                args=(audio, recognizer),
                daemon=True
            ).start()


threading.Thread(target=voice_listener, daemon=True).start()

# ══════════════════════════════════════════
#   SMART ANNOUNCE
#
#   Only speaks when the scene CHANGES:
#   - New object appears that wasn't there
#   - Object count increases (more people)
#   - Very close danger object appears
#   - Same scene repeated max 2 times then silent
#   - Full Gemini scene every 45s
# ══════════════════════════════════════════
def smart_announce(detections, frame, lang, now):
    global last_scene_time

    # Build current scene — count each label
    current_counts = {}
    for d in detections:
        lbl = d['label']
        current_counts[lbl] = current_counts.get(lbl, 0) + 1

    current_labels = set(current_counts.keys())
    last_labels    = scene_state['last_labels']
    last_counts    = scene_state['last_counts']

    # ── Check what changed ──
    new_labels      = current_labels - last_labels        # objects that appeared
    removed_labels  = last_labels - current_labels        # objects that disappeared
    increased       = {l for l in current_labels          # count went up
                       if current_counts.get(l, 0) > last_counts.get(l, 0)}
    got_closer      = {d['label'] for d in detections     # something got very close
                       if d['dist_key'] == 'very_close'
                       and state['last_dist_key'].get(d['label']) != 'very_close'}

    scene_changed = bool(new_labels or increased or got_closer)

    # ── Build natural spoken sentence ──
    def build_sentence(dets):
        if not dets:
            return ""
        # Group by position
        parts = []
        # Closest/most important first
        for d in dets[:4]:
            count = current_counts.get(d['label'], 1)
            if count > 1:
                name = str(count) + ' ' + d['name'] + 's'
            else:
                name = d['name']
            pos  = d['position']
            dist = d['dist_text']
            if pos == 'centre':
                parts.append(name + ' ' + dist)
            else:
                parts.append(name + ' on ' + pos)
        return ', '.join(parts)

    # ── Urgent: danger very close ──
    for d in detections:
        if (d['dist_key'] == 'very_close'
                and d['label'] in ('person', 'car', 'motorcycle',
                                   'bicycle', 'bus', 'truck', 'dog')):
            lt = state['last_announced'].get(d['label'], 0)
            if now - lt > 8:
                count = current_counts.get(d['label'], 1)
                name  = (str(count) + ' ' + d['name'] + 's'
                         if count > 1 else d['name'])
                speak(name + ' very close, be careful!', priority=True)
                state['last_announced'][d['label']] = now
                state['last_dist_key'][d['label']]  = 'very_close'
                # Update scene state
                scene_state['last_labels']  = current_labels
                scene_state['last_counts']  = current_counts.copy()
                return

    # ── Scene changed — speak what's new ──
    if scene_changed:
        scene_hash = str(sorted(current_counts.items()))
        announce_count = scene_state['announce_count'].get(scene_hash, 0)

        # Only announce same scene max 2 times
        if announce_count < 2:
            sentence = build_sentence(detections)
            if sentence:
                speak(sentence)
                scene_state['announce_count'][scene_hash] = announce_count + 1
                scene_state['last_labels']  = current_labels
                scene_state['last_counts']  = current_counts.copy()
                for d in detections:
                    state['last_dist_key'][d['label']] = d['dist_key']
        return

    # ── Nothing changed — stay silent unless 45s passed for full scene ──
    if now - last_scene_time >= ANNOUNCE_COOLDOWN:
        # Reset announce counts so next change is fresh
        scene_state['announce_count'] = {}
        describe_scene_now(frame, lang)


# ══════════════════════════════════════════
#   FLASK
# ══════════════════════════════════════════
flask_app = Flask(__name__)


@flask_app.route('/')
def index():
    return render_template('index.html')


@flask_app.route('/video_feed')
def video_feed():
    def generate():
        while True:
            if current_frame is None:
                time.sleep(0.05)
                continue

            frame = current_frame.copy()
            h, w  = frame.shape[:2]

            if state['app_mode'] in (APP_MODE_SILENT, APP_MODE_ANALYSE):
                lang       = state['language']
                detections = detect_objects(frame, lang)  # YOLO pass

                x1z = int(w * (0.5 - CENTER_ZONE))
                x2z = int(w * (0.5 + CENTER_ZONE))
                overlay = frame.copy()
                cv2.rectangle(overlay, (x1z, 0), (x2z, h), (0, 255, 255), -1)
                cv2.addWeighted(overlay, 0.06, frame, 0.94, 0, frame)
                cv2.line(frame, (x1z, 0), (x1z, h), (0, 255, 255), 1)
                cv2.line(frame, (x2z, 0), (x2z, h), (0, 255, 255), 1)

                for d in detections:
                    x1, y1, x2, y2 = d['box']
                    color = ((0, 0, 255)    if d['dist_key'] == 'very_close'
                             else (0, 130, 255) if d['dist_key'] == 'close'
                             else (0, 210, 0))
                    txt = d['name'] + ' [' + d['position'] + '] ' + d['dist_text']
                    (tw, th), _ = cv2.getTextSize(
                        txt, cv2.FONT_HERSHEY_SIMPLEX, 0.48, 2)
                    cv2.rectangle(frame, (x1, y1), (x2, y2), color, 2)
                    cv2.rectangle(frame,
                                  (x1, max(0, y1 - th - 10)),
                                  (x1 + tw + 8, y1), color, -1)
                    cv2.putText(frame, txt,
                                (x1 + 4, max(10, y1 - 4)),
                                cv2.FONT_HERSHEY_SIMPLEX, 0.48, (0, 0, 0), 2)

                now = time.time()

                # Only announce in ANALYSE mode
                if (state['app_mode'] == APP_MODE_ANALYSE
                        and not busy_event.is_set()
                        and not state['scanning']):
                    smart_announce(detections, frame.copy(), lang, now)

            _, jpeg = cv2.imencode('.jpg', frame, [cv2.IMWRITE_JPEG_QUALITY, 72])
            yield (b'--frame\r\nContent-Type: image/jpeg\r\n\r\n'
                   + jpeg.tobytes() + b'\r\n')
            time.sleep(0.04)

    return Response(generate(),
                    mimetype='multipart/x-mixed-replace; boundary=frame')


@flask_app.route('/capture')
def capture():
    if current_frame is None:
        return jsonify({'error': 'No frame'}), 500
    _, buffer = cv2.imencode('.jpg', current_frame.copy(),
                              [cv2.IMWRITE_JPEG_QUALITY, 90])
    return jsonify({'image': base64.b64encode(buffer).decode('utf-8')})


@flask_app.route('/scan', methods=['POST'])
def scan():
    try:
        data      = request.get_json()
        img_bytes = base64.b64decode(data.get('image'))
        frame     = cv2.imdecode(
            np.frombuffer(img_bytes, np.uint8), cv2.IMREAD_COLOR)
        result    = deep_scan(data.get('language', 'en'))
        return jsonify(result)
    except Exception as e:
        print(f"Scan error: {e}")
        return jsonify({'error': str(e)}), 500


@flask_app.route('/conversation')
def get_conversation():
    return jsonify(state['conversation'])


@flask_app.route('/set_language', methods=['POST'])
def set_language():
    lang = request.get_json().get('language', 'en')
    state['language'] = lang
    state['last_announced'] = {}
    state['last_dist_key']  = {}
    speak("English selected." if lang == 'en' else "Hindi selected.")
    return jsonify({'ok': True})


@flask_app.route('/set_mode', methods=['POST'])
def set_mode():
    state['app_mode'] = request.get_json().get('mode', APP_MODE_SILENT)
    return jsonify({'ok': True})


@flask_app.route('/status')
def status():
    return jsonify({
        'mode':     state['app_mode'],   # silent / analyse / scan
        'language': state['language'],
        'scanning': state['scanning'],
        'speaking': is_speaking,
        'busy':     busy_event.is_set(),
    })


if __name__ == '__main__':
    speak("Netra is ready.")
    flask_app.run(host='0.0.0.0', port=5000, debug=False, threaded=True)