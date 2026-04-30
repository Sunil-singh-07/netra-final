# Netra AI – Smart Assistant for Visually Impaired

Netra is a real-time AI assistant that helps visually impaired users navigate surroundings and identify currency using computer vision, voice interaction, and AI models.

---

##  Features

* Real-time object detection using YOLOv8
* Voice assistant with wake word ("Netra")
* Currency detection using TensorFlow Lite model
* AI verification using Gemini API
* Speech output using Windows TTS
* Scene understanding for navigation assistance

---

##  Tech Stack

* Python, Flask
* OpenCV
* TensorFlow Lite
* YOLOv8 (Ultralytics)
* Google Gemini API
* Groq API
* SpeechRecognition

---

## ⚙️ Setup Instructions

### 1. Clone the repository

```bash
git clone https://github.com/Sunil-singh-07/netra-final.git
cd netra-final
```

### 2. Install dependencies

```bash
pip install -r requirements.txt
```

### 3. Install PyAudio (Windows only)

```bash
pip install pipwin
pipwin install pyaudio
```

### 4. Add API keys

Create a `.env` file in the root folder:

```env
GEMINI_API_KEY=your_key_here
GROQ_API_KEY=your_key_here
```

### 5. Run the project

```bash
python app.py
```

---

##  Notes

* YOLO model (`yolov8n.pt`) will download automatically on first run
* Project currently supports **Windows only** (due to TTS dependency)
* Camera defaults to webcam (can be changed in code)

---

##  Contributing

1. Fork the repository
2. Create a new branch (`feature-name`)
3. Make changes
4. Push and create a Pull Request

---

##  Future Improvements

* Mobile app version
* Offline support
* Multi-language voice assistant
* Faster real-time performance

---

##  Author

Sunil Singh

---
