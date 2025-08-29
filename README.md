# 🎥 Hand Gesture Controlled Presentation System

Control your PowerPoint presentations with just your **hand gestures**! ✋  
This project turns your webcam into a **smart presentation controller**, allowing you to move through slides with simple hand movements instead of a keyboard or clicker.  

The system uses **computer vision (OpenCV + cvzone)** to track your hand and interpret swipe gestures, which are then translated into **PowerPoint commands** via **pywin32**. This enables a smooth, touch-free, and futuristic way to deliver presentations, making it especially useful for:  

- Teachers and professors during lectures  
- Business professionals in client meetings  
- Students giving project presentations  
- Hands-free, one-handed, or remote presentation settings  

---

## 🚀 Features
- Works with **any open PowerPoint presentation** (no need to hardcode file path).  
- **Swipe Gesture Control**:
  - Swipe **Right → Next Slide**
  - Swipe **Left → Previous Slide**
- Works with **any hand** (left or right).  
- Runs using your **built-in or external webcam**.  
- Lightweight and easy to set up.
  
---


## ▶️ Usage

1. Run the script:
   ```bash
   python ppt_gesture_control.py

- Move your hand left or right in front of the camera.
  
- Slide will change accordingly.

- Press Q to quit.

## 📹 Demo

Here’s a quick preview of the hand gesture slide controller in action:

![Demo](demo.gif)


## 🛠️ Tech Stack

- Python
  
- OpenCV
  
- cvzone
  
- pywin32 (for PowerPoint COM automation)

## 📌 Future Improvements

- **Gesture Customization** – Allow users to map their own gestures (e.g., two-finger swipe to jump to a specific slide).  
- **Annotation Mode** – Use finger tracking to draw or highlight directly on the slide during presentations.  
- **Voice + Gesture Hybrid Control** – Combine simple voice commands with gestures for seamless control.  
- **AI-based Gesture Recognition** – Replace rule-based detection with a machine learning model for more natural hand movements.  
- **Cross-Platform Integration** – Extend support to Google Slides, Keynote, and LibreOffice Impress.  
