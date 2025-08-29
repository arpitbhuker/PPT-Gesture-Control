import win32com.client
from cvzone.HandTrackingModule import HandDetector
import cv2

# Connect to PowerPoint Application (use active presentation)
Application = win32com.client.Dispatch("PowerPoint.Application")
Application.Visible = True

if Application.Presentations.Count > 0:
    Presentation = Application.ActivePresentation
else:
    raise Exception("No PowerPoint presentation is currently open!")

print("Controlling Presentation:", Presentation.Name)
Presentation.SlideShowSettings.Run()

# Parameters
width, height = 900, 720
gestureThreshold = 300

# Camera Setup
cap = cv2.VideoCapture(0)
cap.set(3, width)
cap.set(4, height)

# Hand Detector
detectorHand = HandDetector(detectionCon=0.8, maxHands=1)

# Variables
buttonPressed = False
counter = 0
delay = 15
prev_cx = None

while True:
    success, img = cap.read()
    if not success:
        break

    # Find hands
    hands, img = detectorHand.findHands(img)
    
    if hands and not buttonPressed:
        hand = hands[0]
        cx, cy = hand["center"]

        if prev_cx is not None:
            movement = cx - prev_cx

            if abs(movement) > 50:  # Threshold for swipe
                if movement > 0:  # Hand moved right
                    print("Next Slide")
                    Presentation.SlideShowWindow.View.Next()
                else:  # Hand moved left
                    print("Previous Slide")
                    Presentation.SlideShowWindow.View.Previous()

                buttonPressed = True

        prev_cx = cx  # Update position

    if buttonPressed:
        counter += 1
        if counter > delay:
            counter = 0
            buttonPressed = False
            prev_cx = None  # Reset tracking

    cv2.imshow("Presentation Control", img)

    key = cv2.waitKey(1)
    if key == ord('q'):
        break

cap.release()
cv2.destroyAllWindows()
