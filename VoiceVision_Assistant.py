import cv2 as cv
import mediapipe as mp
import pyautogui
import numpy as np
import speech_recognition as sp
import threading
import subprocess
import psutil
import os
import platform
import tkinter as tk

end_program = False
root = tk.Tk()
root.geometry("200x60+1400+100")  # Adjust size and position as needed
label_text = tk.StringVar()
label = tk.Label(root, textvariable=label_text, font=("Helvetica", 12))
label.pack()

class VoiceAssistant:
    def __init__(self):
        self.recognizer = sp.Recognizer()

    def listen(self):
        with sp.Microphone() as mic:
            print("Listening...")
            label_text.set("Listening...")
            self.recognizer.pause_threshold = 1
            self.recognizer.adjust_for_ambient_noise(mic, duration=0.1)
            try:
                audio = self.recognizer.listen(mic, timeout=2)
                return self.recognizer.recognize_google(audio).lower()
            except sp.WaitTimeoutError:
                label_text.set("No speech detected after 2 seconds.")
                print("No speech detected after 2 seconds.")
                return ""

    def terminate(self):
        global end_program
        end_program = True

    def tabs(self):
        pyautogui.keyDown("command" if platform.system() == "Darwin" else "alt")
        pyautogui.press("tab")

    def change(self):
        pyautogui.press("tab")

    def done(self):
        pyautogui.keyUp("command" if platform.system() == "Darwin" else "alt")

    def last_tab(self):
        pyautogui.keyDown("command" if platform.system() == "Darwin" else "alt")
        pyautogui.press("tab", presses=3)
        pyautogui.keyUp("command" if platform.system() == "Darwin" else "alt")

    def execute_command(self, command):
        if command.startswith("open"):
            application_name = command[len("open "):]
            self.open_application(application_name)
        elif command.startswith("close"):
            application_name = command[len("close "):]
            self.close_application(application_name)
        elif command.startswith("type"):
            sentence_to_type = command[len("type "):]
            pyautogui.typewrite(sentence_to_type)
        elif command == "click":
            pyautogui.click()
        elif command == "erase":
            self.erase_last_word()
        elif command == "delete":
            self.erase_line()
        elif command == "select":
            self.select_line()
        elif command == "copy":
            self.copy_line()
        elif command == "paste":
            self.paste_line()
        elif command == "terminate":
            self.terminate()
        elif command == "last tab":
            self.last_tab()

    def open_application(self, application_name):
        username = os.getlogin()
        system = platform.system()
        if system == "Darwin":  # macOS
            applications = {
                "google": "Google Chrome",
                "file": "Finder",
                "calculator": "Calculator",
                "notepad": "TextEdit",
                "code": "Visual Studio Code"
            }
            if application_name.lower() in applications:
                subprocess.Popen(["open", "-a", applications[application_name.lower()]])
                print(f"{application_name} opened successfully.")
            else:
                print(f"Application '{application_name}' not recognized.")
        else:  # Windows
            applications = {
                "google": r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
                "file": r"C:\Windows\explorer.exe",
                "calculator": r"C:\Windows\System32\calc.exe",
                "notepad": r"C:\Windows\System32\notepad.exe",
                "code": r"C:\Users\{}\AppData\Local\Programs\Microsoft VS Code\Code.exe".format(username)
            }
            if application_name.lower() in applications:
                subprocess.Popen([applications[application_name.lower()]], shell=True)
                print(f"{application_name} opened successfully.")
            else:
                print(f"Application '{application_name}' not recognized.")

    def close_application(self, application_name):
        for process in psutil.process_iter():
            try:
                if application_name.lower() in process.name().lower():
                    process.terminate()
                    print(f"{application_name} closed successfully.")
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass

    def erase_last_word(self):
        pyautogui.hotkey("command" if platform.system() == "Darwin" else "ctrl", "backspace")
        print("Last word erased.")

    def select_line(self):
        pyautogui.click()
        pyautogui.keyDown('shift')
        pyautogui.press('home')

    def erase_line(self):
        pyautogui.click()
        pyautogui.press('backspace')
        pyautogui.keyUp('shift')

    def copy_line(self):
        pyautogui.click()
        pyautogui.hotkey("command" if platform.system() == "Darwin" else "ctrl", "c")
        pyautogui.keyUp('shift')

    def paste_line(self):
        pyautogui.click()
        pyautogui.hotkey("command" if platform.system() == "Darwin" else "ctrl", "v")

    def run(self):
        while not end_program:
            try:
                command = self.listen()
                label_text.set(command)
                print(command)
                self.execute_command(command)
            except sp.UnknownValueError:
                label_text.set("Sorry, I couldn't understand what you said. Please try again.")
                print("Sorry, I couldn't understand what you said. Please try again.")
            except sp.RequestError as e:
                print(f"Could not request results from Google Speech Recognition service; {e}")
            except KeyboardInterrupt:
                print("Voice assistant terminated by user.")
                break

class WebcamController:
    def __init__(self):
        self.smoothing = 7
        self.plocX, self.plocY = 0, 0
        self.mpDraw = mp.solutions.drawing_utils
        self.mpFaceMesh = mp.solutions.face_mesh
        self.faceMesh = self.mpFaceMesh.FaceMesh(static_image_mode=False, max_num_faces=1, min_detection_confidence=0.5, min_tracking_confidence=0.5)
        self.cap = cv.VideoCapture(0)
        self.cap.set(3, 640)
        self.cap.set(4, 480)
        self.wScr, self.hScr = pyautogui.size()  # Screen width and height

    def run(self):
        global end_program
        while not end_program:
            success, img = self.cap.read()
            if not success:
                print("Failed to capture image.")
                continue

            img = cv.flip(img, 1)
            imgRGB = cv.cvtColor(img, cv.COLOR_BGR2RGB)
            result = self.faceMesh.process(imgRGB)

            if result.multi_face_landmarks:
                for faceLms in result.multi_face_landmarks:
                    for id, lm in enumerate(faceLms.landmark):
                        h, w, c = img.shape
                        cx, cy = int(lm.x * w), int(lm.y * h)

                        if id == 4:  # Index of the nose landmark
                            x1 = np.interp(cx, (0, w), (0, self.wScr))  # Map to screen width
                            y1 = np.interp(cy, (0, h), (0, self.hScr))  # Map to screen height

                            # Smoothing for cursor movement
                            self.plocX += (x1 - self.plocX) / self.smoothing
                            self.plocY += (y1 - self.plocY) / self.smoothing

                            try:
                                if 0 <= self.plocX <= self.wScr and 0 <= self.plocY <= self.hScr:
                                    pyautogui.moveTo(self.plocX, self.plocY)
                            except pyautogui.FailSafeException:
                                print("PyAutoGUI fail-safe triggered. Mouse moved to a corner of the screen.")

                            cv.circle(img, (cx, cy), 5, (0, 0, 255), cv.FILLED)
            
            cv.namedWindow("Webcam", cv.WINDOW_NORMAL)
          
            if cv.waitKey(1) & 0xFF == ord('q'):
                break

        self.cap.release()
        cv.destroyAllWindows()


if __name__ == "__main__":
    voice_assistant = VoiceAssistant()
    webcam_controller = WebcamController()

    voice_thread = threading.Thread(target=voice_assistant.run)
    webcam_thread = threading.Thread(target=webcam_controller.run)

    voice_thread.start()
    webcam_thread.start()
    root.mainloop()
    voice_thread.join()
    webcam_thread.join()

    cv.destroyAllWindows()
