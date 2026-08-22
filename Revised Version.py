import tkinter as tk
from tkinter import filedialog, messagebox
import os
import cv2
import numpy as np
import pyautogui
import win32com.client
import time
from PIL import Image, ImageTk

# =========================================================
# CONFIG
# =========================================================
WINDOW_WIDTH = 900
WINDOW_HEIGHT = 700
GESTURE_DELAY = 1

# =========================================================
# MAIN CLASS
# =========================================================
class PPTGestureControl:

    def __init__(self, root):

        self.root = root

        self.root.title("PPT Gesture Control")

        self.root.geometry(
            f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}"
        )

        self.root.configure(bg="#111111")

        # Variables
        self.ppt_file = ""
        self.powerpoint = None
        self.presentation = None
        self.cap = None
        self.last_action = 0

        # Screen Size
        self.screen_w, self.screen_h = pyautogui.size()

        # Create GUI
        self.create_widgets()

    # =====================================================
    # GUI
    # =====================================================
    def create_widgets(self):

        title = tk.Label(
            self.root,
            text="PPT Gesture Control",
            font=("Arial", 24, "bold"),
            bg="#111111",
            fg="white"
        )

        title.pack(pady=20)

        select_btn = tk.Button(
            self.root,
            text="Select PPT",
            font=("Arial", 14, "bold"),
            command=self.choose_ppt,
            bg="#00aa00",
            fg="white",
            width=20
        )

        select_btn.pack(pady=10)

        self.file_label = tk.Label(
            self.root,
            text="No File Selected",
            font=("Arial", 12),
            bg="#111111",
            fg="yellow"
        )

        self.file_label.pack(pady=10)

        start_btn = tk.Button(
            self.root,
            text="Start PPT Control",
            font=("Arial", 14, "bold"),
            command=self.start_control,
            bg="#0066ff",
            fg="white",
            width=20
        )

        start_btn.pack(pady=10)

        self.status_label = tk.Label(
            self.root,
            text="Waiting...",
            font=("Arial", 12, "bold"),
            bg="#111111",
            fg="#00ffff"
        )

        self.status_label.pack(pady=10)

        self.camera_label = tk.Label(self.root)

        self.camera_label.pack(pady=20)

    # =====================================================
    # SELECT PPT
    # =====================================================
    def choose_ppt(self):

        file = filedialog.askopenfilename(
            title="Select PPT File",
            filetypes=[("PowerPoint Files", "*.pptx")]
        )

        if file:

            self.ppt_file = file

            self.file_label.config(
                text=os.path.basename(file)
            )

            self.status_label.config(
                text="PPT Selected"
            )

    # =====================================================
    # START CONTROL
    # =====================================================
    def start_control(self):

        if not self.ppt_file:

            messagebox.showerror(
                "Error",
                "Please Select PPT File"
            )

            return

        try:

            self.powerpoint = (
                win32com.client.Dispatch(
                    "PowerPoint.Application"
                )
            )

            self.powerpoint.Visible = 1

            self.presentation = (
                self.powerpoint.Presentations.Open(
                    self.ppt_file
                )
            )

            self.presentation.SlideShowSettings.Run()

            self.status_label.config(
                text="PowerPoint Started"
            )

        except Exception as e:

            messagebox.showerror(
                "PowerPoint Error",
                str(e)
            )

            return

        # Camera
        self.cap = cv2.VideoCapture(0)

        if not self.cap.isOpened():

            messagebox.showerror(
                "Camera Error",
                "Unable to Open Camera"
            )

            return

        self.update_frame()

    # =====================================================
    # UPDATE CAMERA
    # =====================================================
    def update_frame(self):

        success, frame = self.cap.read()

        if not success:

            self.root.after(
                10,
                self.update_frame
            )

            return

        frame = cv2.flip(frame, 1)

        # Convert to HSV
        hsv = cv2.cvtColor(
            frame,
            cv2.COLOR_BGR2HSV
        )

        # =================================================
        # DETECT RED COLOR OBJECT
        # =================================================

        lower_red = np.array([0, 120, 70])

        upper_red = np.array([10, 255, 255])

        mask = cv2.inRange(
            hsv,
            lower_red,
            upper_red
        )

        contours, _ = cv2.findContours(
            mask,
            cv2.RETR_TREE,
            cv2.CHAIN_APPROX_SIMPLE
        )

        if contours:

            largest = max(
                contours,
                key=cv2.contourArea
            )

            area = cv2.contourArea(largest)

            if area > 1000:

                x, y, w, h = cv2.boundingRect(
                    largest
                )

                # Draw Rectangle
                cv2.rectangle(
                    frame,
                    (x, y),
                    (x + w, y + h),
                    (0, 255, 0),
                    2
                )

                center_x = x + w // 2
                center_y = y + h // 2

                # Draw Center Point
                cv2.circle(
                    frame,
                    (center_x, center_y),
                    5,
                    (255, 0, 0),
                    -1
                )

                self.handle_gesture(
                    center_x,
                    center_y
                )

        # Show Camera
        frame_rgb = cv2.cvtColor(
            frame,
            cv2.COLOR_BGR2RGB
        )

        img = Image.fromarray(frame_rgb)

        imgtk = ImageTk.PhotoImage(image=img)

        self.camera_label.imgtk = imgtk

        self.camera_label.configure(
            image=imgtk
        )

        self.root.after(
            10,
            self.update_frame
        )

    # =====================================================
    # HANDLE GESTURE
    # =====================================================
    def handle_gesture(self, x, y):

        current_time = time.time()

        if current_time - self.last_action < GESTURE_DELAY:
            return

        # =================================================
        # LEFT AREA = PREVIOUS SLIDE
        # =================================================
        if x < 200:

            pyautogui.press("left")

            self.status_label.config(
                text="Previous Slide"
            )

            self.last_action = current_time

        # =================================================
        # RIGHT AREA = NEXT SLIDE
        # =================================================
        elif x > 450:

            pyautogui.press("right")

            self.status_label.config(
                text="Next Slide"
            )

            self.last_action = current_time

        # =================================================
        # CENTER AREA = POINTER
        # =================================================
        else:

            screen_x = int(
                x * self.screen_w / 640
            )

            screen_y = int(
                y * self.screen_h / 480
            )

            pyautogui.moveTo(
                screen_x,
                screen_y
            )

            self.status_label.config(
                text="Pointer Mode"
            )

    # =====================================================
    # CLEANUP
    # =====================================================
    def cleanup(self):

        if self.cap:
            self.cap.release()

        cv2.destroyAllWindows()

        if self.powerpoint:
            self.powerpoint.Quit()

        self.root.destroy()

# =========================================================
# MAIN
# =========================================================
if __name__ == "__main__":

    root = tk.Tk()

    app = PPTGestureControl(root)

    root.protocol(
        "WM_DELETE_WINDOW",
        app.cleanup
    )

    root.mainloop()

tkroot.mainloop()
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import cv2
import numpy as np
import pyautogui
import win32com.client
import time
from PIL import Image, ImageTk

# =========================================================
# CONFIG
# =========================================================
WINDOW_WIDTH = 900
WINDOW_HEIGHT = 700
GESTURE_DELAY = 1

# =========================================================
# MAIN CLASS
# =========================================================
class PPTGestureControl:

    def __init__(self, root):

        self.root = root

        self.root.title("PPT Gesture Control")

        self.root.geometry(
            f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}"
        )

        self.root.configure(bg="#111111")

        # Variables
        self.ppt_file = ""
        self.powerpoint = None
        self.presentation = None
        self.cap = None
        self.last_action = 0

        # Screen Size
        self.screen_w, self.screen_h = pyautogui.size()

        # Create GUI
        self.create_widgets()

    # =====================================================
    # GUI
    # =====================================================
    def create_widgets(self):

        title = tk.Label(
            self.root,
            text="PPT Gesture Control",
            font=("Arial", 24, "bold"),
            bg="#111111",
            fg="white"
        )

        title.pack(pady=20)

        select_btn = tk.Button(
            self.root,
            text="Select PPT",
            font=("Arial", 14, "bold"),
            command=self.choose_ppt,
            bg="#00aa00",
            fg="white",
            width=20
        )

        select_btn.pack(pady=10)

        self.file_label = tk.Label(
            self.root,
            text="No File Selected",
            font=("Arial", 12),
            bg="#111111",
            fg="yellow"
        )

        self.file_label.pack(pady=10)

        start_btn = tk.Button(
            self.root,
            text="Start PPT Control",
            font=("Arial", 14, "bold"),
            command=self.start_control,
            bg="#0066ff",
            fg="white",
            width=20
        )

        start_btn.pack(pady=10)

        self.status_label = tk.Label(
            self.root,
            text="Waiting...",
            font=("Arial", 12, "bold"),
            bg="#111111",
            fg="#00ffff"
        )

        self.status_label.pack(pady=10)

        self.camera_label = tk.Label(self.root)

        self.camera_label.pack(pady=20)

    # =====================================================
    # SELECT PPT
    # =====================================================
    def choose_ppt(self):

        file = filedialog.askopenfilename(
            title="Select PPT File",
            filetypes=[("PowerPoint Files", "*.pptx")]
        )

        if file:

            self.ppt_file = file

            self.file_label.config(
                text=os.path.basename(file)
            )

            self.status_label.config(
                text="PPT Selected"
            )

    # =====================================================
    # START CONTROL
    # =====================================================
    def start_control(self):

        if not self.ppt_file:

            messagebox.showerror(
                "Error",
                "Please Select PPT File"
            )

            return

        try:

            self.powerpoint = (
                win32com.client.Dispatch(
                    "PowerPoint.Application"
                )
            )

            self.powerpoint.Visible = 1

            self.presentation = (
                self.powerpoint.Presentations.Open(
                    self.ppt_file
                )
            )

            self.presentation.SlideShowSettings.Run()

            self.status_label.config(
                text="PowerPoint Started"
            )

        except Exception as e:

            messagebox.showerror(
                "PowerPoint Error",
                str(e)
            )

            return

        # Camera
        self.cap = cv2.VideoCapture(0)

        if not self.cap.isOpened():

            messagebox.showerror(
                "Camera Error",
                "Unable to Open Camera"
            )

            return

        self.update_frame()

    # =====================================================
    # UPDATE CAMERA
    # =====================================================
    def update_frame(self):

        success, frame = self.cap.read()

        if not success:

            self.root.after(
                10,
                self.update_frame
            )

            return

        frame = cv2.flip(frame, 1)

        # Convert to HSV
        hsv = cv2.cvtColor(
            frame,
            cv2.COLOR_BGR2HSV
        )

        # =================================================
        # DETECT RED COLOR OBJECT
        # =================================================

        lower_red = np.array([0, 120, 70])

        upper_red = np.array([10, 255, 255])

        mask = cv2.inRange(
            hsv,
            lower_red,
            upper_red
        )

        contours, _ = cv2.findContours(
            mask,
            cv2.RETR_TREE,
            cv2.CHAIN_APPROX_SIMPLE
        )

        if contours:

            largest = max(
                contours,
                key=cv2.contourArea
            )

            area = cv2.contourArea(largest)

            if area > 1000:

                x, y, w, h = cv2.boundingRect(
                    largest
                )

                # Draw Rectangle
                cv2.rectangle(
                    frame,
                    (x, y),
                    (x + w, y + h),
                    (0, 255, 0),
                    2
                )

                center_x = x + w // 2
                center_y = y + h // 2

                # Draw Center Point
                cv2.circle(
                    frame,
                    (center_x, center_y),
                    5,
                    (255, 0, 0),
                    -1
                )

                self.handle_gesture(
                    center_x,
                    center_y
                )

        # Show Camera
        frame_rgb = cv2.cvtColor(
            frame,
            cv2.COLOR_BGR2RGB
        )

        img = Image.fromarray(frame_rgb)

        imgtk = ImageTk.PhotoImage(image=img)

        self.camera_label.imgtk = imgtk

        self.camera_label.configure(
            image=imgtk
        )

        self.root.after(
            10,
            self.update_frame
        )

    # =====================================================
    # HANDLE GESTURE
    # =====================================================
    def handle_gesture(self, x, y):

        current_time = time.time()

        if current_time - self.last_action < GESTURE_DELAY:
            return

        # =================================================
        # LEFT AREA = PREVIOUS SLIDE
        # =================================================
        if x < 200:

            pyautogui.press("left")

            self.status_label.config(
                text="Previous Slide"
            )

            self.last_action = current_time

        # =================================================
        # RIGHT AREA = NEXT SLIDE
        # =================================================
        elif x > 450:

            pyautogui.press("right")

            self.status_label.config(
                text="Next Slide"
            )

            self.last_action = current_time

        # =================================================
        # CENTER AREA = POINTER
        # =================================================
        else:

            screen_x = int(
                x * self.screen_w / 640
            )

            screen_y = int(
                y * self.screen_h / 480
            )

            pyautogui.moveTo(
                screen_x,
                screen_y
            )

            self.status_label.config(
                text="Pointer Mode"
            )

    # =====================================================
    # CLEANUP
    # =====================================================
    def cleanup(self):

        if self.cap:
            self.cap.release()

        cv2.destroyAllWindows()
        import tkinter as tk
        from tkinter import filedialog, messagebox
        import cv2
        import numpy as np
        import pyautogui
        import win32com.client
        import time
        from PIL import Image, ImageTk

        # =========================================================
        # CONFIG
        # =========================================================
        WINDOW_WIDTH = 900
        WINDOW_HEIGHT = 700
        GESTURE_DELAY = 1

        # =========================================================
        # MAIN CLASS
        # =========================================================
        class PPTGestureControl:

            def __init__(self, root):

                self.root = root

                self.root.title("PPT Gesture Control")

                self.root.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")

                # Title Label
                self.label = tk.Label(
                    root,
                    text="PPT Gesture Control System",
                    font=("Arial", 20)
                )

                self.label.pack(pady=20)

                # Open PPT Button
                self.open_btn = tk.Button(
                    root,
                    text="Open PPT",
                    command=self.open_ppt,
                    width=20,
                    height=2
                )

                self.open_btn.pack(pady=10)

                # Start Gesture Button
                self.start_btn = tk.Button(
                    root,
                    text="Start Gesture Control",
                    command=self.start_gesture,
                    width=20,
                    height=2
                )

                self.start_btn.pack(pady=10)

            # =====================================================
            # OPEN POWERPOINT
            # =====================================================
            def open_ppt(self):

                ppt_path = filedialog.askopenfilename(
                    filetypes=[("PowerPoint Files", "*.pptx")]
                )

                if ppt_path:

                    try:
                        self.powerpoint = win32com.client.Dispatch(
                            "PowerPoint.Application"
                        )

                        self.powerpoint.Visible = 1

                        self.presentation = self.powerpoint.Presentations.Open(
                            ppt_path
                        )

                        self.presentation.SlideShowSettings.Run()

                        messagebox.showinfo(
                            "Success",
                            "PowerPoint Opened Successfully"
                        )

                    except Exception as e:

                        messagebox.showerror(
                            "Error",
                            str(e)
                        )

            # =====================================================
            # START GESTURE CONTROL
            # =====================================================
            def start_gesture(self):

                cap = cv2.VideoCapture(0)

                if not cap.isOpened():
                    messagebox.showerror(
                        "Error",
                        "Camera not detected"
                    )

                    return

                while True:

                    ret, frame = cap.read()

                    if not ret:
                        break

                    frame = cv2.flip(frame, 1)

                    cv2.putText(
                        frame,
                        "Press N = Next Slide",
                        (20, 40),
                        cv2.FONT_HERSHEY_SIMPLEX,
                        0.8,
                        (0, 255, 0),
                        2
                    )

                    cv2.putText(
                        frame,
                        "Press P = Previous Slide",
                        (20, 80),
                        cv2.FONT_HERSHEY_SIMPLEX,
                        0.8,
                        (0, 255, 0),
                        2
                    )

                    cv2.putText(
                        frame,
                        "Press Q = Quit",
                        (20, 120),
                        cv2.FONT_HERSHEY_SIMPLEX,
                        0.8,
                        (0, 255, 0),
                        2
                    )

                    cv2.imshow("Gesture Control", frame)

                    key = cv2.waitKey(1) & 0xFF

                    # Next Slide
                    if key == ord('n'):

                        pyautogui.press("right")

                        time.sleep(GESTURE_DELAY)

                    # Previous Slide
                    elif key == ord('p'):

                        pyautogui.press("left")

                        time.sleep(GESTURE_DELAY)

                    # Quit
                    elif key == ord('q'):

                        break

                cap.release()

                cv2.destroyAllWindows()

            # =====================================================
            # CLEANUP
            # =====================================================
            def cleanup(self):

                cv2.destroyAllWindows()

                try:
                    self.presentation.Close()
                    self.powerpoint.Quit()
                except:
                    pass

        # =========================================================
        # CLOSE FUNCTION
        # =========================================================
        def on_closing():

            app.cleanup()

            root.destroy()

        # =========================================================
        # MAIN
        # =========================================================
        if __name__ == "__main__":
            root = tk.Tk()

            app = PPTGestureControl(root)

            root.protocol("WM_DELETE_WINDOW", on_closing)

            root.mainloop()
        if self.powerpoint:
            self.powerpoint.Quit()

        self.root.destroy()


# =========================================================
# MAIN
# =========================================================
if __name__ == "__main__":

    root = tk.Tk()

    app = PPTGestureControl(root)

    root.protocol(
        "WM_DELETE_WINDOW",
        app.cleanup
    )

    root.mainloop()