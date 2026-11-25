"""Optional, local-only webcam gesture input for A.L.P.H.A-6."""

import importlib.util
import threading
import time


class GestureInput:
    def __init__(self, on_gesture, on_status=None, camera_index=0, stable_frames=8):
        self.on_gesture = on_gesture
        self.on_status = on_status
        self.camera_index = camera_index
        self.stable_frames = stable_frames
        self._stop_event = threading.Event()
        self._thread = None

    @staticmethod
    def availability_error():
        missing = []
        for package in ("cv2", "mediapipe"):
            try:
                available = importlib.util.find_spec(package) is not None
            except (ImportError, ValueError):
                available = False
            if not available:
                missing.append(package)
        if missing:
            packages = ", ".join(
                "opencv-python" if package == "cv2" else package
                for package in missing
            )
            return "Gesture mode needs optional packages: " + packages + ". See requirements-gesture.txt."
        return None

    @property
    def is_running(self):
        return self._thread is not None and self._thread.is_alive()

    def start(self):
        if self.is_running:
            return True
        error = self.availability_error()
        if error:
            self._report_status(error)
            return False
        self._stop_event.clear()
        self._thread = threading.Thread(
            target=self._run,
            name="AlphaGestureInput",
            daemon=True,
        )
        self._thread.start()
        return True

    def stop(self):
        self._stop_event.set()

    def _report_status(self, message):
        if self.on_status:
            self.on_status(message)

    def _run(self):
        camera = None
        try:
            import cv2
            import mediapipe as mp

            camera = cv2.VideoCapture(self.camera_index)
            if not camera.isOpened():
                raise RuntimeError("The default camera could not be opened.")
            self._report_status(
                "Gesture camera active. Frames stay on this device and are not saved."
            )
            candidate = None
            candidate_frames = 0
            last_emitted = None
            with mp.solutions.hands.Hands(
                max_num_hands=1,
                min_detection_confidence=0.6,
                min_tracking_confidence=0.6,
            ) as detector:
                while not self._stop_event.is_set():
                    success, frame = camera.read()
                    if not success or frame is None:
                        raise RuntimeError("The gesture camera stopped returning frames.")

                    rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                    result = detector.process(rgb)
                    gesture = None
                    if result.multi_hand_landmarks:
                        gesture = classify_landmarks(result.multi_hand_landmarks[0])

                    if gesture is None:
                        candidate = None
                        candidate_frames = 0
                        last_emitted = None
                        continue

                    if gesture == candidate:
                        candidate_frames += 1
                    else:
                        candidate = gesture
                        candidate_frames = 1

                    if candidate_frames >= self.stable_frames and gesture != last_emitted:
                        last_emitted = gesture
                        self.on_gesture(gesture)
                    time.sleep(0.01)
        except Exception as error:
            if not self._stop_event.is_set():
                self._report_status("Gesture mode unavailable: " + str(error))
        finally:
            if camera is not None:
                try:
                    camera.release()
                except Exception as error:
                    self._report_status("Gesture camera could not be released: " + str(error))
            self._report_status("Gesture camera stopped.")


def classify_landmarks(hand_landmarks):
    """Classify stable open-palm, two-finger, and fist poses from hand landmarks."""
    landmarks = getattr(hand_landmarks, "landmark", hand_landmarks)
    if len(landmarks) < 21:
        return None

    extended = [
        landmarks[tip].y < landmarks[pip].y
        for tip, pip in ((8, 6), (12, 10), (16, 14), (20, 18))
    ]
    if all(extended):
        return "open palm"
    if extended[:2] == [True, True] and extended[2:] == [False, False]:
        return "two fingers"
    if not any(extended):
        return "fist"
    return None
