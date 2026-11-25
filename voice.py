import math
import queue
import statistics
import threading
from collections import deque
from voice_clone import XTTSVoice, load_voice_settings, save_voice_settings
from audio_devices import get_selected_input_device

try:
    import sounddevice as sd
except ModuleNotFoundError:  # pragma: no cover - optional runtime dependency may be absent.
    sd = None

try:
    import speech_recognition as sr
except ModuleNotFoundError:  # pragma: no cover - optional runtime dependency may be absent.
    sr = None

class VoiceEngine:
    def __init__(self, on_error=None):
        self.engine = None
        self.voice_lock = threading.Lock()
        self.settings_lock = threading.Lock()
        self.voice_settings = load_voice_settings()
        self.xtts = XTTSVoice()
        self.on_error = on_error
        self.xtts_error_shown = False
        self.tasks = queue.Queue()
        self.worker = threading.Thread(target=self._voice_worker, daemon=True)
        self.worker.start()

    def _queue_speech(self, text, wait):
        text = str(text or "").strip()
        if not text:
            return
        finished = threading.Event() if wait else None
        self.tasks.put((text, finished))
        if finished:
            finished.wait()

    def speak(self, text):
        self._queue_speech(text, wait=True)

    def speak_async(self, text):
        self._queue_speech(text, wait=False)

    def _voice_worker(self):
        while True:
            task = self.tasks.get()
            if task is None:
                break
            text, finished = task
            try:
                with self.voice_lock:
                    print("A.L.P.H.A - 6:", text)
                    with self.settings_lock:
                        settings = dict(self.voice_settings)
                    if settings["use_xtts"]:
                        self.xtts.speak(text, settings["reference_wav"])
                    else:
                        raise RuntimeError(
                            "Record your voice in Alpha before enabling spoken replies."
                        )
            except Exception as error:
                print("Voice error:", error)
                if not self.xtts_error_shown:
                    self.xtts_error_shown = True
                    if self.on_error:
                        try:
                            self.on_error(str(error))
                        except Exception as callback_error:
                            print("Voice error callback failed:", callback_error)
            finally:
                if finished:
                    finished.set()

    def stop(self):
        self.tasks.put(None)

    def set_voice_settings(self, use_xtts, reference_wav):
        settings = save_voice_settings(use_xtts, reference_wav)
        with self.settings_lock:
            self.voice_settings = settings
            self.xtts_error_shown = False
        return settings

    def get_voice_settings(self):
        with self.settings_lock:
            return dict(self.voice_settings)

    def preload(self):
        threading.Thread(target=self._preload_voice, daemon=True).start()

    def _preload_voice(self):
        try:
            self.xtts.preload()
        except Exception as error:
            print("XTTS background startup error:", error)
            if self.on_error:
                self.on_error(str(error))

    def listen(self):
        if sd is None or sr is None:
            print("Voice input is unavailable because the audio dependencies are not installed.")
            return ""

        try:
            with self.voice_lock:
                print("A.L.P.H.A - 6 is listening...")
                command = recognize_voice()
                print("YOU:", command)
                return command

        except sr.UnknownValueError:
            print("A.L.P.H.A - 6: I couldn't understand that.")
            return ""

        except sr.RequestError as error:
            print("Speech recognition connection error:", error)
            return ""

        except Exception as error:
            print("Voice error:", error)
            return ""


SAMPLE_RATE = 16000
BLOCK_SECONDS = 0.1
CALIBRATION_SECONDS = 0.4
END_SILENCE_SECONDS = 0.6
MAX_PHRASE_SECONDS = 20
PRE_ROLL_BLOCKS = 3


def _audio_level(samples):
    samples = samples.reshape(-1)
    if not len(samples):
        return 0.0
    return math.sqrt(sum(float(sample) ** 2 for sample in samples) / len(samples))


def record_voice_audio():
    """Capture a complete phrase with ambient-noise calibration and silence detection."""
    if sd is None or sr is None:
        raise RuntimeError("Voice input requires sounddevice and SpeechRecognition.")

    block_frames = int(SAMPLE_RATE * BLOCK_SECONDS)
    calibration_count = int(CALIBRATION_SECONDS / BLOCK_SECONDS)
    end_silence_blocks = int(END_SILENCE_SECONDS / BLOCK_SECONDS)
    max_phrase_blocks = int(MAX_PHRASE_SECONDS / BLOCK_SECONDS)
    with sd.InputStream(
        samplerate=SAMPLE_RATE,
        channels=1,
        dtype="int16",
        blocksize=block_frames,
        device=get_selected_input_device(),
    ) as stream:
        calibration = [stream.read(block_frames)[0].copy() for _ in range(calibration_count)]
        ambient_level = statistics.median(_audio_level(block) for block in calibration)
        speech_threshold = max(140.0, ambient_level * 1.8)
        pre_roll = deque(maxlen=PRE_ROLL_BLOCKS)
        phrase = []
        consecutive_speech = 0
        started = False

        for block in calibration:
            pre_roll.append(block)
            consecutive_speech = consecutive_speech + 1 if _audio_level(block) >= speech_threshold else 0
            if consecutive_speech >= 2:
                phrase.extend(pre_roll)
                started = True
                break

        silence_blocks = 0
        for _ in range(max_phrase_blocks):
            if started and silence_blocks >= end_silence_blocks:
                break

            block = stream.read(block_frames)[0].copy()
            level = _audio_level(block)
            if not started:
                pre_roll.append(block)
                consecutive_speech = consecutive_speech + 1 if level >= speech_threshold else 0
                if consecutive_speech >= 2:
                    phrase.extend(pre_roll)
                    started = True
            else:
                phrase.append(block)
                silence_blocks = 0 if level >= speech_threshold * 0.7 else silence_blocks + 1

    if not phrase:
        raise sr.UnknownValueError()

    return sr.AudioData(
        b"".join(block.tobytes() for block in phrase),
        SAMPLE_RATE,
        2,
    )


def recognize_voice(language="en-IN"):
    """Transcribe captured speech with the existing Google online recognizer."""
    if sr is None:
        raise RuntimeError("SpeechRecognition is not installed.")
    recognizer = sr.Recognizer()
    return recognizer.recognize_google(record_voice_audio(), language=language)