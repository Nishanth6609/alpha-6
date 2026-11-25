import json
import os
import tempfile
import threading
import wave
from pathlib import Path


XTTS_MODEL = "tts_models/multilingual/multi-dataset/xtts_v2"
VOICE_SETTINGS_FILE = (
    Path(os.environ.get("APPDATA", str(Path.home())))
    / "A.L.P.H.A-6"
    / "voice_settings.json"
)
VOICE_SAMPLE_FILE = VOICE_SETTINGS_FILE.parent / "my_voice.wav"
VOICE_SAMPLE_SECONDS = 12
CPML_CONSENT_FILE = VOICE_SETTINGS_FILE.parent / "cpml_noncommercial_consent.txt"


def validate_reference_wav(path):
    """Validate a short, uncompressed WAV sample before selecting it for XTTS."""
    wav_path = Path(path).expanduser()
    if not wav_path.is_file() or wav_path.suffix.lower() != ".wav":
        raise ValueError("Choose an existing WAV audio file.")

    try:
        with wave.open(str(wav_path), "rb") as sample:
            if sample.getcomptype() != "NONE":
                raise ValueError("The reference WAV must use uncompressed PCM audio.")
            if sample.getframerate() < 16000 or sample.getframerate() > 48000:
                raise ValueError("The reference WAV sample rate must be 16-48 kHz.")
            if sample.getnchannels() not in (1, 2):
                raise ValueError("The reference WAV must be mono or stereo.")
            duration = sample.getnframes() / sample.getframerate()
    except (wave.Error, EOFError) as error:
        raise ValueError("The selected file is not a readable PCM WAV.") from error

    if duration < 3 or duration > 30:
        raise ValueError("Use a clear voice sample between 3 and 30 seconds.")
    return wav_path.resolve()


def record_reference_voice(path=VOICE_SAMPLE_FILE):
    """Record a short personal voice sample locally for XTTS."""
    try:
        import sounddevice as sd
    except ModuleNotFoundError as error:
        raise RuntimeError("Voice recording requires sounddevice.") from error

    from audio_devices import get_selected_input_device

    device = get_selected_input_device()
    try:
        device_info = (
            sd.query_devices(device, "input")
            if device is not None
            else sd.query_devices(kind="input")
        )
    except Exception as error:
        raise RuntimeError(f"Could not inspect the selected microphone: {error}") from error

    native_rate = int(round(device_info.get("default_samplerate", 0)))
    candidate_rates = []
    for rate in (native_rate, 48000, 44100, 24000, 22050, 16000):
        if 16000 <= rate <= 48000 and rate not in candidate_rates:
            candidate_rates.append(rate)
    sample_rate = None
    for rate in candidate_rates:
        try:
            sd.check_input_settings(
                device=device,
                channels=1,
                dtype="int16",
                samplerate=rate,
            )
        except Exception:
            continue
        sample_rate = rate
        break
    if sample_rate is None:
        raise RuntimeError(
            "The selected microphone does not support a 16–48 kHz mono recording. "
            "Choose a different input in Alpha's MIC SOURCE menu."
        )

    try:
        samples = sd.rec(
            sample_rate * VOICE_SAMPLE_SECONDS,
            samplerate=sample_rate,
            channels=1,
            dtype="int16",
            device=device,
            blocking=True,
        )
    except Exception as error:
        raise RuntimeError(f"Could not record the voice sample: {error}") from error

    import numpy as np

    if samples.size == 0 or int(np.max(np.abs(samples.astype(np.int32)))) < 100:
        raise RuntimeError(
            "The recording was silent or too quiet. Check the selected microphone and try again."
        )

    sample_path = Path(path)
    sample_path.parent.mkdir(parents=True, exist_ok=True)
    temporary_path = sample_path.with_name(sample_path.stem + ".tmp.wav")
    try:
        with wave.open(str(temporary_path), "wb") as sample:
            sample.setnchannels(1)
            sample.setsampwidth(2)
            sample.setframerate(sample_rate)
            sample.writeframes(samples.tobytes())
        temporary_path.replace(sample_path)
    finally:
        if temporary_path.exists():
            temporary_path.unlink()
    return validate_reference_wav(sample_path)


def load_voice_settings(path=VOICE_SETTINGS_FILE):
    """Load optional voice preferences, falling back safely if the file is invalid."""
    try:
        with open(path, "r", encoding="utf-8") as settings_file:
            settings = json.load(settings_file)
        reference_wav = settings.get("reference_wav", "")
        use_xtts = settings.get("use_xtts", False) is True
        if use_xtts:
            reference_wav = str(validate_reference_wav(reference_wav))
        return {"use_xtts": use_xtts, "reference_wav": reference_wav}
    except FileNotFoundError:
        return {"use_xtts": False, "reference_wav": ""}
    except (OSError, ValueError, TypeError, AttributeError) as error:
        print("Voice settings could not be loaded; spoken replies are disabled:", error)
        return {"use_xtts": False, "reference_wav": ""}


def save_voice_settings(use_xtts, reference_wav, path=VOICE_SETTINGS_FILE):
    """Save voice selection without copying or uploading the selected sample."""
    sample_path = str(Path(reference_wav).expanduser()) if reference_wav else ""
    if use_xtts:
        sample_path = str(validate_reference_wav(reference_wav))

    settings_path = Path(path)
    settings_path.parent.mkdir(parents=True, exist_ok=True)
    temporary_path = settings_path.with_suffix(settings_path.suffix + ".tmp")
    try:
        with open(temporary_path, "w", encoding="utf-8") as settings_file:
            json.dump(
                {"use_xtts": bool(use_xtts), "reference_wav": sample_path},
                settings_file,
                indent=2,
            )
        temporary_path.replace(settings_path)
    finally:
        if temporary_path.exists():
            temporary_path.unlink()
    return {"use_xtts": bool(use_xtts), "reference_wav": sample_path}


def has_cpml_consent():
    try:
        return CPML_CONSENT_FILE.read_text(encoding="utf-8").strip() == "agreed"
    except FileNotFoundError:
        return False


def save_cpml_consent():
    CPML_CONSENT_FILE.parent.mkdir(parents=True, exist_ok=True)
    CPML_CONSENT_FILE.write_text("agreed\n", encoding="utf-8")
    os.environ["COQUI_TOS_AGREED"] = "1"


def _trim_reference_silence(reference_path, output_path):
    """Write a PCM16 reference with edge silence removed, if enough speech remains."""
    import numpy as np

    with wave.open(str(reference_path), "rb") as source:
        if source.getcomptype() != "NONE" or source.getsampwidth() != 2:
            return False
        channels = source.getnchannels()
        sample_rate = source.getframerate()
        frames = source.getnframes()
        samples = np.frombuffer(source.readframes(frames), dtype=np.int16)

    if channels > 1:
        samples = samples.reshape(-1, channels)
        level_samples = samples.astype(np.float32).mean(axis=1)
    else:
        level_samples = samples.astype(np.float32)

    frame_size = max(1, int(sample_rate * 0.02))
    levels = np.array([
        np.sqrt(np.mean(np.square(level_samples[start:start + frame_size])))
        for start in range(0, len(level_samples), frame_size)
    ])
    if not len(levels):
        return False

    threshold = max(80.0, float(levels.max()) * 0.08)
    active_frames = np.flatnonzero(levels >= threshold)
    if not len(active_frames):
        return False

    padding = int(sample_rate * 0.25)
    start = max(0, int(active_frames[0]) * frame_size - padding)
    end = min(
        len(level_samples),
        (int(active_frames[-1]) + 1) * frame_size + padding,
    )
    if end - start < sample_rate * 3 or (start == 0 and end == len(level_samples)):
        return False

    trimmed = samples[start:end]
    with wave.open(str(output_path), "wb") as target:
        target.setnchannels(channels)
        target.setsampwidth(2)
        target.setframerate(sample_rate)
        target.writeframes(trimmed.tobytes())
    return True


class XTTSVoice:
    """Optional local XTTS-v2 synthesis; imports and initializes the model on demand."""

    def __init__(self):
        self.model = None
        self.model_lock = threading.Lock()

    def _load_model(self):
        with self.model_lock:
            if self.model is not None:
                return self.model
            if not has_cpml_consent():
                raise RuntimeError(
                    "XTTS requires your agreement to the non-commercial Coqui "
                    "Public Model License. Review https://tts-hub.github.io/cpml "
                    "and accept the license prompt in Alpha to enable speech."
                )
            os.environ["COQUI_TOS_AGREED"] = "1"
            try:
                from TTS.api import TTS
            except ImportError as error:
                raise RuntimeError(
                    "XTTS-v2 is not installed in the app's Python environment."
                ) from error

            try:
                import torch
                use_gpu = torch.cuda.is_available()
            except ImportError:
                use_gpu = False

            self.model = TTS(model_name=XTTS_MODEL, gpu=use_gpu)
        return self.model

    def preload(self):
        """Load XTTS in the background so the first reply does not pay startup cost."""
        self._load_model()

    def speak(self, text, reference_wav):
        if any(character.isalpha() and not character.isascii() for character in text):
            raise RuntimeError(
                "XTTS-v2 currently supports English only; this reply cannot be "
                "synthesized in the enrolled voice."
            )
        wav_path = validate_reference_wav(reference_wav)
        model = self.model or self._load_model()

        with tempfile.NamedTemporaryFile(suffix=".wav", delete=False) as output:
            output_path = Path(output.name)
        reference_temp_path = None
        try:
            with tempfile.NamedTemporaryFile(suffix=".wav", delete=False) as reference_temp:
                reference_temp_path = Path(reference_temp.name)
            if _trim_reference_silence(wav_path, reference_temp_path):
                speaker_wav = reference_temp_path
            else:
                reference_temp_path.unlink(missing_ok=True)
                reference_temp_path = None
                speaker_wav = wav_path
            model.tts_to_file(
                text=text,
                speaker_wav=str(speaker_wav),
                language="en",
                file_path=str(output_path),
            )
            try:
                import numpy as np
                import sounddevice as sd
                from audio_devices import get_selected_output_device
            except ImportError:
                import winsound
                winsound.PlaySound(str(output_path), winsound.SND_FILENAME)
            else:
                with wave.open(str(output_path), "rb") as generated:
                    channels = generated.getnchannels()
                    sample_width = generated.getsampwidth()
                    sample_rate = generated.getframerate()
                    frames = generated.readframes(generated.getnframes())
                if sample_width != 2:
                    raise RuntimeError("XTTS generated unsupported WAV audio.")
                samples = np.frombuffer(frames, dtype=np.int16)
                if channels > 1:
                    samples = samples.reshape(-1, channels)
                sd.play(
                    samples,
                    samplerate=sample_rate,
                    device=get_selected_output_device(),
                    blocking=True,
                )
        finally:
            output_path.unlink(missing_ok=True)
            if reference_temp_path:
                reference_temp_path.unlink(missing_ok=True)
