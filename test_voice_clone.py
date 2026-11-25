import os
import threading
import tempfile
import unittest
import wave
from types import SimpleNamespace
from pathlib import Path
from unittest.mock import patch

import voice
from voice_clone import (
    XTTSVoice,
    _trim_reference_silence,
    has_cpml_consent,
    load_voice_settings,
    record_reference_voice,
    save_cpml_consent,
    save_voice_settings,
    validate_reference_wav,
)


def write_sample(path, duration=4):
    with wave.open(str(path), "wb") as sample:
        sample.setnchannels(1)
        sample.setsampwidth(2)
        sample.setframerate(16000)
        sample.writeframes(b"\0\0" * 16000 * duration)


class VoiceCloneTests(unittest.TestCase):
    def test_accepts_pcm_wav_and_rejects_too_short_sample(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            sample = Path(temp_dir) / "reference.wav"
            write_sample(sample)
            self.assertEqual(validate_reference_wav(sample), sample.resolve())

            write_sample(sample, duration=2)
            with self.assertRaisesRegex(ValueError, "between 3 and 30 seconds"):
                validate_reference_wav(sample)

    def test_reference_silence_is_trimmed_to_preserve_more_speech_for_xtts(self):
        import numpy as np

        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "reference.wav"
            trimmed = Path(temp_dir) / "trimmed.wav"
            samples = np.concatenate((
                np.zeros(16000, dtype=np.int16),
                np.full(7 * 16000, 3000, dtype=np.int16),
                np.zeros(4 * 16000, dtype=np.int16),
            ))
            with wave.open(str(source), "wb") as wav:
                wav.setnchannels(1)
                wav.setsampwidth(2)
                wav.setframerate(16000)
                wav.writeframes(samples.tobytes())

            self.assertTrue(_trim_reference_silence(source, trimmed))
            with wave.open(str(trimmed), "rb") as wav:
                self.assertGreater(wav.getnframes() / wav.getframerate(), 7)
                self.assertLess(wav.getnframes() / wav.getframerate(), 8)

    def test_xtts_receives_the_trimmed_reference_recording(self):
        import numpy as np

        with tempfile.TemporaryDirectory() as temp_dir:
            sample = Path(temp_dir) / "reference.wav"
            samples = np.concatenate((
                np.zeros(16000, dtype=np.int16),
                np.full(7 * 16000, 3000, dtype=np.int16),
                np.zeros(4 * 16000, dtype=np.int16),
            ))
            with wave.open(str(sample), "wb") as wav:
                wav.setnchannels(1)
                wav.setsampwidth(2)
                wav.setframerate(16000)
                wav.writeframes(samples.tobytes())

            captured = {}

            class Model:
                def tts_to_file(self, *, speaker_wav, file_path, **_kwargs):
                    with wave.open(speaker_wav, "rb") as reference:
                        captured["reference_duration"] = (
                            reference.getnframes() / reference.getframerate()
                        )
                    with wave.open(file_path, "wb") as output:
                        output.setnchannels(1)
                        output.setsampwidth(2)
                        output.setframerate(16000)
                        output.writeframes(b"\0\0" * 16000)

            fake_sounddevice = SimpleNamespace(play=lambda *args, **kwargs: None)
            synthesizer = XTTSVoice()
            synthesizer.model = Model()
            with patch.dict("sys.modules", {"sounddevice": fake_sounddevice}), \
                 patch("audio_devices.get_selected_output_device", return_value=None):
                synthesizer.speak("Hello.", sample)

            self.assertGreater(captured["reference_duration"], 7)
            self.assertLess(captured["reference_duration"], 8)

    def test_settings_round_trip_and_keep_sample_when_windows_voice_is_selected(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            sample = Path(temp_dir) / "reference.wav"
            settings_file = Path(temp_dir) / "config" / "voice.json"
            write_sample(sample)

            saved = save_voice_settings(True, sample, settings_file)
            self.assertTrue(saved["use_xtts"])
            self.assertEqual(load_voice_settings(settings_file), saved)

            windows_settings = save_voice_settings(False, sample, settings_file)
            self.assertFalse(windows_settings["use_xtts"])
            self.assertEqual(windows_settings["reference_wav"], str(sample))
            self.assertEqual(load_voice_settings(settings_file), windows_settings)

    def test_xtts_model_is_not_loaded_until_requested(self):
        voice = XTTSVoice()
        self.assertIsNone(voice.model)

    def test_xtts_requires_license_consent_before_loading_model(self):
        voice = XTTSVoice()
        with patch("voice_clone.has_cpml_consent", return_value=False):
            with self.assertRaisesRegex(RuntimeError, "requires your agreement"):
                voice._load_model()
        self.assertIsNone(voice.model)

    def test_license_consent_is_only_saved_when_explicitly_requested(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            consent_file = Path(temp_dir) / "cpml_noncommercial_consent.txt"
            with patch("voice_clone.CPML_CONSENT_FILE", consent_file), \
                 patch.dict(os.environ):
                os.environ.pop("COQUI_TOS_AGREED", None)
                self.assertFalse(has_cpml_consent())
                save_cpml_consent()
                self.assertTrue(has_cpml_consent())
                self.assertEqual(os.environ["COQUI_TOS_AGREED"], "1")

    def test_voice_recording_uses_selected_input_and_saves_valid_local_wav(self):
        import numpy as np

        class FakeSoundDevice:
            @staticmethod
            def rec(frames, **kwargs):
                self.assertEqual(frames, 44100 * 12)
                self.assertEqual(kwargs["samplerate"], 44100)
                self.assertEqual(kwargs["device"], 5)
                self.assertTrue(kwargs["blocking"])
                return np.full((frames, 1), 500, dtype=np.int16)

            @staticmethod
            def query_devices(device, kind):
                self.assertEqual(device, 5)
                self.assertEqual(kind, "input")
                return {"default_samplerate": 44100}

            @staticmethod
            def check_input_settings(**kwargs):
                self.assertEqual(kwargs["samplerate"], 44100)

        with tempfile.TemporaryDirectory() as temp_dir:
            destination = Path(temp_dir) / "my_voice.wav"
            with patch.dict("sys.modules", {"sounddevice": FakeSoundDevice}), \
                 patch("audio_devices.get_selected_input_device", return_value=5):
                saved = record_reference_voice(destination)

            self.assertEqual(saved, destination.resolve())
            with wave.open(str(saved), "rb") as sample:
                self.assertEqual(sample.getframerate(), 44100)
                self.assertEqual(sample.getnframes(), 44100 * 12)

    def test_voice_recording_rejects_unsupported_native_rate_and_uses_supported_rate(self):
        import numpy as np

        class FakeSoundDevice:
            @staticmethod
            def query_devices(_device, _kind):
                return {"default_samplerate": 96000}

            @staticmethod
            def check_input_settings(**kwargs):
                if kwargs["samplerate"] == 48000:
                    return
                raise ValueError("unsupported rate")

            @staticmethod
            def rec(frames, **kwargs):
                self.assertEqual(frames, 48000 * 12)
                self.assertEqual(kwargs["samplerate"], 48000)
                return np.full((frames, 1), 500, dtype=np.int16)

        with tempfile.TemporaryDirectory() as temp_dir:
            destination = Path(temp_dir) / "my_voice.wav"
            with patch.dict("sys.modules", {"sounddevice": FakeSoundDevice}), \
                 patch("audio_devices.get_selected_input_device", return_value=8):
                saved = record_reference_voice(destination)
            with wave.open(str(saved), "rb") as sample:
                self.assertEqual(sample.getframerate(), 48000)

    def test_telugu_is_rejected_before_loading_model(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            sample = Path(temp_dir) / "reference.wav"
            write_sample(sample)
            voice = XTTSVoice()
            with self.assertRaisesRegex(RuntimeError, "supports English only"):
                voice.speak("తెలుగు", sample)
            self.assertIsNone(voice.model)

    def test_voice_engine_queues_synthesis_off_the_calling_thread(self):
        started = threading.Event()
        release = threading.Event()

        class BlockingSynthesizer:
            def speak(self, _text, _reference_wav):
                started.set()
                release.wait()

        with patch.object(
            voice, "load_voice_settings",
            return_value={"use_xtts": True, "reference_wav": "sample.wav"},
        ), patch.object(
            voice, "XTTSVoice", return_value=BlockingSynthesizer()
        ):
            engine = voice.VoiceEngine()
            engine.speak_async("A.L.P.H.A - 6")
            self.assertTrue(started.wait(timeout=2))
            self.assertTrue(engine.worker.is_alive())
            engine.stop()
            release.set()
            engine.worker.join(timeout=2)

        self.assertFalse(engine.worker.is_alive())

    def test_voice_engine_does_not_use_a_different_voice_when_not_enrolled(self):
        error_received = threading.Event()
        errors = []

        class UnusedSynthesizer:
            called = False

            def speak(self, *_args):
                self.called = True

        synthesizer = UnusedSynthesizer()
        with patch.object(
            voice, "load_voice_settings",
            return_value={"use_xtts": False, "reference_wav": ""},
        ), patch.object(
            voice, "XTTSVoice", return_value=synthesizer
        ):
            engine = voice.VoiceEngine(
                on_error=lambda error: (errors.append(error), error_received.set())
            )
            engine.speak_async("Test")
            self.assertTrue(error_received.wait(timeout=2))
            engine.stop()
            engine.worker.join(timeout=2)

        self.assertIn("Record your voice", errors[0])
        self.assertFalse(synthesizer.called)
        self.assertFalse(engine.worker.is_alive())

    def test_dashboard_voice_reports_xtts_errors_without_falling_back(self):
        import gui_app

        error_received = threading.Event()
        errors = []

        class FailingSynthesizer:
            @staticmethod
            def preload():
                return None

            def speak(self, *_args):
                raise RuntimeError("local XTTS failed")

        with patch.object(
            gui_app, "load_voice_settings",
            return_value={"use_xtts": True, "reference_wav": "sample.wav"},
        ), patch.object(
            gui_app, "XTTSVoice", return_value=FailingSynthesizer()
        ):
            engine = gui_app.AlphaVoice(
                on_error=lambda error: (errors.append(error), error_received.set())
            )
            engine.speak("Test")
            self.assertTrue(error_received.wait(timeout=2))
            engine.stop()
            engine.worker.join(timeout=2)

        self.assertIn("local XTTS failed", errors[0])
        self.assertFalse(engine.worker.is_alive())


if __name__ == "__main__":
    unittest.main()
