import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import audio_devices


class FakeSoundDevice:
    def __init__(self, devices, default_input=0):
        self.devices = devices
        self.default = type("Defaults", (), {"device": (default_input, None)})()

    def query_devices(self):
        return self.devices

    @staticmethod
    def query_hostapis():
        return [{"name": "Windows WASAPI"}, {"name": "Windows DirectSound"}]


def input_device(name, host_api=0, channels=1):
    return {
        "name": name,
        "hostapi": host_api,
        "max_input_channels": channels,
    }


class AudioDeviceTests(unittest.TestCase):
    def test_lists_only_input_devices_and_marks_windows_default(self):
        fake_sd = FakeSoundDevice([
            input_device("Laptop microphone"),
            input_device("Earbuds microphone", host_api=1),
            input_device("Speakers", channels=0),
        ], default_input=1)

        with patch.dict("sys.modules", {"sounddevice": fake_sd}):
            devices = audio_devices.list_input_devices()

        self.assertEqual(
            devices,
            [
                {
                    "index": 0,
                    "name": "Laptop microphone",
                    "host_api": "Windows WASAPI",
                    "is_default": False,
                },
                {
                    "index": 1,
                    "name": "Earbuds microphone",
                    "host_api": "Windows DirectSound",
                    "is_default": True,
                },
            ],
        )

    def test_saved_device_resolves_by_name_if_system_indices_change(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "audio_settings.json"
            original_sd = FakeSoundDevice([
                input_device("Laptop microphone"),
                input_device("Earbuds microphone", host_api=1),
            ])
            with patch.dict("sys.modules", {"sounddevice": original_sd}):
                audio_devices.save_selected_input_device(1, path)

            changed_sd = FakeSoundDevice([
                input_device("Other microphone"),
                input_device("Laptop microphone"),
                input_device("Earbuds microphone", host_api=1),
            ])
            with patch.dict("sys.modules", {"sounddevice": changed_sd}):
                self.assertEqual(audio_devices.get_selected_input_device(path), 2)

            self.assertEqual(
                json.loads(path.read_text(encoding="utf-8"))["input_device"],
                {"name": "Earbuds microphone", "host_api": "Windows DirectSound"},
            )

    def test_unavailable_saved_device_is_reported_instead_of_silently_switching(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "audio_settings.json"
            path.write_text(
                json.dumps({
                    "input_device": {
                        "name": "Disconnected earbuds",
                        "host_api": "Windows WASAPI",
                    }
                }),
                encoding="utf-8",
            )
            fake_sd = FakeSoundDevice([input_device("Laptop microphone")])

            with patch.dict("sys.modules", {"sounddevice": fake_sd}):
                with self.assertRaisesRegex(RuntimeError, "Reconnect it"):
                    audio_devices.get_selected_input_device(path)

    def test_output_selection_is_saved_without_erasing_input_selection(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "audio_settings.json"
            fake_sd = FakeSoundDevice([
                input_device("Laptop microphone"),
                {
                    "name": "Earbuds speaker",
                    "hostapi": 1,
                    "max_input_channels": 0,
                    "max_output_channels": 2,
                },
            ])
            with patch.dict("sys.modules", {"sounddevice": fake_sd}):
                audio_devices.save_selected_input_device(0, path)
                audio_devices.save_selected_output_device(1, path)

            settings = json.loads(path.read_text(encoding="utf-8"))
            self.assertEqual(
                settings["input_device"],
                {"name": "Laptop microphone", "host_api": "Windows WASAPI"},
            )
            self.assertEqual(
                settings["output_device"],
                {"name": "Earbuds speaker", "host_api": "Windows DirectSound"},
            )


if __name__ == "__main__":
    unittest.main()
