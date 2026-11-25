import sqlite3
import struct
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import voice
import pc_control
from command_dispatcher import CommandDispatcher


class FakeSamples:
    def __init__(self, value, count=1600):
        self.value = value
        self.count = count

    def reshape(self, _shape):
        return self

    def copy(self):
        return self

    def __len__(self):
        return self.count

    def __iter__(self):
        return iter([self.value] * self.count)

    def tobytes(self):
        return struct.pack("<h", self.value) * self.count


class FakeInputStream:
    def __init__(self, blocks):
        self.blocks = iter(blocks)

    def __enter__(self):
        return self

    def __exit__(self, *_args):
        return False

    def read(self, _frames):
        return next(self.blocks), False


@unittest.skipIf(voice.sd is None or voice.sr is None, "optional voice packages are unavailable")
class VoiceRecognitionTests(unittest.TestCase):
    def test_records_long_phrases_and_stops_after_silence(self):
        blocks = [FakeSamples(100) for _ in range(5)]
        blocks += [FakeSamples(400) for _ in range(70)]
        blocks += [FakeSamples(0) for _ in range(10)]
        blocks += [FakeSamples(400) for _ in range(4)]
        stream = FakeInputStream(blocks)

        with patch.object(voice.sd, "InputStream", return_value=stream) as input_stream, \
             patch.object(voice, "get_selected_input_device", return_value=7):
            audio = voice.record_voice_audio()

        self.assertGreater(len(audio.frame_data), 5 * voice.SAMPLE_RATE * 2)
        self.assertEqual(audio.sample_rate, voice.SAMPLE_RATE)
        self.assertEqual(input_stream.call_args.kwargs["device"], 7)

    def test_recognizer_uses_existing_online_service_and_locale(self):
        audio = object()
        with patch.object(voice, "record_voice_audio", return_value=audio), \
             patch.object(voice.sr, "Recognizer") as recognizer_class:
            recognizer_class.return_value.recognize_google.return_value = "recognized phrase"

            result = voice.recognize_voice(language="en-IN")

        self.assertEqual(result, "recognized phrase")
        recognizer_class.return_value.recognize_google.assert_called_once_with(
            audio, language="en-IN"
        )


class MemoryMigrationTests(unittest.TestCase):
    def test_copies_legacy_database_without_modifying_it(self):
        import gui_app

        with tempfile.TemporaryDirectory() as temp_dir:
            directory = Path(temp_dir)
            legacy_path = directory / "sunday_memory.db"
            new_path = directory / "alpha_memory.db"
            legacy = sqlite3.connect(legacy_path)
            legacy.execute(
                "CREATE TABLE chats (id INTEGER PRIMARY KEY, role TEXT, message TEXT, created_at TEXT)"
            )
            legacy.execute(
                "INSERT INTO chats (role, message, created_at) VALUES (?, ?, ?)",
                ("user", "keep this chat", "2026-01-01"),
            )
            legacy.commit()
            legacy.close()
            original_bytes = legacy_path.read_bytes()

            with patch.object(gui_app, "DB_FILE", new_path), \
                 patch.object(gui_app, "LEGACY_DB_FILE", legacy_path):
                memory = gui_app.AlphaMemory()
                try:
                    self.assertEqual(memory.get_history(), [
                        {"role": "user", "content": "keep this chat"}
                    ])
                finally:
                    memory.db.close()

            self.assertEqual(legacy_path.read_bytes(), original_bytes)
            self.assertTrue(new_path.is_file())


class VoiceCommandRenameTests(unittest.TestCase):
    def test_alpha_wake_name_forms_still_route_approved_commands(self):
        phrases = (
            "alpha open notepad",
            "alpha six, open notepad",
            "hey alpha six, please open notepad",
        )
        with patch.object(pc_control, "open_application", return_value="ok") as open_app:
            for phrase in phrases:
                with self.subTest(phrase=phrase):
                    self.assertEqual(pc_control.run_pc_command(phrase), "ok")
            self.assertEqual(open_app.call_count, len(phrases))

    def test_arbitrary_system_commands_are_not_executed(self):
        import gui_app

        with patch("subprocess.Popen") as popen:
            self.assertIsNone(pc_control.run_pc_command("run powershell to delete files"))
            self.assertIsNone(gui_app.AlphaControl.execute("run powershell to delete files"))
            self.assertIsNone(gui_app.AlphaControl.execute("delete all files"))
        popen.assert_not_called()

    def test_voice_chat_commands_open_chat_only_when_requested(self):
        import gui_app
        from unittest.mock import Mock

        class FakeEntry:
            def __init__(self, value):
                self.value = value

            def get(self):
                return self.value

            def delete(self, *_args):
                self.value = ""

        gui = gui_app.AlphaGUI.__new__(gui_app.AlphaGUI)
        gui.busy = False
        gui.voice_only = True
        gui.command_dispatcher = CommandDispatcher()
        gui.memory = Mock()
        gui.entry = FakeEntry("open chat")
        gui.voice = Mock()
        with patch.object(gui, "show_chat") as show_chat:
            gui.send_message()
        show_chat.assert_called_once_with()
        gui.voice.speak.assert_called_once_with("Opening chat.")

        gui.voice_only = True
        gui.entry = FakeEntry("close chat")
        gui.voice.speak.reset_mock()
        with patch.object(gui, "show_dashboard") as show_dashboard:
            gui.send_message()
        show_dashboard.assert_called_once_with()
        gui.voice.speak.assert_called_once_with("Switching to voice mode.")


if __name__ == "__main__":
    unittest.main()
