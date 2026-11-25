import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

import gui_app
from command_dispatcher import CommandDispatcher
from gesture_input import GestureInput, classify_landmarks


class CommandDispatcherTests(unittest.TestCase):
    def setUp(self):
        self.dispatcher = CommandDispatcher()

    def test_search_aliases_normalize_to_confirmed_search_intents(self):
        intents = [
            self.dispatcher.classify_text("Search for local-first AI"),
            self.dispatcher.classify_text("search the web for local-first AI"),
            self.dispatcher.classify_text("web search local-first AI"),
        ]

        self.assertTrue(all(intent.action == "web_search" for intent in intents))
        self.assertTrue(all(intent.requires_confirmation for intent in intents))
        self.assertEqual({intent.argument for intent in intents}, {"local-first AI"})

    def test_unrecognized_or_shell_like_text_is_not_dispatched(self):
        self.assertIsNone(self.dispatcher.classify_text("run whoami"))
        self.assertIsNone(self.dispatcher.classify_text("open an arbitrary application"))

    def test_sensitive_pc_actions_require_confirmation_and_validate_inputs(self):
        lock = self.dispatcher.classify_text("lock laptop")
        read_file = self.dispatcher.classify_text("read file notes.txt")
        website = self.dispatcher.classify_text("open website example.com")
        invalid_file = self.dispatcher.classify_text("read file ..\\secret.txt")
        invalid_website = self.dispatcher.classify_text("open website javascript:alert(1)")

        self.assertTrue(lock.requires_confirmation)
        self.assertTrue(read_file.requires_confirmation)
        self.assertTrue(website.requires_confirmation)
        self.assertEqual(website.argument, "open website https://example.com")
        self.assertEqual(invalid_file.action, "invalid")
        self.assertEqual(invalid_website.action, "invalid")

    def test_gesture_intents_are_limited_to_safe_ui_and_voice_controls(self):
        self.assertEqual(
            self.dispatcher.classify_gesture("OPEN_PALM").action,
            "start_listening",
        )
        self.assertEqual(
            self.dispatcher.classify_gesture("two fingers").action,
            "show_chat",
        )
        self.assertEqual(
            self.dispatcher.classify_gesture("fist").action,
            "stop_gestures",
        )
        self.assertIsNone(self.dispatcher.classify_gesture("pinch"))


class GestureInputTests(unittest.TestCase):
    @staticmethod
    def landmarks(extended):
        points = [SimpleNamespace(x=0.5, y=0.5) for _ in range(21)]
        for index, (tip, pip) in enumerate(((8, 6), (12, 10), (16, 14), (20, 18))):
            points[tip].y = 0.1 if extended[index] else 0.8
            points[pip].y = 0.4
        return points

    def test_landmark_classifier_recognizes_only_mvp_gestures(self):
        self.assertEqual(
            classify_landmarks(self.landmarks([True, True, True, True])),
            "open palm",
        )
        self.assertEqual(
            classify_landmarks(self.landmarks([True, True, False, False])),
            "two fingers",
        )
        self.assertEqual(
            classify_landmarks(self.landmarks([False, False, False, False])),
            "fist",
        )
        self.assertIsNone(classify_landmarks(self.landmarks([True, False, True, False])))

    def test_start_reports_missing_optional_camera_dependencies(self):
        statuses = []
        gestures = []
        service = GestureInput(gestures.append, statuses.append)
        with patch("gesture_input.importlib.util.find_spec", return_value=None):
            started = service.start()

        self.assertFalse(started)
        self.assertEqual(gestures, [])
        self.assertIn("requirements-gesture.txt", statuses[0])


class ActionAuditTests(unittest.TestCase):
    def test_action_audit_persists_metadata_without_command_parameters(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            database = Path(temp_dir) / "alpha_memory.db"
            with patch.object(gui_app, "DB_FILE", database), patch.object(
                gui_app, "LEGACY_DB_FILE", Path(temp_dir) / "missing.db"
            ):
                memory = gui_app.AlphaMemory()
                intent = CommandDispatcher().classify_text(
                    "read file private-notes.txt"
                )
                memory.log_action("event-1", intent.audit_key, "voice", "denied")
                history = memory.get_action_history()
                memory.close()

        self.assertEqual(history[0][0:3], ("pc.read_file", "voice", "denied"))
        self.assertNotIn("private-notes.txt", repr(history))

    def test_denied_sensitive_action_never_reaches_executor(self):
        events = []
        intent = CommandDispatcher().classify_text("lock laptop")
        ui = SimpleNamespace(
            memory=SimpleNamespace(
                log_action=lambda *event: events.append(event)
            ),
            action_confirmation_text=gui_app.AlphaGUI.action_confirmation_text,
            respond_to_user=lambda _message: None,
            start_pc_action=lambda *_args: self.fail("Denied action was executed"),
        )

        with patch.object(gui_app.messagebox, "askyesno", return_value=False):
            gui_app.AlphaGUI.handle_command_intent(ui, intent, "lock laptop", "voice")

        self.assertEqual([event[3] for event in events], ["requested", "denied"])

    def test_approved_sensitive_action_is_forwarded_to_executor(self):
        events = []
        execution = []
        intent = CommandDispatcher().classify_text("lock laptop")
        ui = SimpleNamespace(
            memory=SimpleNamespace(
                log_action=lambda *event: events.append(event)
            ),
            action_confirmation_text=gui_app.AlphaGUI.action_confirmation_text,
            start_pc_action=lambda *args: execution.append(args),
        )

        with patch.object(gui_app.messagebox, "askyesno", return_value=True):
            gui_app.AlphaGUI.handle_command_intent(ui, intent, "lock laptop", "voice")

        self.assertEqual([event[3] for event in events], ["requested", "approved"])
        self.assertEqual(len(execution), 1)


if __name__ == "__main__":
    unittest.main()
