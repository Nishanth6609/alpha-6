import json
import tempfile
import types
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

import requests
import brain
import gui_app


class FakeMemory:
    def get_history(self):
        return []

    def get_facts(self):
        return []

    def save(self, _role, _message):
        pass


class FakeResponse:
    def __enter__(self):
        return self

    def __exit__(self, *_args):
        return False

    def raise_for_status(self):
        pass

    def iter_lines(self):
        yield json.dumps(
            {"message": {"content": "Answer."}, "done": True}
        ).encode("utf-8")


class AnswerConfigurationTests(unittest.TestCase):
    def test_ollama_readiness_accepts_the_installed_latest_model_tag(self):
        response = Mock()
        response.json.return_value = {"models": [{"name": "llama3.2:latest"}]}
        with patch.object(gui_app.requests, "get", return_value=response) as get:
            gui_app.check_ollama_ready()

        get.assert_called_once_with(
            "http://localhost:11434/api/tags", timeout=(3, 5)
        )

    def test_ollama_readiness_reports_missing_configured_model(self):
        response = Mock()
        response.json.return_value = {"models": [{"name": "qwen2.5:3b"}]}
        with patch.object(gui_app.requests, "get", return_value=response):
            with self.assertRaisesRegex(RuntimeError, "ollama pull llama3.2"):
                gui_app.check_ollama_ready()

    def test_ollama_readiness_retries_a_starting_server(self):
        response = Mock()
        response.json.return_value = {"models": [{"name": "llama3.2:latest"}]}
        with patch.object(
            gui_app.requests,
            "get",
            side_effect=[requests.exceptions.ConnectionError("starting"), response],
        ) as get, patch.object(gui_app.time, "sleep"):
            gui_app.check_ollama_ready()

        self.assertEqual(get.call_count, 2)

    def test_ollama_readiness_reports_server_connection_failure(self):
        with patch.object(
            gui_app.requests,
            "get",
            side_effect=requests.exceptions.ConnectionError("refused"),
        ), patch.object(gui_app.time, "sleep"):
            with self.assertRaisesRegex(RuntimeError, "Connection details: refused"):
                gui_app.check_ollama_ready()

    def test_answers_stream_with_consistent_temperature_and_clear_uncertainty(self):
        with patch.object(gui_app.requests, "post", return_value=FakeResponse()) as post:
            answer = gui_app.AlphaBrain(FakeMemory()).ask("What is 2 + 2?")

        self.assertEqual(answer, "Answer.")
        payload = post.call_args.kwargs["json"]
        self.assertTrue(payload["stream"])
        self.assertEqual(payload["keep_alive"], "10m")
        self.assertEqual(
            payload["options"]["temperature"], gui_app.OLLAMA_TEMPERATURE
        )
        self.assertIn("Do not guess", payload["messages"][0]["content"])

    def test_default_launcher_keeps_ollama_warm_and_uses_consistent_temperature(self):
        client = types.SimpleNamespace(
            chat=Mock(
                return_value=iter([{"message": {"content": "Answer."}}])
            )
        )
        with tempfile.TemporaryDirectory() as temp_dir:
            assistant = brain.AlphaBrain()
            assistant.memory_file = Path(temp_dir) / "memory.json"
            with patch.object(brain, "ollama", client):
                self.assertEqual(assistant.ask("What is 2 + 2?"), "Answer.")

        options = client.chat.call_args.kwargs
        self.assertEqual(
            options["options"]["temperature"], brain.OLLAMA_TEMPERATURE
        )
        self.assertEqual(options["keep_alive"], brain.OLLAMA_KEEP_ALIVE)
        self.assertIn("Do not guess", options["messages"][0]["content"])


if __name__ == "__main__":
    unittest.main()
