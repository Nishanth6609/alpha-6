"""Normalize supported user inputs into explicit, allowlisted command intents."""

from dataclasses import dataclass
from urllib.parse import urlparse


@dataclass(frozen=True)
class CommandIntent:
    action: str
    audit_key: str
    argument: str = ""
    requires_confirmation: bool = False
    message: str = ""


class CommandDispatcher:
    _SEARCH_PREFIXES = ("web search ", "search for ", "search ")
    _SAFE_PC_COMMANDS = {
        "system info",
        "open notepad",
        "open calculator",
        "open explorer",
        "open file explorer",
        "open settings",
        "open browser",
        "open task manager",
        "open taskmanager",
        "open paint",
        "open vscode",
        "open visual studio code",
        "system information",
        "computer information",
        "check my laptop",
        "lock computer",
        "lock my computer",
        "lock laptop",
        "open desktop",
        "open my desktop",
        "open downloads",
        "open my downloads",
        "open documents",
        "open my documents",
        "open pictures",
        "open my pictures",
        "open videos",
        "open music",
    }
    _FOLDERS = {"desktop", "downloads", "documents", "pictures", "videos", "music"}

    def classify_text(self, text):
        if not isinstance(text, str):
            return None

        original = text.strip()
        normalized = " ".join(original.casefold().split())
        if not normalized:
            return None

        if normalized in ("open chat", "show chat", "open conversation"):
            return CommandIntent("show_chat", "ui.show_chat")
        if normalized in ("close chat", "hide chat", "switch to voice mode", "voice mode"):
            return CommandIntent("show_dashboard", "ui.show_dashboard")

        if normalized.startswith(("remember that ", "remember ")):
            fact = original.split(" ", 1)[1].strip()
            if fact.casefold().startswith("that "):
                fact = fact[5:].strip()
            if not fact:
                return CommandIntent("invalid", "memory.remember", message="Tell me what you would like me to remember.")
            return CommandIntent("remember_fact", "memory.remember", argument=fact)
        if normalized in ("what do you remember about me", "show my memory", "show memories"):
            return CommandIntent("show_memory", "memory.show")
        if normalized in ("forget everything you remember", "clear my memory", "delete all memories"):
            return CommandIntent("clear_memory", "memory.clear")
        if normalized.startswith("forget "):
            target = original[7:].strip()
            if not target:
                return CommandIntent("invalid", "memory.forget", message="Tell me which saved fact to forget.")
            return CommandIntent("forget_memory", "memory.forget", argument=target)

        if normalized.startswith("search the web for "):
            query = original[len("search the web for "):].strip()
            return self._web_search_intent(query)

        for prefix in self._SEARCH_PREFIXES:
            if normalized.startswith(prefix):
                query = original[len(prefix):].strip()
                return self._web_search_intent(query)

        if normalized in ("browse my laptop", "browse laptop", "choose a file", "select a file"):
            return CommandIntent("browse_file", "pc.browse_file")

        if normalized in self._SAFE_PC_COMMANDS:
            if normalized in ("lock computer", "lock my computer", "lock laptop"):
                return CommandIntent(
                    "pc_command",
                    "pc.lock",
                    argument=normalized,
                    requires_confirmation=True,
                )
            return CommandIntent(
                "pc_command",
                "pc." + normalized.replace(" ", "_"),
                argument=normalized,
            )

        for prefix, action, confirm in (
            ("read file ", "pc.read_file", True),
            ("open file ", "pc.open_file", True),
            ("find file ", "pc.find_file", True),
        ):
            if normalized.startswith(prefix):
                filename = original[len(prefix):].strip().strip('"').strip("'")
                if (
                    not filename
                    or len(filename) > 180
                    or any(char in filename for char in ("/", "\\"))
                    or any(ord(char) < 32 for char in filename)
                ):
                    return CommandIntent(
                        "invalid",
                        action,
                        message="Provide a filename only (up to 180 characters), without a folder path.",
                    )
                return CommandIntent(
                    "pc_command",
                    action,
                    argument=prefix + filename,
                    requires_confirmation=confirm,
                )

        if normalized.startswith("open website "):
            website = original[len("open website "):].strip()
            if (
                not website
                or len(website) > 2048
                or any(char.isspace() or ord(char) < 32 for char in website)
            ):
                return CommandIntent("invalid", "pc.open_website", message="Provide a website address up to 2048 characters.")
            candidate = website if "://" in website else "https://" + website
            try:
                parsed = urlparse(candidate)
                valid_url = (
                    parsed.scheme in ("http", "https")
                    and bool(parsed.hostname)
                    and parsed.username is None
                    and parsed.password is None
                )
                if valid_url:
                    parsed.port
            except ValueError:
                valid_url = False
            if not valid_url:
                return CommandIntent("invalid", "pc.open_website", message="Only valid HTTP or HTTPS website addresses can be opened.")
            return CommandIntent(
                "pc_command",
                "pc.open_website",
                argument="open website " + candidate,
                requires_confirmation=True,
            )

        if normalized.startswith(("list my ", "list ")):
            phrase = normalized[8:] if normalized.startswith("list my ") else normalized[5:]
            phrase = phrase.replace(" folder", "").replace(" files", "").strip()
            if phrase in self._FOLDERS:
                return CommandIntent(
                    "pc_command",
                    "pc.list_folder",
                    argument="list " + phrase,
                    requires_confirmation=True,
                )
            return None

        return None

    def classify_gesture(self, gesture):
        if not isinstance(gesture, str):
            return None
        normalized = " ".join(gesture.casefold().replace("_", " ").split())
        actions = {
            "open palm": CommandIntent("start_listening", "gesture.listen"),
            "two fingers": CommandIntent("show_chat", "gesture.show_chat"),
            "fist": CommandIntent("stop_gestures", "gesture.stop"),
        }
        return actions.get(normalized)

    @staticmethod
    def _web_search_intent(query):
        if not query:
            return CommandIntent("invalid", "web.search", message="Please enter what you want to search for.")
        if len(query) > 300 or any(ord(char) < 32 for char in query):
            return CommandIntent("invalid", "web.search", message="Search queries must be 300 characters or fewer and contain no control characters.")
        return CommandIntent(
            "web_search",
            "web.search",
            argument=query,
            requires_confirmation=True,
        )
