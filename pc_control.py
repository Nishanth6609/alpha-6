"""
A.L.P.H.A - 6 AI - Safe Windows Control Engine
New additive module. Does not replace existing GUI or brain.
"""

import os
import re
import platform
import shutil
import subprocess
import webbrowser
from urllib.parse import quote_plus


APPROVED_APPS = {
    "notepad": ["notepad.exe"],
    "calculator": ["calc.exe"],
    "file explorer": ["explorer.exe"],
    "explorer": ["explorer.exe"],
}


def open_application(app_name):
    """Open only explicitly approved Windows applications."""

    app_name = app_name.lower().strip()

    if app_name not in APPROVED_APPS:
        return "This application is not in A.L.P.H.A - 6's approved list."

    try:
        subprocess.Popen(
            APPROVED_APPS[app_name],
            shell=False
        )
        return f"Opening {app_name}."

    except OSError as error:
        return f"Could not open {app_name}: {error}"


def open_website(site):
    """Open a small allowlist of websites."""

    allowed_sites = {
        "google": "https://www.google.com",
        "youtube": "https://www.youtube.com",
        "github": "https://github.com",
        "wikipedia": "https://www.wikipedia.org",
    }

    site = site.lower().strip()

    if site not in allowed_sites:
        return "That website is not in the approved list."

    webbrowser.open(allowed_sites[site])
    return f"Opening {site}."


def search_web(query):
    """Open a Google search in the default browser."""

    query = query.strip()

    if not query:
        return "Please provide something to search for."

    if len(query) > 300:
        return "Search query is too long."

    url = "https://www.google.com/search?q=" + quote_plus(query)
    webbrowser.open(url)

    return f"Searching the web for {query}."


def get_system_information():
    """Return basic computer information."""

    try:
        info = [
            f"Operating system: {platform.system()}",
            f"Windows version: {platform.version()}",
            f"Processor: {platform.processor()}",
            f"CPU cores: {os.cpu_count()}",
        ]

        try:
            import psutil

            ram = psutil.virtual_memory()
            disk = shutil.disk_usage(os.path.expanduser("~"))

            info.append(
                f"RAM available: {ram.available / (1024**3):.2f} GB"
            )
            info.append(
                f"RAM usage: {ram.percent}%"
            )
            info.append(
                f"Disk free: {disk.free / (1024**3):.2f} GB"
            )

        except ImportError:
            info.append("Install psutil for additional system information.")

        return "\n".join(info)

    except Exception as error:
        return f"System information error: {error}"


def run_pc_command(command):
    """
    Recognize specific safe commands.
    Return None when the request is not a supported PC action.
    """

    if not isinstance(command, str):
        return None

    text = command.lower().strip()

    # Open approved applications
    match = re.fullmatch(
        r"(?:(?:hey )?alpha(?: (?:6|six))?[, ]*)?(?:please )?open "
        r"(notepad|calculator|file explorer|explorer)",
        text
    )

    if match:
        return open_application(match.group(1))

    # Open approved websites
    match = re.fullmatch(
        r"(?:(?:hey )?alpha(?: (?:6|six))?[, ]*)?(?:please )?open "
        r"(google|youtube|github|wikipedia)",
        text
    )

    if match:
        return open_website(match.group(1))

    # Search the web
    match = re.fullmatch(
        r"(?:(?:hey )?alpha(?: (?:6|six))?[, ]*)?search (?:the )?web for (.+)",
        text
    )

    if match:
        return search_web(match.group(1))

    # Laptop status
    if text in (
        "system information",
        "computer information",
        "check my laptop",
        "alpha system status",
        "alpha 6 system status",
        "alpha six system status",
        "hey alpha six system status"
    ):
        return get_system_information()

    return None