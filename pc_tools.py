import subprocess
import time

try:
    import psutil
except ModuleNotFoundError:  # pragma: no cover - optional dependency may be absent in tests.
    psutil = None

# Fixed allowlist only: no user-provided shell commands are executed.
APP_COMMANDS = {
    "notepad": ["notepad.exe"],
    "calculator": ["calc.exe"],
    "explorer": ["explorer.exe"],
}


def open_approved_app(name):
    key = name.lower().strip()
    if key not in APP_COMMANDS:
        raise ValueError("That application is not in A.L.P.H.A - 6's approved app list.")
    subprocess.Popen(APP_COMMANDS[key], shell=False)


def get_system_snapshot():
    if psutil is None:
        return {"cpu": "N/A", "ram": "N/A", "battery": "N/A", "uptime": "N/A"}

    try:
        cpu = psutil.cpu_percent(interval=None)
    except Exception:
        cpu = 0

    try:
        ram = psutil.virtual_memory().percent
    except Exception:
        ram = 0

    try:
        battery = psutil.sensors_battery()
        battery_text = f"{round(battery.percent)}%" if battery else "N/A"
    except Exception:
        battery_text = "N/A"

    try:
        uptime_seconds = max(0, int(time.time() - psutil.boot_time()))
        hours, rem = divmod(uptime_seconds, 3600)
        minutes = rem // 60
        uptime_text = f"{hours}h {minutes}m"
    except Exception:
        uptime_text = "N/A"

    return {
        "cpu": f"{cpu:.0f}%",
        "ram": f"{ram:.0f}%",
        "battery": battery_text,
        "uptime": uptime_text,
    }
