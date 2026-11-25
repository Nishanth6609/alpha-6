import json
import os
from pathlib import Path


AUDIO_SETTINGS_FILE = (
    Path(os.environ.get("APPDATA", str(Path.home())))
    / "A.L.P.H.A-6"
    / "audio_settings.json"
)


def list_input_devices():
    try:
        import sounddevice as sd
    except ModuleNotFoundError as error:
        raise RuntimeError("Microphone input requires sounddevice.") from error

    devices = sd.query_devices()
    host_apis = sd.query_hostapis()
    try:
        default_device = sd.default.device[0]
    except (AttributeError, IndexError, TypeError):
        default_device = None

    inputs = []
    for index, device in enumerate(devices):
        if int(device.get("max_input_channels", 0)) <= 0:
            continue
        host_api_index = int(device.get("hostapi", -1))
        host_api = (
            host_apis[host_api_index].get("name", "Unknown")
            if 0 <= host_api_index < len(host_apis)
            else "Unknown"
        )
        inputs.append({
            "index": index,
            "name": str(device["name"]),
            "host_api": str(host_api),
            "is_default": index == default_device,
        })
    return inputs


def list_output_devices():
    try:
        import sounddevice as sd
    except ModuleNotFoundError as error:
        raise RuntimeError("Speaker output requires sounddevice.") from error

    devices = sd.query_devices()
    host_apis = sd.query_hostapis()
    try:
        default_device = sd.default.device[1]
    except (AttributeError, IndexError, TypeError):
        default_device = None

    outputs = []
    for index, device in enumerate(devices):
        if int(device.get("max_output_channels", 0)) <= 0:
            continue
        host_api_index = int(device.get("hostapi", -1))
        host_api = (
            host_apis[host_api_index].get("name", "Unknown")
            if 0 <= host_api_index < len(host_apis)
            else "Unknown"
        )
        outputs.append({
            "index": index,
            "name": str(device["name"]),
            "host_api": str(host_api),
            "is_default": index == default_device,
        })
    return outputs


def get_saved_audio_settings(path=AUDIO_SETTINGS_FILE):
    try:
        with open(path, "r", encoding="utf-8") as settings_file:
            return json.load(settings_file)
    except (FileNotFoundError, OSError, ValueError, TypeError):
        return {}


def get_selected_output_device(path=AUDIO_SETTINGS_FILE):
    selection = get_saved_audio_settings(path).get("output_device")
    if selection is None:
        return None
    if not isinstance(selection, dict):
        raise RuntimeError("The saved speaker selection is invalid. Choose an output again.")

    for device in list_output_devices():
        if (
            device["name"] == selection.get("name")
            and device["host_api"] == selection.get("host_api")
        ):
            return device["index"]

    raise RuntimeError(
        f"Selected speaker '{selection.get('name', 'Unknown')}' is unavailable. "
        "Reconnect it or select another speaker in Alpha."
    )


def save_selected_output_device(device_index, path=AUDIO_SETTINGS_FILE):
    selection = None
    if device_index is not None:
        device = next(
            (item for item in list_output_devices() if item["index"] == device_index),
            None,
        )
        if device is None:
            raise ValueError("Select an available speaker output device.")
        selection = {
            "name": device["name"],
            "host_api": device["host_api"],
        }

    settings_path = Path(path)
    settings_path.parent.mkdir(parents=True, exist_ok=True)
    settings = get_saved_audio_settings(settings_path)
    settings["output_device"] = selection
    temporary_path = settings_path.with_suffix(settings_path.suffix + ".tmp")
    try:
        with open(temporary_path, "w", encoding="utf-8") as settings_file:
            json.dump(settings, settings_file, indent=2)
        temporary_path.replace(settings_path)
    finally:
        if temporary_path.exists():
            temporary_path.unlink()


def selected_output_device_name():
    try:
        device_index = get_selected_output_device()
        if device_index is None:
            return "Windows default speaker"
        device = next(
            (item for item in list_output_devices() if item["index"] == device_index),
            None,
        )
        return device["name"] if device else "Selected speaker"
    except RuntimeError as error:
        print("Output device selection:", error)
        return "Windows default speaker"


def get_selected_input_device(path=AUDIO_SETTINGS_FILE):
    try:
        with open(path, "r", encoding="utf-8") as settings_file:
            selection = json.load(settings_file).get("input_device")
    except FileNotFoundError:
        return None
    except (OSError, ValueError, TypeError, AttributeError) as error:
        print("Audio settings could not be loaded; using the Windows default microphone:", error)
        return None

    if selection is None:
        return None
    if not isinstance(selection, dict):
        raise RuntimeError("The saved microphone selection is invalid. Choose an input again.")

    for device in list_input_devices():
        if (
            device["name"] == selection.get("name")
            and device["host_api"] == selection.get("host_api")
        ):
            return device["index"]

    raise RuntimeError(
        f"Selected microphone '{selection.get('name', 'Unknown')}' is unavailable. "
        "Reconnect it or select another microphone in Alpha."
    )


def save_selected_input_device(device_index, path=AUDIO_SETTINGS_FILE):
    selection = None
    if device_index is not None:
        device = next(
            (item for item in list_input_devices() if item["index"] == device_index),
            None,
        )
        if device is None:
            raise ValueError("Select an available microphone input device.")
        selection = {
            "name": device["name"],
            "host_api": device["host_api"],
        }

    settings_path = Path(path)
    settings_path.parent.mkdir(parents=True, exist_ok=True)
    settings = get_saved_audio_settings(settings_path)
    settings["input_device"] = selection
    temporary_path = settings_path.with_suffix(settings_path.suffix + ".tmp")
    try:
        with open(temporary_path, "w", encoding="utf-8") as settings_file:
            json.dump(settings, settings_file, indent=2)
        temporary_path.replace(settings_path)
    finally:
        if temporary_path.exists():
            temporary_path.unlink()


def selected_input_device_name():
    device_index = get_selected_input_device()
    if device_index is None:
        try:
            import sounddevice as sd
            device_index = sd.default.device[0]
            if device_index is not None:
                return str(sd.query_devices(device_index, "input")["name"])
        except (ImportError, IndexError, KeyError, TypeError, ValueError):
            pass
        return "Windows default microphone"
    device = next(
        (item for item in list_input_devices() if item["index"] == device_index),
        None,
    )
    return device["name"] if device else "Selected microphone"
