import os
import json

CONFIG_FILENAME = "config.json"

config = {
    # Existing fields
    "TEST_NO": None,
    "TEST_DATE": None,
    "PROJECT": None,

    # New global fields
    "TEST_NAME": None,
    "REPORT_NO": None,
    "TEST_ID": None,
    "WO_NO": None,
    "OEM": None,
    "PROGRAM": None,
    "PURPOSE": None,

    # EVA fields
    "DUMMY_PCT": None,
    "SENSOR": None,

    # Seat Configuration
    "SEAT_COUNT": 1,

    # Per-Seat Dynamic Fields (Lists of length 5)
    "SMP_ID": [None, None, None, None, None],
    "TEST_SAMPLE": [None, None, None, None, None],
}


def _get_tempfiles_dir():
    root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(root_dir, "tempfiles")


def save_config():
    """Config'i tempfiles/config.json'a kaydeder."""
    tempfiles_dir = _get_tempfiles_dir()
    os.makedirs(tempfiles_dir, exist_ok=True)
    path = os.path.join(tempfiles_dir, CONFIG_FILENAME)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)
    return path


def load_config(directory=None):
    """Config'i JSON dosyasından yükler. directory verilmezse tempfiles/ kullanılır."""
    if directory is None:
        directory = _get_tempfiles_dir()
    path = os.path.join(directory, CONFIG_FILENAME)
    if not os.path.exists(path):
        return False
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    for key, val in data.items():
        if key in config:
            config[key] = val
    # Backward compat
    if config.get("PROGRAM"):
        config["PROJECT"] = config["PROGRAM"]
    return True
