"""
KDL configuration storage.
Stores small app settings in JSON under the user's AppData folder.
DB passwords are encrypted locally using Windows DPAPI (win32crypt).
"""

import json
import os
import base64
import logging
from typing import Any, Dict
import copy

try:
    import win32crypt
except ImportError:
    win32crypt = None


APP_DIR_NAME = "KDL"
SETTINGS_FILE_NAME = "settings.json"
logger = logging.getLogger(__name__)


def _settings_path_candidates():
    """Ordered candidate paths for settings storage."""
    candidates = []

    appdata = os.getenv("APPDATA")
    if appdata:
        candidates.append(os.path.join(appdata, APP_DIR_NAME, SETTINGS_FILE_NAME))

    local_appdata = os.getenv("LOCALAPPDATA")
    if local_appdata:
        candidates.append(os.path.join(local_appdata, APP_DIR_NAME, SETTINGS_FILE_NAME))

    home = os.path.expanduser("~")
    candidates.append(os.path.join(home, ".kdl", SETTINGS_FILE_NAME))
    candidates.append(os.path.join(os.getcwd(), ".kdl", SETTINGS_FILE_NAME))

    # Keep order stable while removing duplicates.
    deduped = []
    seen = set()
    for p in candidates:
        norm = os.path.normcase(os.path.normpath(p))
        if norm in seen:
            continue
        seen.add(norm)
        deduped.append(p)
    return deduped


def _settings_read_path() -> str:
    """Return newest existing settings path, else preferred primary path."""
    existing = []
    for path in _settings_path_candidates():
        if os.path.exists(path):
            try:
                mtime = os.path.getmtime(path)
            except Exception:
                mtime = 0.0
            existing.append((mtime, path))
    if existing:
        existing.sort(key=lambda item: item[0], reverse=True)
        return existing[0][1]
    return _settings_path_candidates()[0]


def _settings_write_path() -> str:
    """Return first writable path by trying candidates in priority order."""
    for path in _settings_path_candidates():
        base_dir = os.path.dirname(path)
        try:
            os.makedirs(base_dir, exist_ok=True)
            with open(path, "a", encoding="utf-8"):
                pass
            return path
        except Exception:
            continue
    # Final fallback; caller handles write errors.
    return _settings_path_candidates()[-1]


def _encrypt_password(pwd: str) -> str:
    if not pwd:
        return pwd
    if not win32crypt:
        # Never persist raw passwords when DPAPI is unavailable.
        logger.warning("win32crypt unavailable — password will not be saved (DPAPI required)")
        return ""
    try:
        encrypted_bytes = win32crypt.CryptProtectData(
            pwd.encode('utf-8'), 'KDL_DB_PWD', None, None, None, 0
        )
        # Prefix with DPAPI_ to identify encrypted strings
        return "DPAPI_" + base64.b64encode(encrypted_bytes).decode('utf-8')
    except Exception:
        # Fail closed: do not store plaintext on encryption failure.
        logger.warning("DPAPI encryption failed — password will not be saved", exc_info=True)
        return ""


def _decrypt_password(enc_pwd: str) -> str:
    if not enc_pwd or not win32crypt:
        return enc_pwd
    if str(enc_pwd).startswith("DPAPI_"):
        try:
            b64_str = enc_pwd[6:]
            _, decrypted_bytes = win32crypt.CryptUnprotectData(
                base64.b64decode(b64_str.encode('utf-8')), None, None, None, 0
            )
            return decrypted_bytes.decode('utf-8')
        except Exception:
            # Preserve original value so it is not silently wiped on next save.
            logger.warning(
                "Failed to decrypt DPAPI password token from settings; preserving stored token.",
                exc_info=True,
            )
            return enc_pwd
    # Fallback to returning the plain string if it wasn't encrypted
    return enc_pwd


def load_settings() -> Dict[str, Any]:
    path = _settings_read_path()
    if not os.path.exists(path):
        return {}

    try:
        with open(path, "r", encoding="utf-8") as fh:
            data = json.load(fh)
        if not isinstance(data, dict):
            return {}
            
        # Decrypt passwords
        profiles = data.get("database", {}).get("profiles", [])
        for p in profiles:
            if "password" in p and p["password"]:
                p["password"] = _decrypt_password(p["password"])
                
        return data
    except Exception as e:
        print(f"Error loading settings: {e}")
        return {}


def save_settings(data: Dict[str, Any]) -> bool:
    path = _settings_write_path()
    try:
        # Deep copy to avoid mutating the in-memory state while encrypting
        save_data = copy.deepcopy(data)
        
        # Encrypt passwords
        profiles = save_data.get("database", {}).get("profiles", [])
        for p in profiles:
            if "password" in p and p["password"] and not str(p["password"]).startswith("DPAPI_"):
                p["password"] = _encrypt_password(p["password"])
                
        with open(path, "w", encoding="utf-8") as fh:
            json.dump(save_data, fh, indent=2, ensure_ascii=True)
        return True
    except Exception as e:
        print(f"Error saving settings: {e}")
        return False
