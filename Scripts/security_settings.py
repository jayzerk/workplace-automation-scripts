"""Local security settings and explicit Windows startup registration."""

from __future__ import annotations

import json
import os
import re
from copy import deepcopy
from pathlib import Path
from urllib.parse import urlsplit

try:
    import winreg
except ImportError:  # pragma: no cover - the packaged application targets Windows
    winreg = None


DEFAULT_ALLOWED_DOMAINS = ["docs.google.com"]
STARTUP_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
STARTUP_VALUE_NAME = "ExcelToolsQueueBoard"
DOMAIN_PATTERN = re.compile(
    r"^(?:\*\.)?(?:[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.)+"
    r"[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?$",
    re.IGNORECASE,
)
SAFE_LOCAL_EXTENSIONS = {
    ".bmp",
    ".csv",
    ".docx",
    ".gif",
    ".jpeg",
    ".jpg",
    ".json",
    ".m4a",
    ".md",
    ".mov",
    ".mp3",
    ".mp4",
    ".pdf",
    ".png",
    ".pptx",
    ".tif",
    ".tiff",
    ".tsv",
    ".txt",
    ".wav",
    ".webp",
    ".xlsx",
}


def normalize_allowed_domain(value: str) -> str:
    """Validate and normalize one exact or wildcard domain entry."""
    domain = value.strip().casefold().rstrip(".")
    if not domain:
        raise ValueError("Allowed domains cannot contain blank entries.")
    try:
        ascii_domain = domain.encode("idna").decode("ascii")
    except UnicodeError as error:
        raise ValueError(f"Invalid allowed domain: {value}") from error
    if not DOMAIN_PATTERN.fullmatch(ascii_domain):
        raise ValueError(
            f"Invalid allowed domain '{value}'. Enter a domain only, such as "
            "example.com or *.example.com."
        )
    return ascii_domain


def normalize_allowed_domains(values: list[str]) -> list[str]:
    normalized: list[str] = []
    for value in values:
        domain = normalize_allowed_domain(value)
        if domain not in normalized:
            normalized.append(domain)
    return normalized


def _domain_is_allowed(hostname: str, allowed_domains: list[str]) -> bool:
    host = hostname.casefold().rstrip(".")
    for entry in allowed_domains:
        if entry.startswith("*."):
            suffix = entry[2:]
            if host.endswith("." + suffix):
                return True
        elif host == entry:
            return True
    return False


def validate_link(
    value: str,
    allowed_domains: list[str],
    *,
    enforce_allowlist: bool = True,
) -> str:
    """Validate a safe local document path or allowlisted HTTPS URL."""
    link = value.strip()
    if not link:
        return ""

    if link.startswith(("\\\\", "//")):
        raise ValueError("Network-share and Windows device paths are blocked.")

    path = Path(link)
    if path.is_absolute():
        if path.suffix.casefold() not in SAFE_LOCAL_EXTENSIONS:
            raise ValueError(
                "That local file type is blocked. Executables, scripts, shortcuts, "
                "archives, and macro-enabled Office files are not allowed."
            )
        return str(path)

    parsed = urlsplit(link)
    if parsed.scheme.casefold() != "https" or not parsed.hostname:
        raise ValueError("Web links must use HTTPS, or select an absolute local file.")
    if parsed.username or parsed.password:
        raise ValueError("Web links containing embedded credentials are blocked.")
    try:
        port = parsed.port
    except ValueError as error:
        raise ValueError("The web link contains an invalid port.") from error
    if port not in (None, 443):
        raise ValueError("Only the standard HTTPS port is allowed.")

    hostname = parsed.hostname.encode("idna").decode("ascii").casefold()
    if enforce_allowlist and not _domain_is_allowed(hostname, allowed_domains):
        raise ValueError(
            f"The domain '{hostname}' is not allowed. Add it through Allowed Domains."
        )
    return link


class AppSettingsStore:
    """Persist application security settings with atomic file replacement."""

    def __init__(self, settings_file: str | Path):
        self.settings_file = Path(settings_file)
        self.allowed_domains = list(DEFAULT_ALLOWED_DOMAINS)

    def load(self) -> None:
        if not self.settings_file.exists():
            return
        try:
            payload = json.loads(self.settings_file.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as error:
            raise ValueError(f"Could not read security settings: {error}") from error
        raw_domains = payload.get("allowed_domains") if isinstance(payload, dict) else None
        if not isinstance(raw_domains, list) or not all(
            isinstance(value, str) for value in raw_domains
        ):
            raise ValueError("Security settings contain an invalid domain list.")
        self.allowed_domains = normalize_allowed_domains(raw_domains)

    def save(self) -> None:
        self.settings_file.parent.mkdir(parents=True, exist_ok=True)
        temporary_file = self.settings_file.with_suffix(
            self.settings_file.suffix + ".tmp"
        )
        payload = {
            "version": 1,
            "allowed_domains": self.allowed_domains,
        }
        try:
            temporary_file.write_text(
                json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
            )
            os.replace(temporary_file, self.settings_file)
        except OSError:
            try:
                temporary_file.unlink(missing_ok=True)
            except OSError:
                pass
            raise

    def snapshot(self) -> list[str]:
        return deepcopy(self.allowed_domains)

    def restore(self, snapshot: list[str]) -> None:
        self.allowed_domains = list(snapshot)


def get_startup_enabled(expected_command: str) -> bool:
    """Return whether this app's exact command is registered at user sign-in."""
    if winreg is None:
        return False
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, STARTUP_KEY) as key:
            value, value_type = winreg.QueryValueEx(key, STARTUP_VALUE_NAME)
    except OSError:
        return False
    return value_type == winreg.REG_SZ and value == expected_command


def set_startup_enabled(command: str, enabled: bool) -> None:
    """Create or remove only this app's visible per-user startup entry."""
    if winreg is None:
        raise OSError("Windows startup registration is unavailable on this system.")
    if enabled:
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, STARTUP_KEY) as key:
            winreg.SetValueEx(key, STARTUP_VALUE_NAME, 0, winreg.REG_SZ, command)
        return

    try:
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            STARTUP_KEY,
            0,
            winreg.KEY_QUERY_VALUE | winreg.KEY_SET_VALUE,
        ) as key:
            existing_command, value_type = winreg.QueryValueEx(
                key,
                STARTUP_VALUE_NAME,
            )
            if value_type != winreg.REG_SZ or existing_command != command:
                raise OSError(
                    "The startup entry no longer points to this exact app command, "
                    "so it was not removed."
                )
            winreg.DeleteValue(key, STARTUP_VALUE_NAME)
    except FileNotFoundError:
        pass
