"""Lemon Squeezy license validation with activation caching.

Flow:
    1. Check local cache (~/.xlsx-fixer/license.json)
    2. Cache < 7 days → use cached (skip network)
    3. Cache 7-30 days → try validate online, grace period if network fails
    4. Cache > 30 days or no cache → must validate online
    5. No cache at all → activate first, then cache

Zero external deps — uses urllib.request (stdlib).
"""

from __future__ import annotations

import hashlib
import json
import platform
import time
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

LS_VALIDATE_URL = "https://api.lemonsqueezy.com/v1/licenses/validate"
LS_ACTIVATE_URL = "https://api.lemonsqueezy.com/v1/licenses/activate"

CACHE_DIR = Path.home() / ".xlsx-fixer"
CACHE_FILE = CACHE_DIR / "license.json"

CACHE_TTL = 7 * 86400       # 7 days — revalidate after this
GRACE_PERIOD = 30 * 86400   # 30 days — allow offline if cache exists
REQUEST_TIMEOUT = 10         # seconds


def _instance_name() -> str:
    """Stable per-machine identifier."""
    node = platform.node() or "unknown"
    return hashlib.sha256(node.encode()).hexdigest()[:16]


def _read_cache() -> dict | None:
    """Read cached license validation result."""
    if not CACHE_FILE.exists():
        return None
    try:
        data = json.loads(CACHE_FILE.read_text())
        if "validated_at" in data and "valid" in data:
            return data
    except (json.JSONDecodeError, KeyError):
        pass
    return None


def _write_cache(result: dict, key_hash: str) -> None:
    """Write license validation result to cache."""
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    cache = {
        "valid": result.get("valid", False),
        "validated_at": time.time(),
        "key_hash": key_hash,
        "variant": result.get("meta", {}).get("variant_name", "unknown"),
        "product": result.get("meta", {}).get("product_name", "unknown"),
    }
    CACHE_FILE.write_text(json.dumps(cache, indent=2))


def _ls_request(url: str, license_key: str) -> dict:
    """POST to Lemon Squeezy API. Returns parsed JSON."""
    body = urllib.parse.urlencode({
        "license_key": license_key,
        "instance_name": _instance_name(),
    }).encode()
    req = urllib.request.Request(
        url,
        data=body,
        headers={"Accept": "application/json"},
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=REQUEST_TIMEOUT) as resp:
        return json.loads(resp.read())


def _key_hash(license_key: str) -> str:
    """Hash the key so we can detect key changes without storing plaintext."""
    return hashlib.sha256(license_key.encode()).hexdigest()[:16]


def validate_license(license_key: str) -> tuple[bool, str]:
    """Validate a Lemon Squeezy license key.

    Returns:
        (is_valid, message) tuple.
    """
    kh = _key_hash(license_key)
    cache = _read_cache()

    # If cache exists for THIS key, check age
    if cache and cache.get("key_hash") == kh:
        age = time.time() - cache["validated_at"]

        # Fresh cache — skip network
        if age < CACHE_TTL:
            if cache["valid"]:
                return True, f"Licensed ({cache.get('variant', 'Pro')})"
            return False, "License invalid (cached)"

        # Stale cache — try online, allow grace period on failure
        try:
            result = _ls_request(LS_VALIDATE_URL, license_key)
            _write_cache(result, kh)
            if result.get("valid"):
                return True, f"Licensed ({result.get('meta', {}).get('variant_name', 'Pro')})"
            return False, result.get("error", "License invalid")
        except (urllib.error.URLError, OSError, json.JSONDecodeError):
            if age < GRACE_PERIOD:
                return True, f"Licensed ({cache.get('variant', 'Pro')}) [offline grace]"
            return False, "License expired — unable to revalidate (offline >30 days)"

    # No cache or different key — must go online
    # Try validate first (works if already activated)
    try:
        result = _ls_request(LS_VALIDATE_URL, license_key)
        if result.get("valid"):
            _write_cache(result, kh)
            return True, f"Licensed ({result.get('meta', {}).get('variant_name', 'Pro')})"

        # Not yet activated — try activation
        if result.get("error") == "license_key is not activated.":
            result = _ls_request(LS_ACTIVATE_URL, license_key)
            if result.get("activated"):
                _write_cache(result, kh)
                variant = result.get("meta", {}).get("variant_name", "Pro")
                return True, f"Activated ({variant})"
            return False, result.get("error", "Activation failed")

        return False, result.get("error", "License invalid")

    except urllib.error.HTTPError as e:
        if e.code == 404:
            return False, "Invalid license key"
        return False, f"Validation error (HTTP {e.code})"
    except (urllib.error.URLError, OSError):
        return False, "Unable to validate license — check your internet connection"
    except json.JSONDecodeError:
        return False, "Invalid response from license server"


def read_key_file(path: Path) -> str | None:
    """Read a license key from a file (first non-empty line)."""
    try:
        for line in path.read_text().splitlines():
            stripped = line.strip()
            if stripped and not stripped.startswith("#"):
                return stripped
    except (FileNotFoundError, PermissionError):
        pass
    return None
