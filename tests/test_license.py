"""Tests for license validation module.

Tests use mocking to avoid hitting the real Lemon Squeezy API.
"""

import json
import time
from pathlib import Path
from unittest.mock import patch, MagicMock

import pytest

from xlsx_fixer.license import (
    CACHE_FILE,
    CACHE_TTL,
    GRACE_PERIOD,
    _instance_name,
    _key_hash,
    _read_cache,
    _write_cache,
    read_key_file,
    validate_license,
)


@pytest.fixture(autouse=True)
def isolate_cache(tmp_path, monkeypatch):
    """Redirect cache to tmp_path so tests don't pollute real config."""
    cache_dir = tmp_path / ".xlsx-fixer"
    cache_file = cache_dir / "license.json"
    monkeypatch.setattr("xlsx_fixer.license.CACHE_DIR", cache_dir)
    monkeypatch.setattr("xlsx_fixer.license.CACHE_FILE", cache_file)
    return cache_file


class TestInstanceName:
    def test_returns_string(self):
        name = _instance_name()
        assert isinstance(name, str)
        assert len(name) == 16

    def test_stable(self):
        assert _instance_name() == _instance_name()


class TestKeyHash:
    def test_returns_string(self):
        h = _key_hash("test-key-123")
        assert isinstance(h, str)
        assert len(h) == 16

    def test_different_keys_different_hashes(self):
        assert _key_hash("key-a") != _key_hash("key-b")


class TestCache:
    def test_read_empty(self, isolate_cache):
        assert _read_cache() is None

    def test_write_and_read(self, isolate_cache):
        result = {"valid": True, "meta": {"variant_name": "Commercial"}}
        _write_cache(result, "abc123")
        cache = _read_cache()
        assert cache is not None
        assert cache["valid"] is True
        assert cache["key_hash"] == "abc123"
        assert cache["variant"] == "Commercial"
        assert "validated_at" in cache

    def test_corrupt_cache_returns_none(self, isolate_cache):
        isolate_cache.parent.mkdir(parents=True, exist_ok=True)
        isolate_cache.write_text("not json")
        assert _read_cache() is None


class TestReadKeyFile:
    def test_reads_first_line(self, tmp_path):
        f = tmp_path / "key"
        f.write_text("MY-LICENSE-KEY-123\n")
        assert read_key_file(f) == "MY-LICENSE-KEY-123"

    def test_skips_comments(self, tmp_path):
        f = tmp_path / "key"
        f.write_text("# comment\n\nACTUAL-KEY\n")
        assert read_key_file(f) == "ACTUAL-KEY"

    def test_strips_whitespace(self, tmp_path):
        f = tmp_path / "key"
        f.write_text("  KEY-WITH-SPACES  \n")
        assert read_key_file(f) == "KEY-WITH-SPACES"

    def test_missing_file(self, tmp_path):
        assert read_key_file(tmp_path / "nope") is None

    def test_empty_file(self, tmp_path):
        f = tmp_path / "key"
        f.write_text("")
        assert read_key_file(f) is None


class TestValidateLicense:
    def _mock_ls_response(self, valid=True, variant="Commercial", error=None, activated=False):
        """Build a mock LS API response."""
        resp = {
            "valid": valid,
            "error": error,
            "meta": {"variant_name": variant, "product_name": "xlsx-fixer Pro"},
        }
        if activated:
            resp["activated"] = True
        return resp

    @patch("xlsx_fixer.license._ls_request")
    def test_valid_key_online(self, mock_req, isolate_cache):
        mock_req.return_value = self._mock_ls_response(valid=True)
        valid, msg = validate_license("good-key")
        assert valid is True
        assert "Licensed" in msg
        # Should have written cache
        assert isolate_cache.exists()

    @patch("xlsx_fixer.license._ls_request")
    def test_invalid_key_online(self, mock_req, isolate_cache):
        mock_req.return_value = self._mock_ls_response(valid=False, error="Invalid key")
        valid, msg = validate_license("bad-key")
        assert valid is False
        assert "Invalid key" in msg

    @patch("xlsx_fixer.license._ls_request")
    def test_fresh_cache_skips_network(self, mock_req, isolate_cache):
        # Pre-populate cache
        result = self._mock_ls_response(valid=True)
        kh = _key_hash("cached-key")
        _write_cache(result, kh)

        valid, msg = validate_license("cached-key")
        assert valid is True
        # Should NOT have called the network
        mock_req.assert_not_called()

    @patch("xlsx_fixer.license._ls_request")
    def test_stale_cache_revalidates(self, mock_req, isolate_cache):
        # Pre-populate stale cache (8 days old)
        result = self._mock_ls_response(valid=True)
        kh = _key_hash("stale-key")
        _write_cache(result, kh)
        # Backdate the cache
        cache = json.loads(isolate_cache.read_text())
        cache["validated_at"] = time.time() - (8 * 86400)
        isolate_cache.write_text(json.dumps(cache))

        mock_req.return_value = self._mock_ls_response(valid=True, variant="Enterprise")
        valid, msg = validate_license("stale-key")
        assert valid is True
        assert "Enterprise" in msg
        mock_req.assert_called_once()

    @patch("xlsx_fixer.license._ls_request")
    def test_stale_cache_grace_period_on_network_failure(self, mock_req, isolate_cache):
        # Pre-populate cache (8 days old — stale but within 30-day grace)
        result = self._mock_ls_response(valid=True, variant="Pro")
        kh = _key_hash("grace-key")
        _write_cache(result, kh)
        cache = json.loads(isolate_cache.read_text())
        cache["validated_at"] = time.time() - (8 * 86400)
        isolate_cache.write_text(json.dumps(cache))

        mock_req.side_effect = OSError("no network")
        valid, msg = validate_license("grace-key")
        assert valid is True
        assert "offline grace" in msg

    @patch("xlsx_fixer.license._ls_request")
    def test_expired_cache_no_grace(self, mock_req, isolate_cache):
        # Pre-populate cache (35 days old — past grace period)
        result = self._mock_ls_response(valid=True)
        kh = _key_hash("expired-key")
        _write_cache(result, kh)
        cache = json.loads(isolate_cache.read_text())
        cache["validated_at"] = time.time() - (35 * 86400)
        isolate_cache.write_text(json.dumps(cache))

        mock_req.side_effect = OSError("no network")
        valid, msg = validate_license("expired-key")
        assert valid is False
        assert "offline" in msg.lower()

    @patch("xlsx_fixer.license._ls_request")
    def test_activation_flow(self, mock_req, isolate_cache):
        # First call: validate returns "not activated"
        # Second call: activate succeeds
        mock_req.side_effect = [
            {"valid": False, "error": "license_key is not activated."},
            self._mock_ls_response(valid=True, activated=True),
        ]
        valid, msg = validate_license("new-key")
        assert valid is True
        assert "Activated" in msg
        assert mock_req.call_count == 2

    @patch("xlsx_fixer.license._ls_request")
    def test_network_failure_no_cache(self, mock_req, isolate_cache):
        mock_req.side_effect = OSError("no network")
        valid, msg = validate_license("offline-key")
        assert valid is False
        assert "internet" in msg.lower()

    @patch("xlsx_fixer.license._ls_request")
    def test_different_key_invalidates_cache(self, mock_req, isolate_cache):
        # Cache for key A
        result = self._mock_ls_response(valid=True)
        _write_cache(result, _key_hash("key-a"))

        # Validate with key B — should go online
        mock_req.return_value = self._mock_ls_response(valid=True, variant="Team")
        valid, msg = validate_license("key-b")
        assert valid is True
        mock_req.assert_called_once()
