# -*- coding: utf-8 -*-
"""job-hunt-il shared config loader.

Every script reads user-specific data (paths, profile, search titles) from
config.json instead of hardcoding it. Resolution order:
  1. JOB_HUNT_IL_CONFIG environment variable (full path to a config.json)
  2. config.json in the skill root (one level up from scripts/)

Created by `setup.py` from config.example.json. Nothing personal ships in the
repo; config.json is gitignored.
"""
import json
import os
import sys

_SKILL_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_DEFAULT_CONFIG = os.path.join(_SKILL_ROOT, "config.json")


def config_path() -> str:
    return os.environ.get("JOB_HUNT_IL_CONFIG", _DEFAULT_CONFIG)


def load() -> dict:
    path = config_path()
    if not os.path.exists(path):
        sys.exit(
            f"ERROR: config not found at {path}.\n"
            "Run the setup first:  python scripts/setup.py"
        )
    with open(path, encoding="utf-8") as f:
        cfg = json.load(f)
    # Resolve workdir-relative paths. NOTE: "tracker" is deliberately excluded
    # here — since the Google Sheets switch it holds a spreadsheet ID (or ""
    # until setup.py creates one), not a filesystem path.
    wd = cfg.get("workdir", "")
    for key in ("profile_md", "positioning_md", "hebrew_cv_template",
                "output_dir", "google_credentials", "google_token"):
        val = cfg.get(key, "")
        if val and not os.path.isabs(val):
            cfg[key] = os.path.normpath(os.path.join(wd, val))
    return cfg


def api_key(cfg: dict) -> str:
    """Anthropic API key from the env var named in config (never stored in files)."""
    env_name = cfg.get("api_key_env", "ANTHROPIC_API_KEY")
    key = os.environ.get(env_name, "")
    if not key and os.name == "nt":
        # Windows: freshly-set user env vars aren't visible to already-open
        # terminals; fall back to reading the registry directly.
        try:
            import winreg
            reg = winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Environment")
            key, _ = winreg.QueryValueEx(reg, env_name)
            winreg.CloseKey(reg)
        except Exception:
            pass
    return key


def profile_text(cfg: dict) -> str:
    """The user's career profile (facts) — injected into every scoring and
    tailoring prompt."""
    path = cfg["profile_md"]
    if not os.path.exists(path):
        sys.exit(f"ERROR: profile not found at {path}. Run setup, or create it "
                 f"from templates/profile_template.md")
    with open(path, encoding="utf-8") as f:
        return f.read()


def positioning_text(cfg: dict) -> str:
    """How to FRAME the facts (optional but strongly recommended)."""
    path = cfg.get("positioning_md", "")
    if path and os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            return f.read()
    return ""
