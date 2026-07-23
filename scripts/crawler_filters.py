#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Shared filters for the crawlers (Jobify360, LinkedIn).

One place so the crawlers can never drift apart. Each crawler imports
`should_skip` and calls it right before writing a candidate row.

Rules:
1. SCORE FLOOR: drop anything under config `min_score`.
2. LOCATION EXCLUDES: drop jobs whose location matches any entry in config
   `exclude_locations` (user's list — e.g. cities beyond commute range).
   Latin entries match on word boundaries (so "omer" can't fire inside
   "customer"); Hebrew entries match as substrings (Hebrew prefixes like
   ב/ה/ל attach to the word, so substring is what we want).
   When the source gives no location field, we fall back to scanning the
   title + description text — less precise, but the only signal available.
"""

import re
import jh_config

_cfg = jh_config.load()
MIN_SCORE = _cfg.get("min_score", 65)
_EXCLUDES = [str(p).lower() for p in _cfg.get("exclude_locations", []) if str(p).strip()]

_LATIN = [p for p in _EXCLUDES if p.isascii()]
_HEBREW = [p for p in _EXCLUDES if not p.isascii()]
_LATIN_RE = re.compile(r"\b(" + "|".join(re.escape(p) for p in _LATIN) + r")\b") if _LATIN else None


def _matches_exclude(text):
    text = (text or "").lower()
    if _LATIN_RE and _LATIN_RE.search(text):
        return True
    return any(p in text for p in _HEBREW)


def is_excluded_location(location="", title="", jd=""):
    """True if the job is in a location the user excluded."""
    if not _EXCLUDES:
        return False
    if (location or "").strip():
        # Trust the location field; don't text-scan (avoids false positives).
        return _matches_exclude(location)
    # No location captured — scan title + JD as the only available signal.
    return _matches_exclude(f"{title} {jd}")


def should_skip(score, location="", title="", jd="", company="", **_legacy):
    # **_legacy absorbs kwargs from the original pipeline (is_international,
    # require_english, ...) so the ported crawlers keep working unchanged.
    """Decide whether to skip writing a candidate row.

    Returns (skip: bool, reason: str). reason is "" when skip is False.
    """
    try:
        s = int(score)
    except (TypeError, ValueError):
        s = 0

    if s < MIN_SCORE:
        return True, f"score {s} < {MIN_SCORE}"
    if is_excluded_location(location, title, jd):
        return True, "excluded location"
    return False, ""
