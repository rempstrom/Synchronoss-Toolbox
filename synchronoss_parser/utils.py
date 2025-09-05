"""Utility helpers for Synchronoss scripts."""

from __future__ import annotations


def normalize_phone_number(value: str) -> str:
    """Return ``value`` stripped down to just digits.

    Parameters
    ----------
    value: str
        A phone number that may contain spaces, ``+`` signs, dashes or other
        punctuation.

    Returns
    -------
    str
        Only the numeric digits found in ``value``.
    """

    if not value:
        return ""
    return "".join(ch for ch in str(value) if ch.isdigit())
