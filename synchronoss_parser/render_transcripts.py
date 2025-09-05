#!/usr/bin/env python3
"""
Chat Transcript Renderer

Reads message CSVs from a folder structure like:

messages/
  20240120.csv
  20241121.csv
  ...
  attachments/
    mms/
      in/
        2024-01-20/
          <files>
      out/
        2024-01-20/
          <files>
    rcs/
      in/
        2024-11-21/
          <files>
      out/
        2024-11-21/
          <files>

Each CSV must have columns: Date, Type, Direction, Attachments, Body, Sender, Recipients, Message ID

This script generates one HTML transcript per CSV (chat-bubble style), plus an index.html.
Attachments are embedded inline when possible (images/audio/video) and otherwise linked.
The attachment lookup path is:
  messages/attachments/{Type}/{Direction}/{YYYY-MM-DD}/{AttachmentFileName}
where YYYY-MM-DD is derived from the CSV filename (e.g., 20241121.csv -> 2024-11-21).

Usage:
  render-transcripts --in messages --out transcripts [--contacts-xlsx contacts.xlsx]

Notes:
- HTML keeps relative links to your existing attachments (no copying).
- Dates are sorted chronologically using best-effort parsing; the raw Date string is also shown.
- "out" messages are right-aligned (sent), "in" are left-aligned (received).
- Safe to run multiple times; outputs are overwritten.
"""

import argparse
import csv
import html
import json
import os
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple

import pandas as pd


# ------------------------- Config & Utilities -------------------------

IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".gif", ".webp", ".bmp"}
VIDEO_EXTS = {".mp4", ".webm", ".ogg", ".mov", ".m4v"}
AUDIO_EXTS = {".mp3", ".wav", ".m4a", ".aac", ".flac", ".ogg"}
INLINE_TEXT_EXTS = {".vcard", ".vcf"}  # small text-like files we might show inline

CSV_DATE_FROM_FILENAME_FMT = "%Y%m%d"
ATTACHMENT_FOLDER_DATE_FMT = "%Y-%m-%d"

CSS_STYLES = """
:root {
  --bg: #1e293b;           /* slate-800 */
  --panel: #1f2937;        /* gray-800 */
  --text: #f3f4f6;         /* gray-100 */
  --muted: #cbd5e1;        /* slate-300 */
  --sent: #4ade80;         /* green-400 */
  --sent-contrast: #064e3b;/* green-900 */
  --recv: #374151;         /* gray-700 */
  --bubble-radius: 16px;
  --max-width: 980px;
}
* { box-sizing: border-box; }
body {
  margin: 0; padding: 0; background: var(--bg); color: var(--text);
  font: 15px/1.5 system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif;
}
.header {
  position: sticky; top: 0; z-index: 10; background: linear-gradient(180deg, rgba(17,24,39,0.95), rgba(17,24,39,0.7));
  backdrop-filter: blur(6px); border-bottom: 1px solid #1f2937; padding: 10px 16px;
}
.container { max-width: var(--max-width); margin: 0 auto; padding: 12px 16px 80px; }
.thread-meta { color: var(--muted); font-size: 12px; margin-top: 2px; }
.search-bar {
  margin-top: 8px;
}
.search-input {
  width: 100%;
  max-width: 400px;
  padding: 6px 8px;
  border-radius: 8px;
  border: 1px solid #374151;
  background: var(--panel);
  color: var(--text);
}

.message {
  display: flex; margin: 8px 0; gap: 10px; align-items: flex-end;
}
.bubble {
  max-width: 74%; padding: 10px 12px; border-radius: var(--bubble-radius);
  word-wrap: break-word; overflow-wrap: anywhere; box-shadow: 0 2px 10px rgba(0,0,0,0.2);
}
.sent { margin-left: auto; justify-content: flex-end; }
.sent .bubble { background: var(--sent); color: var(--sent-contrast); }
.sent .meta  { text-align: right; }
.sent .sender { text-align: right; }
.received .bubble { background: var(--recv); }

.sender { font-size: 16px; font-weight: 700; margin-bottom: 4px; }
.meta { font-size: 11px; color: var(--muted); margin-top: 6px; }
.body-text { white-space: pre-wrap; }
.missing { color: var(--muted); font-style: italic; }
.attachments { margin-top: 8px; display: grid; gap: 8px; }
.attachment img { max-width: 100%; border-radius: 12px; display: block; }
.attachment video, .attachment audio { width: 100%; outline: none; }
.attachment a { color: #93c5fd; text-decoration: none; word-break: break-all; }
.attachment a:hover { text-decoration: underline; }

.day-divider { text-align: center; margin: 18px 0; color: var(--muted); font-size: 12px; }
.footer { position: fixed; bottom: 0; left: 0; right: 0; padding: 8px 16px; background: linear-gradient(0deg, rgba(17,24,39,0.95), rgba(17,24,39,0.6)); border-top: 1px solid #1f2937; }
.footer a { color: #93c5fd; text-decoration: none; }
.footer a:hover { text-decoration: underline; }
.code { font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace; font-size: 12px; }
"""

INDEX_CSS = """
body { background:#1e293b; color:#f3f4f6; font: 15px/1.5 system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif; }
.container { max-width: 900px; margin: 0 auto; padding: 24px 16px; }
h1 { margin: 0 0 8px; font-size: 24px; }
.subtitle { color:#cbd5e1; margin-bottom: 16px; font-size: 13px; }
.search-bar { margin-bottom: 16px; }
.search-input { width:100%; max-width:400px; padding:6px 8px; border-radius:8px; border:1px solid #374151; background:#1f2937; color:#f3f4f6; }
.list { display: grid; gap: 10px; }
.item { background:#1f2937; border:1px solid #374151; border-radius: 12px; padding: 12px; }
.item a { color:#93c5fd; text-decoration:none; font-weight:600; }
.item a:hover { text-decoration: underline; }
.meta { color:#cbd5e1; font-size: 12px; margin-top: 4px; }
"""

@dataclass
class Message:
    date_raw: str
    date_dt: Optional[datetime]
    msg_type: str
    direction: str
    attachments: List[str]
    body: str
    sender: str
    recipients: str
    message_id: str
    attachment_day: Optional[str] = None


def parse_csv_date(value: str) -> Optional[datetime]:
    if not value:
        return None
    s = value.strip()
    # Try ISO-8601 with Z
    try:
        if s.endswith("Z"):
            return datetime.fromisoformat(s.replace("Z", "+00:00"))
    except Exception:
        pass
    # Try plain fromisoformat
    try:
        return datetime.fromisoformat(s)
    except Exception:
        pass
    # Try common formats
    fmts = [
        "%Y-%m-%d %H:%M:%S",
        "%m/%d/%Y %H:%M:%S",
        "%m/%d/%Y %I:%M:%S %p",
        "%Y-%m-%dT%H:%M:%S",
        "%Y-%m-%dT%H:%M:%S.%f%z",
        "%Y-%m-%dT%H:%M:%S%z",
    ]
    for f in fmts:
        try:
            return datetime.strptime(s, f)
        except Exception:
            continue
    # Fallback: epoch seconds or ms?
    try:
        n = int(s)
        if n > 10_000_000_000:  # likely ms
            return datetime.fromtimestamp(n/1000, tz=timezone.utc)
        return datetime.fromtimestamp(n, tz=timezone.utc)
    except Exception:
        return None


def split_attachments(field: str) -> List[str]:
    if not field:
        return []
    s = field.strip()
    # Try JSON list first
    if (s.startswith("[") and s.endswith("]")) or (s.startswith("{") and s.endswith("}")):
        try:
            parsed = json.loads(s)
            if isinstance(parsed, list):
                return [str(x).strip() for x in parsed if str(x).strip()]
            if isinstance(parsed, dict) and "files" in parsed:
                return [str(x).strip() for x in parsed["files"] if str(x).strip()]
        except Exception:
            pass
    # Fallback: split on common delimiters (semicolon strongest for your data)
    parts = []
    for chunk in s.replace("|", ";").replace(",", ";").split(";"):
        c = chunk.strip()
        if c:
            parts.append(c)
    return parts


def derive_attachment_day_from_csv_name(csv_path: Path) -> Optional[str]:
    """Convert 20241121.csv -> 2024-11-21"""
    try:
        stem = csv_path.stem  # e.g., 20241121
        dt = datetime.strptime(stem, CSV_DATE_FROM_FILENAME_FMT)
        return dt.strftime(ATTACHMENT_FOLDER_DATE_FMT)
    except Exception:
        return None


def classify_ext(path: Path) -> str:
    ext = path.suffix.lower()
    if ext in IMAGE_EXTS:
        return "image"
    if ext in VIDEO_EXTS:
        return "video"
    if ext in AUDIO_EXTS:
        return "audio"
    if ext in INLINE_TEXT_EXTS:
        return "inline_text"
    return "other"


def safe_text(s: Optional[str]) -> str:
    if s is None:
        return ""
    return html.escape(str(s))


def normalize_phone_number(number: str) -> str:
    """Return a canonical form for a phone number.

    All non-digit characters are stripped and a leading ``1`` (US/Canada
    country code) is removed when the result would otherwise be eleven
    digits. Examples::

        normalize_phone_number("+1 111-222-3333") -> "1112223333"
        normalize_phone_number("(111) 222-3333") -> "1112223333"
        normalize_phone_number("+12223334444") -> "2223334444"

    Supported input formats therefore include ``+12223334444``,
    ``111-222-3333``, ``(111) 222-3333``, ``+1 111-222-3333`` and
    ``1112223333``.
    """

    digits = "".join(ch for ch in str(number) if ch.isdigit())
    if len(digits) == 11 and digits.startswith("1"):
        digits = digits[1:]
    return digits


def build_contact_lookup(xlsx_path: Optional[str]) -> Callable[[str], str]:
    """Return a lookup function mapping phone numbers to contact names."""
    mapping: Dict[str, str] = {}
    if xlsx_path:
        try:
            df = pd.read_excel(xlsx_path)
            for _, row in df.iterrows():
                first = str(row.get("firstname") or "").strip()
                last = str(row.get("lastname") or "").strip()
                numbers = str(row.get("phone_numbers") or "").split(";")
                name = f"{first} {last}".strip()
                if not name:
                    continue
                for num in numbers:
                    digits = normalize_phone_number(num)
                    if digits:
                        mapping[digits] = name
        except Exception:
            pass

    def lookup(number: str) -> str:
        digits = normalize_phone_number(number)
        return mapping.get(digits, number)

    return lookup


def build_attachment_path(messages_root: Path, msg_type: str, direction: str, day_str: str, filename: str) -> Path:
    return messages_root / "attachments" / msg_type / direction / day_str / filename


def relpath_for_html(from_file: Path, to_target: Path) -> str:
    try:
        return os.path.relpath(to_target, start=from_file.parent).replace(os.sep, "/")
    except Exception:
        return str(to_target).replace(os.sep, "/")


def load_messages_from_csv(csv_file: Path, contact_lookup: Callable[[str], str] = lambda x: x) -> List[Message]:
    msgs: List[Message] = []
    day_folder = derive_attachment_day_from_csv_name(csv_file)
    with csv_file.open("r", encoding="utf-8", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            date_raw = (row.get("Date") or "").strip()
            date_dt = parse_csv_date(date_raw)
            msg_type = (row.get("Type") or "").strip().lower()
            direction = (row.get("Direction") or "").strip().lower()
            attachments_field = row.get("Attachments")
            attachments = split_attachments(attachments_field) if attachments_field else []
            body = row.get("Body") or ""
            sender = contact_lookup(row.get("Sender") or "")
            raw_recip = row.get("Recipients") or ""
            recip_parts = []
            for part in raw_recip.replace(",", ";").split(";"):
                p = part.strip()
                if p:
                    recip_parts.append(contact_lookup(p))
            recipients = "; ".join(recip_parts)
            message_id = row.get("Message ID") or ""
            msgs.append(
                Message(
                    date_raw,
                    date_dt,
                    msg_type,
                    direction,
                    attachments,
                    body,
                    sender,
                    recipients,
                    message_id,
                    day_folder,
                )
            )
    # Sort chronologically with stable fallback to raw string
    msgs.sort(key=lambda m: (m.date_dt or datetime.max.replace(tzinfo=timezone.utc), m.date_raw))
    return msgs


# ------------------------- Grouping -------------------------

def sanitize_participants(participants: Tuple[str, ...]) -> str:
    if not participants:
        return "chat"
    cleaned = ["".join(ch for ch in p if ch.isalnum()) for p in participants]
    cleaned = [c or "unknown" for c in cleaned]
    return "-".join(cleaned) or "chat"


def group_messages_by_chat(messages: List[Message], target: str) -> Dict[Tuple[str, ...], List[Message]]:
    groups: Dict[Tuple[str, ...], List[Message]] = {}
    for m in messages:
        if m.msg_type in {"sms", "mms", "rcs"}:
            if m.direction == "out" and not m.sender:
                m.sender = target
            if m.direction == "in" and not m.recipients:
                m.recipients = target
        recips: List[str] = []
        if m.recipients:
            for part in m.recipients.replace(",", ";").split(";"):
                p = part.strip()
                if p:
                    recips.append(p)
        participants = set(recips)
        if m.sender:
            participants.add(m.sender)
        if target:
            participants.add(target)
        key = tuple(sorted(participants))
        groups.setdefault(key, []).append(m)
    for lst in groups.values():
        lst.sort(key=lambda mm: (mm.date_dt or datetime.max.replace(tzinfo=timezone.utc), mm.date_raw))
    return groups


# ------------------------- HTML Rendering -------------------------

def render_thread_html(
    messages_root: Path,
    out_file: Path,
    msgs: List[Message],
    participants: List[str],
    target_number: str,
    contact_lookup: Callable[[str], str] = lambda x: x,
) -> Tuple[int, int]:
    total = len(msgs)
    with_attachments = 0

    disp_participants = [contact_lookup(p) for p in participants]
    title = f"Chat – {', '.join(disp_participants)}"

    parts: List[str] = []
    parts.append("<!DOCTYPE html>")
    parts.append("<html lang=\"en\">")
    parts.append("<head>")
    parts.append("<meta charset=\"utf-8\">")
    parts.append("<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">")
    parts.append(f"<title>{html.escape(title)}</title>")
    parts.append("<style>" + CSS_STYLES + "</style>")
    parts.append("</head>")
    parts.append("<body>")

    parts.append("<div class=\"header\">")
    parts.append("  <div class=\"container\">")
    parts.append(f"    <div><strong>{html.escape(title)}</strong></div>")
    target_disp = contact_lookup(target_number)
    meta_line = f"Target: {html.escape(target_disp)}<br>Participants: {html.escape(', '.join(disp_participants))}"
    parts.append(f"    <div class=\"thread-meta\">{meta_line}</div>")
    parts.append("    <div class=\"search-bar\"><input id=\"search\" class=\"search-input\" placeholder=\"Search messages\"></div>")
    parts.append("  </div>")
    parts.append("</div>")

    parts.append("<div class=\"container\">")

    current_day: Optional[str] = None

    for m in msgs:
        # Day divider (based on local date of parsed datetime if available, else raw)
        day_label = None
        if m.date_dt:
            day_label = m.date_dt.astimezone().strftime("%A, %B %d, %Y")
        elif m.date_raw:
            # Try to grab just the date part
            day_label = m.date_raw.split("T")[0]
        if day_label and day_label != current_day:
            current_day = day_label
            parts.append(f"<div class=\"day-divider\">{html.escape(current_day)}</div>")

        side_class = "sent" if m.direction == "out" else "received"
        parts.append(f"<div class=\"message {side_class}\">")
        parts.append("  <div class=\"bubble\">")

        sender = safe_text(contact_lookup(m.sender))
        if sender:
            parts.append(f"    <div class=\"sender\">{sender}</div>")

        body_html = safe_text(m.body)
        if body_html:
            parts.append(f"    <div class=\"body-text\">{body_html}</div>")
        else:
            msg_type = (m.msg_type or "").upper()
            placeholder = (
                f"NO DATA IN CSV FOR THIS {msg_type} MESSAGE - LOG ONLY" if msg_type else "NO DATA IN CSV FOR THIS MESSAGE - LOG ONLY"
            )
            parts.append(f"    <div class=\"body-text missing\">{placeholder}</div>")

        # Attachments
        attachment_snippets: List[str] = []
        if m.attachments:
            for fname in m.attachments:
                if not fname or fname.lower() in {"null", "null.txt", "none", "(null)", "aaaa"}:
                    continue
                if not m.attachment_day:
                    continue
                target = build_attachment_path(
                    messages_root, m.msg_type or "", m.direction or "", m.attachment_day, fname
                )
                if not target.exists():
                    # Some exports stash attachments without the dated subfolder; check fallback path
                    alt = messages_root / "attachments" / (m.msg_type or "") / (m.direction or "") / fname
                    chosen = target if target.exists() else (alt if alt.exists() else None)
                else:
                    chosen = target

                if chosen is None:
                    # Show a small missing-note so you know there *was* an attachment reference
                    attachment_snippets.append(
                        f"<div class=\"attachment\"><em class=\"meta\">(missing attachment: {html.escape(fname)})</em></div>"
                    )
                    continue

                kind = classify_ext(chosen)
                rel = relpath_for_html(out_file, chosen)
                if kind == "image":
                    attachment_snippets.append(
                        f"<div class=\"attachment\"><img loading=\"lazy\" src=\"{rel}\" alt=\"{html.escape(fname)}\"></div>"
                    )
                elif kind == "video":
                    attachment_snippets.append(
                        f"<div class=\"attachment\"><video controls preload=\"metadata\" src=\"{rel}\"></video></div>"
                    )
                elif kind == "audio":
                    attachment_snippets.append(
                        f"<div class=\"attachment\"><audio controls preload=\"metadata\" src=\"{rel}\"></audio></div>"
                    )
                elif kind == "inline_text":
                    try:
                        text = chosen.read_text(encoding="utf-8", errors="replace")
                        text = html.escape(text)
                        attachment_snippets.append(
                            f"<div class=\"attachment\"><pre class=\"code\" style=\"white-space:pre-wrap\">{text}</pre></div>"
                        )
                    except Exception:
                        attachment_snippets.append(
                            f"<div class=\"attachment\"><a href=\"{rel}\" download>{html.escape(fname)}</a></div>"
                        )
                else:
                    attachment_snippets.append(
                        f"<div class=\"attachment\"><a href=\"{rel}\" download>{html.escape(fname)}</a></div>"
                    )
            if attachment_snippets:
                with_attachments += 1
                parts.append("    <div class=\"attachments\">")
                parts.extend(["      " + s for s in attachment_snippets])
                parts.append("    </div>")

        msg_type = (m.msg_type or "unknown").upper()
        if msg_type == "MMS" and not attachment_snippets:
            parts.append(
                f"    <div class=\"body-text missing\">NO {msg_type} ATTACHMENT AVAILABLE - LOG ONLY</div>"
            )

        # Meta line
        if m.date_dt:
            local_str = m.date_dt.astimezone().strftime("%Y-%m-%d %H:%M:%S %Z")
            parts.append(
                f"    <div class=\"meta\">{html.escape(local_str)} · {html.escape(m.direction)} · {html.escape(m.msg_type)}</div>"
            )
        else:
            parts.append(
                f"    <div class=\"meta\">{html.escape(m.date_raw)} · {html.escape(m.direction)} · {html.escape(m.msg_type)}</div>"
            )

        parts.append("  </div>")  # bubble
        parts.append("</div>")    # message

    parts.append("</div>")  # container

    parts.append("<div class=\"footer\">")
    parts.append("  <div class=\"container\">Return to <a href=\"index.html\">index</a></div>")
    parts.append("</div>")
    parts.append("<script>")
    parts.append("const s=document.getElementById('search');")
    parts.append("s&&s.addEventListener('input',e=>{const q=e.target.value.toLowerCase();document.querySelectorAll('.message').forEach(m=>{m.style.display=m.textContent.toLowerCase().includes(q)?'':'none';});});")
    parts.append("</script>")
    parts.append("</body></html>")

    out_file.parent.mkdir(parents=True, exist_ok=True)
    out_file.write_text("\n".join(parts), encoding="utf-8")

    return total, with_attachments


# ------------------------- Index Page -------------------------

def write_index(out_dir: Path, entries: List[Tuple[str, str, int, int]]):
    # entries: list of (title, rel_path, msg_count, with_attach_count)
    lines: List[str] = []
    lines.append("<!DOCTYPE html>")
    lines.append("<html lang=\"en\"><head><meta charset=\"utf-8\"><meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">")
    lines.append("<title>Message Transcripts</title>")
    lines.append(f"<style>{INDEX_CSS}</style>")
    lines.append("</head><body>")
    lines.append("<div class=\"container\">")
    lines.append("  <h1>Message Transcripts</h1>")
    lines.append("  <div class=\"subtitle\">One HTML per chat. Click to view. (Times shown in local system timezone inside each transcript.)</div>")
    lines.append("  <div class=\"search-bar\"><input id=\"search\" class=\"search-input\" placeholder=\"Search chats\"></div>")
    lines.append("  <div class=\"list\">")
    for title, rel, c, ca in entries:
        lines.append("    <div class=\"item\">")
        lines.append(f"      <a href=\"{html.escape(rel)}\">{html.escape(title)}</a>")
        lines.append(f"      <div class=\"meta\">Messages: {c} · Messages with attachments: {ca}</div>")
        lines.append("    </div>")
    lines.append("  </div>")
    lines.append("</div>")
    lines.append("<script>")
    lines.append("const s=document.getElementById('search');")
    lines.append("s&&s.addEventListener('input',e=>{const q=e.target.value.toLowerCase();document.querySelectorAll('.item').forEach(it=>{it.style.display=it.textContent.toLowerCase().includes(q)?'':'none';});});")
    lines.append("</script>")
    lines.append("</body></html>")
    (out_dir / "index.html").write_text("\n".join(lines), encoding="utf-8")


# ------------------------- Main -------------------------

def main():
    ap = argparse.ArgumentParser(description="Render chat transcripts from CSVs into HTML.")
    ap.add_argument("--in", dest="in_dir", required=True, help="Input root folder (expects CSVs inside, plus attachments/...) e.g. messages")
    ap.add_argument("--out", dest="out_dir", required=True, help="Output folder for HTML transcripts, e.g. transcripts")
    ap.add_argument("--target-number", default="", help="Phone number of the target user")
    ap.add_argument(
        "--contacts-xlsx",
        dest="contacts_xlsx",
        default="",
        help="Path to Excel file mapping phone numbers to contacts",
    )
    args = ap.parse_args()

    target = args.target_number
    lookup = build_contact_lookup(args.contacts_xlsx)

    messages_root = Path(args.in_dir).resolve()
    out_root = Path(args.out_dir).resolve()
    out_root.mkdir(parents=True, exist_ok=True)

    csv_files = sorted(messages_root.glob("*.csv"))
    if not csv_files:
        print(f"No CSV files found in {messages_root}")
        return

    all_msgs: List[Message] = []
    call_records: List[Message] = []

    for csv_file in csv_files:
        msgs = load_messages_from_csv(csv_file, lookup)
        for m in msgs:
            if m.msg_type == "call":
                if not m.sender:
                    m.sender = target
                if not m.recipients:
                    m.recipients = target
                call_records.append(m)
            else:
                all_msgs.append(m)

    grouped = group_messages_by_chat(all_msgs, target)

    index_entries: List[Tuple[str, str, int, int]] = []
    for participants, msgs in grouped.items():
        title = f"Chat – {', '.join(participants)}"
        key = sanitize_participants(participants)
        out_file = out_root / f"chat-{key}.html"
        total, with_attachments = render_thread_html(
            messages_root, out_file, msgs, list(participants), target, lookup
        )
        rel = os.path.relpath(out_file, start=out_root).replace(os.sep, "/")
        index_entries.append((title, rel, total, with_attachments))
        print(f"Rendered chat {', '.join(participants)}: {total} messages ({with_attachments} with attachments)")

    write_index(out_root, index_entries)

    from openpyxl import Workbook

    call_log_path = out_root / "Call Log.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.append(["Date", "Direction", "Sender", "Recipients", "Message ID"])
    for m in call_records:
        if m.date_dt:
            date_str = m.date_dt.astimezone().strftime("%Y-%m-%d %H:%M:%S %Z")
        else:
            date_str = m.date_raw
        ws.append([date_str, m.direction, m.sender, m.recipients, m.message_id])
    wb.save(call_log_path)

    print(f"\nDone. Open: {out_root / 'index.html'}")


if __name__ == "__main__":
    main()
