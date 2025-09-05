# Synchronoss Parser

Utility scripts for working with Synchronoss data exports.

This toolbox converts Verizon/Synchronoss backups into more useful formats. It can compile media
files, convert contact lists, merge call logs with contacts, render HTML message transcripts and
produce attachment inventories. A Tkinter GUI exposes these features for non‑technical users.

Phone numbers are normalized by removing punctuation and an optional leading `1` country code so
contacts match regardless of formatting.

## Sample data layout

Scripts expect exports arranged like:

```
Call Log/
  call_log.csv
messages/
  <date>.csv
  attachments/
    mms/in/<YYYY-MM-DD>/<files>
    mms/out/<YYYY-MM-DD>/<files>
```

## Installation

Install the package and its dependencies with pip:

```bash
pip install .
```

Tkinter may require additional OS-level packages.

## Scripts

### collect_media.py
Copy media from a Verizon Mobile backup into a single folder and log EXIF metadata to Excel.

```bash
collect-media
```

### collect_attachments.py
Collect message attachments from a Synchronoss export into a single folder and log metadata to
Excel.

The script scans the `messages/attachments/` directory:

```
messages/
  attachments/
    mms/in/<YYYY-MM-DD>/<files>
    mms/out/<YYYY-MM-DD>/<files>
```

All files are copied to `Compiled Attachments/` and their details recorded in
`Compiled Attachments/compiled_attachment_log/compiled_attachment_log.xlsx`.

```bash
collect-attachments
```

### contacts_to_excel.py
Convert a `contacts.txt` export to an Excel spreadsheet.

```bash
contacts-to-excel --input contacts.txt --output contacts.xlsx
```

### merge_contacts_logs.py
Annotate a call log CSV with caller and recipient names using a contacts Excel file.

```bash
merge-contacts-logs --call-log "Call Log/call_log.csv" --contacts-xlsx contacts.xlsx
```

### render_transcripts.py
Render chat-bubble style HTML transcripts from message CSVs and link attachments. Optionally supply
a contacts Excel file for name lookups.

```bash
render-transcripts --in messages --out transcripts --contacts-xlsx contacts.xlsx
```

### decrypt_unzip.py
Decrypt a top-level Synchronoss archive and expand any nested encrypted archives. The function
recursively decrypts files ending in `.gpg` (e.g. `.zip.gpg`, `.txt.gpg`) using `gpg` and unpacks
every archive it finds. `gpg` must be installed and available on your `PATH`. Processed `.zip` and
`.gpg` archives are removed after extraction by default, but the original input archive is always
preserved. Pass `cleanup=False` to retain all archives.

```python
from pathlib import Path
from synchronoss_parser.decrypt_unzip import decrypt_and_unzip

decrypt_and_unzip(Path("export.zip"), "passphrase")
decrypt_and_unzip(Path("export.zip"), "passphrase", cleanup=False)  # keep archives
```

### toolbox_gui.py
Tkinter GUI that wraps Collect Media, Collect Attachments, Contacts to Excel, Render Transcripts and
Decrypt & Unzip workflows.

The Decrypt & Unzip tab accepts the path to an encrypted archive (``.zip``) and the password used
to decrypt it. Any nested ``.gpg`` files are automatically decrypted, and any decrypted archives are
expanded. The tab requires a local `gpg` installation to be present on the system path.

```bash
toolbox-gui
```

### collect_media_gui.py
Simple GUI wrapper for `collect_media.py`.

```bash
collect-media-gui
```

### attachment_log.py
Generate an inventory of all message attachments. The script scans message CSVs and
records each attachment along with its sender and recipients, producing an Excel
workbook and an HTML table with thumbnails.

```bash
attachment-log --messages messages --out "Attachment Log"
```

## Building Windows Executables

Install PyInstaller:

```bash
pip install pyinstaller
```

Build one-file, windowed executables for the GUI scripts using the provided
helper after installing the package:

```bash
build-exe
```

The command invokes PyInstaller on `toolbox_gui.spec` (and
`collect_media_gui.spec` if present). PyInstaller places the resulting
executables in `dist/` and bundles Python and all required libraries, so
recipients don't need Python installed. On Windows, run them by
double-clicking or from a Command Prompt:

```bash
dist\\toolbox_gui.exe
dist\\collect_media_gui.exe
```

### Verifying the build

Run the generated `.exe` on a clean Windows machine without Python
installed to confirm the bundled interpreter works. Copy the file from
`dist/` to the test system and launch it; the application should open
normally without additional dependencies.

## Testing

Run the test suite with:

```bash
pytest
```

