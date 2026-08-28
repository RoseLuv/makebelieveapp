"""
Booking Importer — desktop GUI

Reads booking rows from an Excel file and submits each one to the
Make Believe Group booking endpoint, driven from a simple point-and-click
window instead of the command line.

Buttons:
  - Select File...   choose the .xlsx to import
  - Set Columns...   map each required field to a spreadsheet column
  - Send             run the import (respects the Dry Run checkbox)
  - Cancel           stop a run that's in progress, after the current row

The import itself runs on a background thread so the window stays
responsive, with live per-row logging and a progress bar. This file is a
plain Tkinter app (no external UI framework), which keeps it compatible
with PyInstaller/py2app packaging into a Windows .exe or macOS .app later.
"""

from __future__ import annotations

import csv
import queue
import re
import threading
import time
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Callable, Optional

import openpyxl
import requests
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

# --------------------------------------------------------------------------
# Configuration
# --------------------------------------------------------------------------

URL = "https://www.makebelievegroup.co.uk/wp-admin/admin-ajax.php"
CALENDAR_ID = "202"
REQUEST_TIMEOUT_SECONDS = 15
SECONDS_BETWEEN_REQUESTS = 20
MAX_RETRIES = 3
RETRY_BACKOFF_SECONDS = 5

HEADER = {
    "User-Agent": "Mozilla/5.0",
    "Accept": "*/*",
    "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
    "X-Requested-With": "XMLHttpRequest",
    "Origin": "https://www.makebelievegroup.co.uk",
    "Referer": "https://www.makebelievegroup.co.uk/school/make-believe-hayes-harlington/",
}

MONTH_CONVERSION = {
    "JAN": "01", "FEB": "02", "MAR": "03",
    "APR": "04", "MAY": "05", "JUN": "06",
    "JUL": "07", "AUG": "08", "SEP": "09",
    "OCT": "10", "NOV": "11", "DEC": "12",
}

# Matches "11th April 2026", "1 Jan26", "3rd Dec, 2027", etc.
WRITTEN_MONTH_DATE_PATTERN = re.compile(
    r"^(\d{1,2})(?:st|nd|rd|th)?\s+([A-Za-z]+)\.?,?\s*(\d{2,4})$",
    re.IGNORECASE,
)

FIELD_NAMES = [
    "Child Name",
    "Parent Name",
    "Date",
    "Group Time",
    "Child Birthday",
    "Email",
    "Phone Number",
    "Postcode",
    "Source",
]


# --------------------------------------------------------------------------
# Parsing helpers (unchanged from the CLI version)
# --------------------------------------------------------------------------

class RowParseError(Exception):
    """Raised when a single row can't be turned into a valid payload."""


def parseName(name) -> tuple[str, str]:
    if name is None or not str(name).strip():
        raise RowParseError("Name is empty or missing")
    parts = str(name).split()
    if len(parts) == 1:
        return parts[0], parts[0]
    return parts[0], " ".join(parts[1:])


def parseWrittenMonthDate(text: str) -> Optional[str]:
    match = WRITTEN_MONTH_DATE_PATTERN.match(text)
    if not match:
        return None

    day, monthWord, year = match.groups()
    monthKey = monthWord[:3].upper()
    month = MONTH_CONVERSION.get(monthKey)
    if month is None:
        raise RowParseError(f"Unrecognised month name {monthWord!r} in date: {text!r}")

    day = day.zfill(2)
    year = "20" + year if len(year) == 2 else year

    try:
        datetime(int(year), int(month), int(day))
    except ValueError as exc:
        raise RowParseError(f"Invalid calendar date: {text!r} ({exc})") from exc

    return f"{year}-{month}-{day}"


def parseDate(value) -> str:
    """
    Accepts either:
      - a datetime/date object straight from openpyxl,
      - a string in DD.MM.YYYY / DD/MM/YYYY / DD-MM-YYYY / DD MM YYYY / DDMMYYYY, or
      - a string with a written month name, e.g. '11th April 2026' or '1 Jan26'.
    """
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")

    if value is None or not str(value).strip():
        raise RowParseError("Date is empty or missing")

    text = str(value).strip()

    writtenMonthResult = parseWrittenMonthDate(text)
    if writtenMonthResult is not None:
        return writtenMonthResult

    if "." in text:
        dateParts = text.split(".")
    elif "/" in text:
        dateParts = text.split("/")
    elif "-" in text:
        dateParts = text.split("-")
    elif " " in text:
        dateParts = text.split(" ")
    elif text.isdigit() and len(text) == 8:
        dateParts = [text[:2], text[2:4], text[4:]]
    else:
        raise RowParseError(f"Unrecognised date format: {value!r}")

    if len(dateParts) != 3:
        raise RowParseError(f"Unrecognised date format: {value!r}")

    day, month, year = dateParts

    if not (day.isnumeric() and month.isnumeric() and year.isnumeric()):
        raise RowParseError(f"Non-numeric date component: {value!r}")

    if len(day) not in (1, 2) or len(month) not in (1, 2) or len(year) not in (2, 4):
        raise RowParseError(f"Unexpected date component lengths: {value!r}")

    day = day.zfill(2)
    month = month.zfill(2)
    year = "20" + year if len(year) == 2 else year

    try:
        datetime(int(year), int(month), int(day))
    except ValueError as exc:
        raise RowParseError(f"Invalid calendar date: {value!r} ({exc})") from exc

    return f"{year}-{month}-{day}"


def parseGroupTime(groupTime) -> str:
    if groupTime is None or not str(groupTime).strip():
        return "1000-1300"
    if "1st group" in str(groupTime):
        return "1000-1130"
    return "1130-1300"


def parseGroup(groupTime: str) -> str:
    if groupTime == "1000-1130":
        return "Infants"
    if groupTime == "1130-1300":
        return "Infants+"
    return "Juniors & Seniors"


def parseUnixTime(registerDate: str, groupTime: str) -> int:
    year, month, day = (int(part) for part in registerDate.split("-"))
    hour, minute = (10, 0) if groupTime[:2] == "10" else (11, 30)
    dt = datetime(year, month, day, hour, minute, tzinfo=timezone.utc)
    return int(dt.timestamp())


def cellText(value) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    return str(value)


# --------------------------------------------------------------------------
# Row -> payload
# --------------------------------------------------------------------------

@dataclass
class RowResult:
    rowNumber: int
    status: str  # "sent", "dry_run", "http_error", "parse_error", "cancelled"
    detail: str = ""


def buildPayload(row, columnInfo: dict[str, int]) -> dict:
    childFirstName, childLastName = parseName(row[columnInfo["Child Name"]].value)
    parentFirstName, parentLastName = parseName(row[columnInfo["Parent Name"]].value)
    registerDate = parseDate(row[columnInfo["Date"]].value)
    groupTime = parseGroupTime(row[columnInfo["Group Time"]].value)
    group = parseGroup(groupTime)
    groupUnixTime = parseUnixTime(registerDate, groupTime)

    childBirthday = cellText(row[columnInfo["Child Birthday"]].value)
    email = cellText(row[columnInfo["Email"]].value)
    phoneNumber = cellText(row[columnInfo["Phone Number"]].value)
    postcode = cellText(row[columnInfo["Postcode"]].value)
    source = cellText(row[columnInfo["Source"]].value)

    if not email:
        raise RowParseError("Email is empty or missing")

    return {
        "guest_name": parentFirstName,
        "guest_surname": parentLastName,
        "guest_email": email,
        "action": "booked_add_appt",
        "customer_type": "guest",
        "is_fe_form": "true",
        "total_appts": "1",
        "appoinment": "0",
        "calendar_id": CALENDAR_ID,
        "title": group,
        "date": registerDate,
        "timestamp": groupUnixTime,
        "timeslot": groupTime,
        "single-line-text-label---2611670___required": childFirstName,
        "single-line-text-label---5862683___required": childLastName,
        "single-line-text-label---1962235___required": childBirthday,
        "single-line-text-label---7180961___required": phoneNumber,
        "single-line-text-label---5004481___required": postcode,
        "single-line-text-label---4933958___required": source,
    }


# --------------------------------------------------------------------------
# Network
# --------------------------------------------------------------------------

def sendPayload(payload: dict) -> requests.Response:
    lastExc: Optional[Exception] = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            response = requests.post(
                url=URL, data=payload, headers=HEADER, timeout=REQUEST_TIMEOUT_SECONDS
            )
            return response
        except requests.RequestException as exc:
            lastExc = exc
            if attempt < MAX_RETRIES:
                time.sleep(RETRY_BACKOFF_SECONDS * attempt)
    assert lastExc is not None
    raise lastExc


def convertColumnValue(columnValue: str) -> int:
    columnValue = columnValue.strip()
    if not columnValue:
        raise ValueError("Column value is empty")

    if columnValue.isdigit():
        columnIndex = int(columnValue)
    else:
        letters = "".join(ch for ch in columnValue.upper() if ch.isalpha())
        if not letters:
            raise ValueError(f"Invalid column value: {columnValue}")
        columnIndex = 0
        for letter in letters:
            columnIndex = columnIndex * 26 + (ord(letter) - 64)

    if columnIndex <= 0:
        raise ValueError("Column index must be greater than 0")
    return columnIndex - 1


# --------------------------------------------------------------------------
# Excel + row loop
#
# logCallback / progressCallback let the caller (the GUI) observe progress
# without this function touching any widgets directly - important because
# widgets can only safely be touched from the main thread, and this
# function is designed to run on a background thread.
# --------------------------------------------------------------------------

def sendRows(
    path: Path,
    columnInfo: dict[str, int],
    dryRun: bool,
    logCallback: Callable[[str], None] = print,
    progressCallback: Optional[Callable[[int, int], None]] = None,
    cancelEvent: Optional[threading.Event] = None,
) -> list[RowResult]:
    workBookObject = openpyxl.load_workbook(path)
    sheetObject = workBookObject.active

    dataRows = list(sheetObject.iter_rows(min_row=2))
    totalRows = len(dataRows)
    results: list[RowResult] = []

    for index, row in enumerate(dataRows, start=1):
        rowNumber = index + 1  # +1 to account for the header row

        if cancelEvent is not None and cancelEvent.is_set():
            results.append(RowResult(rowNumber, "cancelled", "Run cancelled by user"))
            logCallback(f"[row {rowNumber}] CANCELLED - stopping before this row")
            break

        if progressCallback is not None:
            progressCallback(index, totalRows)

        if all(cell.value is None for cell in row):
            continue  # skip fully blank rows

        try:
            payload = buildPayload(row, columnInfo)
        except RowParseError as exc:
            results.append(RowResult(rowNumber, "parse_error", str(exc)))
            logCallback(f"[row {rowNumber}] SKIPPED (parse error): {exc}")
            continue

        if dryRun:
            results.append(RowResult(rowNumber, "dry_run", repr(payload)))
            logCallback(f"[row {rowNumber}] DRY RUN: {payload}")
            continue

        try:
            response = sendPayload(payload)
        except requests.RequestException as exc:
            results.append(RowResult(rowNumber, "http_error", f"request failed: {exc}"))
            logCallback(f"[row {rowNumber}] FAILED to send: {exc}")
            time.sleep(SECONDS_BETWEEN_REQUESTS)
            continue

        if response.ok:
            results.append(RowResult(rowNumber, "sent", f"HTTP {response.status_code}"))
            logCallback(f"[row {rowNumber}] sent (HTTP {response.status_code})")
        else:
            bodySnippet = response.text[:200].replace("\n", " ")
            results.append(
                RowResult(rowNumber, "http_error", f"HTTP {response.status_code}: {bodySnippet}")
            )
            logCallback(f"[row {rowNumber}] server rejected (HTTP {response.status_code}): {bodySnippet}")

        time.sleep(SECONDS_BETWEEN_REQUESTS)

    if progressCallback is not None:
        progressCallback(totalRows, totalRows)

    return results


def writeLog(results: list[RowResult], sourcePath: Path) -> Path:
    logPath = sourcePath.with_name(
        f"{sourcePath.stem}_import_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    )
    with logPath.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["rowNumber", "status", "detail"])
        for r in results:
            writer.writerow([r.rowNumber, r.status, r.detail])
    return logPath


# --------------------------------------------------------------------------
# GUI
# --------------------------------------------------------------------------

class ColumnMappingDialog(tk.Toplevel):
    """Modal form for mapping each field to a spreadsheet column."""

    def __init__(self, parent: tk.Tk, existingColumnInfo: Optional[dict[str, int]]):
        super().__init__(parent)
        self.title("Map Spreadsheet Columns")
        self.resizable(False, False)
        self.result: Optional[dict[str, int]] = None

        instructions = (
            "Enter the column for each field, either as a letter (A, B, C...) "
            "or a number counting from 1."
        )
        ttk.Label(self, text=instructions, wraplength=340, justify="left").grid(
            row=0, column=0, columnspan=2, padx=12, pady=(12, 8), sticky="w"
        )

        self.entryVars: dict[str, tk.StringVar] = {}
        for i, fieldName in enumerate(FIELD_NAMES, start=1):
            ttk.Label(self, text=f"{fieldName}:").grid(
                row=i, column=0, sticky="e", padx=(12, 6), pady=3
            )
            entryVar = tk.StringVar()
            if existingColumnInfo and fieldName in existingColumnInfo:
                entryVar.set(str(existingColumnInfo[fieldName] + 1))
            ttk.Entry(self, textvariable=entryVar, width=10).grid(
                row=i, column=1, sticky="w", padx=(0, 12), pady=3
            )
            self.entryVars[fieldName] = entryVar

        buttonRow = len(FIELD_NAMES) + 1
        buttonFrame = ttk.Frame(self)
        buttonFrame.grid(row=buttonRow, column=0, columnspan=2, pady=(10, 12))
        ttk.Button(buttonFrame, text="OK", command=self._onSubmit).pack(side="left", padx=6)
        ttk.Button(buttonFrame, text="Cancel", command=self.destroy).pack(side="left", padx=6)

        self.bind("<Return>", lambda event: self._onSubmit())
        self.transient(parent)
        self.grab_set()

    def _onSubmit(self) -> None:
        try:
            parsedColumns = {
                fieldName: convertColumnValue(entryVar.get())
                for fieldName, entryVar in self.entryVars.items()
            }
        except ValueError as exc:
            messagebox.showerror("Invalid column", str(exc), parent=self)
            return
        self.result = parsedColumns
        self.destroy()


class BookingImporterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Booking Importer")
        self.resizable(True, True)
        self.minsize(560, 480)

        self.selectedFilePath: Optional[Path] = None
        self.columnInfo: Optional[dict[str, int]] = None
        self.dryRunVar = tk.BooleanVar(value=True)
        self.isRunning = False
        self.cancelEvent = threading.Event()
        self.messageQueue: "queue.Queue[tuple]" = queue.Queue()

        self._buildWidgets()
        self._updateSendButtonState()

    # ---- widget construction -------------------------------------------------

    def _buildWidgets(self) -> None:
        pad = {"padx": 10, "pady": 6}

        setupFrame = ttk.LabelFrame(self, text="1. Setup")
        setupFrame.pack(fill="x", **pad)

        fileRow = ttk.Frame(setupFrame)
        fileRow.pack(fill="x", padx=8, pady=(8, 4))
        self.selectFileButton = ttk.Button(
            fileRow, text="Select File...", command=self._onSelectFile
        )
        self.selectFileButton.pack(side="left")
        self.fileLabel = ttk.Label(fileRow, text="No file selected", foreground="#555555")
        self.fileLabel.pack(side="left", padx=10)

        columnsRow = ttk.Frame(setupFrame)
        columnsRow.pack(fill="x", padx=8, pady=(4, 8))
        self.setColumnsButton = ttk.Button(
            columnsRow, text="Set Columns...", command=self._onSetColumns
        )
        self.setColumnsButton.pack(side="left")
        self.columnsLabel = ttk.Label(columnsRow, text="Columns not set", foreground="#555555")
        self.columnsLabel.pack(side="left", padx=10)

        optionsFrame = ttk.LabelFrame(self, text="2. Options")
        optionsFrame.pack(fill="x", **pad)
        ttk.Checkbutton(
            optionsFrame,
            text="Dry run (parse and log rows, but never actually send anything)",
            variable=self.dryRunVar,
        ).pack(anchor="w", padx=8, pady=8)

        actionFrame = ttk.LabelFrame(self, text="3. Run")
        actionFrame.pack(fill="x", **pad)

        actionButtonsRow = ttk.Frame(actionFrame)
        actionButtonsRow.pack(fill="x", padx=8, pady=(8, 4))
        self.sendButton = ttk.Button(actionButtonsRow, text="Send", command=self._onSend)
        self.sendButton.pack(side="left")
        self.cancelButton = ttk.Button(
            actionButtonsRow, text="Cancel", command=self._onCancel, state="disabled"
        )
        self.cancelButton.pack(side="left", padx=(8, 0))

        progressRow = ttk.Frame(actionFrame)
        progressRow.pack(fill="x", padx=8, pady=(4, 8))
        self.progressBar = ttk.Progressbar(progressRow, mode="determinate")
        self.progressBar.pack(fill="x", side="left", expand=True)
        self.progressLabel = ttk.Label(progressRow, text="0 / 0 rows")
        self.progressLabel.pack(side="left", padx=(10, 0))

        logFrame = ttk.LabelFrame(self, text="Log")
        logFrame.pack(fill="both", expand=True, **pad)
        self.logText = scrolledtext.ScrolledText(
            logFrame, height=14, state="disabled", font=("Courier New", 10)
        )
        self.logText.pack(fill="both", expand=True, padx=8, pady=8)

        self.statusLabel = ttk.Label(self, text="Ready.", anchor="w", relief="sunken")
        self.statusLabel.pack(fill="x", side="bottom")

    # ---- button handlers -------------------------------------------------

    def _onSelectFile(self) -> None:
        selected = filedialog.askopenfilename(
            title="Select file location",
            filetypes=[("Excel Workbook", "*.xlsx")],
        )
        if not selected:
            return
        self.selectedFilePath = Path(selected)
        self.fileLabel.config(text=self.selectedFilePath.name, foreground="black")
        self._setStatus(f"Selected file: {self.selectedFilePath}")
        self._updateSendButtonState()

    def _onSetColumns(self) -> None:
        dialog = ColumnMappingDialog(self, self.columnInfo)
        self.wait_window(dialog)
        if dialog.result is not None:
            self.columnInfo = dialog.result
            summary = ", ".join(
                f"{name}={index + 1}" for name, index in self.columnInfo.items()
            )
            self.columnsLabel.config(text=summary, foreground="black")
            self._setStatus("Column mapping saved.")
        self._updateSendButtonState()

    def _onSend(self) -> None:
        if self.selectedFilePath is None or self.columnInfo is None:
            messagebox.showwarning(
                "Missing setup",
                "Please select a file and set the column mapping first.",
            )
            return

        if self.dryRunVar.get():
            confirmed = True
        else:
            confirmed = messagebox.askyesno(
                "Confirm send",
                "Dry run is OFF. This will actually submit rows to the live "
                "server. Continue?",
            )
        if not confirmed:
            return

        self.isRunning = True
        self.cancelEvent = threading.Event()
        self._clearLog()
        self.progressBar["value"] = 0
        self.progressLabel.config(text="0 / 0 rows")
        self._setStatus("Running...")
        self._setControlsEnabled(running=True)

        workerThread = threading.Thread(
            target=self._runImportOnBackgroundThread,
            args=(self.selectedFilePath, self.columnInfo, self.dryRunVar.get(), self.cancelEvent),
            daemon=True,
        )
        workerThread.start()
        self.after(100, self._pollQueue)

    def _onCancel(self) -> None:
        self.cancelEvent.set()
        self.cancelButton.config(state="disabled")
        self._setStatus("Cancelling after the current row...")

    # ---- background thread work -------------------------------------------------

    def _runImportOnBackgroundThread(
        self,
        filePath: Path,
        columnInfo: dict[str, int],
        dryRun: bool,
        cancelEvent: threading.Event,
    ) -> None:
        def logCallback(message: str) -> None:
            self.messageQueue.put(("log", message))

        def progressCallback(current: int, total: int) -> None:
            self.messageQueue.put(("progress", current, total))

        try:
            results = sendRows(
                filePath,
                columnInfo,
                dryRun=dryRun,
                logCallback=logCallback,
                progressCallback=progressCallback,
                cancelEvent=cancelEvent,
            )
            logPath = writeLog(results, filePath)
            self.messageQueue.put(("done", results, logPath))
        except Exception as exc:  # surface any unexpected error to the UI instead of crashing silently
            self.messageQueue.put(("error", str(exc)))

    # ---- queue polling (runs on the main/UI thread) -------------------------------------------------

    def _pollQueue(self) -> None:
        try:
            while True:
                message = self.messageQueue.get_nowait()
                kind = message[0]

                if kind == "log":
                    self._appendLog(message[1])
                elif kind == "progress":
                    _, current, total = message
                    self.progressBar["maximum"] = max(total, 1)
                    self.progressBar["value"] = current
                    self.progressLabel.config(text=f"{current} / {total} rows")
                elif kind == "done":
                    _, results, logPath = message
                    self._finishRun(results, logPath)
                elif kind == "error":
                    self._finishRunWithError(message[1])
        except queue.Empty:
            pass

        if self.isRunning:
            self.after(100, self._pollQueue)

    def _finishRun(self, results: list[RowResult], logPath: Path) -> None:
        self.isRunning = False
        self._setControlsEnabled(running=False)

        sentCount = sum(1 for r in results if r.status == "sent")
        dryRunCount = sum(1 for r in results if r.status == "dry_run")
        failedCount = sum(1 for r in results if r.status in ("http_error", "parse_error"))
        cancelledCount = sum(1 for r in results if r.status == "cancelled")

        summaryLines = [
            "",
            "--- Summary ---",
            f"Sent:      {sentCount}",
            f"Dry run:   {dryRunCount}",
            f"Failed:    {failedCount}",
        ]
        if cancelledCount:
            summaryLines.append("Run was cancelled before finishing.")
        summaryLines.append(f"Log written to: {logPath}")
        self._appendLog("\n".join(summaryLines))

        self._setStatus(f"Done. Sent {sentCount}, dry-run {dryRunCount}, failed {failedCount}.")
        messagebox.showinfo(
            "Import finished",
            f"Sent: {sentCount}\nDry run: {dryRunCount}\nFailed: {failedCount}\n\n"
            f"Log written to:\n{logPath}",
        )

    def _finishRunWithError(self, errorMessage: str) -> None:
        self.isRunning = False
        self._setControlsEnabled(running=False)
        self._appendLog(f"\nUNEXPECTED ERROR: {errorMessage}")
        self._setStatus("Run failed with an unexpected error.")
        messagebox.showerror("Unexpected error", errorMessage)

    # ---- small UI helpers -------------------------------------------------

    def _setControlsEnabled(self, running: bool) -> None:
        self.sendButton.config(state="disabled" if running else "normal")
        self.selectFileButton.config(state="disabled" if running else "normal")
        self.setColumnsButton.config(state="disabled" if running else "normal")
        self.cancelButton.config(state="normal" if running else "disabled")
        if not running:
            self._updateSendButtonState()

    def _updateSendButtonState(self) -> None:
        ready = self.selectedFilePath is not None and self.columnInfo is not None
        if not self.isRunning:
            self.sendButton.config(state="normal" if ready else "disabled")

    def _clearLog(self) -> None:
        self.logText.config(state="normal")
        self.logText.delete("1.0", tk.END)
        self.logText.config(state="disabled")

    def _appendLog(self, text: str) -> None:
        self.logText.config(state="normal")
        self.logText.insert(tk.END, text + "\n")
        self.logText.see(tk.END)
        self.logText.config(state="disabled")

    def _setStatus(self, text: str) -> None:
        self.statusLabel.config(text=text)


def main() -> None:
    app = BookingImporterApp()
    app.mainloop()


if __name__ == "__main__":
    main()