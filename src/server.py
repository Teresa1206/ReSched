"""Local website for the ReSched timetable input and scheduling pipeline.

Run with the project virtual environment:
    .venv/bin/python src/server.py

The server listens only on 127.0.0.1.
"""

from __future__ import annotations

import json
import mimetypes
import re
import subprocess
import sys
import tempfile
import threading
import webbrowser
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import unquote, urlparse
from uuid import uuid4

import pandas as pd

ROOT = Path(__file__).resolve().parent
PROJECT = ROOT.parent
INPUT_PAGE = ROOT / "input.html"
SCHEDULE_PAGE = PROJECT / "schedule.html"
PREVIEW_PAGE = ROOT / "preview_schedule.html"
RUNS = ROOT / "runs"
RUNS.mkdir(exist_ok=True)
FILES: dict[str, Path] = {}
PORT = 8765


def as_text(value: object) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value).strip()


def make_key(path: Path) -> str:
    key = uuid4().hex
    FILES[key] = path
    return key


def safe_name(value: str, fallback: str) -> str:
    name = Path(value).name
    name = re.sub(r"[^A-Za-z0-9._() -]", "_", name)
    return name or fallback


def workbook_from_payload(payload: dict, directory: Path) -> Path:
    sheets = payload.get("sheets", [])
    if not isinstance(sheets, list) or not sheets:
        raise ValueError("The input page has no data to save.")
    output = directory / f"{uuid4().hex}_input.xlsx"
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for item in sheets:
            if not isinstance(item, list) or len(item) != 2 or not item[1]:
                continue
            name, grid = item
            if not isinstance(grid, list) or not grid:
                continue
            header, *values = grid
            pd.DataFrame(values, columns=header).to_excel(
                writer, sheet_name=str(name)[:31], index=False
            )
    return output


def run_script(script: str, args: list[str], log_path: Path) -> subprocess.CompletedProcess:
    command = [sys.executable, str(ROOT / script), *args]
    completed = subprocess.run(
        command,
        cwd=ROOT,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=600,
    )
    log_path.write_text(
        f"$ {' '.join(command)}\n\nSTDOUT\n{completed.stdout}\n\nSTDERR\n{completed.stderr}\n",
        encoding="utf-8",
    )
    print(f"{script} exited with {completed.returncode}; log: {log_path}")
    return completed


class Handler(BaseHTTPRequestHandler):
    server_version = "ReSchedLocal/1.0"

    def log_message(self, format: str, *args: object) -> None:
        return

    def send_body(self, body: bytes, content_type: str, status: int = HTTPStatus.OK) -> None:
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Cache-Control", "no-store, max-age=0")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def send_json(self, payload: dict, status: int = HTTPStatus.OK) -> None:
        self.send_body(
            json.dumps(payload, ensure_ascii=False).encode("utf-8"),
            "application/json; charset=utf-8",
            status,
        )

    def read_json(self) -> dict:
        size = int(self.headers.get("Content-Length", "0"))
        if not size:
            return {}
        return json.loads(self.rfile.read(size).decode("utf-8"))

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path in {"/", "/index.html", "/input.html"}:
            self.send_body(INPUT_PAGE.read_bytes(), "text/html; charset=utf-8")
            return
        if parsed.path == "/schedule.html":
            if not SCHEDULE_PAGE.is_file():
                self.send_error(HTTPStatus.NOT_FOUND, "schedule.html not found")
                return
            self.send_body(SCHEDULE_PAGE.read_bytes(), "text/html; charset=utf-8")
            return
        if parsed.path == "/preview.html":
            if not PREVIEW_PAGE.is_file():
                self.send_error(HTTPStatus.NOT_FOUND, "No timetable has been generated yet")
                return
            self.send_body(PREVIEW_PAGE.read_bytes(), "text/html; charset=utf-8")
            return
        if parsed.path.startswith("/downloads/"):
            key = unquote(parsed.path.removeprefix("/downloads/"))
            path = FILES.get(key)
            if not path or not path.is_file():
                self.send_error(HTTPStatus.NOT_FOUND, "File not found")
                return
            content_type = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
            self.send_body(path.read_bytes(), content_type)
            return
        self.send_error(HTTPStatus.NOT_FOUND)

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        try:
            if parsed.path == "/api/sync-input":
                payload = self.read_json()
                directory = RUNS / uuid4().hex
                directory.mkdir(parents=True, exist_ok=True)
                path = workbook_from_payload(payload, directory)
                self.send_json({
                    "ok": True,
                    "sourceKey": make_key(path),
                    "name": path.name,
                    "log": "Input saved.",
                })
                return

            match = re.fullmatch(r"/api/run/([123])", parsed.path)
            if match:
                self.handle_run(match.group(1))
                return

            self.send_error(HTTPStatus.NOT_FOUND)
        except (ValueError, OSError, json.JSONDecodeError) as error:
            self.send_json({"ok": False, "log": str(error)}, HTTPStatus.BAD_REQUEST)

    def handle_run(self, step: str) -> None:
        source_key = self.headers.get("X-Source-Key", "")
        directory = RUNS / uuid4().hex
        directory.mkdir(parents=True, exist_ok=True)

        if step == "1" and not source_key:
            size = int(self.headers.get("Content-Length", "0"))
            if not size:
                self.send_json({"ok": False, "log": "Choose an Excel file or use the form below."}, HTTPStatus.BAD_REQUEST)
                return
            name = safe_name(unquote(self.headers.get("X-Filename", "timetable-input.xlsx")), "timetable-input.xlsx")
            if Path(name).suffix.lower() != ".xlsx":
                self.send_json({"ok": False, "log": "Choose an .xlsx file."}, HTTPStatus.BAD_REQUEST)
                return
            input_path = directory / f"{uuid4().hex}_{name}"
            input_path.write_bytes(self.rfile.read(size))
            source_key = make_key(input_path)
        else:
            input_path = FILES.get(source_key)
            if not input_path or not input_path.is_file():
                self.send_json({"ok": False, "log": "The previous step file was not found. Run the step again."}, HTTPStatus.BAD_REQUEST)
                return
            directory = input_path.parent

        log_path = directory / f"step{step}.log"

        if step == "1":
            completed = run_script("step1_validate_input.py", [str(input_path)], log_path)
            if completed.returncode != 0:
                self.send_json({
                    "ok": False,
                    "log": "Input validation failed. Check the server log for details.",
                    "sourceKey": source_key,
                }, HTTPStatus.UNPROCESSABLE_ENTITY)
                return
            self.send_json({
                "ok": True,
                "log": "Input validation passed. Run step 2 to calculate room distances and prepare the model.",
                "sourceKey": source_key,
                "downloads": [],
            })
            return

        if step == "2":
            travel_path = input_path.with_name(f"{input_path.stem}_travel.xlsx")
            completed = run_script("program1.py", [str(input_path), str(travel_path)], log_path)
            if completed.returncode != 0:
                self.send_json({"ok": False, "log": "Room distance calculation failed. Check the server log for details."}, HTTPStatus.UNPROCESSABLE_ENTITY)
                return
            prepared_path = travel_path.with_name(f"{travel_path.stem}_prepared.xlsx")
            completed = run_script("step2_prepare_model.py", [str(travel_path)], log_path)
            if completed.returncode != 0:
                self.send_json({"ok": False, "log": "Model preparation failed. Check the server log for details."}, HTTPStatus.UNPROCESSABLE_ENTITY)
                return
            downloads = [
                {"name": travel_path.name, "url": f"/downloads/{make_key(travel_path)}"},
                {"name": prepared_path.name, "url": f"/downloads/{make_key(prepared_path)}"},
            ]
            self.send_json({
                "ok": True,
                "log": "Room distances and model data prepared.",
                "nextSourceKey": make_key(prepared_path),
                "downloads": downloads,
            })
            return

        if step == "3":
            prepared_path = input_path
            travel_path = prepared_path.with_name(f"{prepared_path.stem.removesuffix('_prepared')}.xlsx")
            schedule_xlsx = prepared_path.with_name(f"{prepared_path.stem}_schedule.xlsx")
            schedule_json = prepared_path.with_name(f"{prepared_path.stem}_schedule.json")
            completed = run_script("step3_schedule.py", [str(prepared_path)], log_path)
            if completed.returncode != 0:
                self.send_json({"ok": False, "log": "Scheduling failed. Check the server log for details."}, HTTPStatus.UNPROCESSABLE_ENTITY)
                return

            template = SCHEDULE_PAGE if SCHEDULE_PAGE.is_file() else PREVIEW_PAGE
            completed = run_script(
                "build_schedule.py",
                [str(schedule_json), str(prepared_path), str(travel_path), str(template), str(PREVIEW_PAGE)],
                directory / "build.log",
            )
            if completed.returncode != 0:
                self.send_json({"ok": False, "log": "The timetable was scheduled, but the schedule page could not be built."}, HTTPStatus.UNPROCESSABLE_ENTITY)
                return

            try:
                result = json.loads(schedule_json.read_text(encoding="utf-8"))
                lesson_count = len(result.get("Schedule", []))
            except (OSError, json.JSONDecodeError):
                lesson_count = 0
            downloads = [
                {"name": schedule_xlsx.name, "url": f"/downloads/{make_key(schedule_xlsx)}"},
                {"name": schedule_json.name, "url": f"/downloads/{make_key(schedule_json)}"},
            ]
            self.send_json({
                "ok": True,
                "log": f"Timetable scheduled successfully with {lesson_count} lessons.",
                "scheduleUrl": "/preview.html",
                "downloads": downloads,
            })
            return

        self.send_json({"ok": False, "log": "Unknown step."}, HTTPStatus.BAD_REQUEST)


def main() -> None:
    if not INPUT_PAGE.is_file():
        raise SystemExit(f"Input page not found: {INPUT_PAGE}")
    server = ThreadingHTTPServer(("127.0.0.1", PORT), Handler)
    threading.Timer(0.6, lambda: webbrowser.open(f"http://127.0.0.1:{PORT}/")).start()
    print(f"ReSched local website: http://127.0.0.1:{PORT}/", flush=True)
    print("Keep this window open while using the website.", flush=True)
    server.serve_forever()


if __name__ == "__main__":
    main()
