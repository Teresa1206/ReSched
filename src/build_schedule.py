"""Build the schedule.html data block from a Step 3 schedule result.

Usage:
    python build_schedule.py <schedule.json> <prepared.xlsx> <input.xlsx> \
        <template.html> <output.html>

The page keeps its existing five-day, seven-period layout.  The generator
fills SCHEDULE_DATA and PERIOD_TIMES and preserves every other part of the
template.
"""

from __future__ import annotations

import json
import re
import sys
from datetime import time
from pathlib import Path

import pandas as pd

DAY_NUMBERS = {
    "monday": 1,
    "mon": 1,
    "tuesday": 2,
    "tue": 2,
    "tues": 2,
    "wednesday": 3,
    "wed": 3,
    "thursday": 4,
    "thu": 4,
    "thur": 4,
    "thurs": 4,
    "friday": 5,
    "fri": 5,
}
PERIOD_COUNT = 7


def clean(value: object) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value).strip()


def records(path: str | Path, sheet: str) -> list[dict[str, object]]:
    frame = pd.read_excel(path, sheet_name=sheet, dtype=object).fillna("")
    return frame.to_dict("records")


def day_number(value: object) -> int | None:
    text = clean(value).lower()
    if text in DAY_NUMBERS:
        return DAY_NUMBERS[text]
    try:
        number = int(float(text))
    except ValueError:
        return None
    return number if 1 <= number <= 5 else None


def period_number(value: object) -> int | None:
    text = clean(value)
    match = re.search(r"(\d+)$", text)
    if not match:
        return None
    number = int(match.group(1))
    return number if 1 <= number <= PERIOD_COUNT else None


def format_time(value: object) -> str:
    if isinstance(value, time):
        return f"{value.hour}:{value.minute:02d}"
    text = clean(value)
    match = re.match(r"^(\d{1,2}):(\d{2})", text)
    if not match:
        return ""
    return f"{int(match.group(1))}:{match.group(2)}"


def build_data(
    schedule_path: str | Path,
    prepared_path: str | Path,
    input_path: str | Path,
) -> tuple[list[dict[str, object]], dict[int, str]]:
    result = json.loads(Path(schedule_path).read_text(encoding="utf-8"))
    meetings = result.get("StudentMeetings", [])
    if not isinstance(meetings, list):
        meetings = []

    section_courses: dict[str, str] = {}
    for row in records(prepared_path, "Sections"):
        section_id = clean(row.get("SectionID"))
        base = clean(row.get("BaseCourseName")) or clean(row.get("CourseName")) or clean(row.get("CourseID"))
        option = clean(row.get("CourseOption"))
        section_courses[section_id] = f"{base} {option}".strip() if option else base

    room_names = {
        clean(row.get("RoomID")): clean(row.get("RoomName")) or clean(row.get("RoomID"))
        for row in records(input_path, "Classrooms")
    }

    students: dict[str, dict[str, object]] = {}
    for meeting in meetings:
        student_id = clean(meeting.get("StudentID"))
        if not student_id:
            continue
        student = students.setdefault(student_id, {
            "name": clean(meeting.get("StudentName")) or student_id,
            "block": clean(meeting.get("StudentBlock")),
            "lessons": [],
        })
        day = day_number(meeting.get("Day"))
        period = period_number(meeting.get("TimeSlotID"))
        if day is None or period is None:
            continue
        section_id = clean(meeting.get("SectionID"))
        course = section_courses.get(section_id) or clean(meeting.get("CourseDisplay")) or section_id
        room_id = clean(meeting.get("RoomID"))
        student["lessons"].append({
            "course": course,
            "teacher": clean(meeting.get("TeacherName")),
            "room": room_names.get(room_id, room_id),
            "day": day,
            "period": period,
        })

    ordered = []
    for student_id in sorted(students):
        student = students[student_id]
        student["lessons"].sort(key=lambda item: (item["day"], item["period"], item["course"]))
        ordered.append(student)

    times: dict[int, str] = {}
    for row in records(input_path, "TimeSlots"):
        period = period_number(row.get("TimeSlotID"))
        if period is None or period in times:
            continue
        start = format_time(row.get("StartTime"))
        end = format_time(row.get("EndTime"))
        if start and end:
            times[period] = f"{start}–{end}"
    return ordered, dict(sorted(times.items()))


def render(students: list[dict[str, object]], times: dict[int, str], template: str) -> str:
    schedule_block = "const SCHEDULE_DATA = " + json.dumps(
        {"students": students}, ensure_ascii=False, indent=2
    ) + ";"

    lines = ["  const PERIOD_TIMES = {"]
    for index, period in enumerate(sorted(times)):
        suffix = "," if index < len(times) - 1 else ""
        lines.append(f'    {period}: {json.dumps(times[period], ensure_ascii=False)}{suffix}')
    lines.append("  };")
    period_block = "\n".join(lines)

    updated, schedule_count = re.subn(
        r"const SCHEDULE_DATA = \{[\s\S]*?\n\};", schedule_block, template, count=1
    )
    if schedule_count != 1:
        raise ValueError("Could not find the SCHEDULE_DATA block in the template.")
    updated, period_count = re.subn(
        r"  const PERIOD_TIMES = \{[\s\S]*?\n  \};", period_block, updated, count=1
    )
    if period_count != 1:
        raise ValueError("Could not find the PERIOD_TIMES block in the template.")
    return updated


def main() -> int:
    if len(sys.argv) != 6:
        print("Usage: python build_schedule.py <schedule.json> <prepared.xlsx> <input.xlsx> <template.html> <output.html>")
        return 2
    schedule_path, prepared_path, input_path, template_path, output_path = sys.argv[1:]
    students, times = build_data(schedule_path, prepared_path, input_path)
    template = Path(template_path).read_text(encoding="utf-8")
    updated = render(students, times, template)
    Path(output_path).write_text(updated, encoding="utf-8")
    print(f"Built {output_path} with {len(students)} students and {sum(len(s['lessons']) for s in students)} lessons.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
