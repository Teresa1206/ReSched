"""Step 3: create a feasible weekly timetable with OR-Tools CP-SAT.

Usage:
    python step3_schedule.py timetable-input_prepared.xlsx

The input is produced by step2_prepare_model.py.  This solver uses the
workbook's time, room, teacher and student-conflict constraints, and decides
each split A/B enrollment during optimization.  Building movement is retained
in the inputs for the next optimization layer.
"""

from __future__ import annotations

import json
import os
import re
import sys
import time
from collections import defaultdict
from pathlib import Path

LOCAL_DEPS = Path(__file__).with_name("_deps")
if LOCAL_DEPS.is_dir():
    sys.path.insert(0, str(LOCAL_DEPS))

import pandas as pd
from ortools.sat.python import cp_model
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]


def clean(value: object) -> str:
    return "" if pd.isna(value) else str(value).strip()


def records(frame: pd.DataFrame) -> list[dict[str, str]]:
    return [{name: clean(value) for name, value in row.items()} for row in frame.to_dict("records")]


def to_int(value: str) -> int:
    return int(float(value))


def slot_number(slot_id: str) -> int:
    match = re.search(r"(\d+)$", slot_id)
    return int(match.group(1)) if match else 999


def clock_seconds(value: str) -> int:
    """Convert the Excel HH:MM time used by TimeSlots into seconds."""
    hours, minutes = clean(value).split(":", 1)
    return int(hours) * 3600 + int(minutes) * 60


def timetable_course_family(row: dict[str, str]) -> str:
    """The course identity shown on an A/B student timetable.

    Levels are teaching sections, not separate student-facing courses: ENG S
    and ENG H both appear as ENG.  Science is the exception because a student
    takes two distinct science subjects, so PHY/CHEM/BIO must remain distinct.
    PE and electives likewise remain their course category here; their options
    are parallel sections of that category.
    """
    course_id = clean(row.get("CourseID", ""))
    if course_id.upper() == "SCI":
        subject = clean(row.get("SelectionGroup", ""))
        if not subject:
            subject = clean(row.get("CourseOption", "")).split(" ")[0]
        return f"SCI-{subject}" if subject else "SCI"
    return course_id or clean(row.get("BaseCourseName", "")) or clean(row.get("CourseName", ""))


def timetable_course_label(row: dict[str, str]) -> str:
    """Course-name label for the six A/B grids; never includes level."""
    family = timetable_course_family(row)
    name = clean(row.get("BaseCourseName", "")) or clean(row.get("CourseName", "")) or clean(row.get("CourseID", ""))
    if family.startswith("SCI-"):
        # Science needs the subject so the two science lessons remain distinguishable.
        return f"{name} {family.removeprefix('SCI-')}".strip()
    return name


HEADER_FILL = PatternFill(fill_type=None)
TITLE_FILL = PatternFill(fill_type=None)
THIN = Side(style="thin", color="000000")


def style_table_sheet(worksheet) -> None:
    """Make compact detail sheets readable without changing their data."""
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions
    worksheet.sheet_view.showGridLines = False
    for cell in worksheet[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    worksheet.row_dimensions[1].height = 28
    for column_cells in worksheet.columns:
        width = max((len(str(cell.value or "")) for cell in column_cells), default=8) + 2
        worksheet.column_dimensions[column_cells[0].column_letter].width = min(max(width, 11), 48)
        for cell in column_cells:
            cell.alignment = Alignment(vertical="top", wrap_text=True)


def write_entity_matrix(
    worksheet,
    rows: list[dict[str, str]],
    entity_key: str,
    title_for,
    text_for,
    times: list[dict[str, str]],
    available_meetings: set[tuple[str, str]],
) -> None:
    """Write all teachers or rooms into one compact weekly timetable table."""
    worksheet.sheet_view.showGridLines = False
    by_entity: dict[str, list[dict[str, str]]] = defaultdict(list)
    for row in rows:
        by_entity[row[entity_key]].append(row)
    periods = [(day, item) for day in DAYS for item in times]
    total_columns = 1 + len(periods)
    worksheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_columns)
    title = worksheet.cell(row=1, column=1, value="全部教师课表" if entity_key == "TeacherID" else "全部教室课表")
    title.font = Font(size=14, bold=True)
    title.alignment = Alignment(horizontal="center", vertical="center")
    worksheet.row_dimensions[1].height = 26

    worksheet.merge_cells(start_row=2, start_column=1, end_row=3, end_column=1)
    entity_header = worksheet.cell(row=2, column=1, value="教师" if entity_key == "TeacherID" else "教室")
    entity_header.font = Font(bold=True)
    entity_header.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    entity_header.border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
    column = 2
    for day in DAYS:
        start_column = column
        for item in times:
            cell = worksheet.cell(row=3, column=column, value=f"{item['TimeSlotID']}\n{item['StartTime']}-{item['EndTime']}")
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
            column += 1
        worksheet.merge_cells(start_row=2, start_column=start_column, end_row=2, end_column=column - 1)
        cell = worksheet.cell(row=2, column=start_column, value=day)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
    worksheet.row_dimensions[2].height = 22
    worksheet.row_dimensions[3].height = 36

    for row_number, entity_id in enumerate(sorted(by_entity), start=4):
        entity_rows = by_entity[entity_id]
        cell = worksheet.cell(row=row_number, column=1, value=title_for(entity_id, entity_rows))
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        cell.border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
        lookup = {(row["Day"], row["TimeSlotID"]): row for row in entity_rows}
        for column, (day, item) in enumerate(periods, start=2):
            lesson = lookup.get((day, item["TimeSlotID"]))
            value = text_for(lesson) if lesson else ("不上课" if (day, item["TimeSlotID"]) not in available_meetings else "")
            cell = worksheet.cell(row=row_number, column=column, value=value)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
        worksheet.row_dimensions[row_number].height = 52
    worksheet.freeze_panes = "B4"
    worksheet.column_dimensions["A"].width = 22
    for column in range(2, total_columns + 1):
        worksheet.column_dimensions[get_column_letter(column)].width = 18


def write_block_grid(
    worksheet,
    rows: list[dict[str, str]],
    block: str,
    type_label: str,
    science_subjects: set[str] | None,
    times: list[dict[str, str]],
    available_meetings: set[tuple[str, str]],
    visible_sections: set[str],
) -> None:
    """Write one A or B course timetable; students match it to their own choices."""
    worksheet.sheet_view.showGridLines = False
    worksheet.merge_cells("A1:F1")
    title = worksheet["A1"]
    title.value = f"{block} 组 · {type_label}（共同课表；同课程的 level / 选项合并显示）"
    title.font = Font(size=14, bold=True)
    title.alignment = Alignment(horizontal="center", vertical="center")
    worksheet.row_dimensions[1].height = 26
    headings = ["时间\nTime", "周一\nMonday", "周二\nTuesday", "周三\nWednesday", "周四\nThursday", "周五\nFriday"]
    for column, heading in enumerate(headings, 1):
        cell = worksheet.cell(row=2, column=column, value=heading)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
    worksheet.row_dimensions[2].height = 34
    lookup: dict[tuple[str, str], list[dict[str, str]]] = defaultdict(list)
    for row in rows:
        if row["SectionID"] not in visible_sections:
            continue
        subject = clean(row.get("SelectionGroup", "")) or clean(row.get("CourseOption", "")).split(" ")[0]
        if science_subjects and row["CourseID"].upper() == "SCI" and subject not in science_subjects:
            continue
        if row["Block"] in {"", block}:
            lookup[(row["Day"], row["TimeSlotID"])].append(row)
    for offset, item in enumerate(times, 3):
        time_cell = worksheet.cell(row=offset, column=1, value=f"{item['TimeSlotID']}\n{item['StartTime']}-{item['EndTime']}")
        time_cell.font = Font(bold=True)
        time_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        time_cell.border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
        for column, day in enumerate(DAYS, 2):
            entries = lookup.get((day, item["TimeSlotID"]), [])
            by_course: dict[str, list[dict[str, str]]] = defaultdict(list)
            for entry in entries:
                by_course[timetable_course_family(entry)].append(entry)
            # This is a student-facing A/B grid, not a section roster.  One
            # cell says only what the student studies; levels, teachers and
            # rooms stay in the detailed sheets.  Thus ENG S + ENG H renders
            # simply as "ENG", never as two courses.
            text = "\n".join(
                timetable_course_label(course_entries[0])
                for _family, course_entries in sorted(by_course.items())
            )
            if not text and (day, item["TimeSlotID"]) not in available_meetings:
                text = "不上课"
            cell = worksheet.cell(row=offset, column=column, value=text)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
        worksheet.row_dimensions[offset].height = max(66, 48 * max(1, max((len(lookup.get((day, item["TimeSlotID"]), [])) for day in DAYS), default=1)))
    worksheet.column_dimensions["A"].width = 18
    for column in range(2, 7):
        worksheet.column_dimensions[get_column_letter(column)].width = 29


def student_enrollment_inputs(tables: dict[str, list[dict[str, str]]]) -> tuple[dict[str, set[str]], list[dict[str, str]]]:
    """Return fixed enrollments plus A/B pairs that CP-SAT must decide."""
    sections = tables["Sections"]
    students = {row["StudentID"]: row for row in tables["SectionStudents"]}
    candidates: dict[str, set[str]] = defaultdict(set)
    for row in tables["SectionStudentCandidates"]:
        candidates[row["SectionID"]].add(row["StudentID"])
    fixed: dict[str, set[str]] = defaultdict(set)
    for row in tables["SectionStudents"]:
        fixed[row["SectionID"]].add(row["StudentID"])

    by_key: dict[tuple[str, str], dict[str, str]] = defaultdict(dict)
    for section in sections:
        if section["Block"] in {"A", "B"}:
            by_key[(section["CourseID"], section["CourseOption"])][section["Block"]] = section["SectionID"]

    split_groups: list[dict[str, str | set[str]]] = []
    for (course_id, option), blocks in by_key.items():
        if set(blocks) != {"A", "B"}:
            raise ValueError("每个 A/B 分班课程小项必须同时具备 A 和 B 两个小节。")
        eligible = sorted(candidates[blocks["A"]] | candidates[blocks["B"]])
        split_groups.append({
            "CourseID": course_id,
            "CourseOption": option,
            "A": blocks["A"],
            "B": blocks["B"],
            "Students": set(eligible),
        })
    return fixed, split_groups


def solve(prepared_path: str | Path) -> tuple[dict[str, pd.DataFrame], list[str]]:
    # One wall-clock allowance for the full model build + solve.  Individual
    # CP-SAT limits alone are misleading because building the model can itself
    # take noticeable time on a Type-heavy input.
    whole_run_started = time.monotonic()
    whole_run_budget_seconds = max(45, min(120, int(os.environ.get("RESCHED_TOTAL_CORE_SECONDS", "75"))))
    path = Path(prepared_path)
    if not path.is_file():
        raise ValueError(f"找不到第二步结果：{path}")
    needed = ["Sections", "SectionStudents", "SectionStudentCandidates", "EligibleTeachers", "CandidateRooms", "AvailableMeetingTimes", "TeacherAvailability", "Students", "HomeroomBlockBalance"]
    workbook = pd.ExcelFile(path)
    missing = [name for name in needed if name not in workbook.sheet_names]
    if missing:
        raise ValueError(f"第二步结果缺少工作表：{', '.join(missing)}")
    tables = {name: records(pd.read_excel(path, sheet_name=name, dtype=object).fillna("")) for name in needed}
    # New Step 2 files include the matrix generated from program1_time_cost's
    # walking model.  Older prepared files remain readable, just without the
    # soft movement objective.
    if "RoomTravel" in workbook.sheet_names:
        tables["RoomTravel"] = records(pd.read_excel(path, sheet_name="RoomTravel", dtype=object).fillna(""))
    sections = tables["Sections"]
    student_homeroom = {row["StudentID"]: row["HomeroomID"] for row in tables["Students"]}
    student_names = {row["StudentID"]: row.get("StudentName", "") for row in tables["Students"]}
    block_limits = {
        row["HomeroomID"]: (to_int(row["MinStudentsPerBlock"]), to_int(row["MaxStudentsPerBlock"]))
        for row in tables["HomeroomBlockBalance"]
    }
    if not sections:
        raise ValueError("Sections 为空，无法排课。")

    fixed_memberships, split_groups = student_enrollment_inputs(tables)
    teachers_by_section: dict[str, list[str]] = defaultdict(list)
    teacher_names: dict[str, str] = {}
    for row in tables["EligibleTeachers"]:
        teachers_by_section[row["SectionID"]].append(row["TeacherID"])
        teacher_names[row["TeacherID"]] = row.get("TeacherName", "")
    rooms_by_section: dict[str, list[str]] = defaultdict(list)
    for row in tables["CandidateRooms"]:
        rooms_by_section[row["SectionID"]].append(row["RoomID"])
    travel_seconds = {
        (row["FromRoomID"], row["ToRoomID"]): int(round(float(row["EstimatedTravelSeconds"])))
        for row in tables.get("RoomTravel", [])
    }
    travel_cost = {
        (row["FromRoomID"], row["ToRoomID"]): int(round(float(row["OptimizationCost"])))
        for row in tables.get("RoomTravel", [])
    }
    availability = {(row["TeacherID"], row["Day"], row["TimeSlotID"]): to_int(row["Available"]) for row in tables["TeacherAvailability"]}
    meetings = sorted(tables["AvailableMeetingTimes"], key=lambda row: (DAYS.index(row["Day"]), slot_number(row["TimeSlotID"])))
    if not meetings:
        raise ValueError("没有可排时段。")

    # Explain definite staffing and parallel-course impossibilities before solving.
    preflight_issues: list[str] = []
    preflight_suggestions: list[str] = []
    diagnostic_rows: list[dict[str, str]] = []
    adjustment_actions: list[dict[str, str]] = []
    # This travels with every adjustment plan.  It distinguishes a confirmed
    # input defect from a merely unfinished CP-SAT search.
    search_info: dict[str, object] = {
        "Status": "未开始搜索",
        "ConfiguredLimitSeconds": 0,
        "ActualSearchSeconds": 0,
        "Rounds": 0,
        "CandidateConflictCuts": 0,
        "Meaning": "已在排课前发现确定的输入硬约束问题；尚未进入 OR-Tools 搜索。",
    }
    section_lookup = {section["SectionID"]: section for section in sections}
    teacher_available_slots: dict[str, list[str]] = defaultdict(list)
    for meeting in meetings:
        for teacher_id in teacher_names:
            if availability.get((teacher_id, meeting["Day"], meeting["TimeSlotID"]), 0) == 1:
                teacher_available_slots[teacher_id].append(f"{meeting['Day']} {meeting['TimeSlotID']}")

    def add_adjustment(
        priority: str, section_id: str, input_sheet: str, action: str,
        current_teachers: list[str] | None = None,
        allow_add_teacher: bool = False,
    ) -> None:
        current_teachers = current_teachers or []
        teacher_text = "；".join(
            f"{teacher_id} {teacher_names.get(teacher_id, '')}：{', '.join(teacher_available_slots.get(teacher_id, [])) or '无可用时段'}"
            for teacher_id in current_teachers
        ) or "当前没有可教教师"
        adjustment_actions.append({
            "Priority": priority,
            "SectionID": section_id,
            "CourseID": section_lookup.get(section_id, {}).get("CourseID", ""),
            # Suggestions are for editing inputs.  A/B and internal Type are
            # solver-only details, so never expose them in the action title.
            "Course": adjustment_course_label(section_lookup.get(section_id, {})) or section_id,
            "InputSheet": input_sheet,
            "Action": action,
            "CurrentTeachersAndAvailableTimes": teacher_text,
            # The UI must not offer “add teacher” merely because a timed
            # search was inconclusive.  It appears only for a measured
            # staffing/availability shortage.
            "AllowAddTeacher": allow_add_teacher,
        })

    def adjustment_course_label(section: dict[str, str]) -> str:
        """Human input label without solver-only A/B or Type suffixes."""
        base = clean(section.get("BaseCourseName", "")) or clean(section.get("CourseName", ""))
        option = clean(section.get("CourseOption", ""))
        homeroom = clean(section.get("HomeroomID", ""))
        block = clean(section.get("Block", ""))
        # Step 2 names split sections as e.g. "ENG H A".  That trailing A/B
        # is an internal scheduling decision, never part of an input-facing
        # course label.
        if block and base.endswith(f" {block}"):
            base = base[: -(len(block) + 1)].strip()
        suffix = option or homeroom
        # CourseName commonly already contains its option ("ENG H").  Do not
        # print it twice in a suggestion or resource report.
        if suffix and (base == suffix or base.endswith(f" {suffix}")):
            return base
        return " ".join(item for item in [base, suffix] if item).strip()

    def resource_analysis() -> list[dict[str, object]]:
        """Report edit-level staffing, including whole-grade parallel cohorts."""
        # A teacher's entered availability is only an upper bound.  First
        # reserve lessons that no other eligible teacher can cover; what
        # remains is still optimistic because shared courses are not deducted.
        sole_load: dict[str, int] = defaultdict(int)
        for section in sections:
            eligible = teachers_by_section[section["SectionID"]]
            if len(eligible) == 1:
                sole_load[eligible[0]] += to_int(section["PeriodsPerWeek"])

        def capacity_for(group_sections: list[dict[str, str]], teachers: list[str]) -> tuple[int, int]:
            group_ids = {section["SectionID"] for section in group_sections}
            raw = sum(len(teacher_available_slots.get(teacher, [])) for teacher in teachers)
            residual = 0
            for teacher in teachers:
                reserved_elsewhere = sum(
                    to_int(section["PeriodsPerWeek"])
                    for section in sections
                    if section["SectionID"] not in group_ids and teachers_by_section[section["SectionID"]] == [teacher]
                )
                residual += max(0, len(teacher_available_slots.get(teacher, [])) - reserved_elsewhere)
            return raw, residual

        def teacher_text(teachers: list[str]) -> str:
            return "；".join(f"{teacher} {teacher_names.get(teacher, '')}".strip() for teacher in teachers) or "无"

        grouped: dict[tuple[str, str, str], list[dict[str, str]]] = defaultdict(list)
        for section in sections:
            grouped[(section["CourseID"], clean(section.get("CourseOption", "")), clean(section.get("HomeroomID", "")))].append(section)
        result: list[dict[str, object]] = []
        for (_course_id, _option, _homeroom), group_sections in grouped.items():
            label = adjustment_course_label(group_sections[0]) or group_sections[0]["SectionID"]
            teachers = sorted({teacher for section in group_sections for teacher in teachers_by_section[section["SectionID"]]})
            demand = sum(to_int(section["PeriodsPerWeek"]) for section in group_sections)
            raw_capacity, residual_capacity = capacity_for(group_sections, teachers)
            same_time = clean(group_sections[0].get("AllGradeSameTime", "")) in {"1", "1.0"}
            if not teachers:
                status = "硬缺口：没有被指定为可教该课程/level 的教师"
            elif demand > residual_capacity:
                status = f"硬缺口：本项共需 {demand} 节；扣除其他唯一任教后，理论余量仅 {residual_capacity} 节"
            else:
                status = "未发现本项单独的教师硬缺口"
            result.append({
                "Scope": "单项",
                "Course": label,
                "RequiredPeriods": demand,
                "ParallelTeachersNeeded": 1,
                "EligibleTeachers": teacher_text(teachers),
                "EligibleTeacherCount": len(teachers),
                "EnteredAvailableSlots": raw_capacity,
                "ResidualCapacity": residual_capacity,
                "Status": status,
            })

        # Whole-grade simultaneous courses must be checked as one cohort.
        # Four electives at the same time are not four independent rows: they
        # require a one-to-one matching of four sections to four teachers.
        cohorts: dict[str, list[dict[str, str]]] = defaultdict(list)
        for section in sections:
            if clean(section.get("AllGradeSameTime", "")) in {"1", "1.0"}:
                cohorts[section["CourseID"]].append(section)
        for course_id, cohort in cohorts.items():
            is_cas = clean(cohort[0].get("HomeroomID", "")) != ""
            teachers = sorted({teacher for section in cohort for teacher in teachers_by_section[section["SectionID"]]})
            raw_capacity, residual_capacity = capacity_for(cohort, teachers)
            assignments: dict[str, str] = {}
            def assign(section_id: str, seen: set[str]) -> bool:
                for teacher in teachers_by_section[section_id]:
                    if teacher in seen:
                        continue
                    seen.add(teacher)
                    if teacher not in assignments or assign(assignments[teacher], seen):
                        assignments[teacher] = section_id
                        return True
                return False
            matched = sum(assign(section["SectionID"], set()) for section in cohort)
            members = "、".join(dict.fromkeys(adjustment_course_label(section) or section["SectionID"] for section in cohort))
            base = adjustment_course_label(cohort[0]).split(" ")[0] or course_id
            if matched < len(cohort):
                unmatched = [adjustment_course_label(section) or section["SectionID"] for section in cohort if section["SectionID"] not in set(assignments.values())]
                status = f"硬缺口：同一时间需 {len(cohort)} 位不同教师，目前最多可匹配 {matched} 位；未覆盖：{'、'.join(unmatched) or '请检查任教关系'}"
            else:
                status = (
                    f"CAS 行政班分配满足：{len(cohort)} 个行政班均由各自班主任在同一时间承担"
                    if is_cas else
                    f"并行任教关系满足：{len(cohort)} 个小项可由 {len(cohort)} 位不同教师同时承担"
                )
            result.append({
                "Scope": "CAS 行政班固定并行" if is_cas else "全年级统一并行组",
                "Course": f"{base}（各行政班班主任同时上课）" if is_cas else f"{base}（全年级统一时间）",
                "RequiredPeriods": max(to_int(section["PeriodsPerWeek"]) for section in cohort),
                "ParallelTeachersNeeded": len(cohort),
                "EligibleTeachers": teacher_text(teachers),
                "EligibleTeacherCount": len(teachers),
                "EnteredAvailableSlots": raw_capacity,
                "ResidualCapacity": residual_capacity,
                "Status": status,
                "Members": members,
            })
        return sorted(result, key=lambda row: ("硬缺口" not in str(row["Status"]), str(row["Course"])))

    def write_adjustment_plan(summary: str) -> Path:
        """Persist concrete failure advice so the local page can render it."""
        output = path.with_name(f"{path.stem}_adjustment_plan.json")
        output.write_text(json.dumps({
            "Summary": summary,
            "SearchInfo": search_info,
            "ResourceAnalysis": resource_analysis(),
            "Actions": adjustment_actions,
            "TeacherAvailability": [
                {
                    "TeacherID": teacher_id,
                    "TeacherName": teacher_names.get(teacher_id, ""),
                    "AvailableTimes": teacher_available_slots.get(teacher_id, []),
                }
                for teacher_id in sorted(teacher_names)
            ],
        }, ensure_ascii=False, indent=2), encoding="utf-8")
        return output

    def add_search_space_adjustments() -> None:
        """Give actionable advice even when CP-SAT times out before a conflict is cut.

        A timeout is not a proof of infeasibility, but it must never leave the
        adjustment panel empty.  Rank the tightest real course sections using
        the entered teacher availability and candidate-room counts.
        """
        if len(adjustment_actions) >= 5:
            return
        ranked: list[tuple[float, str, int, int, int]] = []
        for section in sections:
            section_id = section["SectionID"]
            teachers = teachers_by_section[section_id]
            available_times = {
                (meeting["Day"], meeting["TimeSlotID"])
                for meeting in meetings
                if any(availability.get((teacher_id, meeting["Day"], meeting["TimeSlotID"]), 0) == 1 for teacher_id in teachers)
            }
            required = to_int(section["PeriodsPerWeek"])
            room_count = len(rooms_by_section[section_id])
            # Higher score = more constrained: fewer eligible people, fewer
            # available times per required lesson, and fewer rooms.
            score = (20 if not teachers else 8 / len(teachers)) + required / max(1, len(available_times)) + 3 / max(1, room_count)
            ranked.append((score, section_id, len(teachers), len(available_times), room_count))
        seen_courses: set[str] = {f"{item.get('CourseID', '')}|" for item in adjustment_actions}
        for _score, section_id, teacher_count, time_count, room_count in sorted(ranked, reverse=True):
            section = section_lookup[section_id]
            required = to_int(section["PeriodsPerWeek"])
            # A fixed sports field, art room, or CAS homeroom is normal.  One
            # teacher + one room is sufficient for one course option.  Do not
            # turn a merely narrow-but-adequate input into a false “add a
            # teacher / room” action card.  Whole-grade parallel resource
            # shortages are already checked separately as a cohort above.
            if teacher_count > 0 and time_count >= required and room_count > 0:
                continue
            course_key = f"{section['CourseID']}|{section.get('CourseOption', '')}"
            if course_key in seen_courses:
                continue
            seen_courses.add(course_key)
            teacher_ids = "、".join(teachers_by_section[section_id]) or "无"
            teacher_phrase = "没有可教教师" if teacher_count == 0 else f"当前可教教师为 {teacher_ids}（共 {teacher_count} 位），合计覆盖 {time_count} 个可排时段"
            room_phrase = "没有候选教室" if room_count == 0 else f"可用候选教室 {room_count} 间"
            # Safe default for the “resources are adequate but search timed
            # out” case.  It is a status note, not an instruction to add a
            # teacher.
            input_sheet = "Courses / TeacherAssignments / TeacherAvailability"
            allow_add_teacher = False
            if teacher_count == 0:
                next_step = "这是已确认的教师任教缺口：在教师编辑页为至少 1 位现有或新增教师勾选该课程/level。"
                input_sheet = "TeacherAssignments"
                allow_add_teacher = True
            elif time_count < required:
                next_step = (
                    f"现有教师的时段还差至少 {required - time_count} 节：先把下列教师补充为 Available=1，"
                    "只有无法增加时段时，才需要另一位可教该课程/level 的教师。"
                )
                input_sheet = "TeacherAvailability / TeacherAssignments"
                allow_add_teacher = True
            elif room_count == 0:
                next_step = "这是已确认的场地缺口：请取消过严的地点限定，或在课程小项中加入至少一间候选教室。"
                input_sheet = "CourseOptions"
                allow_add_teacher = False
            else:
                # Reached only for an actual measured deficit above.
                next_step = "请按已定位的输入缺口修改后再运行。"
            add_adjustment(
                "检查结果", section_id, input_sheet,
                f"“{adjustment_course_label(section) or section_id}”每周需 {required} 节；{teacher_phrase}；{room_phrase}。"
                + next_step,
                teachers_by_section[section_id],
                allow_add_teacher,
            )
            if len(adjustment_actions) >= 5:
                break
    # Science subjects remain useful display information.  The actual schedule
    # Type now comes from Step 2's combined eligible-course selections, not
    # from Science alone.
    science_subjects_by_student: dict[str, set[str]] = defaultdict(set)
    for section_id, student_ids in fixed_memberships.items():
        section = section_lookup[section_id]
        if section["CourseID"].upper() != "SCI":
            continue
        subject = clean(section.get("SelectionGroup", "")) or clean(section["CourseOption"]).split(" ")[0]
        for student_id in student_ids:
            science_subjects_by_student[student_id].add(subject)
    prepared_student_types = {
        row["StudentID"]: clean(row.get("ScheduleType", ""))
        for row in tables["Students"]
    }
    prepared_type_descriptions = {
        row["StudentID"]: clean(row.get("TypeDescription", ""))
        for row in tables["Students"]
    }
    student_atype = {
        student_id: prepared_student_types.get(student_id) or "T1"
        for student_id in student_homeroom
    }
    for section in sections:
        section_id = section["SectionID"]
        if not teachers_by_section[section_id]:
            preflight_issues.append(f"{section_id} 没有可教教师")
            preflight_suggestions.append(f"在教师任教表中为 {section_id} 指定至少一位教师。")
            add_adjustment("必须", section_id, "TeacherAssignments", f"新增至少 1 位可教“{adjustment_course_label(section) or section_id}”的教师，并在 TeacherAssignments 中勾选该课程/level。", allow_add_teacher=True)
        if not rooms_by_section[section_id]:
            preflight_issues.append(f"{section_id} 没有可用教室")
            preflight_suggestions.append(f"为 {section_id} 增加候选教室，或检查限定楼/教室。")
        available_times = {
            (meeting["Day"], meeting["TimeSlotID"])
            for meeting in meetings
            if any(availability.get((teacher_id, meeting["Day"], meeting["TimeSlotID"]), 0) == 1 for teacher_id in teachers_by_section[section_id])
        }
        required = to_int(section["PeriodsPerWeek"])
        if teachers_by_section[section_id] and len(available_times) < required:
            preflight_issues.append(f"{section_id} 每周需 {required} 节，但其可教教师合计仅覆盖 {len(available_times)} 个可排时段")
            preflight_suggestions.append(f"增加 {section_id} 的教师可用时段，或增加另一位可教教师。")
            add_adjustment(
                "必须", section_id, "TeacherAvailability / TeacherAssignments",
                f"该小节还差至少 {required - len(available_times)} 个不同可排时段。优先把下列现有教师在本周其他时段改为 Available=1；若不能增加时段，新增 1 位可教该课的教师。",
                teachers_by_section[section_id],
                True,
            )

    compulsory_load: dict[str, int] = defaultdict(int)
    sole_teacher_sections: dict[str, list[dict[str, str]]] = defaultdict(list)
    for section in sections:
        eligible = teachers_by_section[section["SectionID"]]
        if len(eligible) == 1:
            compulsory_load[eligible[0]] += to_int(section["PeriodsPerWeek"])
            sole_teacher_sections[eligible[0]].append(section)
    available_load: dict[str, int] = defaultdict(int)
    for (teacher_id, _day, _slot), available in availability.items():
        available_load[teacher_id] += available
    for teacher_id, load in compulsory_load.items():
        available = available_load[teacher_id]
        if load <= available:
            continue
        excess = load - available
        name = teacher_names.get(teacher_id, "")
        display_name = f"{teacher_id}（{name}）" if name else teacher_id
        preflight_issues.append(f"{display_name} 被唯一指定承担 {load} 节，但只有 {available} 个可用时段，至少超出 {excess} 节")
        candidates = sorted(sole_teacher_sections[teacher_id], key=lambda section: to_int(section["PeriodsPerWeek"]))
        released, move_items = 0, []
        for section in candidates:
            released += to_int(section["PeriodsPerWeek"])
            move_items.append(adjustment_course_label(section) or section["SectionID"])
            if released >= excess:
                break
        preflight_suggestions.append(
            f"要保留现有时段，至少把 {', '.join(move_items)}（共 {released} 节）改由另一位教师任教；"
            f"或者把 {display_name} 的可用时段至少增加 {excess} 节。"
        )
        for section in candidates:
            add_adjustment(
                "必须", section["SectionID"], "TeacherAssignments / TeacherAvailability",
                f"{display_name} 的唯一承担量超过可用时段。为此小节增加另一位可教教师，或把 {display_name} 至少增加 {excess} 节可用时间。",
                [teacher_id],
                True,
            )

    # Each student's total requested periods cannot exceed the week's usable slots.
    student_sections: dict[str, set[str]] = defaultdict(set)
    pressure_labels: dict[str, str] = {
        section_id: section["CourseName"] or section_id for section_id, section in section_lookup.items()
    }
    pressure_is_split: dict[str, bool] = {section_id: False for section_id in section_lookup}
    for section_id, student_ids in fixed_memberships.items():
        for student_id in student_ids:
            student_sections[student_id].add(section_id)
    for group in split_groups:
        representative = f"{group['A']}/{group['B']}"
        section_a = section_lookup[str(group["A"])]
        pressure_labels[representative] = f"{section_a['CourseID']} {section_a['CourseOption']}（A/B）"
        pressure_is_split[representative] = True
        for student_id in group["Students"]:
            student_sections[student_id].add(representative)
    for student_id, section_ids in student_sections.items():
        load = sum(
            to_int(section_lookup[section_id]["PeriodsPerWeek"])
            if section_id in section_lookup else to_int(section_lookup[section_id.split("/")[0]]["PeriodsPerWeek"])
            for section_id in section_ids
        )
        if load > len(meetings):
            preflight_issues.append(f"学生 {student_id} 需要 {load} 节课，但一周仅有 {len(meetings)} 个可排时段")
            preflight_suggestions.append("减少该学生的一门课程或课程节数，或增加全校可排时段。")

    # Surface the actual section pairs that cannot overlap for the largest groups of students.
    pair_pressure: dict[tuple[str, str], int] = defaultdict(int)
    for section_ids in student_sections.values():
        ordered = sorted(section_ids)
        for left_index, left in enumerate(ordered):
            for right in ordered[left_index + 1:]:
                pair_pressure[(left, right)] += 1
    for (left, right), shared in sorted(pair_pressure.items(), key=lambda item: item[1], reverse=True)[:8]:
        if shared < 2:
            continue
        left_split, right_split = pressure_is_split.get(left, False), pressure_is_split.get(right, False)
        if left_split and right_split:
            detail = (
                f"{pressure_labels.get(left, left)} 与 {pressure_labels.get(right, right)}：共有 {shared} 名学生同时选中。"
                "它们在同一全局 block 内不能同一时段（A 对 A、B 对 B）；A 对 B 仍可并行。"
            )
        elif left_split or right_split:
            split_label = pressure_labels.get(left if left_split else right, left if left_split else right)
            fixed_label = pressure_labels.get(right if left_split else left, right if left_split else left)
            detail = (
                f"{fixed_label} 与 {split_label}：共有 {shared} 名学生同时选中。若两者重叠，"
                f"这些学生必须全部被分到 {split_label} 的相反 block，因而会显著压缩 A/B 分班空间。"
            )
        else:
            detail = f"{pressure_labels.get(left, left)} 与 {pressure_labels.get(right, right)}：共有 {shared} 名学生同时选中，两个固定小节不能同一时段。"
        diagnostic_rows.append({
            "Category": "学生冲突压力",
            "Detail": detail,
            "Suggestion": "这是排课器自动处理的冲突强度，不是错误。若无解，优先增加其中一门课的教师可用时段或取消不必要的统一时间。",
        })

    sections_by_course: dict[str, list[dict[str, str]]] = defaultdict(list)
    for section in sections:
        sections_by_course[section["CourseID"]].append(section)
    for course_id, course_sections in sections_by_course.items():
        if clean(course_sections[0]["AllGradeSameTime"]) not in {"1", "1.0"}:
            continue
        teacher_union = {teacher for section in course_sections for teacher in teachers_by_section[section["SectionID"]]}
        room_union = {room for section in course_sections for room in rooms_by_section[section["SectionID"]]}
        if len(teacher_union) < len(course_sections):
            labels = "、".join(adjustment_course_label(section) or section["SectionID"] for section in course_sections)
            current_teachers = "、".join(f"{teacher}{'（'+teacher_names.get(teacher,'')+'）' if teacher_names.get(teacher,'') else ''}" for teacher in sorted(teacher_union)) or "无"
            missing_teachers = len(course_sections) - len(teacher_union)
            preflight_issues.append(f"{course_id} 要求全年级同一时间：{labels} 必须同时上课，但目前仅 {current_teachers} 可教，缺少至少 {missing_teachers} 位并行教师")
            preflight_suggestions.append(f"若保留 {course_id} 的统一时间，请给其中 {missing_teachers} 个小节分别增加不同教师（共需至少 {len(course_sections)} 位）；否则取消“全年级统一时间”。")
            for section in course_sections:
                add_adjustment(
                    "必须", section["SectionID"], "TeacherAssignments / Courses",
                    f"“全年级统一时间”要求 {len(course_sections)} 个小节同时上课。当前至少缺 {missing_teachers} 位并行教师；为该小节新增可教教师，或在 Courses 取消 AllGradeSameTime。",
                    teachers_by_section[section["SectionID"]],
                    True,
                )
        if len(room_union) < len(course_sections):
            missing_rooms = len(course_sections) - len(room_union)
            preflight_issues.append(f"{course_id} 要求全年级同一时间，但 {len(course_sections)} 个并行小节只有 {len(room_union)} 间候选教室，缺少至少 {missing_rooms} 间")
            preflight_suggestions.append(f"为 {course_id} 增加至少 {missing_rooms} 间可用教室，或取消“全年级统一时间”。")

    # A required consecutive pair needs one teacher who is available for adjacent slots.
    for section in sections:
        if clean(section["RequireConsecutive"]) not in {"1", "1.0"} or to_int(section["PeriodsPerWeek"]) != 2:
            continue
        has_pair = False
        for day in DAYS:
            slots = sorted([meeting["TimeSlotID"] for meeting in meetings if meeting["Day"] == day], key=slot_number)
            for first, second in zip(slots, slots[1:]):
                if slot_number(second) != slot_number(first) + 1:
                    continue
                if any(availability.get((teacher, day, first), 0) and availability.get((teacher, day, second), 0) for teacher in teachers_by_section[section["SectionID"]]):
                    has_pair = True
        if not has_pair:
            preflight_issues.append(f"{section['SectionID']} 必须连堂，但没有同一教师可用的相邻时段")
            preflight_suggestions.append(f"增加 {section['SectionID']} 教师的连续可用时段，或取消连堂规则。")
            add_adjustment(
                "必须", section["SectionID"], "TeacherAvailability / Courses",
                "至少把一位可教教师的相邻两节课设为 Available=1；若课程不必连堂，可在 Courses 取消 RequireConsecutive。",
                teachers_by_section[section["SectionID"]],
            )

    ab_students_by_homeroom: dict[str, set[str]] = defaultdict(set)
    for group in split_groups:
        for student_id in group["Students"]:
            ab_students_by_homeroom[student_homeroom.get(student_id, "")].add(student_id)
    for homeroom_id, student_ids in ab_students_by_homeroom.items():
        min_a, max_a = block_limits.get(homeroom_id, (len(student_ids) // 2, (len(student_ids) + 1) // 2))
        if min_a > max_a or min_a < 0 or max_a > len(student_ids):
            preflight_issues.append(f"{homeroom_id} 的 A 组人数范围是 {min_a}–{max_a}，但可参与 A/B 分班的学生只有 {len(student_ids)} 人")
            preflight_suggestions.append(f"把 {homeroom_id} 的 A 组人数下限/上限改为 0–{len(student_ids)} 的有效范围；通常可设为 {len(student_ids)//2}–{(len(student_ids)+1)//2}。")
    if preflight_issues:
        suggestions = list(dict.fromkeys(preflight_suggestions + ["检查后重新导出输入，依次运行第一步、第二步、第三步。"] ))
        plan_path = write_adjustment_plan("输入存在可在排课前确定的硬性不足；请按“调整方案”逐项修改原始输入后重新运行。")
        pressure_report = ""
        if diagnostic_rows:
            pressure_report = "\n\n学生冲突压力（仅用于解释；排课器会自动处理，不需要逐条修改）：\n- " + "\n- ".join(row["Detail"] for row in diagnostic_rows[:5])
        raise ValueError("不可行原因：\n- " + "\n- ".join(preflight_issues) + "\n\n建议：\n- " + "\n- ".join(suggestions) + pressure_report + f"\n\n已生成可操作的调整方案：{plan_path.name}")

    model = cp_model.CpModel()
    variables: dict[tuple[str, int, str, str], cp_model.IntVar] = {}
    time_vars: dict[tuple[str, int], list[cp_model.IntVar]] = defaultdict(list)
    room_time: dict[tuple[str, int], list[cp_model.IntVar]] = defaultdict(list)
    teacher_time: dict[tuple[str, int], list[cp_model.IntVar]] = defaultdict(list)
    teacher_selected: dict[tuple[str, str], cp_model.IntVar] = {}
    room_selected: dict[tuple[str, str], cp_model.IntVar] = {}

    for section in sections:
        section_id = section["SectionID"]
        teachers = teachers_by_section[section_id]
        rooms = rooms_by_section[section_id]
        if not teachers:
            raise ValueError(f"{section_id} 没有可教教师。")
        if not rooms:
            raise ValueError(f"{section_id} 没有可用教室。")
        selected_teachers = []
        for teacher_id in teachers:
            chosen = model.NewBoolVar(f"teacher_{section_id}_{teacher_id}")
            teacher_selected[(section_id, teacher_id)] = chosen
            selected_teachers.append(chosen)
        model.Add(sum(selected_teachers) == 1)
        selected_rooms = []
        for room_id in rooms:
            chosen = model.NewBoolVar(f"room_{section_id}_{room_id}")
            room_selected[(section_id, room_id)] = chosen
            selected_rooms.append(chosen)
        model.Add(sum(selected_rooms) == 1)
        # A consecutive two-period section can only use slots belonging to an
        # adjacent pair for which at least one *same* teacher is available in
        # both periods.  Remove every other time before x variables are made.
        consecutive_domain: set[int] | None = None
        if clean(section.get("RequireConsecutive", "")) in {"1", "1.0"} and to_int(section["PeriodsPerWeek"]) == 2:
            consecutive_domain = set()
            for first_index in range(len(meetings) - 1):
                second_index = first_index + 1
                first, second = meetings[first_index], meetings[second_index]
                if first["Day"] != second["Day"] or slot_number(second["TimeSlotID"]) != slot_number(first["TimeSlotID"]) + 1:
                    continue
                if any(
                    availability.get((teacher_id, first["Day"], first["TimeSlotID"]), 0) == 1
                    and availability.get((teacher_id, second["Day"], second["TimeSlotID"]), 0) == 1
                    for teacher_id in teachers
                ):
                    consecutive_domain.update({first_index, second_index})
        for time_index, meeting in enumerate(meetings):
            if consecutive_domain is not None and time_index not in consecutive_domain:
                continue
            for teacher_id in teachers:
                if availability.get((teacher_id, meeting["Day"], meeting["TimeSlotID"]), 0) != 1:
                    continue
                for room_id in rooms:
                    variable = model.NewBoolVar(f"x_{section_id}_{time_index}_{teacher_id}_{room_id}")
                    variables[(section_id, time_index, teacher_id, room_id)] = variable
                    model.Add(variable <= teacher_selected[(section_id, teacher_id)])
                    model.Add(variable <= room_selected[(section_id, room_id)])
                    time_vars[(section_id, time_index)].append(variable)
                    room_time[(room_id, time_index)].append(variable)
                    teacher_time[(teacher_id, time_index)].append(variable)

    # A whole-grade simultaneous course is a parallel cohort.  Reserve
    # different teachers and rooms for its different options/classes before
    # time search starts.  This is logically implied by their shared meeting
    # time, but expressing it here prunes impossible teacher/room combinations
    # immediately instead of discovering them after a time assignment.
    sections_by_simultaneous_course: dict[str, list[dict[str, str]]] = defaultdict(list)
    for section in sections:
        if clean(section.get("AllGradeSameTime", "")) in {"1", "1.0"}:
            sections_by_simultaneous_course[section["CourseID"]].append(section)
    for cohort in sections_by_simultaneous_course.values():
        for left_index, left in enumerate(cohort):
            for right in cohort[left_index + 1:]:
                left_id, right_id = left["SectionID"], right["SectionID"]
                for teacher_id in set(teachers_by_section[left_id]) & set(teachers_by_section[right_id]):
                    model.Add(teacher_selected[(left_id, teacher_id)] + teacher_selected[(right_id, teacher_id)] <= 1)
                for room_id in set(rooms_by_section[left_id]) & set(rooms_by_section[right_id]):
                    model.Add(room_selected[(left_id, room_id)] + room_selected[(right_id, room_id)] <= 1)

    section_time_used: dict[tuple[str, int], cp_model.IntVar] = {}
    for section in sections:
        section_id = section["SectionID"]
        required = to_int(section["PeriodsPerWeek"])
        all_variables: list[cp_model.IntVar] = []
        for time_index in range(len(meetings)):
            options = time_vars[(section_id, time_index)]
            if options:
                used = model.NewBoolVar(f"used_{section_id}_{time_index}")
                model.Add(sum(options) == used)
                section_time_used[(section_id, time_index)] = used
                all_variables.extend(options)
            else:
                section_time_used[(section_id, time_index)] = model.NewConstant(0)
        model.Add(sum(all_variables) == required)
        if not all_variables:
            raise ValueError(f"{section_id} 没有教师可用的排课组合。")

    for options in room_time.values():
        model.Add(sum(options) <= 1)
    for options in teacher_time.values():
        model.Add(sum(options) <= 1)

    # Teacher assignment is fixed for a whole course section, so use those
    # choices to measure weekly teaching load.  This is deliberately a soft
    # second-phase target: it never prevents a feasible timetable, but avoids
    # needlessly giving all work to the first eligible teacher.
    teacher_weekly_load: dict[str, cp_model.LinearExpr] = {}
    total_required_periods = sum(to_int(section["PeriodsPerWeek"]) for section in sections)
    max_teacher_load = model.NewIntVar(0, total_required_periods, "max_teacher_weekly_load")
    for teacher_id in teacher_names:
        terms = [
            to_int(section["PeriodsPerWeek"]) * teacher_selected[(section["SectionID"], teacher_id)]
            for section in sections
            if (section["SectionID"], teacher_id) in teacher_selected
        ]
        teacher_weekly_load[teacher_id] = sum(terms) if terms else 0
        model.Add(max_teacher_load >= teacher_weekly_load[teacher_id])

    # One A/B decision belongs to each student, not to each course.  A student
    # assigned to A remains in A for every course whose SplitAB setting is on.
    enrollment: dict[tuple[str, str], cp_model.IntVar] = {}
    ab_students = sorted({student_id for group in split_groups for student_id in group["Students"]})
    student_in_a: dict[str, cp_model.IntVar] = {
        student_id: model.NewBoolVar(f"global_block_A_{student_id}") for student_id in ab_students
    }
    students_by_homeroom: dict[str, list[str]] = defaultdict(list)
    for student_id in ab_students:
        students_by_homeroom[student_homeroom.get(student_id, "")].append(student_id)
    for homeroom_id, student_ids in students_by_homeroom.items():
        min_a, max_a = block_limits.get(homeroom_id, (len(student_ids) // 2, (len(student_ids) + 1) // 2))
        if min_a > max_a or min_a < 0 or max_a > len(student_ids):
            raise ValueError(f"{homeroom_id} 的 A/B 分班人数范围无效：{min_a}–{max_a}，可分学生数为 {len(student_ids)}。")
        model.Add(sum(student_in_a[student_id] for student_id in student_ids) >= min_a)
        model.Add(sum(student_in_a[student_id] for student_id in student_ids) <= max_a)
        # A/B labels are interchangeable when the permitted range is
        # symmetric.  Fixing one stable representative in A removes the
        # mirror image where every A/B student is swapped, without changing
        # any meaningful timetable choice.
        if min_a + max_a == len(student_ids) and max_a > 0:
            model.Add(student_in_a[min(student_ids)] == 1)

    for group in split_groups:
        section_a, section_b = str(group["A"]), str(group["B"])
        for student_id in group["Students"]:
            enrollment[(student_id, section_a)] = student_in_a[student_id]
            enrollment[(student_id, section_b)] = student_in_a[student_id].Not()

    # Student conflicts are represented once, at Block + Type timetable level
    # below.  Every student in the same Type has the same selected course
    # families; global and A/B courses are included in every applicable Type.
    # Building per-student attendance literals here therefore duplicated the
    # exact same rule thousands of times.  The exported result is still
    # independently checked student by student before it is accepted.

    # Do not create travel variables yet.  Their teacher × room × adjacent-slot
    # cross-product is large, especially after adding realistic multi-subject
    # teacher capability.  Phase 1 must only build what it needs to find a
    # valid timetable; travel variables are appended only after that succeeds.

    # All options of a whole-grade simultaneous course share exactly the same meeting times.
    by_course: dict[str, list[dict[str, str]]] = defaultdict(list)
    for section in sections:
        by_course[section["CourseID"]].append(section)
    for course_sections in by_course.values():
        if clean(course_sections[0]["AllGradeSameTime"]) not in {"1", "1.0"}:
            continue
        anchor = course_sections[0]["SectionID"]
        for section in course_sections[1:]:
            for time_index in range(len(meetings)):
                model.Add(section_time_used[(section["SectionID"], time_index)] == section_time_used[(anchor, time_index)])

    # A SelectionGroup represents one student-facing subject with several
    # teaching levels.  For example, CHEM S and CHEM H are different sections
    # but the same Chemistry lesson band, so they must happen simultaneously.
    # The signature keeps A/B and Type timetables separate while synchronizing
    # only the levels that belong to the same subject within that timetable.
    by_subject_band: dict[tuple[str, str, str, str], list[dict[str, str]]] = defaultdict(list)
    for section in sections:
        subject = clean(section.get("SelectionGroup", ""))
        if not subject:
            continue
        by_subject_band[(
            clean(section.get("CourseID", "")),
            clean(section.get("Block", "")),
            clean(section.get("Type", "")),
            subject,
        )].append(section)
    for subject_sections in by_subject_band.values():
        if len(subject_sections) < 2:
            continue
        anchor = subject_sections[0]["SectionID"]
        for section in subject_sections[1:]:
            for time_index in range(len(meetings)):
                model.Add(section_time_used[(section["SectionID"], time_index)] == section_time_used[(anchor, time_index)])

    # Two-period consecutive courses (for example CAS) must use adjacent slots on one day.
    for section in sections:
        if clean(section["RequireConsecutive"]) not in {"1", "1.0"}:
            continue
        if to_int(section["PeriodsPerWeek"]) != 2:
            raise ValueError(f"{section['SectionID']} 的连堂规则目前只支持每周 2 节。")
        pairs: list[cp_model.IntVar] = []
        for first in range(len(meetings) - 1):
            second = first + 1
            same_day = meetings[first]["Day"] == meetings[second]["Day"]
            adjacent = slot_number(meetings[second]["TimeSlotID"]) == slot_number(meetings[first]["TimeSlotID"]) + 1
            if not (same_day and adjacent):
                continue
            pair = model.NewBoolVar(f"consecutive_{section['SectionID']}_{first}")
            model.Add(pair <= section_time_used[(section["SectionID"], first)])
            model.Add(pair <= section_time_used[(section["SectionID"], second)])
            pairs.append(pair)
        if not pairs:
            raise ValueError(f"{section['SectionID']} 没有可用的相邻时段。")
        model.Add(sum(pairs) == 1)

    # Do not invent a hard “one meeting per day” rule for ordinary courses.
    # The input only makes consecutive lessons mandatory when explicitly
    # marked.  Forbidding same-day repeats was shrinking the feasible search
    # space without user authorization and causing otherwise valid timetables
    # to be reported as unschedulable.  Distribution remains a result-quality
    # hint in the independent evaluator, not a condition for feasibility.

    # Build each Block + Type student timetable directly into the model.
    # Earlier versions discovered these clashes only after a complete solve,
    # added a few cuts, and solved again.  Here every applicable course family
    # is known in advance: a global course applies to every Type, an A/B course
    # applies to its block, and a Type course applies only to its own Type.
    # Levels/options in the same student-facing family remain allowed together.
    type_blocks = ["A", "B"] if any(clean(section.get("Block", "")) for section in sections) else [""]
    type_labels_for_model = sorted(set(student_atype.values())) or [""]
    upfront_type_constraints = 0
    for block in type_blocks:
        for type_label in type_labels_for_model:
            by_family: dict[str, list[str]] = defaultdict(list)
            for section in sections:
                section_block = clean(section.get("Block", ""))
                section_type = clean(section.get("Type", ""))
                if section_block and section_block != block:
                    continue
                if section_type and section_type != type_label:
                    continue
                by_family[timetable_course_family(section)].append(section["SectionID"])
            # SelectionGroup equality above makes every level in a family use
            # the same time.  One representative is therefore sufficient.
            representatives = [sorted(section_ids)[0] for section_ids in by_family.values()]
            for time_index in range(len(meetings)):
                model.Add(sum(section_time_used[(section_id, time_index)] for section_id in representatives) <= 1)
                upfront_type_constraints += 1

    # Let CP-SAT's portfolio choose the branching order.  A hand-written
    # course-first order looked reasonable, but on type-heavy inputs it can
    # commit to a broad elective/Type decision too early and starve the more
    # constrained parallel cohorts.  The solver's native portfolio is more
    # reliable for finding *any* feasible timetable first.

    # Retained only as an independent post-solve verifier / diagnostics source.
    # The actual Block + Type rule is now fully present above before the first
    # CP-SAT solve.
    def type_conflict_analysis(solved: cp_model.CpSolver) -> tuple[set[tuple[str, str, int]], list[str]]:
        current_sections: dict[str, set[str]] = defaultdict(set)
        for section_id, student_ids in fixed_memberships.items():
            for student_id in student_ids:
                current_sections[student_id].add(section_id)
        for group in split_groups:
            for student_id in group["Students"]:
                block = "A" if solved.Value(student_in_a[student_id]) else "B"
                current_sections[student_id].add(str(group[block]))
        by_type_period: dict[tuple[str, str, int], dict[str, set[str]]] = defaultdict(lambda: defaultdict(set))
        for student_id, section_ids in current_sections.items():
            if student_id not in student_in_a:
                continue
            block = "A" if solved.Value(student_in_a[student_id]) else "B"
            type_label = student_atype.get(student_id, "Type 未识别")
            for section_id in section_ids:
                for time_index in range(len(meetings)):
                    if solved.Value(section_time_used[(section_id, time_index)]):
                        by_type_period[(block, type_label, time_index)][section_id].add(student_id)
        resources: dict[tuple[str, int], tuple[str, str]] = {}
        for (section_id, time_index, teacher_id, room_id), variable in variables.items():
            if solved.Value(variable):
                resources[(section_id, time_index)] = (teacher_id, room_id)
        cuts: set[tuple[str, str, int]] = set()
        reports: list[str] = []
        for (block, type_label, time_index), section_students in by_type_period.items():
            ordered = sorted(section_students)
            course_families = sorted({timetable_course_family(section_lookup[section_id]) for section_id in ordered})
            if len(course_families) > 1:
                details = []
                for section_id in ordered:
                    section = section_lookup[section_id]
                    teacher_id, room_id = resources.get((section_id, time_index), ("未知教师", "未知教室"))
                    students = sorted(section_students[section_id])
                    sample_students = "、".join(
                        f"{student_id} {student_names.get(student_id, '')}".strip()
                        for student_id in students[:4]
                    )
                    more = f" 等 {len(students)} 人" if len(students) > 4 else ""
                    details.append(
                        f"{section['CourseName']}（{section_id}；{teacher_id} {teacher_names.get(teacher_id, '')}；{room_id}；受影响：{sample_students}{more}）"
                    )
                reports.append(
                    f"{block} {type_label} {meetings[time_index]['Day']} {meetings[time_index]['TimeSlotID']}："
                    + "；".join(details)
                )
            for left_index, left in enumerate(ordered):
                for right in ordered[left_index + 1:]:
                    if timetable_course_family(section_lookup[left]) != timetable_course_family(section_lookup[right]):
                        cuts.add((left, right, time_index))
        return cuts, reports

    def type_conflict_explanation(solved: cp_model.CpSolver, left: str, right: str) -> tuple[str, bool]:
        """State whether a Type collision is caused by raw timetable volume.

        A collision in a candidate is not automatically a staffing deficit.
        Calculate the minimum number of course-family periods required by the
        affected student timetable and compare it with the actual usable
        weekly slots before suggesting any input change.
        """
        sections_by_student: dict[str, set[str]] = defaultdict(set)
        for section_id, student_ids in fixed_memberships.items():
            for student_id in student_ids:
                sections_by_student[student_id].add(section_id)
        for group in split_groups:
            for student_id in group["Students"]:
                block = "A" if solved.Value(student_in_a[student_id]) else "B"
                sections_by_student[student_id].add(str(group[block]))
        # A Type collision does not require that one *individual* student has
        # both named sections: the shared timetable can contain one course
        # needed by one member and another course needed by another member.
        # Therefore calculate the lower bound from the whole Block + Type
        # group, exactly as the Type timetable itself is defined.
        sections_by_type: dict[tuple[str, str], set[str]] = defaultdict(set)
        for student_id, section_ids in sections_by_student.items():
            if student_id not in student_in_a:
                continue
            block = "A" if solved.Value(student_in_a[student_id]) else "B"
            type_label = student_atype.get(student_id, "Type 未识别")
            sections_by_type[(block, type_label)].update(section_ids)

        candidates: list[tuple[str, str, int, int]] = []
        for (block, type_label), section_ids in sections_by_type.items():
            if left not in section_ids or right not in section_ids:
                continue
            family_periods: dict[str, int] = {}
            for section_id in section_ids:
                section = section_lookup[section_id]
                family = timetable_course_family(section)
                family_periods[family] = max(
                    family_periods.get(family, 0), to_int(section["PeriodsPerWeek"])
                )
            candidates.append((block, type_label, sum(family_periods.values()), len(meetings)))

        if not candidates:
            return (
                "当前候选中出现不同课程同一时段，但尚未能定位对应的 Block + Type 课表；"
                "这不能据此判断教师或总课时不足。",
                False,
            )

        # Explain the tightest affected timetable, not an arbitrary one.
        block, type_label, required, available = max(candidates, key=lambda item: item[2] - item[3])
        group_name = f"{block} block／{type_label}"
        if required > available:
            return (
                f"{group_name} 这张学生课表最低需要 {required} 节，但一周只有 {available} 个可排时段，"
                f"确定少 {required - available} 节；这是总课时超过可排时段的硬缺口。",
                True,
            )
        return (
            f"{group_name} 这张学生课表最低需要 {required} 节，一周可排 {available} 节，"
            f"尚有 {available - required} 节余量；总课时没有超过可排时段。"
            "本次只是课程时段组合尚未排开，并不能说明教师人数或教师可用时间不足。",
            False,
        )

    applied_type_cuts: set[tuple[str, str, int]] = set()
    last_type_reports: list[str] = []
    solver: cp_model.CpSolver | None = None
    status = cp_model.UNKNOWN
    # Type clashes are hard constraints from the start, so phase 1 requires
    # one bounded feasibility solve rather than a repair-and-resolve cycle.
    requested_search_seconds = max(30, min(120, int(os.environ.get("RESCHED_SEARCH_SECONDS", "75"))))
    # Feasibility is the priority.  Keep only a short optional travel pass so
    # it cannot take a quarter of the only search budget away from a timetable
    # that has not been found yet.
    reserved_movement_seconds = max(3, min(15, int(os.environ.get("RESCHED_MOVE_OPTIMIZE_SECONDS", "5"))))
    remaining_for_search = whole_run_budget_seconds - (time.monotonic() - whole_run_started) - reserved_movement_seconds
    search_budget_seconds = max(1, min(requested_search_seconds, int(remaining_for_search)))
    search_started = time.monotonic()
    type_deadline = search_started + search_budget_seconds
    rounds_attempted = 0
    for _round in range(1):
        remaining_seconds = type_deadline - time.monotonic()
        if remaining_seconds <= 0.5:
            status = cp_model.UNKNOWN
            break
        solver = cp_model.CpSolver()
        solver.parameters.max_time_in_seconds = max(1, remaining_seconds)
        solver.parameters.num_search_workers = 8
        solver.parameters.random_seed = _round + 1
        rounds_attempted += 1
        status = solver.Solve(model)
        if status not in {cp_model.OPTIMAL, cp_model.FEASIBLE}:
            # UNKNOWN only means this round did not finish.  It is not an
            # infeasibility proof and must consume the remaining search budget
            # before we offer input changes.
            if status == cp_model.UNKNOWN and time.monotonic() < type_deadline - 0.5:
                continue
            break
        _all_cuts, last_type_reports = type_conflict_analysis(solver)
        break
    actual_search_seconds = round(time.monotonic() - search_started, 1)
    if status == cp_model.UNKNOWN:
        search_meaning = "已达到搜索时间/轮次上限，尚未证明输入不可行；不要仅凭此结果增加教师或修改课程。"
        search_status = "搜索未完成"
    elif status == cp_model.INFEASIBLE:
        search_meaning = "OR-Tools 已证明当前模型不可行；调整方案中的“必须”项目才是需要修改的硬性缺口。"
        search_status = "已证明不可行"
    elif status == cp_model.MODEL_INVALID:
        search_meaning = "模型无效；这不是输入资源不足的结论。"
        search_status = "模型无效"
    else:
        search_meaning = "已找到候选解，并继续进行结果复核。"
        search_status = "已找到候选解"
    search_info.update({
        "Status": search_status,
        "ConfiguredLimitSeconds": search_budget_seconds,
        "ActualSearchSeconds": actual_search_seconds,
        "Rounds": rounds_attempted,
        "CandidateConflictCuts": len(applied_type_cuts),
        "Meaning": search_meaning,
    })
    if solver is None or status not in {cp_model.OPTIMAL, cp_model.FEASIBLE}:
        if status == cp_model.UNKNOWN:
            for left, right, _time_index in sorted(applied_type_cuts)[:8]:
                explanation, _is_volume_short = type_conflict_explanation(solver, left, right)
                add_adjustment(
                    "搜索冲突（未证明输入缺口）", left, "Courses / TeacherAvailability / CourseOptions",
                    f"“{adjustment_course_label(section_lookup[left])}”与“{adjustment_course_label(section_lookup[right])}”在候选课表中仍重叠。{explanation} 请先继续搜索或检查不必要的统一时间、连堂和地点限定；不要仅凭这一项增加教师。",
                    teachers_by_section[left],
                )
            add_search_space_adjustments()
            timeout_summary = (
                "搜索在限定时间内未能消除已定位的课程时间冲突；以下按课程列出可直接扩充的教师、时段或场地资源。"
                if applied_type_cuts else
                "搜索在限定时间内尚未得到足以定位单一冲突的候选课表；以下按现有教师、可用时段和候选教室最紧张的课程，给出可直接修改的输入项。"
            )
            write_adjustment_plan(timeout_summary)
            conflict_detail = "\n- " + "\n- ".join(last_type_reports[:8]) if last_type_reports else "\n- 尚未取得可用于定位的候选冲突。"
            raise ValueError(
                f"排课尚未完成：已在 {search_budget_seconds} 秒内搜索，仍未找到同时消除以下实际冲突的解；这不等于已证明输入不可行。\n\n"
                "当前需拆开的具体冲突（来自上一轮候选课表）："
                f"{conflict_detail}\n\n"
                "说明：这仍不是已证明的输入不可行。每张调整卡会说明该 Type 课表的最低课时与可排时段；"
                "只有明确显示总课时不足或资源硬缺口时，才需要修改输入。"
            )
        if status == cp_model.MODEL_INVALID:
            raise ValueError("排课模型无效：请重新运行第一步、第二步后再试；若仍出现，请保留 Excel 和错误信息以便检查模型。")
        for left, right, _time_index in sorted(applied_type_cuts)[:8]:
            explanation, is_volume_short = type_conflict_explanation(solver, left, right)
            add_adjustment(
                "必须：总课时不足" if is_volume_short else "冲突定位中",
                left, "Courses / TeacherAvailability / CourseOptions",
                f"“{adjustment_course_label(section_lookup[left])}”与“{adjustment_course_label(section_lookup[right])}”无法错开。{explanation}",
                teachers_by_section[left],
            )
        add_search_space_adjustments()
        write_adjustment_plan("当前输入被证明无法满足所有硬约束；请优先完成以下具体修改后再运行。")
        conflict_detail = "\n- " + "\n- ".join(last_type_reports[:8]) if last_type_reports else ""
        raise ValueError(
            "确定不可行：在现有输入下，无法同时满足课程冲突与其他硬约束。\n\n"
            "导致当前搜索分支冲突的课程："
            f"{conflict_detail}\n\n"
            "建议：\n"
            "- 在上述课程中为其中一门增加可教教师、教师可用时段或候选教室。\n"
            "- 若 CAS / 自选课需要在不同小项并行，保留它们同 CourseID 的合并显示；不要把它们拆成不同课程类别。\n"
            "- 增加全校可排时段，或把学生划分为更细的 Type。"
        )
    remaining_type_cuts, remaining_type_reports = type_conflict_analysis(solver)
    if remaining_type_cuts:
        for left, right, _time_index in sorted(remaining_type_cuts)[:8]:
            explanation, is_volume_short = type_conflict_explanation(solver, left, right)
            add_adjustment(
                "必须：总课时不足" if is_volume_short else "结果复核冲突",
                left, "Courses / TeacherAvailability / CourseOptions",
                f"结果复核发现“{adjustment_course_label(section_lookup[left])}”与“{adjustment_course_label(section_lookup[right])}”仍在同一时段出现。{explanation}"
                + (" 请增加全校可排时段，或减少该 Type 课表的课程节数。" if is_volume_short else " 这不是增加教师的依据；请先检查统一时间、连堂、地点限定，或让排课器继续搜索。"),
                teachers_by_section[left],
            )
        add_search_space_adjustments()
        write_adjustment_plan("结果复核拒绝导出：以下课程仍造成同一时段的跨课程并行。")
        detail = "\n- " + "\n- ".join(remaining_type_reports[:8])
        raise ValueError(
            "排课结果复核失败：仍存在同一时段的跨课程并行，已拒绝导出该结果。冲突示例："
            f"{detail}\n\n"
            "每一张调整卡都已列出该 Type 课表的最低课时和可排时段。\n"
            "- 若卡片显示“总课时不足”：增加全校可排时段，或减少该 Type 课表的课程节数。\n"
            "- 若卡片显示“总课时没有超过”：这不是教师短缺；先检查不必要的统一时间、连堂、地点限定，或继续搜索。\n"
            "- 只有资源分析另行显示教师/教室硬缺口时，才按对应的教师或教室建议修改输入。"
        )

    # Phase 2: improve only the teacher room transitions, after all hard
    # constraints and Type conflicts have already passed.  A timeout here
    # keeps the valid phase-1 solution; it never turns a usable timetable into
    # a failure merely because the nicer room arrangement took too long.
    movement_optimization_seconds = 0
    movement_optimized = False
    remaining_for_movement = whole_run_budget_seconds - (time.monotonic() - whole_run_started)
    if remaining_for_movement >= 3:
        movement_optimization_seconds = min(reserved_movement_seconds, int(remaining_for_movement))
        soft_travel_terms: list[cp_model.LinearExpr] = []
        if travel_seconds:
            # Add penalties only for transitions that *actually occur* in the
            # first feasible timetable.  This makes travel a real soft
            # objective while avoiding the enormous all-rooms × all-rooms
            # cross-product that previously stalled large instances.
            current_by_teacher_day: dict[tuple[str, str], list[tuple[int, str, cp_model.IntVar]]] = defaultdict(list)
            for (_section_id, time_index, teacher_id, room_id), variable in variables.items():
                if solver.Value(variable) == 1:
                    current_by_teacher_day[(teacher_id, meetings[time_index]["Day"])].append((time_index, room_id, variable))
            for (teacher_id, day), lessons in current_by_teacher_day.items():
                lessons.sort(key=lambda item: slot_number(meetings[item[0]]["TimeSlotID"]))
                for (first_index, first_room, first_var), (second_index, second_room, second_var) in zip(lessons, lessons[1:]):
                    if slot_number(meetings[second_index]["TimeSlotID"]) != slot_number(meetings[first_index]["TimeSlotID"]) + 1:
                        continue
                    gap = max(0, clock_seconds(meetings[second_index]["StartTime"]) - clock_seconds(meetings[first_index]["EndTime"]))
                    required = travel_seconds.get((first_room, second_room), 0)
                    shortfall = max(0, required - gap)
                    if not shortfall:
                        continue
                    active = model.NewBoolVar(f"observed_move_{teacher_id}_{first_index}_{second_index}_{first_room}_{second_room}")
                    model.Add(active <= first_var)
                    model.Add(active <= second_var)
                    model.Add(active >= first_var + second_var - 1)
                    # The penalty includes both total distance and the missing
                    # seconds, so the next solve favours a genuinely easier
                    # movement path when it can do so without breaking rules.
                    soft_travel_terms.append(int(travel_cost.get((first_room, second_room), required) + shortfall * 10) * active)
        # A one-period lower maximum teaching load is worth more than any
        # travel tie-breaker.  For larger instances travel is evaluated after
        # scheduling instead of expanding a second enormous CP-SAT model.
        model.Minimize(max_teacher_load * 100000 + sum(soft_travel_terms))
        movement_solver = cp_model.CpSolver()
        movement_solver.parameters.max_time_in_seconds = movement_optimization_seconds
        movement_solver.parameters.num_search_workers = 8
        movement_solver.parameters.random_seed = 101
        movement_status = movement_solver.Solve(model)
        if movement_status in {cp_model.OPTIMAL, cp_model.FEASIBLE}:
            solver = movement_solver
            movement_optimized = True
        else:
            # Keep the already-verified phase-1 solver and its valid solution.
            # Movement quality is optional; feasibility is not.
            pass

    search_info["MovementOptimizationSeconds"] = movement_optimization_seconds
    search_info["MovementOptimized"] = movement_optimized

    block_assignments = [
        {
            "StudentID": student_id,
            "HomeroomID": student_homeroom.get(student_id, ""),
            "Block": "A" if solver.Value(student_in_a[student_id]) else "B",
            "AType": student_atype.get(student_id, "Type 未识别"),
            "TypeDescription": prepared_type_descriptions.get(student_id, ""),
            "ScienceSubjects": " + ".join(sorted(science_subjects_by_student.get(student_id, set()))),
        }
        for student_id in ab_students
    ]
    assigned_block = {row["StudentID"]: row["Block"] for row in block_assignments}
    schedule: list[dict[str, str]] = []
    for (section_id, time_index, teacher_id, room_id), variable in variables.items():
        if solver.Value(variable) != 1:
            continue
        meeting = meetings[time_index]
        section = next(row for row in sections if row["SectionID"] == section_id)
        schedule.append({
            "SectionID": section_id, "CourseID": section["CourseID"], "CourseOption": section["CourseOption"], "SelectionGroup": section.get("SelectionGroup", ""), "Block": section["Block"], "BaseCourseName": section.get("BaseCourseName", section["CourseID"]),
            "CourseName": section["CourseName"], "Day": meeting["Day"],
            "TimeSlotID": meeting["TimeSlotID"], "StartTime": meeting["StartTime"], "EndTime": meeting["EndTime"],
            "RoomID": room_id, "TeacherID": teacher_id, "TeacherName": teacher_names.get(teacher_id, ""),
        })
    schedule.sort(key=lambda row: (DAYS.index(row["Day"]), slot_number(row["TimeSlotID"]), row["SectionID"]))
    schedule_by_section: dict[str, list[dict[str, str]]] = defaultdict(list)
    for row in schedule:
        schedule_by_section[row["SectionID"]].append(row)
    course_summary = []
    for section in sections:
        section_id = section["SectionID"]
        meetings_for_section = schedule_by_section[section_id]
        course_summary.append({
            "SectionID": section_id,
            "CourseName": section["CourseName"],
            "Block": section["Block"],
            "PeriodsPerWeek": section["PeriodsPerWeek"],
            "Teacher": " / ".join(sorted({f"{row['TeacherID']} {row['TeacherName']}".strip() for row in meetings_for_section})),
            "Room": " / ".join(sorted({row["RoomID"] for row in meetings_for_section})),
            "Meetings": "；".join(f"{row['Day']} {row['TimeSlotID']} ({row['StartTime']}-{row['EndTime']})" for row in meetings_for_section),
        })

    all_student_sections = [
        {"StudentID": student_id, "SectionID": section_id, "Block": section_lookup[section_id]["Block"], "HomeroomID": student_homeroom.get(student_id, "")}
        for section_id, student_ids in fixed_memberships.items() for student_id in student_ids
    ]
    for group in split_groups:
        for student_id in group["Students"]:
            block = assigned_block[student_id]
            all_student_sections.append({
                "StudentID": student_id,
                "SectionID": str(group[block]),
                "Block": block,
                "HomeroomID": student_homeroom.get(student_id, ""),
            })
    type_sections: dict[tuple[str, str], set[str]] = defaultdict(set)
    for assignment in all_student_sections:
        student_id = assignment["StudentID"]
        block = assigned_block.get(student_id, "")
        type_label = student_atype.get(student_id, "Type 未识别")
        if block in {"A", "B"}:
            type_sections[(block, type_label)].add(assignment["SectionID"])
    type_section_membership = [
        {"Block": block, "AType": type_label, "SectionID": section_id}
        for (block, type_label), section_ids in sorted(type_sections.items())
        for section_id in sorted(section_ids)
    ]
    student_timetable = []
    for assignment in all_student_sections:
        section = section_lookup[assignment["SectionID"]]
        meetings_for_section = schedule_by_section[assignment["SectionID"]]
        student_timetable.append({
            "StudentID": assignment["StudentID"], "HomeroomID": assignment["HomeroomID"],
            "StudentBlock": assigned_block.get(assignment["StudentID"], ""),
            "SectionBlock": assignment["Block"], "SectionID": assignment["SectionID"],
            "CourseName": section["CourseName"], "Teacher": " / ".join(sorted({f"{row['TeacherID']} {row['TeacherName']}".strip() for row in meetings_for_section})),
            "Room": " / ".join(sorted({row["RoomID"] for row in meetings_for_section})),
            "Meetings": "；".join(f"{row['Day']} {row['TimeSlotID']}" for row in meetings_for_section),
        })
    student_timetable.sort(key=lambda row: (row["StudentID"], row["CourseName"], row["SectionID"]))
    student_a = [row for row in student_timetable if assigned_block.get(row["StudentID"]) == "A"]
    student_b = [row for row in student_timetable if assigned_block.get(row["StudentID"]) == "B"]

    student_meetings = []
    for assignment in all_student_sections:
        for meeting in schedule_by_section[assignment["SectionID"]]:
            student_meetings.append({
                "StudentID": assignment["StudentID"], "StudentName": student_names.get(assignment["StudentID"], ""),
                "StudentBlock": assigned_block.get(assignment["StudentID"], ""), "HomeroomID": assignment["HomeroomID"],
                "Day": meeting["Day"], "TimeSlotID": meeting["TimeSlotID"], "StartTime": meeting["StartTime"], "EndTime": meeting["EndTime"],
                "CourseDisplay": (meeting["BaseCourseName"] or meeting["CourseID"]) + (f" {assignment['Block']}" if assignment["Block"] in {"A", "B"} else ""),
                "TeacherName": meeting["TeacherName"], "RoomID": meeting["RoomID"], "SectionID": meeting["SectionID"],
            })

    # Report the actual movement outcome separately from feasibility.  A
    # shortfall is visible to the user but never changes an audit pass into a
    # hard failure, because movement is intentionally a soft preference.
    travel_review: list[dict[str, str | int]] = []
    def collect_travel(entity_kind: str, entity_key: str, rows: list[dict[str, str]]) -> None:
        by_day: dict[str, list[dict[str, str]]] = defaultdict(list)
        for row in rows:
            by_day[row["Day"]].append(row)
        for day, day_rows in by_day.items():
            ordered = sorted(day_rows, key=lambda row: slot_number(row["TimeSlotID"]))
            for first, second in zip(ordered, ordered[1:]):
                if slot_number(second["TimeSlotID"]) != slot_number(first["TimeSlotID"]) + 1:
                    continue
                required = travel_seconds.get((first["RoomID"], second["RoomID"]), 0)
                gap = max(0, clock_seconds(second["StartTime"]) - clock_seconds(first["EndTime"]))
                travel_review.append({
                    "EntityType": entity_kind, "EntityID": entity_key, "Day": day,
                    "FromPeriod": first["TimeSlotID"], "FromRoomID": first["RoomID"],
                    "ToPeriod": second["TimeSlotID"], "ToRoomID": second["RoomID"],
                    "BreakSeconds": gap, "EstimatedTravelSeconds": required,
                    "ShortfallSeconds": max(0, required - gap),
                    "OptimizationCost": travel_cost.get((first["RoomID"], second["RoomID"]), required),
                })
    for teacher_id in teacher_names:
        collect_travel("教师", teacher_id, [row for row in schedule if row["TeacherID"] == teacher_id])
    for student_id in student_homeroom:
        collect_travel("学生", student_id, [row for row in student_meetings if row["StudentID"] == student_id])
    travel_shortfalls = [row for row in travel_review if int(row["ShortfallSeconds"]) > 0]

    # A feasible schedule can still leave free periods.  This is not a missing
    # course assignment unless the input explicitly requires a full timetable,
    # but users need the exact gap in order to decide whether to add courses.
    scheduled_periods_by_student: dict[str, int] = defaultdict(int)
    for meeting in student_meetings:
        scheduled_periods_by_student[meeting["StudentID"]] += 1
    available_periods_per_student = len(meetings)
    student_load = []
    for student_id in sorted(student_homeroom):
        scheduled_periods = scheduled_periods_by_student[student_id]
        free_periods = max(0, available_periods_per_student - scheduled_periods)
        student_load.append({
            "StudentID": student_id,
            "StudentName": student_names.get(student_id, ""),
            "HomeroomID": student_homeroom.get(student_id, ""),
            "Block": assigned_block.get(student_id, ""),
            "AType": student_atype.get(student_id, "Type 未识别"),
            "AvailablePeriods": available_periods_per_student,
            "ScheduledPeriods": scheduled_periods,
            "FreePeriods": free_periods,
            "AdditionalPeriodsNeededForFullTimetable": free_periods,
        })
    load_by_student = {row["StudentID"]: row for row in student_load}
    for assignment in block_assignments:
        load = load_by_student.get(assignment["StudentID"], {})
        assignment["AvailablePeriods"] = load.get("AvailablePeriods", available_periods_per_student)
        assignment["ScheduledPeriods"] = load.get("ScheduledPeriods", 0)
        assignment["FreePeriods"] = load.get("FreePeriods", available_periods_per_student)
        assignment["AdditionalPeriodsNeededForFullTimetable"] = load.get("AdditionalPeriodsNeededForFullTimetable", available_periods_per_student)

    teacher_timetable = sorted(schedule, key=lambda row: (row["TeacherID"], DAYS.index(row["Day"]), slot_number(row["TimeSlotID"])))
    room_timetable = sorted(schedule, key=lambda row: (row["RoomID"], DAYS.index(row["Day"]), slot_number(row["TimeSlotID"])))
    teacher_load = []
    for teacher_id in sorted(teacher_names):
        taught = sum(1 for row in schedule if row["TeacherID"] == teacher_id)
        available = available_load.get(teacher_id, 0)
        teacher_load.append({"TeacherID": teacher_id, "TeacherName": teacher_names[teacher_id], "ScheduledPeriods": taught, "AvailablePeriods": available, "FreePeriods": available - taught})

    balance_counts: dict[str, dict[str, int]] = defaultdict(lambda: {"A": 0, "B": 0})
    for assignment in block_assignments:
        balance_counts[assignment["HomeroomID"]][assignment["Block"]] += 1
    max_balance_gap = max((abs(counts["A"] - counts["B"]) for counts in balance_counts.values()), default=0)
    def audit_row(rule: str, passed: bool, detail: str) -> dict[str, str]:
        return {"Rule": rule, "Status": "通过" if passed else "未通过", "Detail": detail}
    section_counts = {section_id: len(rows) for section_id, rows in schedule_by_section.items()}
    allowed_meeting_keys = {(meeting["Day"], meeting["TimeSlotID"]) for meeting in meetings}
    schedule_uses_allowed_times = all((row["Day"], row["TimeSlotID"]) in allowed_meeting_keys for row in schedule)
    teacher_is_eligible = all(row["TeacherID"] in teachers_by_section[row["SectionID"]] for row in schedule)
    room_is_eligible = all(row["RoomID"] in rooms_by_section[row["SectionID"]] for row in schedule)
    teacher_is_available = all(availability.get((row["TeacherID"], row["Day"], row["TimeSlotID"]), 0) == 1 for row in schedule)
    enrolled_sections_by_student: dict[str, set[str]] = defaultdict(set)
    for assignment in all_student_sections:
        enrolled_sections_by_student[assignment["StudentID"]].add(assignment["SectionID"])
    fixed_enrollment_ok = all(section_id in enrolled_sections_by_student[student_id] for section_id, students in fixed_memberships.items() for student_id in students)
    split_enrollment_ok = all(
        str(group["A"]) in enrolled_sections_by_student[student_id] or str(group["B"]) in enrolled_sections_by_student[student_id]
        for group in split_groups for student_id in group["Students"]
    )
    block_limits_ok = all(
        block_limits.get(homeroom_id, (0, len(counts)))[0] <= counts["A"] <= block_limits.get(homeroom_id, (0, len(counts)))[1]
        for homeroom_id, counts in balance_counts.items()
    )
    students_with_free_periods = [row for row in student_load if row["FreePeriods"] > 0]
    average_free_periods = (
        sum(row["FreePeriods"] for row in student_load) / len(student_load)
        if student_load else 0
    )
    max_free_student = max(student_load, key=lambda row: (row["FreePeriods"], row["StudentID"]), default=None)
    simultaneous_ok = True
    for course_sections in by_course.values():
        if clean(course_sections[0]["AllGradeSameTime"]) not in {"1", "1.0"}:
            continue
        expected = {(row["Day"], row["TimeSlotID"]) for row in schedule_by_section[course_sections[0]["SectionID"]]}
        simultaneous_ok = simultaneous_ok and all({(row["Day"], row["TimeSlotID"]) for row in schedule_by_section[section["SectionID"]]} == expected for section in course_sections[1:])
    consecutive_ok = True
    for section in sections:
        if clean(section["RequireConsecutive"]) not in {"1", "1.0"}:
            continue
        rows = schedule_by_section[section["SectionID"]]
        consecutive_ok = consecutive_ok and len(rows) == 2 and rows[0]["Day"] == rows[1]["Day"] and abs(slot_number(rows[0]["TimeSlotID"]) - slot_number(rows[1]["TimeSlotID"])) == 1
    type_period_courses: dict[tuple[str, str, str, str], set[str]] = defaultdict(set)
    for assignment in all_student_sections:
        student_id = assignment["StudentID"]
        block = assigned_block.get(student_id, "")
        type_label = student_atype.get(student_id, "Type 未识别")
        for meeting in schedule_by_section[assignment["SectionID"]]:
            type_period_courses[(block, type_label, meeting["Day"], meeting["TimeSlotID"])].add(
                timetable_course_family(meeting)
            )
    type_period_conflicts = {
        key: sorted(course_ids) for key, course_ids in type_period_courses.items()
        if len(course_ids) > 1
    }
    type_conflict_detail = "每一个 A/B-Type 课表的同一格最多出现一个实际被该组学生选到的课程类别；ENG/MATH 等课程的 S/H level 会合并为同一门课，SCI 的 PHY/CHEM/BIO 保持为不同课程。"
    if type_period_conflicts:
        examples = [
            f"{block} {type_label} {day} {slot}: {', '.join(course_ids)}"
            for (block, type_label, day, slot), course_ids in list(sorted(type_period_conflicts.items()))[:5]
        ]
        type_conflict_detail = "发现 Type 并行冲突：" + "；".join(examples)
    audit_rows = [
        audit_row("课程节数", all(section_counts.get(section["SectionID"], 0) == to_int(section["PeriodsPerWeek"]) for section in sections), "每个课程小节的实际上课节数与输入要求一致。"),
        audit_row("全校不可上课时段", schedule_uses_allowed_times, "没有课程排进 NoClassTimes 指定的全校不可上课时段。"),
        audit_row("教师不冲突", len({(row["TeacherID"], row["Day"], row["TimeSlotID"]) for row in schedule}) == len(schedule), "同一教师在同一时段最多上一节课。"),
        audit_row("教师任教范围", teacher_is_eligible, "每一节课均由 TeacherAssignments 中指定可教该小节的教师承担。"),
        audit_row("教师可用时间", teacher_is_available, "没有课程落在 TeacherAvailability 标记为不可用的时段。"),
        audit_row("教室不冲突", len({(row["RoomID"], row["Day"], row["TimeSlotID"]) for row in schedule}) == len(schedule), "同一教室在同一时段最多安排一节课。"),
        audit_row("教室适用范围", room_is_eligible, "每一节课均使用 CandidateRooms 中允许的教室，并遵守课程地点限定。"),
        audit_row("学生不冲突", len({(row["StudentID"], row["Day"], row["TimeSlotID"]) for row in student_meetings}) == len(student_meetings), "每位学生在同一时段最多上一节课。"),
        audit_row("Block + Type 无跨课程并行", not type_period_conflicts, type_conflict_detail),
        audit_row("学生选课对应", fixed_enrollment_ok and split_enrollment_ok, "每位学生都已进入其固定选课小节，A/B 课程则恰好进入同一全局 block 的对应小节。"),
        audit_row("固定教师与教室", all(len({row["TeacherID"] for row in rows}) == 1 and len({row["RoomID"] for row in rows}) == 1 for rows in schedule_by_section.values()), "同一课程小节固定一位教师和一间教室。"),
        audit_row("全局 A/B 分班", all(assigned_block.get(row["StudentID"], "") == row["Block"] for row in block_assignments), "每位学生在所有 A/B 课程中保持同一 block。"),
        audit_row("行政班 A/B 人数范围", block_limits_ok, "每个 Homeroom 的 A 组人数均落在 HomeroomBlockBalance 输入的下限与上限内。"),
        audit_row("全年级统一时间", simultaneous_ok, "勾选统一时间的课程，其所有小节使用完全相同的时段。"),
        audit_row("必须连堂", consecutive_ok, "勾选连堂的课程安排在同一天相邻节次。"),
        {
            "Rule": "非连堂课程分散",
            "Status": "提示",
            "Detail": "输入没有要求非连堂课程必须跨天；同日重复不会使课表无效。可按教学偏好在结果分析中查看。",
        },
        audit_row("周课表时间轴", True, "每一张周课表的左侧时间轴只显示一组 P1–P7；全校停课格会明确写为“不上课”。"),
        {
            "Rule": "学生满课缺口",
            "Status": "提示",
            "Detail": (
                f"{len(students_with_free_periods)}/{len(student_load)} 名学生尚未排满。"
                f"按每人 {available_periods_per_student} 个可上课时段计算，平均还可增加 {average_free_periods:.1f} 节；"
                f"最大缺口为 {max_free_student['StudentID']} 的 {max_free_student['FreePeriods']} 节。"
            ) if max_free_student else "没有学生数据。",
        },
        {
            "Rule": "楼间移动时间（软约束）",
            "Status": ("已优化" if movement_optimized else "已复算") if travel_seconds else "未覆盖",
            "Detail": (
                f"{'已在第二阶段优化教师相邻课移动' if movement_optimized else '基础可行解已优先导出，教师移动优化在时间上限内未完成'}；"
                f"结果页另行复算了 {len(travel_review)} 次师生相邻课转换，其中 {len(travel_shortfalls)} 次课间不足。"
                "通行时间不会使课表判为无效。"
                if travel_seconds else
                "第二步结果没有 RoomTravel 工作表；请重新运行第二步后再进行移动时间软优化。"
            ),
        },
    ]
    failed_audits = [row["Rule"] for row in audit_rows if row["Status"] == "未通过"]
    uncovered_audits = [row["Rule"] for row in audit_rows if row["Status"] == "未覆盖"]
    busiest_teacher = max(teacher_load, key=lambda row: (row["ScheduledPeriods"], row["TeacherID"]), default=None)
    busiest_room_count = max((sum(1 for row in schedule if row["RoomID"] == room_id), room_id) for room_id in {row["RoomID"] for row in schedule}) if schedule else (0, "—")
    review_rows = []
    if failed_audits:
        review_rows.append({
            "Category": "必须处理",
            "Finding": f"合规检查未通过：{'、'.join(failed_audits)}。此结果不应作为正式课表使用。",
            "Input adjustment for a better result": "打开“合规检查”工作表查看逐条失败原因；修正输入后重新运行第三步。",
        })
    else:
        passed_audit_count = sum(row["Status"] == "通过" for row in audit_rows)
        coverage_note = f"；尚未覆盖：{'、'.join(uncovered_audits)}" if uncovered_audits else ""
        review_rows.append({
            "Category": "合规结果",
            "Finding": f"本次共 {len(schedule)} 节课；{passed_audit_count} 项已建模规则均通过复核，包括不可上课时段、教师可用时间、课程地点、学生选课与 A/B 人数范围{coverage_note}。",
            "Input adjustment for a better result": "这是一份可用的基线结果；修改任何输入后应重新运行三步并以新的合规检查为准。",
        })
    if busiest_teacher:
        ratio = (busiest_teacher["ScheduledPeriods"] / busiest_teacher["AvailablePeriods"]) if busiest_teacher["AvailablePeriods"] else 0
        review_rows.append({
            "Category": "教师负荷",
            "Finding": f"负荷最高的是 {busiest_teacher['TeacherID']} {busiest_teacher['TeacherName']}：已排 {busiest_teacher['ScheduledPeriods']} 节 / 可用 {busiest_teacher['AvailablePeriods']} 时段（{ratio:.0%}）。",
            "Input adjustment for a better result": "若希望降低该教师压力，为其部分课程增加另一位可教教师，或增加其可用时段。",
        })
    if max_free_student:
        review_rows.append({
            "Category": "学生满课缺口",
            "Finding": f"每位学生每周有 {available_periods_per_student} 个可上课时段；当前 {len(students_with_free_periods)}/{len(student_load)} 名学生未排满，平均还缺 {average_free_periods:.1f} 节。{max_free_student['StudentID']} 缺口最大，尚可增加 {max_free_student['FreePeriods']} 节。",
            "Input adjustment for a better result": "如学校要求学生满课，在 Courses 中增加课程/节数，或在学生选课中增加可选课程；详见“学生课时”工作表的每位学生缺口。",
        })
    review_rows.append({
        "Category": "教室使用",
        "Finding": f"使用最频繁的教室是 {busiest_room_count[1]}，共 {busiest_room_count[0]} 节。每节课均在课程允许的候选教室中安排。",
        "Input adjustment for a better result": "若需均衡场地使用，为该课程增加可替代教室；如场地必须固定，保持课程小项的地点限定。",
    })
    review_rows.append({
        "Category": "A/B 分班",
        "Finding": f"A/B 由求解器按行政班一次决定，最大人数差为 {max_balance_gap}；学生 Type 由符合条件课程的实际选课组合决定。",
        "Input adjustment for a better result": "通过 HomeroomBlockBalance 调整每个行政班 A 组人数下限/上限；不要把学生手动指定到 A 或 B。",
    })
    review_rows.append({
        "Category": "楼间移动（软优化）",
        "Finding": (
            f"先完成全部硬约束，再{'优化教师相邻课的通行成本' if movement_optimized else '保留基础可行解'}；"
            f"结果复算共评估 {len(travel_review)} 次师生相邻课转换，仍有 {len(travel_shortfalls)} 次课间不足。"
            if travel_seconds else "本次输入未提供 RoomTravel，未进行移动时间优化。"
        ),
        "Input adjustment for a better result": "课间不足不会使排课失败。若要继续降低它，增加相邻时段间隔、为课程加入同楼候选教室，或减少不必要的跨楼场地限定。",
    })
    notes = ["已用 OR-Tools 找到可行课表。", "A/B 学生由 OR-Tools 按行政班一次决定，并在所有 A/B 课程中保持同一 block。", "同一课程小节固定一位教师和一间教室；非连堂课程不会在同一天重复。", "先完成硬约束，再在限定时间内优化教师移动；学生移动仅在结果页复算评分。"]
    diagnostic_rows.insert(0, {
        "Category": "排课结果",
        "Detail": "已完成教师、教室、学生冲突、统一时间和连堂规则检查，并找到可行解；教师移动只在第二阶段限时优化，学生移动在结果页复算。",
        "Suggestion": "若要继续改善通行时间，优先增加同楼候选教室或增加课间间隔，不影响基础可行性。",
    })
    return {
        "Schedule": pd.DataFrame(schedule),
        "ScheduleTimeSlots": pd.DataFrame(meetings),
        "CourseSummary": pd.DataFrame(course_summary),
        "Students_A": pd.DataFrame(student_a),
        "Students_B": pd.DataFrame(student_b),
        "StudentTimetable": pd.DataFrame(student_timetable),
        "StudentMeetings": pd.DataFrame(student_meetings),
        # Keep the raw matrix in the JSON result so the web-page assessment can
        # independently recompute movement, rather than echoing this solver's
        # TravelSoftReview rows.
        "RoomTravel": pd.DataFrame(tables.get("RoomTravel", [])),
        "TravelSoftReview": pd.DataFrame(travel_review),
        "TeacherTimetable": pd.DataFrame(teacher_timetable),
        "RoomTimetable": pd.DataFrame(room_timetable),
        "TeacherLoad": pd.DataFrame(teacher_load),
        "StudentLoad": pd.DataFrame(student_load),
        "StudentBlockAssignments": pd.DataFrame(block_assignments),
        "TypeSectionMembership": pd.DataFrame(type_section_membership),
        "ScheduleReview": pd.DataFrame(review_rows),
        "Diagnostics": pd.DataFrame(diagnostic_rows),
        "AuditChecks": pd.DataFrame(audit_rows),
        "SolverNotes": pd.DataFrame({"Note": notes}),
    }, notes


def main() -> int:
    if len(sys.argv) != 2:
        print("用法：python step3_schedule.py <第二步准备结果.xlsx>")
        return 2
    input_path = Path(sys.argv[1])
    try:
        tables, notes = solve(input_path)
    except (ValueError, ModuleNotFoundError) as error:
        print(f"排课失败：{error}")
        return 1
    failed_audits = [row["Rule"] for row in records(tables["AuditChecks"]) if row["Status"] == "未通过"]
    if failed_audits:
        print("排课结果复核失败，已拒绝导出：" + "、".join(failed_audits))
        return 1
    output_xlsx = input_path.with_name(f"{input_path.stem}_schedule.xlsx")
    output_json = input_path.with_name(f"{input_path.stem}_schedule.json")
    with pd.ExcelWriter(output_xlsx, engine="openpyxl") as writer:
        sheet_names = {
            "CourseSummary": "课程总览", "TeacherLoad": "教师负荷", "StudentLoad": "学生课时",
            "StudentBlockAssignments": "学生分组", "TypeSectionMembership": "Type课程范围", "TravelSoftReview": "移动软优化", "ScheduleReview": "结果建议",
            "AuditChecks": "合规检查", "SolverNotes": "说明",
        }
        for name, sheet_name in sheet_names.items():
            table = tables.get(name, pd.DataFrame())
            table.to_excel(writer, sheet_name=sheet_name, index=False)
            style_table_sheet(writer.sheets[sheet_name])

        # AvailableMeetingTimes is one record per day × period.  It is the solver's
        # calendar, not the vertical axis of a weekly grid.  Using it directly
        # repeated P1–P7 once for every weekday in the exported timetable.
        available_meetings = {(row["Day"], row["TimeSlotID"]) for row in records(tables["ScheduleTimeSlots"])}
        unique_times: dict[str, dict[str, str]] = {}
        for row in records(tables["ScheduleTimeSlots"]):
            unique_times.setdefault(row["TimeSlotID"], {
                "TimeSlotID": row["TimeSlotID"],
                "StartTime": row["StartTime"],
                "EndTime": row["EndTime"],
            })
        times = sorted(unique_times.values(), key=lambda row: slot_number(row["TimeSlotID"]))
        teacher_rows = records(tables["TeacherTimetable"])
        room_rows = records(tables["RoomTimetable"])
        schedule_rows = records(tables["Schedule"])
        type_sections_by_key: dict[tuple[str, str], set[str]] = defaultdict(set)
        for row in records(tables["TypeSectionMembership"]):
            type_sections_by_key[(row["Block"], row["AType"])].add(row["SectionID"])
        type_descriptions = {
            row["AType"]: row.get("TypeDescription", "")
            for row in records(tables["StudentBlockAssignments"])
            if row.get("AType", "")
        }
        # Types are no longer assumed to be the three science combinations.
        # Export exactly the Block + Type timetables that actually contain
        # students, and put the combined selection signature in the title.
        for block, type_label in sorted(type_sections_by_key, key=lambda item: (item[0], item[1])):
            safe_type = re.sub(r"[^A-Za-z0-9]", "", type_label) or "Type"
            title = f"{type_label} · {type_descriptions.get(type_label, '')}".strip(" ·")
            worksheet = writer.book.create_sheet(f"{block}组-{safe_type}"[:31])
            write_block_grid(
                worksheet, schedule_rows, block, title, None, times, available_meetings,
                type_sections_by_key[(block, type_label)],
            )
        write_entity_matrix(
            writer.book.create_sheet("教师课表"), teacher_rows, "TeacherID",
            lambda teacher_id, rows: f"{teacher_id}  {rows[0]['TeacherName']}",
            lambda row: f"{row['CourseName']}\n{row['RoomID']}", times, available_meetings,
        )
        write_entity_matrix(
            writer.book.create_sheet("教室课表"), room_rows, "RoomID",
            lambda room_id, _rows: room_id,
            lambda row: f"{row['CourseName']}\n{row['TeacherName']}", times, available_meetings,
        )
        for worksheet in writer.book.worksheets:
            worksheet.sheet_properties.pageSetUpPr.fitToPage = True
            worksheet.page_setup.fitToWidth = 1
            worksheet.page_setup.fitToHeight = 0
            worksheet.page_margins.left = 0.25
            worksheet.page_margins.right = 0.25
    with output_json.open("w", encoding="utf-8") as handle:
        json.dump({name: records(table) for name, table in tables.items()}, handle, ensure_ascii=False, indent=2)
    print(f"排课完成：{output_xlsx}")
    print(f"同时生成：{output_json}")
    print(f"共排定 {len(tables['Schedule'])} 个上课时段。")
    for note in notes:
        print(f"提示：{note}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except ValueError as error:
        # User-facing pipeline failures are expected validation/scheduling
        # outcomes.  Do not send a Python traceback into the web page.
        print(str(error))
        raise SystemExit(1)
    except Exception:
        print("第三步发生内部错误，未生成课表。请保留输入并重新运行；页面不会显示技术堆栈。")
        raise SystemExit(1)
