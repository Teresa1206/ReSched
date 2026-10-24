"""Independently re-evaluate an exported ReSched timetable.

This intentionally ignores AuditChecks, ScheduleReview and TravelSoftReview in
the schedule JSON.  It reads the Step 2 rules and the actual Schedule rows,
then recomputes all reported checks from scratch.
"""

from __future__ import annotations

import json
import sys
from collections import defaultdict
from pathlib import Path

import pandas as pd

from step3_schedule import DAYS, clean, records, slot_number, student_enrollment_inputs, timetable_course_family


def truth(value: object) -> bool:
    return clean(value) in {"1", "1.0", "true", "True", "TRUE"}


def audit(rule: str, passed: bool, detail: str) -> dict[str, str]:
    return {"Rule": rule, "Status": "通过" if passed else "未通过", "Detail": detail}


def evaluate(prepared_path: str | Path, schedule_path: str | Path) -> dict[str, object]:
    prepared_path, schedule_path = Path(prepared_path), Path(schedule_path)
    book = pd.ExcelFile(prepared_path)
    needed = ["Sections", "SectionStudents", "SectionStudentCandidates", "EligibleTeachers", "CandidateRooms", "AvailableMeetingTimes", "TeacherAvailability", "Students", "HomeroomBlockBalance"]
    missing = [name for name in needed if name not in book.sheet_names]
    if missing:
        raise ValueError("第二步文件缺少工作表：" + "、".join(missing))
    tables = {name: records(pd.read_excel(prepared_path, sheet_name=name, dtype=object).fillna("")) for name in needed}
    if "RoomTravel" in book.sheet_names:
        tables["RoomTravel"] = records(pd.read_excel(prepared_path, sheet_name="RoomTravel", dtype=object).fillna(""))
    result = json.loads(schedule_path.read_text(encoding="utf-8"))
    schedule = [{key: clean(value) for key, value in row.items()} for row in result.get("Schedule", [])]
    if not schedule:
        raise ValueError("排课结果 JSON 没有 Schedule，无法复算。")

    sections = tables["Sections"]
    section_by_id = {row["SectionID"]: row for row in sections}
    unknown_sections = [row.get("SectionID", "") for row in schedule if row.get("SectionID", "") not in section_by_id]
    if unknown_sections:
        raise ValueError("结果含有第二步文件不存在的课程小节：" + "、".join(sorted(set(unknown_sections))[:5]))
    teachers = defaultdict(set)
    for row in tables["EligibleTeachers"]:
        teachers[row["SectionID"]].add(row["TeacherID"])
    rooms = defaultdict(set)
    for row in tables["CandidateRooms"]:
        rooms[row["SectionID"]].add(row["RoomID"])
    allowed_times = {(row["Day"], row["TimeSlotID"]) for row in tables["AvailableMeetingTimes"]}
    available = {(row["TeacherID"], row["Day"], row["TimeSlotID"]): truth(row["Available"]) for row in tables["TeacherAvailability"]}
    students = {row["StudentID"]: row for row in tables["Students"]}
    block_by_student = {clean(row.get("StudentID")): clean(row.get("Block")) for row in result.get("StudentBlockAssignments", [])}
    fixed, split_groups = student_enrollment_inputs(tables)
    student_sections: dict[str, set[str]] = defaultdict(set)
    for section_id, ids in fixed.items():
        for student_id in ids:
            student_sections[student_id].add(section_id)
    for group in split_groups:
        for student_id in group["Students"]:
            block = block_by_student.get(student_id, "")
            if block not in {"A", "B"}:
                continue
            student_sections[student_id].add(str(group[block]))

    by_section = defaultdict(list)
    for row in schedule:
        by_section[row["SectionID"]].append(row)
    student_meetings = []
    for student_id, section_ids in student_sections.items():
        for section_id in section_ids:
            for row in by_section[section_id]:
                student_meetings.append({"StudentID": student_id, "SectionID": section_id, "Day": row["Day"], "TimeSlotID": row["TimeSlotID"], "RoomID": row["RoomID"]})

    counts_ok = all(len(by_section[section["SectionID"]]) == int(float(section.get("PeriodsPerWeek") or 0)) for section in sections)
    teacher_keys = [(row["TeacherID"], row["Day"], row["TimeSlotID"]) for row in schedule]
    room_keys = [(row["RoomID"], row["Day"], row["TimeSlotID"]) for row in schedule]
    student_keys = [(row["StudentID"], row["Day"], row["TimeSlotID"]) for row in student_meetings]
    teachers_ok = all(row["TeacherID"] in teachers[row["SectionID"]] for row in schedule)
    rooms_ok = all(row["RoomID"] in rooms[row["SectionID"]] for row in schedule)
    availability_ok = all(available.get((row["TeacherID"], row["Day"], row["TimeSlotID"]), False) for row in schedule)
    time_ok = all((row["Day"], row["TimeSlotID"]) in allowed_times for row in schedule)

    simultaneous_ok = True
    for course_id in {row["CourseID"] for row in sections if truth(row.get("AllGradeSameTime"))}:
        members = [row for row in sections if row["CourseID"] == course_id]
        expected = {(row["Day"], row["TimeSlotID"]) for row in by_section[members[0]["SectionID"]]} if members else set()
        simultaneous_ok = simultaneous_ok and all({(row["Day"], row["TimeSlotID"]) for row in by_section[member["SectionID"]]} == expected for member in members[1:])
    consecutive_ok = True
    for section in sections:
        if not truth(section.get("RequireConsecutive")):
            continue
        rows = by_section[section["SectionID"]]
        consecutive_ok = consecutive_ok and len(rows) == 2 and rows[0]["Day"] == rows[1]["Day"] and abs(slot_number(rows[0]["TimeSlotID"]) - slot_number(rows[1]["TimeSlotID"])) == 1

    balances = defaultdict(lambda: {"A": 0, "B": 0})
    for student_id, block in block_by_student.items():
        if student_id in students and block in {"A", "B"}:
            balances[students[student_id]["HomeroomID"]][block] += 1
    limits = {row["HomeroomID"]: (int(float(row.get("MinStudentsPerBlock") or 0)), int(float(row.get("MaxStudentsPerBlock") or 10**6))) for row in tables["HomeroomBlockBalance"]}
    balance_ok = all(low <= values["A"] <= high for homeroom, values in balances.items() for low, high in [limits.get(homeroom, (0, 10**6))])
    ab_ok = all(block_by_student.get(student_id) in {"A", "B"} for group in split_groups for student_id in group["Students"])

    type_period_families: dict[tuple[str, str, str, str], set[str]] = defaultdict(set)
    type_by_student = {clean(row.get("StudentID")): clean(row.get("AType")) for row in result.get("StudentBlockAssignments", [])}
    for student_id, section_ids in student_sections.items():
        for section_id in section_ids:
            family = timetable_course_family(section_by_id[section_id])
            for row in by_section[section_id]:
                type_period_families[(block_by_student.get(student_id, ""), type_by_student.get(student_id, ""), row["Day"], row["TimeSlotID"])].add(family)
    type_conflicts = {key: sorted(value) for key, value in type_period_families.items() if len(value) > 1}

    repeated_same_day = []
    for section in sections:
        if truth(section.get("RequireConsecutive")):
            continue
        days = [row["Day"] for row in by_section[section["SectionID"]]]
        if len(days) != len(set(days)):
            repeated_same_day.append(section["SectionID"])
    type_detail = "每个 A/B + Type 课表的同一格最多出现一个学生实际需要的课程类别。"
    if type_conflicts:
        samples = [f"{key[0]} {key[1]} {key[2]} {key[3]}：{'、'.join(value)}" for key, value in list(sorted(type_conflicts.items()))[:4]]
        type_detail = "冲突：" + "；".join(samples)
    audit_rows = [
        audit("课程节数", counts_ok, "按当前 Schedule 重新统计每个 SectionID 的课次。"),
        audit("全校不可上课时段", time_ok, "逐条核对当前 Schedule 是否仅使用 AvailableMeetingTimes。"),
        audit("教师不冲突", len(set(teacher_keys)) == len(teacher_keys), "从当前 Schedule 重建教师×时段占用。"),
        audit("教师任教范围", teachers_ok, "当前每条课次的教师均须在 EligibleTeachers 中。"),
        audit("教师可用时间", availability_ok, "当前每条课次均按 TeacherAvailability 重新核对。"),
        audit("教室不冲突", len(set(room_keys)) == len(room_keys), "从当前 Schedule 重建教室×时段占用。"),
        audit("教室适用范围", rooms_ok, "当前每条课次的教室均须在 CandidateRooms 中。"),
        audit("学生不冲突", len(set(student_keys)) == len(student_keys), "由第二步选课、当前 A/B 分组与当前 Schedule 重新生成学生课次后核对。"),
        audit("Block + Type 无跨课程并行", not type_conflicts, type_detail),
        audit("全局 A/B 分班", ab_ok and balance_ok, "从当前 StudentBlockAssignments 重新核对 A/B 完整性及每个行政班人数范围。"),
        audit("全年级统一时间", simultaneous_ok, "重新比较勾选统一时间课程的所有小节时段集合。"),
        audit("必须连堂", consecutive_ok, "重新检查 RequireConsecutive 小节是否同一天连续两节。"),
        {"Rule": "非连堂课程分散", "Status": "提示", "Detail": "这不是输入中的硬性要求，因此不会阻止课表导出。" + (f" 当前有 {len(repeated_same_day)} 个小节同日重复，可按需优化。" if repeated_same_day else " 当前没有同日重复。")},
    ]
    scheduled_by_student = defaultdict(int)
    for row in student_meetings:
        scheduled_by_student[row["StudentID"]] += 1
    total_slots = len(allowed_times)
    student_load = [{"StudentID": student_id, "StudentName": student.get("StudentName", ""), "ScheduledPeriods": scheduled_by_student[student_id], "AvailablePeriods": total_slots, "FreePeriods": max(0, total_slots - scheduled_by_student[student_id]), "AdditionalPeriodsNeededForFullTimetable": max(0, total_slots - scheduled_by_student[student_id])} for student_id, student in sorted(students.items())]
    teacher_load = []
    teacher_names = {row["TeacherID"]: row.get("TeacherName", "") for row in tables["TeacherAvailability"]}
    for teacher_id in sorted({row["TeacherID"] for row in tables["TeacherAvailability"]}):
        capacity = sum(1 for key, ok in available.items() if key[0] == teacher_id and ok)
        assigned = sum(1 for row in schedule if row["TeacherID"] == teacher_id)
        teacher_load.append({"TeacherID": teacher_id, "TeacherName": teacher_names.get(teacher_id, ""), "ScheduledPeriods": assigned, "AvailablePeriods": capacity, "FreePeriods": capacity - assigned})
    return {
        "AuditChecks": audit_rows,
        "TeacherLoad": teacher_load,
        "StudentLoad": student_load,
        "StudentMeetings": student_meetings,
        "StudentBlockAssignments": result.get("StudentBlockAssignments", []),
        "RoomTravel": tables.get("RoomTravel", []),
        # Return the canonical day × period clock from Step 2.  The browser
        # uses this instead of trusting any Excel-converted time text carried
        # by individual Schedule rows.
        "ScheduleTimeSlots": tables["AvailableMeetingTimes"],
        "Schedule": schedule,
        "Recomputed": True,
    }


def main() -> int:
    # The local web server parses this program's stdout as UTF-8 JSON.  On
    # Windows a redirected console can otherwise default to a legacy codepage
    # and corrupt every Chinese audit message.
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8")
    if len(sys.argv) != 3:
        print("用法：python step3_evaluate_result.py <第二步准备文件.xlsx> <第三步结果_schedule.json>", file=sys.stderr)
        return 2
    try:
        print(json.dumps(evaluate(sys.argv[1], sys.argv[2]), ensure_ascii=False))
        return 0
    except (ValueError, KeyError, OSError, json.JSONDecodeError) as error:
        print(f"结果复算失败：{error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
