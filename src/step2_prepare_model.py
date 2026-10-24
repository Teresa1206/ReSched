"""Step 2: turn a validated ReSched input workbook into solver-ready tables.

Usage:
    python step2_prepare_model.py <timetable-input.xlsx>

This program does not schedule classes. It expands the user-entered choices
into sections, enrollments, eligible teachers, candidate rooms, and usable
weekly meeting times. Step 3 will give these tables to OR-Tools.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any

import pandas as pd

from step1_validate_input import _clean, validate_input


DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]


def records(frame: pd.DataFrame) -> list[dict[str, str]]:
    return [{key: _clean(value) for key, value in row.items()} for row in frame.to_dict("records")]


def prepare(input_path: str | Path) -> tuple[dict[str, pd.DataFrame], list[str]]:
    result = validate_input(input_path)
    if not result.is_valid:
        raise ValueError(result.report())

    t = {name: records(frame) for name, frame in result.tables.items()}
    courses = {row["CourseID"]: row for row in t["Courses"]}
    options = {course_id: [] for course_id in courses}
    option_rules: dict[tuple[str, str], dict[str, str]] = {}
    for row in t["CourseOptions"]:
        options[row["CourseID"]].append(row["OptionName"])
        option_rules[(row["CourseID"], row["OptionName"])] = row

    students = {row["StudentID"]: row for row in t["Students"]}
    rooms = t["Classrooms"]
    special_buildings = {
        row["BuildingID"] for row in t["Buildings"]
        if row.get("IsSpecialBuilding", "") in {"1", "1.0"}
    }
    room_buildings = {row["RoomID"]: row["BuildingID"] for row in rooms}
    building_distances = {
        (row["FromBuildingID"], row["ToBuildingID"]): float(row["DistanceMeters"])
        for row in t["BuildingDistances"]
    }
    building_distances.update({(right, left): distance for (left, right), distance in list(building_distances.items())})
    homeroom_rooms = {row["HomeroomID"]: row["RoomID"] for row in t["Homerooms"]}
    teachers = t["Teachers"]
    teacher_names = {row["TeacherID"]: row["TeacherName"] for row in teachers}
    selections: dict[tuple[str, str], set[str]] = {}
    for row in t["StudentSelections"]:
        selections.setdefault((row["StudentID"], row["CourseID"]), set()).add(row["CourseOption"])

    # Teacher input describes teaching capability only: course + level/option.
    # Block and Type are decisions made by the scheduler, never teacher inputs.
    assignment_teachers: dict[tuple[str, str], list[str]] = {}
    for row in t["TeacherAssignments"]:
        assignment_teachers.setdefault((row["CourseID"], row["CourseOption"]), []).append(row["TeacherID"])
    cas_teacher = {row["HomeroomID"]: row["TeacherID"] for row in t["TeacherCAS"]}

    sections: list[dict[str, Any]] = []
    section_students: list[dict[str, str]] = []
    section_student_candidates: list[dict[str, str]] = []
    section_teachers: list[dict[str, str]] = []
    candidate_rooms: list[dict[str, str]] = []
    # Program 1 writes TravelInconvenience with every unordered room pair.
    # Use its estimated walking seconds and inconvenience score directly when
    # present; otherwise fall back to the built-in travel model.
    program_travel: dict[tuple[str, str], tuple[float, float]] = {}
    try:
        book = pd.ExcelFile(input_path)
        if "TravelInconvenience" in book.sheet_names:
            travel_frame = pd.read_excel(
                input_path, sheet_name="TravelInconvenience", dtype=object
            ).fillna("")
            for travel_row in travel_frame.to_dict("records"):
                room_a = _clean(travel_row.get("room_a"))
                room_b = _clean(travel_row.get("room_b"))
                if not room_a or not room_b:
                    continue
                seconds = float(travel_row.get("estimated_walk_seconds") or 0)
                cost = float(travel_row.get("inconvenience_score") or 0)
                program_travel[(room_a, room_b)] = (seconds, cost)
                program_travel[(room_b, room_a)] = (seconds, cost)
    except (ValueError, OSError, KeyError):
        program_travel = {}

    room_travel: list[dict[str, str | float]] = []
    for first in rooms:
        for second in rooms:
            programmed = program_travel.get((first["RoomID"], second["RoomID"]))
            if programmed is not None:
                estimated_seconds, optimization_cost = programmed
            else:
                if first["RoomID"] == second["RoomID"]:
                    horizontal = 0.0
                    floor_difference = 0
                    cross_building = 0
                else:
                    floor_difference = abs(int(float(first["Floor"])) - int(float(second["Floor"])))
                    cross_building = int(first["BuildingID"] != second["BuildingID"])
                    horizontal = (
                        abs(int(float(first["RoomOrdinal"])) - int(float(second["RoomOrdinal"]))) * 10.0
                        if not cross_building else
                        building_distances.get((first["BuildingID"], second["BuildingID"]), 0.0)
                    )
                estimated_seconds = horizontal / 1.2 + floor_difference * 20.0
                optimization_cost = estimated_seconds + (120.0 if cross_building else 0.0) + floor_difference * 15.0
            room_travel.append({
                "FromRoomID": first["RoomID"], "ToRoomID": second["RoomID"],
                "EstimatedTravelSeconds": round(estimated_seconds, 1),
                "OptimizationCost": round(optimization_cost, 1),
            })
    notices: list[str] = list(result.warnings)

    def add_section(course: dict[str, str], block: str, schedule_type: str, option: str, enrolled: list[str], eligible: list[str], homeroom: str = "", candidates: list[str] | None = None) -> None:
        suffix = option or homeroom or "ALL"
        section_parts = [course["CourseID"], option or homeroom or "ALL"]
        if block:
            section_parts.append(block)
        if schedule_type:
            section_parts.append(schedule_type)
        section_id = "-".join(section_parts)
        # Course IDs remain stable keys; all user-facing schedules use CourseName.
        section_label = " ".join([f"{course['CourseName'] or course['CourseID']} {option or homeroom or 'ALL'}", block]).strip()
        rule = option_rules.get((course["CourseID"], option), {})
        fixed_building = rule.get("RequiredBuildingID", "")
        fixed_room = rule.get("RequiredRoomID", "")
        homeroom_room = homeroom_rooms.get(homeroom, "")
        # CAS is an administrative-class activity: each homeroom stays in its
        # own room with its own homeroom teacher.  It is deliberately not
        # treated as a pool of interchangeable rooms, even though every class
        # meets in the same CAS time band.
        if course["CourseType"] == "cas" and homeroom_room:
            allowed_rooms = [room for room in rooms if room["RoomID"] == homeroom_room]
        else:
            allowed_rooms = [
                room for room in rooms
                if (not fixed_room or room["RoomID"] == fixed_room)
                and (not fixed_building or room["BuildingID"] == fixed_building)
                # A special building (such as the stadium or art center) is available
                # only when this course option explicitly names that building/room.
                and (room["BuildingID"] not in special_buildings or bool(fixed_building or fixed_room))
            ]
        student_pool = candidates if candidates is not None else enrolled
        sections.append({
            "SectionID": section_id,
            "CourseID": course["CourseID"],
            "BaseCourseName": course["CourseName"],
            "CourseName": section_label,
            "Block": block,
            "Type": schedule_type,
            "CourseOption": option,
            "SelectionGroup": rule.get("SelectionGroup", ""),
            "HomeroomID": homeroom,
            "HomeroomBuildingID": room_buildings.get(homeroom_room, ""),
            "HomeroomRoomID": homeroom_room,
            "PeriodsPerWeek": course["PeriodsPerWeek"],
            "AllGradeSameTime": course["AllGradeSameTime"],
            "RequireConsecutive": course["RequireConsecutive"],
            "EnrollmentMode": "A/B decision" if candidates is not None else "Fixed",
            "StudentPoolCount": len(student_pool),
            "RequiredBuildingID": fixed_building,
            "RequiredRoomID": fixed_room,
        })
        for student_id in enrolled:
            section_students.append({"SectionID": section_id, "StudentID": student_id})
        for student_id in candidates or []:
            section_student_candidates.append({"SectionID": section_id, "StudentID": student_id})
        for teacher_id in eligible:
            section_teachers.append({
                "SectionID": section_id,
                "TeacherID": teacher_id,
                "TeacherName": teacher_names.get(teacher_id, ""),
            })
        for room in allowed_rooms:
            candidate_rooms.append({"SectionID": section_id, "RoomID": room["RoomID"], "BuildingID": room["BuildingID"]})
        if not eligible:
            notices.append(f"{section_id} 没有可教教师；第 3 步无法为它排课。")
        if not allowed_rooms:
            notices.append(f"{section_id} 没有符合地点要求的教室；第 3 步无法为它排课。")

    # Type is decided by the *combined selections* of every course that is
    # neither A/B-split, whole-grade simultaneous, nor CAS.  Thus MOD2 is a
    # genuine Type input alongside Science; PE disappears from the signature
    # whenever it is set to whole-grade simultaneous.  A level inside one
    # SelectionGroup (CHEM S/H) does not create a new Type.
    type_courses = [
        course for course in courses.values()
        if course["CourseType"] in {"level", "choice"}
        and course["SplitAB"] not in {"1", "1.0"}
        and course["AllGradeSameTime"] not in {"1", "1.0"}
        and course["CourseType"] != "cas"
    ]
    type_signatures: dict[str, tuple[tuple[str, tuple[str, ...]], ...]] = {}
    for student_id in sorted(students):
        parts: list[tuple[str, tuple[str, ...]]] = []
        for course in type_courses:
            course_id = course["CourseID"]
            groups = sorted({
                option_rules.get((course_id, option), {}).get("SelectionGroup", "") or option
                for option in selections.get((student_id, course_id), set())
            })
            # Only selected groups matter.  This makes different MOD2 choices
            # different Types while CHEM S and CHEM H remain one Chemistry group.
            parts.append((course_id, tuple(groups)))
        type_signatures[student_id] = tuple(parts)
    unique_signatures = sorted(set(type_signatures.values())) or [tuple()]
    signature_labels = {signature: f"T{index + 1}" for index, signature in enumerate(unique_signatures)}
    student_types = {student_id: signature_labels[signature] for student_id, signature in type_signatures.items()}
    type_labels = [signature_labels[signature] for signature in unique_signatures]
    type_descriptions = {
        signature_labels[signature]: " · ".join(
            f"{course_id}: {' + '.join(groups) if groups else '无选项'}"
            for course_id, groups in signature
        ) or "共同课程"
        for signature in unique_signatures
    }

    for course in courses.values():
        offered_blocks = ["A", "B"] if course["SplitAB"] in {"1", "1.0"} else [""]
        uses_type = course in type_courses
        offered_types = type_labels if uses_type else [""]
        course_type = course["CourseType"]
        if course_type in {"level", "choice"}:
            for block in offered_blocks:
                for schedule_type in offered_types:
                    for option in options[course["CourseID"]]:
                        selected = sorted(student_id for student_id in students if option in selections.get((student_id, course["CourseID"]), set()) and (not schedule_type or student_types[student_id] == schedule_type))
                        # Do not create a phantom section for an option nobody
                        # in this Type selected.  Such a section would consume
                        # a teacher/room and make the model look harder than
                        # the actual student demand.
                        if not selected:
                            continue
                        eligible = assignment_teachers.get((course["CourseID"], option), [])
                        if block:
                            add_section(course, block, schedule_type, option, [], eligible, candidates=selected)
                        else:
                            add_section(course, block, schedule_type, option, selected, eligible)
        elif course_type == "required":
            for block in offered_blocks:
                for schedule_type in offered_types:
                    eligible = assignment_teachers.get((course["CourseID"], ""), [])
                    enrolled = sorted(student_id for student_id in students if not schedule_type or student_types[student_id] == schedule_type)
                    add_section(course, block, schedule_type, "", enrolled, eligible)
        elif course_type == "cas":
            for block in offered_blocks:
                for homeroom_id in sorted({student["HomeroomID"] for student in students.values()}):
                    enrolled = sorted(
                        student_id for student_id, student in students.items()
                        if student["HomeroomID"] == homeroom_id
                    )
                    eligible = [cas_teacher[homeroom_id]] if homeroom_id in cas_teacher else []
                    add_section(course, block, "", "", enrolled, eligible, homeroom_id)

    blocked = {(row["Day"], row["TimeSlotID"]) for row in t["NoClassTimes"]}
    available_times = [
        {"Day": day, "TimeSlotID": slot["TimeSlotID"], "StartTime": slot["StartTime"], "EndTime": slot["EndTime"]}
        for day in DAYS for slot in t["TimeSlots"]
        if (day, slot["TimeSlotID"]) not in blocked
    ]

    available_lookup = {
        (row["TeacherID"], row["Day"], row["TimeSlotID"]): int(float(row["Available"]))
        for row in t["TeacherAvailability"]
    }
    teacher_availability = [
        {
            "TeacherID": teacher["TeacherID"],
            "Day": item["Day"],
            "TimeSlotID": item["TimeSlotID"],
            "Available": available_lookup[(teacher["TeacherID"], item["Day"], item["TimeSlotID"])],
        }
        for teacher in teachers for item in available_times
    ]
    for section in sections:
        section_id = section["SectionID"]
        eligible = [row["TeacherID"] for row in section_teachers if row["SectionID"] == section_id]
        if eligible and not any(
            row["Available"] == 1 and row["TeacherID"] in eligible for row in teacher_availability
        ):
            notices.append(f"{section_id} 的可教教师在所有全校可排时段都不可用；第 3 步无法为它排课。")

    return {
        "Sections": pd.DataFrame(sections),
        "SectionStudents": pd.DataFrame(section_students),
        "SectionStudentCandidates": pd.DataFrame(section_student_candidates),
        "EligibleTeachers": pd.DataFrame(section_teachers),
        "CandidateRooms": pd.DataFrame(candidate_rooms),
        "RoomTravel": pd.DataFrame(room_travel),
        "AvailableMeetingTimes": pd.DataFrame(available_times),
        "TeacherAvailability": pd.DataFrame(teacher_availability),
        "HomeroomBlockBalance": pd.DataFrame(t["HomeroomBlockBalance"]),
        "Students": pd.DataFrame([
            {**student, "ScheduleType": student_types.get(student_id, "T1"), "TypeDescription": type_descriptions.get(student_types.get(student_id, "T1"), "")}
            for student_id, student in students.items()
        ]),
        "PreparationNotes": pd.DataFrame({"Note": notices}),
    }, notices


def main() -> int:
    if len(sys.argv) != 2:
        print("用法：python step2_prepare_model.py <timetable-input.xlsx>")
        return 2
    input_path = Path(sys.argv[1])
    try:
        tables, notices = prepare(input_path)
    except ValueError as error:
        print(error)
        return 1
    output_xlsx = input_path.with_name(f"{input_path.stem}_prepared.xlsx")
    output_json = input_path.with_name(f"{input_path.stem}_prepared.json")
    with pd.ExcelWriter(output_xlsx, engine="openpyxl") as writer:
        for name, table in tables.items():
            table.to_excel(writer, sheet_name=name, index=False)
    with output_json.open("w", encoding="utf-8") as handle:
        json.dump({name: records(table) for name, table in tables.items()}, handle, ensure_ascii=False, indent=2)
    print(f"准备完成：{output_xlsx}")
    print(f"同时生成：{output_json}")
    print(f"共生成 {len(tables['Sections'])} 个待排课程小节。")
    for notice in notices:
        print(f"提示：{notice}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
