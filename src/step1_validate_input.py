"""Step 1 of the timetable pipeline: load, clean, and validate exported input.

Usage:
    python step1_validate_input.py timetable-input.xlsx

The program deliberately stops before travel-time calculation and OR-Tools.
It accepts the normalized workbook exported by ReSched_Track1_Input_Schema.html.
"""

from __future__ import annotations

import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Iterable, List, Set

import pandas as pd


SHEETS: Dict[str, List[str]] = {
    "Buildings": ["BuildingID", "BuildingName", "Floors", "IsSpecialBuilding"],
    "Classrooms": ["RoomID", "BuildingID", "Floor", "RoomOrdinal", "RoomName"],
    "BuildingDistances": ["FromBuildingID", "ToBuildingID", "DistanceMeters"],
    "TimeSlots": ["TimeSlotID", "StartTime", "EndTime"],
    "NoClassTimes": ["Day", "TimeSlotID", "Reason"],
    "Homerooms": ["HomeroomID", "HomeroomName", "RoomID"],
    "HomeroomBlockBalance": ["HomeroomID", "MinStudentsPerBlock", "MaxStudentsPerBlock"],
    "Courses": [
        "CourseID",
        "CourseName",
        "SplitAB",
        "PeriodsPerWeek",
        "CourseType",
        "OptionCount",
        "StudentPickCount",
        "TypeCount",
        "AllGradeSameTime",
        "RequireConsecutive",
    ],
    "CourseOptions": [
        "CourseID",
        "OptionNumber",
        "OptionName",
        "RequiredBuildingID",
        "RequiredRoomID",
        "SelectionGroup",
        "Level",
    ],
    "Teachers": ["TeacherID", "TeacherName"],
    "TeacherAssignments": ["TeacherID", "CourseID", "CourseOption"],
    "TeacherAvailability": ["TeacherID", "Day", "TimeSlotID", "Available"],
    "TeacherCAS": ["TeacherID", "HomeroomID"],
    "Students": ["StudentID", "StudentName", "HomeroomID"],
    "StudentSelections": ["StudentID", "CourseID", "CourseOption"],
}

COURSE_TYPES = {"level", "choice", "required", "cas"}
DAYS = {"Monday", "Tuesday", "Wednesday", "Thursday", "Friday"}
BLOCKS = {"A", "B"}
OPTIONAL_LEGACY_COLUMNS: Dict[str, Set[str]] = {
    "CourseOptions": {"SelectionGroup", "Level"},
    # Existing workbooks remain readable; a missing flag means an ordinary building.
    "Buildings": {"IsSpecialBuilding"},
}


@dataclass
class ValidationResult:
    tables: Dict[str, pd.DataFrame]
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)

    @property
    def is_valid(self) -> bool:
        return not self.errors

    def report(self) -> str:
        lines = ["输入校验通过。" if self.is_valid else f"输入校验失败：{len(self.errors)} 项错误。"]
        lines.extend(f"错误：{message}" for message in self.errors)
        lines.extend(f"提示：{message}" for message in self.warnings)
        return "\n".join(lines)


def _clean(value: object) -> str:
    """Turn an Excel cell into a stable, trimmed ID/text value."""
    if pd.isna(value):
        return ""
    return str(value).strip()


def _read_workbook(path: Path) -> ValidationResult:
    excel = pd.ExcelFile(path)
    missing = [sheet for sheet in SHEETS if sheet not in excel.sheet_names]
    result = ValidationResult(tables={})
    if missing:
        result.errors.append(f"缺少工作表：{', '.join(missing)}")
        return result

    for sheet, required_columns in SHEETS.items():
        frame = pd.read_excel(path, sheet_name=sheet, dtype=object).fillna("")
        legacy_columns = OPTIONAL_LEGACY_COLUMNS.get(sheet, set())
        absent = [column for column in required_columns if column not in frame.columns and column not in legacy_columns]
        if absent:
            result.errors.append(f"{sheet} 缺少列：{', '.join(absent)}")
            continue
        if sheet == "Buildings" and "IsSpecialBuilding" not in frame.columns:
            frame["IsSpecialBuilding"] = 0
            result.warnings.append("Buildings 未包含 IsSpecialBuilding；已按所有楼均为普通楼处理。")
        if sheet == "CourseOptions" and {"SelectionGroup", "Level"} - set(frame.columns):
            # Older exports only have flat labels.  Infer a group from a shared
            # first word such as PHY H / PHY S, while ordinary options stay alone.
            labels = frame["OptionName"].map(_clean)
            course_ids = frame["CourseID"].map(_clean)
            prefixes = labels.map(lambda label: label.split(maxsplit=1)[0] if " " in label else "")
            counts = pd.DataFrame({"course": course_ids, "prefix": prefixes}).query("prefix != ''").value_counts()
            frame["SelectionGroup"] = [
                prefix if prefix and counts.get((course_id, prefix), 0) > 1 else label
                for course_id, prefix, label in zip(course_ids, prefixes, labels)
            ]
            frame["Level"] = [
                label.split(maxsplit=1)[1] if prefix and counts.get((course_id, prefix), 0) > 1 else ""
                for course_id, prefix, label in zip(course_ids, prefixes, labels)
            ]
            result.warnings.append("CourseOptions 未单列学科组；已自动识别如 PHY H / PHY S 的学科分组。")
        result.tables[sheet] = frame[required_columns].copy()
    return result


def _ids(frame: pd.DataFrame, column: str) -> Set[str]:
    return set(frame[column].map(_clean)) - {""}


def _duplicate_values(frame: pd.DataFrame, column: str) -> Iterable[str]:
    values = frame[column].map(_clean)
    return values[(values != "") & values.duplicated()].unique()


def _require_unique(result: ValidationResult, sheet: str, column: str) -> None:
    for value in _duplicate_values(result.tables[sheet], column):
        result.errors.append(f"{sheet}.{column} 存在重复值：{value}")


def _require_nonempty(result: ValidationResult, sheet: str, columns: List[str]) -> None:
    frame = result.tables[sheet]
    for index, row in frame.iterrows():
        for column in columns:
            if not _clean(row[column]):
                result.errors.append(f"{sheet} 第 {index + 2} 行：{column} 不能为空")


def _check_foreign_keys(
    result: ValidationResult,
    sheet: str,
    column: str,
    valid_ids: Set[str],
    target: str,
    allow_blank: bool = False,
) -> None:
    for index, value in enumerate(result.tables[sheet][column].map(_clean), start=2):
        if not value and allow_blank:
            continue
        if value not in valid_ids:
            result.errors.append(f"{sheet} 第 {index} 行：{column}={value or '空'} 未在 {target} 中找到")


def _as_positive_int(value: object) -> int | None:
    try:
        number = int(float(value))
        return number if number > 0 else None
    except (TypeError, ValueError):
        return None


def _check_time(value: object) -> bool:
    try:
        pd.to_datetime(_clean(value), format="%H:%M")
        return True
    except ValueError:
        return False


def validate_input(path: str | Path) -> ValidationResult:
    """Load an exported workbook and validate all relationships needed by step 2."""
    workbook = Path(path)
    if not workbook.is_file():
        return ValidationResult(tables={}, errors=[f"找不到输入文件：{workbook}"])

    result = _read_workbook(workbook)
    if result.errors:
        return result

    t = result.tables
    for sheet, column in [
        ("Buildings", "BuildingID"),
        ("Classrooms", "RoomID"),
        ("TimeSlots", "TimeSlotID"),
        ("Homerooms", "HomeroomID"),
        ("Courses", "CourseID"),
        ("Teachers", "TeacherID"),
        ("Students", "StudentID"),
    ]:
        _require_nonempty(result, sheet, [column])
        _require_unique(result, sheet, column)

    building_ids = _ids(t["Buildings"], "BuildingID")
    room_ids = _ids(t["Classrooms"], "RoomID")
    slot_ids = _ids(t["TimeSlots"], "TimeSlotID")
    homeroom_ids = _ids(t["Homerooms"], "HomeroomID")
    course_ids = _ids(t["Courses"], "CourseID")
    teacher_ids = _ids(t["Teachers"], "TeacherID")
    student_ids = _ids(t["Students"], "StudentID")

    _check_foreign_keys(result, "Classrooms", "BuildingID", building_ids, "Buildings")
    _check_foreign_keys(result, "Homerooms", "RoomID", room_ids, "Classrooms")
    _check_foreign_keys(result, "HomeroomBlockBalance", "HomeroomID", homeroom_ids, "Homerooms")
    _check_foreign_keys(result, "Students", "HomeroomID", homeroom_ids, "Homerooms")
    _check_foreign_keys(result, "NoClassTimes", "TimeSlotID", slot_ids, "TimeSlots")
    _check_foreign_keys(result, "CourseOptions", "CourseID", course_ids, "Courses")
    _check_foreign_keys(result, "TeacherAssignments", "TeacherID", teacher_ids, "Teachers")
    _check_foreign_keys(result, "TeacherAssignments", "CourseID", course_ids, "Courses")
    _check_foreign_keys(result, "TeacherAvailability", "TeacherID", teacher_ids, "Teachers")
    _check_foreign_keys(result, "TeacherAvailability", "TimeSlotID", slot_ids, "TimeSlots")
    _check_foreign_keys(result, "TeacherCAS", "TeacherID", teacher_ids, "Teachers")
    _check_foreign_keys(result, "TeacherCAS", "HomeroomID", homeroom_ids, "Homerooms")
    _check_foreign_keys(result, "StudentSelections", "StudentID", student_ids, "Students")
    _check_foreign_keys(result, "StudentSelections", "CourseID", course_ids, "Courses")

    for index, row in t["Buildings"].iterrows():
        if _as_positive_int(row["Floors"]) is None:
            result.errors.append(f"Buildings 第 {index + 2} 行：Floors 必须为正整数")
        if _clean(row["IsSpecialBuilding"]) not in {"0", "0.0", "1", "1.0"}:
            result.errors.append(f"Buildings 第 {index + 2} 行：IsSpecialBuilding 必须为 0 或 1")

    for index, row in t["Classrooms"].iterrows():
        floor = _as_positive_int(row["Floor"])
        ordinal = _as_positive_int(row["RoomOrdinal"])
        if floor is None or ordinal is None:
            result.errors.append(f"Classrooms 第 {index + 2} 行：Floor 与 RoomOrdinal 必须为正整数")
            continue
        expected = f"{_clean(row['BuildingID'])}-{floor * 100 + ordinal}"
        if _clean(row["RoomID"]) != expected:
            result.errors.append(f"Classrooms 第 {index + 2} 行：RoomID 应为 {expected}")

    distance_pairs: Set[frozenset[str]] = set()
    for index, row in t["BuildingDistances"].iterrows():
        origin, destination = _clean(row["FromBuildingID"]), _clean(row["ToBuildingID"])
        if origin == destination:
            result.errors.append(f"BuildingDistances 第 {index + 2} 行：起点楼与终点楼不能相同")
        pair = frozenset((origin, destination))
        if len(pair) == 2 and pair in distance_pairs:
            result.errors.append(f"BuildingDistances 第 {index + 2} 行：楼间距离重复：{origin} / {destination}")
        distance_pairs.add(pair)
        try:
            if float(row["DistanceMeters"]) < 0:
                raise ValueError
        except (TypeError, ValueError):
            result.errors.append(f"BuildingDistances 第 {index + 2} 行：DistanceMeters 必须为非负数")

    for index, row in t["TimeSlots"].iterrows():
        start, end = _clean(row["StartTime"]), _clean(row["EndTime"])
        if not _check_time(start) or not _check_time(end) or start >= end:
            result.errors.append(f"TimeSlots 第 {index + 2} 行：开始/结束时间无效")

    for index, row in t["NoClassTimes"].iterrows():
        if _clean(row["Day"]) not in DAYS:
            result.errors.append(f"NoClassTimes 第 {index + 2} 行：Day 必须为周一至周五")

    availability_pairs: Set[tuple[str, str, str]] = set()
    for index, row in t["TeacherAvailability"].iterrows():
        teacher_id, day, slot_id = _clean(row["TeacherID"]), _clean(row["Day"]), _clean(row["TimeSlotID"])
        pair = (teacher_id, day, slot_id)
        if day not in DAYS:
            result.errors.append(f"TeacherAvailability 第 {index + 2} 行：Day 必须为周一至周五")
        if pair in availability_pairs:
            result.errors.append(f"TeacherAvailability 第 {index + 2} 行：教师、星期、时段组合重复")
        availability_pairs.add(pair)
        try:
            if int(float(row["Available"])) not in {0, 1}:
                raise ValueError
        except (TypeError, ValueError):
            result.errors.append(f"TeacherAvailability 第 {index + 2} 行：Available 只能是 0 或 1")

    expected_availability = {(teacher_id, day, slot_id) for teacher_id in teacher_ids for day in DAYS for slot_id in slot_ids}
    missing_availability = expected_availability - availability_pairs
    if missing_availability:
        result.errors.append(f"TeacherAvailability 缺少 {len(missing_availability)} 个教师/星期/时段组合")

    course_rows = { _clean(row["CourseID"]): row for _, row in t["Courses"].iterrows() }
    options_by_course: Dict[str, Set[str]] = {course_id: set() for course_id in course_ids}
    selection_group_by_option: Dict[tuple[str, str], str] = {}
    for index, row in t["CourseOptions"].iterrows():
        course_id, option = _clean(row["CourseID"]), _clean(row["OptionName"])
        if not option:
            result.errors.append(f"CourseOptions 第 {index + 2} 行：OptionName 不能为空")
        elif option in options_by_course.get(course_id, set()):
            result.errors.append(f"CourseOptions 第 {index + 2} 行：同一课程的小项重复：{option}")
        options_by_course.setdefault(course_id, set()).add(option)
        selection_group = _clean(row["SelectionGroup"])
        if not selection_group:
            result.errors.append(f"CourseOptions 第 {index + 2} 行：SelectionGroup 不能为空")
        selection_group_by_option[(course_id, option)] = selection_group
        required_building, required_room = _clean(row["RequiredBuildingID"]), _clean(row["RequiredRoomID"])
        if required_building and required_building not in building_ids:
            result.errors.append(f"CourseOptions 第 {index + 2} 行：RequiredBuildingID 不存在")
        if required_room and required_room not in room_ids:
            result.errors.append(f"CourseOptions 第 {index + 2} 行：RequiredRoomID 不存在")
        if required_building and required_room:
            room = t["Classrooms"].loc[t["Classrooms"]["RoomID"].map(_clean) == required_room].iloc[0]
            if _clean(room["BuildingID"]) != required_building:
                result.errors.append(f"CourseOptions 第 {index + 2} 行：限定教室不在限定楼内")

    for course_id, row in course_rows.items():
        course_type = _clean(row["CourseType"])
        option_count = _as_positive_int(row["OptionCount"]) or 0
        pick_count = _as_positive_int(row["StudentPickCount"]) or 0
        actual_options = len(options_by_course.get(course_id, set()))
        if course_type not in COURSE_TYPES:
            result.errors.append(f"Courses：{course_id} 的 CourseType 无效")
        if _as_positive_int(row["PeriodsPerWeek"]) is None:
            result.errors.append(f"Courses：{course_id} 的 PeriodsPerWeek 必须为正整数")
        for column in ["SplitAB"]:
            try:
                if int(float(row[column])) not in {0, 1}:
                    raise ValueError
            except (TypeError, ValueError):
                result.errors.append(f"Courses：{course_id} 的 {column} 只能是 0 或 1")
        split = _clean(row["SplitAB"]) in {"1", "1.0"}
        same_time = _clean(row["AllGradeSameTime"]) in {"1", "1.0"}
        try:
            type_count = int(float(row["TypeCount"] or 0))
        except (TypeError, ValueError):
            type_count = -1
            result.errors.append(f"Courses：{course_id} 的 TypeCount 必须为非负整数")
        if split and same_time:
            result.errors.append(f"Courses：{course_id} 选择全年级统一时间时不能 A/B 分班")
        if split and course_type in {"required", "cas"}:
            result.errors.append(f"Courses：{course_id} 为普通必修或 CAS，不能 A/B 分班")
        if not split and not same_time and course_type != "cas" and type_count < 2:
            result.errors.append(f"Courses：{course_id} 未分 A/B 且非全年级统一时间，TypeCount 至少应为 2")
        if (split or same_time or course_type == "cas") and type_count:
            result.errors.append(f"Courses：{course_id} 已分 A/B、全年级统一或 CAS 时，TypeCount 必须为 0")
        for column in ["AllGradeSameTime", "RequireConsecutive"]:
            try:
                if int(float(row[column])) not in {0, 1}:
                    raise ValueError
            except (TypeError, ValueError):
                result.errors.append(f"Courses：{course_id} 的 {column} 只能是 0 或 1")
        if course_type in {"level", "choice"}:
            if option_count != actual_options:
                result.errors.append(f"Courses：{course_id} 声明 {option_count} 项，但实际有 {actual_options} 项")
            if not 1 <= pick_count <= option_count:
                result.errors.append(f"Courses：{course_id} 的 StudentPickCount 超出范围")
            if course_type == "level" and pick_count != 1:
                result.errors.append(f"Courses：{course_id} 为 level 课程，StudentPickCount 必须为 1")
        elif actual_options:
            result.errors.append(f"Courses：{course_id} 不是 level/choice，却有 CourseOptions")

    option_pairs = {(course_id, option) for course_id, values in options_by_course.items() for option in values}
    for sheet, id_column in [("TeacherAssignments", "TeacherID"), ("StudentSelections", "StudentID")]:
        for index, row in t[sheet].iterrows():
            course_id, option = _clean(row["CourseID"]), _clean(row["CourseOption"])
            if sheet == "TeacherAssignments" and _clean(course_rows.get(course_id, {}).get("CourseType", "")) == "required" and not option:
                continue
            if (course_id, option) not in option_pairs:
                result.errors.append(f"{sheet} 第 {index + 2} 行：课程小项不存在：{course_id} / {option}")

    if t["TeacherCAS"].duplicated(subset=["TeacherID"], keep=False).any():
        for teacher_id in t["TeacherCAS"].loc[t["TeacherCAS"].duplicated(subset=["TeacherID"], keep=False), "TeacherID"].map(_clean).unique():
            result.errors.append(f"TeacherCAS：教师 {teacher_id} 同时担任多个 CAS 班主任")
    if t["TeacherCAS"].duplicated(subset=["HomeroomID"], keep=False).any():
        for homeroom_id in t["TeacherCAS"].loc[t["TeacherCAS"].duplicated(subset=["HomeroomID"], keep=False), "HomeroomID"].map(_clean).unique():
            result.errors.append(f"TeacherCAS：行政班 {homeroom_id} 同时指定了多位班主任")
    if any(_clean(row["CourseType"]) == "cas" for row in course_rows.values()):
        cas_homerooms = set(t["TeacherCAS"]["HomeroomID"].map(_clean))
        missing_cas_homerooms = homeroom_ids - cas_homerooms
        if missing_cas_homerooms:
            result.errors.append(
                "TeacherCAS：CAS 需要每个行政班都有一位班主任；尚未指定："
                + "、".join(sorted(missing_cas_homerooms))
            )

    selected_by_student_course = t["StudentSelections"].groupby(["StudentID", "CourseID"], dropna=False)
    for (student_id, course_id), group in selected_by_student_course:
        student_id, course_id = _clean(student_id), _clean(course_id)
        if course_id not in course_rows:
            continue
        expected = _as_positive_int(course_rows[course_id]["StudentPickCount"]) or 0
        unique_options = set(group["CourseOption"].map(_clean))
        selected_groups = {
            selection_group_by_option.get((course_id, option), option)
            for option in unique_options
        }
        if len(unique_options) != expected:
            result.errors.append(
                f"StudentSelections：学生 {student_id} 在 {course_id} 中选了 {len(unique_options)} 项，应选 {expected} 项"
            )
        if len(selected_groups) != expected:
            result.errors.append(
                f"StudentSelections：学生 {student_id} 在 {course_id} 中从 {len(selected_groups)} 个选择组中选课，应选 {expected} 个不同学科/组"
            )

    for index, row in t["HomeroomBlockBalance"].iterrows():
        if _clean(row["HomeroomID"]) not in homeroom_ids:
            result.errors.append(f"HomeroomBlockBalance 第 {index + 2} 行：HomeroomID 不存在")

    for course_id, row in course_rows.items():
        if _clean(row["CourseType"]) not in {"level", "choice"}:
            continue
        enrolled = set(t["StudentSelections"].loc[t["StudentSelections"]["CourseID"].map(_clean) == course_id, "StudentID"].map(_clean))
        eligible_students = student_ids
        missing = eligible_students - enrolled
        if missing:
            result.warnings.append(f"{course_id} 有 {len(missing)} 名学生尚未完成选择")

    return result


def main() -> int:
    if len(sys.argv) != 2:
        print("用法：python step1_validate_input.py <timetable-input.xlsx>")
        return 2
    result = validate_input(sys.argv[1])
    print(result.report())
    return 0 if result.is_valid else 1


if __name__ == "__main__":
    raise SystemExit(main())
