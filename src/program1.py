"""
读取 timetable-input Excel，构建结构化模型，计算教室通行不便度，导出多 sheet Excel。

输出（13 个 sheet）：
- 原始输入中除 NoClassTimes / HomeroomBlockBalance / TeacherAvailability / TeacherCAS 外的 11 个 sheet
- TravelInconvenience：每两间教室一行（无序对），含跨楼标记与加倍信息
- TravelMatrix：教室 × 教室 方阵，数值已含跨楼加倍

跨楼惩罚：
- 基础分 = 水平距离 × WEIGHT_HORIZONTAL + 楼层差 × WEIGHT_FLOOR
- 跨楼时基础分再 × CROSS_BUILDING_MULTIPLIER（默认 2.0），以抑制跨楼排课
"""

import math
import sys
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

import pandas as pd

INPUT_FILE = "timetable-input (9).xlsx"
OUTPUT_FILE = "school_data_with_travel.xlsx"

# ---------- 权重与常量 ----------
WEIGHT_HORIZONTAL = 1.0
WEIGHT_FLOOR = 15.0
WALK_SPEED_MPS = 1.2
ROOM_ORDINAL_METERS = 10.0      # 同楼 RoomOrdinal 每单位间距（米）
DEFAULT_CAPACITY = 18

# 跨楼惩罚：不希望学生跨楼，跨楼时最终分数 × 该倍数
CROSS_BUILDING_MULTIPLIER = 2.0

# Keep every input sheet in the output so the scheduler can read the same
# workbook after the travel tables are added.
EXCLUDED_SHEETS: set[str] = set()


# ============================================================
# 1. 结构化模型
# ============================================================

@dataclass
class Classroom:
    room_id: str
    building_id: str
    building_name: str
    floor: int
    room_ordinal: int
    room_name: str = ""
    capacity: int = DEFAULT_CAPACITY


@dataclass
class TimeSlot:
    slot_id: str
    start: str
    end: str


@dataclass
class Course:
    course_id: str            # 唯一 ID：如 "ENG:S" / "CAS"
    base_course_id: str       # 原始 CourseID：如 "ENG"
    option: str               # 原始 CourseOption：如 "S"；无选项时为空
    course_name: str
    periods_per_week: int
    course_type: str
    required_room_id: Optional[str] = None


@dataclass
class Teacher:
    teacher_id: str
    name: str
    course_ids: List[str] = field(default_factory=list)
    homeroom_id: Optional[str] = None


@dataclass
class Student:
    student_id: str
    name: str
    homeroom_id: str
    grade: int
    course_ids: List[str] = field(default_factory=list)


@dataclass
class SchoolData:
    classrooms: Dict[str, Classroom] = field(default_factory=dict)
    time_slots: Dict[str, TimeSlot] = field(default_factory=dict)
    courses: Dict[str, Course] = field(default_factory=dict)
    teachers: Dict[str, Teacher] = field(default_factory=dict)
    students: Dict[str, Student] = field(default_factory=dict)
    building_distance: Dict[Tuple[str, str], float] = field(default_factory=dict)


# ============================================================
# 2. 读取输入
# ============================================================

def _course_unique_id(base_id: str, option) -> str:
    """把 (CourseID, CourseOption) 合成唯一 ID。"""
    if option is None or (isinstance(option, float) and math.isnan(option)):
        return base_id
    option = str(option).strip()
    return f"{base_id}:{option}" if option else base_id


def load_school_data(path: str) -> Tuple[SchoolData, Dict[str, pd.DataFrame]]:
    sheets = pd.read_excel(path, sheet_name=None)
    print(f"Loaded {len(sheets)} sheets: {list(sheets.keys())}")

    school = SchoolData()

    buildings_df = sheets["Buildings"]
    classrooms_df = sheets["Classrooms"]
    dist_df = sheets["BuildingDistances"]
    slots_df = sheets["TimeSlots"]
    courses_df = sheets["Courses"]
    options_df = sheets["CourseOptions"]
    teachers_df = sheets["Teachers"]
    assignments_df = sheets["TeacherAssignments"]
    cas_df = sheets["TeacherCAS"]
    students_df = sheets["Students"]
    selections_df = sheets["StudentSelections"]

    # 2.1 楼名映射
    building_name = dict(zip(buildings_df["BuildingID"], buildings_df["BuildingName"]))

    # 2.2 楼间距离（对称化）
    for _, row in dist_df.iterrows():
        a = row["FromBuildingID"]
        b = row["ToBuildingID"]
        d = float(row["DistanceMeters"])
        school.building_distance[(a, b)] = d
        school.building_distance[(b, a)] = d

    # 2.3 教室
    for _, row in classrooms_df.iterrows():
        bid = row["BuildingID"]
        room = Classroom(
            room_id=row["RoomID"],
            building_id=bid,
            building_name=building_name[bid],
            floor=int(row["Floor"]),
            room_ordinal=int(row["RoomOrdinal"]),
            room_name=str(row.get("RoomName", "")),
            capacity=DEFAULT_CAPACITY,
        )
        school.classrooms[room.room_id] = room

    # 2.4 时间段（P1–P7，不带星期）
    for _, row in slots_df.iterrows():
        school.time_slots[row["TimeSlotID"]] = TimeSlot(
            slot_id=row["TimeSlotID"],
            start=str(row["StartTime"]),
            end=str(row["EndTime"]),
        )

    # 2.5 课程（科目 + 选项 → 唯一 Course）
    for _, c_row in courses_df.iterrows():
        base_id = c_row["CourseID"]
        options = options_df[options_df["CourseID"] == base_id]
        if len(options) == 0:
            # 无选项科目（如 CAS）
            school.courses[base_id] = Course(
                course_id=base_id,
                base_course_id=base_id,
                option="",
                course_name=str(c_row["CourseName"]),
                periods_per_week=int(c_row["PeriodsPerWeek"]),
                course_type=str(c_row["CourseType"]),
            )
        else:
            for _, o_row in options.iterrows():
                option = str(o_row["OptionName"])
                unique_id = f"{base_id}:{option}"
                req = o_row.get("RequiredRoomID")
                req = str(req).strip() if pd.notna(req) and str(req).strip() else None
                school.courses[unique_id] = Course(
                    course_id=unique_id,
                    base_course_id=base_id,
                    option=option,
                    course_name=f"{c_row['CourseName']} - {option}",
                    periods_per_week=int(c_row["PeriodsPerWeek"]),
                    course_type=str(c_row["CourseType"]),
                    required_room_id=req,
                )

    # 2.6 教师
    for _, t_row in teachers_df.iterrows():
        school.teachers[t_row["TeacherID"]] = Teacher(
            teacher_id=t_row["TeacherID"],
            name=str(t_row["TeacherName"]),
        )
    for _, a_row in assignments_df.iterrows():
        tid = a_row["TeacherID"]
        unique_id = _course_unique_id(a_row["CourseID"], a_row["CourseOption"])
        if tid in school.teachers and unique_id not in school.teachers[tid].course_ids:
            school.teachers[tid].course_ids.append(unique_id)
    for _, row in cas_df.iterrows():
        tid = row["TeacherID"]
        if tid in school.teachers:
            school.teachers[tid].homeroom_id = str(row["HomeroomID"])

    # 2.7 学生（年级从 HomeroomID 前缀解析）
    for _, s_row in students_df.iterrows():
        hid = str(s_row["HomeroomID"])
        grade = int(hid.split("-")[0])
        school.students[s_row["StudentID"]] = Student(
            student_id=s_row["StudentID"],
            name=str(s_row["StudentName"]),
            homeroom_id=hid,
            grade=grade,
        )
    for _, sel in selections_df.iterrows():
        sid = sel["StudentID"]
        unique_id = _course_unique_id(sel["CourseID"], sel["CourseOption"])
        if sid in school.students and unique_id not in school.students[sid].course_ids:
            school.students[sid].course_ids.append(unique_id)

    return school, sheets


# ============================================================
# 3. 通行不便度（含跨楼加倍）
# ============================================================

def _is_cross_building(r1: Classroom, r2: Classroom) -> bool:
    return r1.building_id != r2.building_id


def _horizontal_distance(school: SchoolData, r1: Classroom, r2: Classroom) -> float:
    """同楼用 RoomOrdinal 差 × 10 米；跨楼查 BuildingDistances。"""
    if r1.building_id == r2.building_id:
        return abs(r1.room_ordinal - r2.room_ordinal) * ROOM_ORDINAL_METERS
    return school.building_distance[(r1.building_id, r2.building_id)]


def travel_inconvenience(school: SchoolData, r1: Classroom, r2: Classroom) -> float:
    """
    两间教室之间的通行不便度分数。

    基础分 = 水平距离 × WEIGHT_HORIZONTAL + 楼层差 × WEIGHT_FLOOR
    跨楼时基础分再 × CROSS_BUILDING_MULTIPLIER，以抑制跨楼排课。
    """
    if r1.room_id == r2.room_id:
        return 0.0

    horizontal = _horizontal_distance(school, r1, r2)
    floor_diff = abs(r1.floor - r2.floor)
    base_score = horizontal * WEIGHT_HORIZONTAL + floor_diff * WEIGHT_FLOOR

    if _is_cross_building(r1, r2):
        return base_score * CROSS_BUILDING_MULTIPLIER
    return base_score


def travel_inconvenience_with_estimate(
    school: SchoolData, r1: Classroom, r2: Classroom
) -> Dict[str, float]:
    """
    返回不便度详情。
    - inconvenience_score 是已含跨楼加倍后的最终分数
    - horizontal_distance_m / floor_difference 仍为真实值
    - cross_building_multiplier 用于说明是否被加倍
    """
    score = travel_inconvenience(school, r1, r2)

    if r1.room_id == r2.room_id:
        horizontal, floor_diff = 0.0, 0
    else:
        horizontal = _horizontal_distance(school, r1, r2)
        floor_diff = abs(r1.floor - r2.floor)

    cross = _is_cross_building(r1, r2) and r1.room_id != r2.room_id
    multiplier = CROSS_BUILDING_MULTIPLIER if cross else 1.0

    est = score / WALK_SPEED_MPS if WALK_SPEED_MPS > 0 else float("inf")

    return {
        "inconvenience_score": round(score, 2),
        "horizontal_distance_m": round(horizontal, 1),
        "floor_difference": floor_diff,
        "cross_building": int(cross),
        "cross_building_multiplier": multiplier,
        "estimated_walk_seconds": round(est, 1),
    }


# ============================================================
# 4. 构建输出表
# ============================================================

def build_travel_frames(school: SchoolData):
    rooms = list(school.classrooms.values())

    # 4.1 长表：每两间教室一行（无序对）
    rows = []
    for i in range(len(rooms)):
        for j in range(i + 1, len(rooms)):
            r1, r2 = rooms[i], rooms[j]
            d = travel_inconvenience_with_estimate(school, r1, r2)
            rows.append({
                "room_a": r1.room_id,
                "building_a": r1.building_name,
                "floor_a": r1.floor,
                "room_b": r2.room_id,
                "building_b": r2.building_name,
                "floor_b": r2.floor,
                "cross_building": d["cross_building"],
                "cross_building_multiplier": d["cross_building_multiplier"],
                "inconvenience_score": d["inconvenience_score"],
                "horizontal_distance_m": d["horizontal_distance_m"],
                "floor_difference": d["floor_difference"],
                "estimated_walk_seconds": d["estimated_walk_seconds"],
            })
    df_travel = pd.DataFrame(rows)

    # 4.2 方阵（数值已含跨楼加倍）
    room_ids = [r.room_id for r in rooms]
    matrix = pd.DataFrame(0.0, index=room_ids, columns=room_ids)
    for i, r1 in enumerate(rooms):
        for j, r2 in enumerate(rooms):
            matrix.iloc[i, j] = round(travel_inconvenience(school, r1, r2), 2)
    matrix.index.name = "room_id"
    return df_travel, matrix


# ============================================================
# 5. 主流程
# ============================================================

def main():
    input_file = sys.argv[1] if len(sys.argv) > 1 else INPUT_FILE
    output_file = sys.argv[2] if len(sys.argv) > 2 else OUTPUT_FILE
    school, sheets = load_school_data(input_file)

    print(f"Rooms: {len(school.classrooms)}")
    print(f"Time slots: {len(school.time_slots)}")
    print(f"Courses (subject + option): {len(school.courses)}")
    print(f"Teachers: {len(school.teachers)}")
    print(f"Students: {len(school.students)}")

    df_travel, matrix = build_travel_frames(school)

    output_sheets = {name: df for name, df in sheets.items() if name not in EXCLUDED_SHEETS}
    output_sheets["TravelInconvenience"] = df_travel
    output_sheets["TravelMatrix"] = matrix

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        for name, df in output_sheets.items():
            df.to_excel(writer, sheet_name=name, index=(name == "TravelMatrix"))

    print(f"\nExported: {output_file} with {len(output_sheets)} sheets")
    for name, df in output_sheets.items():
        print(f"  {name}: {df.shape[0]} rows x {df.shape[1]} columns")


if __name__ == "__main__":
    main()
