# ReSched — SHSID 排课系统

本项目用于生成、检查和临时调整学校课表。当前版本是一个自包含的本地 Web 服务：在输入页录入或导入数据，依次运行三步流水线，生成课表预览和可下载结果。

## 环境要求

- Python 3.11+（本项目在 Python 3.12 上验证）
- 依赖：`pandas`、`openpyxl`、`ortools`

如果使用项目自带的 `.venv`，依赖已经装好。重新安装：

```bash
.venv/bin/python -m pip install pandas openpyxl ortools
```

## 启动服务

在项目根目录（本 README 所在目录）执行：

```bash
.venv/bin/python src/server.py
```

如果没有 `.venv`，也可以用系统 Python：

```bash
python3 src/server.py
```

服务启动后监听：

```
http://127.0.0.1:8765/
```

浏览器通常会自动打开；如果没有，手动访问上面的地址。

> 页面依赖本地服务提供的接口，请通过 `http://127.0.0.1:8765/` 访问，不要直接双击打开 `src/input.html`。

## 停止服务

```bash
# 在运行服务的终端里按 Ctrl+C
# 或者：
pkill -f "src/server.py"
```

检查端口是否仍被占用：

```bash
lsof -nP -iTCP:8765 -sTCP:LISTEN
```

如果端口被占用，先 kill 对应进程，或修改 `src/server.py` 里的 `PORT`（默认 8765）。

## 页面操作流程

打开 `http://127.0.0.1:8765/` 后：

1. **加载示例数据**
   - 点击 **Load schedulable example (15 teachers / 40 students)**。
   - 这只是把示例输入填进表单，不会自动排课。

2. **确认输入来源**
   - 滚动到 **Run the schedule** 区域。
   - 确认 **Step 1 input source** 显示 **Using the form below**。
   - 如果你之前选过 Excel 文件，点 **Use the form below** 切回表单数据。

3. **运行第一步：校验输入**
   - 点击 **Run step 1**（`step1_validate_input.py`）。
   - 校验通过后，**Run step 2** 才会解锁。

4. **运行第二步：计算教室距离并准备模型**
   - 点击 **Run step 2**。
   - 这一步先运行 `program1.py`（生成 `TravelInconvenience` / `TravelMatrix`），再运行 `step2_prepare_model.py`。
   - 成功后 **Run step 3** 解锁。

5. **运行第三步：OR-Tools 排课**
   - 点击 **Run step 3**。
   - `step3_schedule.py` 分配时间、教室、教师和 A/B 学生，随后 `build_schedule.py` 用 `schedule.html` 模板生成预览页。

6. **查看结果**
   - **View generated timetable** 在 Step 3 成功前是灰色不可点击的；成功后可点击，打开 `/preview.html`。
   - 在第三步结果区域下载：
     - `*_schedule.xlsx`
     - `*_schedule.json`

## 顶部按钮说明

| 按钮 | 作用 |
|---|---|
| **Save to this browser** | 把当前输入保存到浏览器 `localStorage`，刷新页面后仍在 |
| **Export Excel** | 导出规范化的 `timetable-input.xlsx`；这是 **Step 1 可以直接读取**的输入文件 |
| **Download JSON** | 导出 `timetable-input.json`，是**输入数据快照**，不是排课结果；页面没有“导入 JSON”按钮 |
| **Load schedulable example** | 载入合成示例数据（15 位教师 / 40 位学生） |
| **Clear all input** | 清空当前输入 |

> 想保存或复用输入，优先用 **Export Excel**；以后在 **Step 1 input source** 里选择该 `.xlsx` 即可。
>
> Step 3 之后下载的 `*_schedule.json` 才是**排课结果**，可用于独立复核、分析或临时调课。

## 生成的文件

服务运行时会在 `src/runs/<运行ID>/` 下保存中间文件和日志：

| 文件 | 说明 |
|---|---|
| `<ID>_input.xlsx` | 页面表单同步生成的输入工作簿 |
| `<ID>_input_travel.xlsx` | `program1.py` 生成的教室通行矩阵 |
| `<ID>_input_travel_prepared.xlsx` / `.json` | `step2_prepare_model.py` 生成的求解器输入 |
| `<ID>_input_travel_prepared_schedule.xlsx` | 排课结果表 |
| `<ID>_input_travel_prepared_schedule.json` | 排课结果 JSON |
| `step1.log` / `step2.log` / `step3.log` / `build.log` | 各步日志 |

`src/preview_schedule.html` 是最近一次生成的课表预览页面（每次 Step 3 覆盖）。

## 命令行单独运行

各步骤也可以脱离 Web 服务单独执行（在项目根目录运行）：

```bash
# 第一步：校验输入
.venv/bin/python src/step1_validate_input.py timetable-input.xlsx

# 第二步：通行矩阵 + 求解器输入
.venv/bin/python src/program1.py timetable-input.xlsx timetable-input_travel.xlsx
.venv/bin/python src/step2_prepare_model.py timetable-input_travel.xlsx

# 第三步：排课
.venv/bin/python src/step3_schedule.py timetable-input_travel_prepared.xlsx

# 独立复算排课结果
.venv/bin/python src/step3_evaluate_result.py \
    timetable-input_travel_prepared.xlsx \
    timetable-input_travel_prepared_schedule.json

# 生成课表预览页
.venv/bin/python src/build_schedule.py \
    timetable-input_travel_prepared_schedule.json \
    timetable-input_travel_prepared.xlsx \
    timetable-input_travel.xlsx \
    schedule.html src/preview_schedule.html
```

## 常见问题

### `ModuleNotFoundError: No module named 'pandas'`

依赖没装到当前 Python，改用项目 venv：

```bash
.venv/bin/python -m pip install pandas openpyxl ortools
.venv/bin/python src/server.py
```

### `OSError: [Errno 48] Address already in use`

端口 8765 已有服务在跑：

```bash
lsof -nP -iTCP:8765 -sTCP:LISTEN
pkill -f "src/server.py"
```

然后重新启动。

### `python: command not found`

macOS 上可能没有 `python` 命令，改用 `python3` 或直接使用虚拟环境：

```bash
python3 src/server.py
# 或
.venv/bin/python src/server.py
```

### 打开 `/preview.html` 显示 404

预览页只在第三步成功后生成；先完成 Step 3，再点击 **View generated timetable**。

### 第三步失败

查看对应运行的日志：

```bash
src/runs/<运行ID>/step3.log
```

把日志里的错误信息发出来即可继续定位。

## 数据说明

- 页面里的 **15 teachers / 40 students** 是合成示例，仅用于验证流程，不是真实 SHSID 数据。
- 真实排课需要按 `step1` 的 schema 准备完整数据：`Buildings`、`Classrooms`、`BuildingDistances`、`TimeSlots`、`NoClassTimes`、`Homerooms`、`HomeroomBlockBalance`、`Courses`、`CourseOptions`、`Teachers`、`TeacherAssignments`、`TeacherAvailability`、`TeacherCAS`、`Students`、`StudentSelections`。
- `src/input_original.html` 是中文原版输入页的备份，当前流程使用 `src/input.html`。

## 相关文件

- `Track 1 Problem Statement.pdf` — 题目与评审要求
- `src/input.html` — 实际使用的输入页面（英文）
- `schedule.html` — 课表页面模板（预览页由它生成）
