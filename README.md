# ReSched — SHSID 排课系统

本项目用于生成、检查和临时调整学校课表。当前在 macOS 上可用的是 `src/` 目录；`ReScheduler/` 是较早的中文 Windows 分发版。

## 目录说明

| 路径 | 说明 |
|---|---|
| `src/` | 当前使用的版本：英文 UI、自包含 Web 服务、包含 `program1.py` 通行矩阵和 `build_schedule.py` |
| `ReScheduler/` | 较早的中文 Windows 版本：含第四步临时调课，但 Web 服务依赖缺失的 `temporary_web.py`，`.cmd` 里硬编码了 Windows 路径 |
| `current-html/SHSID.html` | 当前学校系统导出的**一位学生**课表快照，不能直接导入排课流程 |
| `Track 1 Problem Statement.pdf` | 题目说明 |
| `start.command` | macOS 启动脚本（自动检查 Python 版本和依赖） |

## 环境要求

- Python 3.11+（本项目在 Python 3.12 上验证）
- 依赖：`pandas`、`openpyxl`、`ortools`

如果使用项目自带的 `.venv`，依赖已经装好。重新安装：

```bash
cd /Users/teresayou/Documents/myws/test-project
.venv/bin/python -m pip install pandas openpyxl ortools
```

`ReScheduler/requirements.txt` 也列出了同样的依赖。

## 启动服务

### 方式 A：直接命令（当前使用）

```bash
cd /Users/teresayou/Documents/myws/test-project
python src/server.py
```

如果系统里没有 `python` 命令，用：

```bash
.venv/bin/python src/server.py
# 或
python3 src/server.py
```

### 方式 B：macOS 启动脚本

```bash
cd /Users/teresayou/Documents/myws/test-project
./start.command
```

也可以直接在 Finder 里双击 `start.command`。

启动后服务监听：

```
http://127.0.0.1:8765/
```

浏览器通常会自动打开；如果没有，手动访问上面的地址。

### 停止服务

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
   - 点击 **Run step 1**。
   - 校验通过后，**Run step 2** 才会解锁。

4. **运行第二步：准备模型 + 计算教室距离**
   - 点击 **Run step 2**。
   - 这一步会运行 `program1.py`（生成 `TravelInconvenience`），再运行 `step2_prepare_model.py`。
   - 成功后 **Run step 3** 解锁。

5. **运行第三步：OR-Tools 排课**
   - 点击 **Run step 3**。
   - 求解器会分配时间、教室、教师和 A/B 学生，然后由 `build_schedule.py` 生成课表页面。

6. **查看结果**
   - **View generated timetable** 在 Step 3 成功前是灰色不可点击的；Step 3 成功后会变成蓝色可用状态。
   - 点击它打开 `/preview.html`。
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
| `*_input.xlsx` | 页面表单同步生成的输入工作簿 |
| `*_travel.xlsx` | `program1.py` 生成的教室通行矩阵 |
| `*_travel_prepared.xlsx` | `step2_prepare_model.py` 生成的求解器输入 |
| `*_schedule.xlsx` | 排课结果表 |
| `*_schedule.json` | 排课结果 JSON |
| `step1.log` / `step2.log` / `step3.log` | 各步日志 |

`src/preview_schedule.html` 是最近一次生成的课表预览页面。

## 常见问题

### `ModuleNotFoundError: No module named 'pandas'`

依赖没装到当前 Python：用项目 venv：

```bash
cd /Users/teresayou/Documents/myws/test-project
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

### `ModuleNotFoundError: No module named 'temporary_web'`

你运行的是 `ReScheduler/resched_web_server.py`。该模块缺失，且不是 pip 包。改用：

```bash
python src/server.py
```

### `python: command not found`

macOS 上可能只有 `python3`：

```bash
python3 src/server.py
# 或
.venv/bin/python src/server.py
```

### 第三步失败

查看日志：

```bash
src/runs/<运行ID>/step3.log
```

把日志里的错误信息发出来即可继续定位。

## 数据说明

- 页面里的 **15 teachers / 40 students** 是合成示例，仅用于验证流程，不是真实 SHSID 数据。
- 真实排课需要按 `step1` 的 schema 准备完整数据：`Buildings`、`Classrooms`、`BuildingDistances`、`TimeSlots`、`NoClassTimes`、`Homerooms`、`HomeroomBlockBalance`、`Courses`、`CourseOptions`、`Teachers`、`TeacherAssignments`、`TeacherAvailability`、`TeacherCAS`、`Students`、`StudentSelections`。
- `current-html/SHSID.html` 只包含一位学生的课表，不能直接导入；需要从学校系统导出完整数据后重新录入/转换。

## 相关文件

- `Track 1 Problem Statement.pdf` — 题目与评审要求
- `ReScheduler/requirements.txt` — 依赖清单
- `src/input.html` — 实际使用的输入页面（英文）
- `src/input_original.html` — 中文原版输入页面备份（不使用）
