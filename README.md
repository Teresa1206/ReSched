# ReSched — SHSID Timetable System

A self-contained local web app for generating, reviewing, and temporarily adjusting school timetables. Enter or import data on the input page, run the three-step pipeline, and get a timetable preview plus downloadable results.

## Requirements

- Python 3.11+ (verified on Python 3.12)
- Dependencies: `pandas`, `openpyxl`, `ortools`

The bundled `.venv` already has the dependencies installed. To reinstall them:

```bash
.venv/bin/python -m pip install pandas openpyxl ortools
```

## Start the server

Run from the project root (the directory containing this README):

```bash
.venv/bin/python src/server.py
```

If you do not have `.venv`, you can use the system Python:

```bash
python3 src/server.py
```

The server listens on:

```
http://127.0.0.1:8765/
```

Your browser usually opens automatically; if not, open the URL above manually.

> The pages rely on the local server's API. Open `http://127.0.0.1:8765/` instead of double-clicking `src/input.html`.

## Stop the server

```bash
# Press Ctrl+C in the terminal running the server, or:
pkill -f "src/server.py"
```

Check whether the port is still in use:

```bash
lsof -nP -iTCP:8765 -sTCP:LISTEN
```

If the port is occupied, kill the process or change `PORT` in `src/server.py` (default: 8765).

## Workflow

After opening `http://127.0.0.1:8765/`:

1. **Load example data**
   - Click **Load schedulable example (15 teachers / 40 students)**.
   - This only fills the form with sample input; it does not run the scheduler.

2. **Confirm the input source**
   - Scroll to the **Run the schedule** section.
   - Make sure **Step 1 input source** shows **Using the form below**.
   - If you previously chose an Excel file, click **Use the form below** to switch back to the form data.

3. **Run step 1: validate input**
   - Click **Run step 1** (`step1_validate_input.py`).
   - **Run step 2** unlocks only after validation passes.

4. **Run step 2: compute room distances and prepare the model**
   - Click **Run step 2**.
   - This runs `program1.py` first (generating `TravelInconvenience` / `TravelMatrix`), then `step2_prepare_model.py`.
   - **Run step 3** unlocks on success.

5. **Run step 3: OR-Tools scheduling**
   - Click **Run step 3**.
   - `step3_schedule.py` assigns times, rooms, teachers, and A/B students; `build_schedule.py` then renders the preview page from the `schedule.html` template.

6. **View results**
   - **View generated timetable** is greyed out until Step 3 succeeds; once it does, clicking it opens `/preview.html`.
   - Download from the step 3 result area:
     - `*_schedule.xlsx`
     - `*_schedule.json`

## Top buttons

| Button | Action |
|---|---|
| **Save to this browser** | Saves the current input to browser `localStorage`; it survives a page refresh |
| **Export Excel** | Exports the normalized `timetable-input.xlsx`, which **Step 1 can read directly** |
| **Download JSON** | Exports `timetable-input.json`, an **input snapshot** rather than a scheduling result; the page has no "import JSON" button |
| **Load schedulable example** | Loads synthetic sample data (15 teachers / 40 students) |
| **Clear all input** | Clears the current input |

> To save or reuse input, prefer **Export Excel**; later select that `.xlsx` in **Step 1 input source**.
>
> The `*_schedule.json` downloaded after Step 3 is the **scheduling result**, suitable for independent review, analysis, or temporary adjustments.

## Generated files

Each run writes intermediate files and logs to `src/runs/<run ID>/`:

| File | Description |
|---|---|
| `<ID>_input.xlsx` | Input workbook generated from the page form |
| `<ID>_input_travel.xlsx` | Room travel matrix produced by `program1.py` |
| `<ID>_input_travel_prepared.xlsx` / `.json` | Solver input produced by `step2_prepare_model.py` |
| `<ID>_input_travel_prepared_schedule.xlsx` | Timetable result workbook |
| `<ID>_input_travel_prepared_schedule.json` | Timetable result JSON |
| `step1.log` / `step2.log` / `step3.log` / `build.log` | Logs for each step |

`src/preview_schedule.html` is the preview page for the latest run (overwritten on every Step 3).

## Run the steps from the command line

Each step can also run outside the web server (from the project root):

```bash
# Step 1: validate input
.venv/bin/python src/step1_validate_input.py timetable-input.xlsx

# Step 2: travel matrix + solver input
.venv/bin/python src/program1.py timetable-input.xlsx timetable-input_travel.xlsx
.venv/bin/python src/step2_prepare_model.py timetable-input_travel.xlsx

# Step 3: scheduling
.venv/bin/python src/step3_schedule.py timetable-input_travel_prepared.xlsx

# Independently re-evaluate the timetable
.venv/bin/python src/step3_evaluate_result.py \
    timetable-input_travel_prepared.xlsx \
    timetable-input_travel_prepared_schedule.json

# Build the timetable preview page
.venv/bin/python src/build_schedule.py \
    timetable-input_travel_prepared_schedule.json \
    timetable-input_travel_prepared.xlsx \
    timetable-input_travel.xlsx \
    schedule.html src/preview_schedule.html
```

## FAQ

### `ModuleNotFoundError: No module named 'pandas'`

The dependencies are not installed for the current Python. Use the project venv:

```bash
.venv/bin/python -m pip install pandas openpyxl ortools
.venv/bin/python src/server.py
```

### `OSError: [Errno 48] Address already in use`

Port 8765 is already in use:

```bash
lsof -nP -iTCP:8765 -sTCP:LISTEN
pkill -f "src/server.py"
```

Then restart the server.

### `python: command not found`

On macOS there may be no `python` command; use `python3` or the virtual environment:

```bash
python3 src/server.py
# or
.venv/bin/python src/server.py
```

### `/preview.html` returns 404

The preview page is generated only after Step 3 succeeds. Complete Step 3 first, then click **View generated timetable**.

### Step 3 fails

Check the log for the corresponding run:

```bash
src/runs/<run ID>/step3.log
```

Send the error from the log to continue troubleshooting.

## Data notes

- The **15 teachers / 40 students** on the page is synthetic sample data for validating the workflow, not real SHSID data.
- Real scheduling requires complete data following the `step1` schema: `Buildings`, `Classrooms`, `BuildingDistances`, `TimeSlots`, `NoClassTimes`, `Homerooms`, `HomeroomBlockBalance`, `Courses`, `CourseOptions`, `Teachers`, `TeacherAssignments`, `TeacherAvailability`, `TeacherCAS`, `Students`, `StudentSelections`.
- `src/input_original.html` is a backup of the original input page in Chinese; the current workflow uses `src/input.html`.

## Related files

- `Track 1 Problem Statement.pdf` — problem statement and judging requirements
- `src/input.html` — the input page in use (English)
- `schedule.html` — timetable page template (the preview page is generated from it)
