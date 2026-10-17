# Phase 1 plan: from VCL spec to Excel batch, with no hand-written VBA

**Status: proposal, waiting for Boss's approval. No code until it's approved.**

Phase 1 (Nov 2026, roadmap) turns the hand-built pilot into a pipeline. Goal: one command
builds a paste-ready batch and its expected results straight from the specs:

```
python harness/batch_builder.py --batch 001 specs/task_01.vcl specs/task_02.vcl specs/task_03.vcl specs/task_04.vcl
```

Proposed Gate 1: batch_001 and batch_002 (all 10 specs) pass in Excel 2016, and every
measure matches the expected file.

## 1. The pipeline

```
spec.vcl ── parser + validator ──┬── op compiler ───── the procedure (VBA)
             (errors stop it)    ├── test compiler ─── one test procedure per TEST
                                 ├── fixtures.py ───── fixture builders (Excel-proven in the pilot)
                                 └── simulator ─────── expected results, in Python
ForgeTest.bas engine + all of the above ── lint ──> batches/batch_NNN.bas + .expected.csv
```

The pilot proved the batch shape in Excel. Phase 1 makes a program write that shape instead
of me.

## 2. OP to VBA (`compiler/codegen.py` + `templates/ops/`)

- **One template per op**, in `templates/ops/<OP>.bas`, with named slots. For example:
  - `SET_VALUE.bas`: `{target}.Value = {value}`
  - `DICT_ADD.bas`: `ForgeDictAdd {dict}, {key}, {value}, "{mode}"`
  - `LOOP_ROWS.bas`: `For {var} = {from} To {to}{step}` ... `Next {var}` (dir=up adds `Step -1`)
- **Structural ops** (IF, ELSE, LOOP_ROWS, FOR_EACH, SELECT_CASE) wrap their nested steps; the
  compiler walks the parser's OP tree.
- **Value functions are expression rules, not templates**:
  - `CELL($ws, r, c)` becomes `ws.Cells(r, c).Value` when read, and `ws.Cells(r, c)` when it is a target.
  - `COL` becomes `ForgeCol`, `IS_NUMBER` becomes `ForgeIsNumber`, and so on.
  - `TRIM(x)` becomes `Trim$(CStr(x))`.
  - Arithmetic on a cell value is wrapped in `CDbl()`, only where the spec guards it with IS_NUMBER (rule 15).
- **Procedure skeleton**, exactly the pilot's `BuildShortageList`:
  - Dims from ARGS and VAR. A Dictionary becomes a late-bound `Object` with case-insensitive keys.
  - A PRE block that logs once and returns -1.
  - PERF save and apply.
  - Numbered body lines (10, 20, ...) so `Erl` works.
  - The standard Done/Fail handler, which restores PERF.
- **Templates to build:** the 10 specs use only 13 ops (LET, IF, ELSE, SET_VALUE, LOOP_ROWS,
  DICT_ADD, DICT_GET, RETURN, CLEAR, COPY_ROW, FORMAT, SORT, FOR_EACH) and 14 value functions.
  The roadmap's "15 templates" are those 13 plus DELETE_ROW and EXIT_IF from the Orchestrator example.
- **Not covered yet:** an op without a template stops the build with a clear message. NOTE lines
  are never compiled; such a spec is marked "needs model" (March), not guessed.

## 3. Assertion to VBA (`harness/testgen.py`)

- **Each TEST becomes one test procedure in the pilot's shape:**
  1. Begin the test.
  2. Build the fixture.
  3. Take the PRE_ measures.
  4. `ForgeCall` the procedure `runs=n` times.
  5. Run one `ForgeCheck` per AND-part.
  6. End the test. The test procedure also has its own error handler.
- **Each AND-part is `left op right`.** Each side is arithmetic over measures, RESULT, ERROR_COUNT
  and numbers. Each check is named exactly as written in the spec, so Excel's Description column
  lines up with the expected file.
- **Measures map one to one onto engine helpers:**
  - `COUNTA(S![H])` becomes `ForgeCountA(ForgeHeaderRange(wb, "S", "H"))`.
  - `VALUE(S![H]2)` becomes `ForgeValue(ForgeHeaderCell(...))`.
  - `FILL_COUNT(..., #FF0000)` becomes `ForgeFillCount(..., "#FF0000")`.
  - COUNTIF, COUNTIFS, SUM and SUMIFS go through WorksheetFunction.
- **`PRE_` measures** are taken once before the first run. Identical ones are shared.
- **Test IDs** run in order through the batch (`B001-T01` ... `B001-T15`). The Spec column names
  the spec, and each test procedure's comment gives the spec line.

## 4. Fixtures

- `harness/fixtures.py` is reused as it is, since the pilot proved it in Excel.
- There is one builder per fixture file, shared by all tests that use it. Every test still gets a fresh workbook.
- batch_001 is the first batch to use the `'` text guard (values like `W-01`). That is new in
  Excel, so it is worth watching.

## 5. Expected results: a VCL simulator (`harness/simulate.py`)

- **What it does:** it runs a spec on the fixture data in Python. It holds a small sheet model
  (values and fill colours) and gives every op the same meaning as its VBA template.
- **Excel's rules, where they matter:**
  - Text that looks like a number becomes a number when it is written to a cell.
  - COUNTIF and COUNTIFS match criteria the way Excel does.
  - Sorts put blanks last.
  - Dictionary keys ignore case.
- **Output:** the measures and PASS/FAIL for every TEST, written to `batch_NNN.expected.csv` like the pilot's.
- **Two jobs:**
  1. Judge Excel's results line by line.
  2. Catch a bad spec or a compiler bug in seconds, before it costs a round trip to the work PC.
     If the simulator says a TEST fails, we don't send that batch.
- **TODAY:** task_06, task_09 and task_10 depend on it, so Python and Excel need the same date.
  See decision 2.

## 6. Testing without Excel

| Piece | Unit tests | Proven-in-Excel reference |
|---|---|---|
| op compiler | a fixed expected VBA snippet per op and value function | compiling task_01 must reproduce the pilot's `BuildShortageList` |
| test compiler | a fixed expected output per assertion form | compiling task_01's TESTs must reproduce the pilot's 4 test procedures |
| simulator | per-op meaning; Excel criteria and sort rules | must match the pilot's Excel numbers in `data/results/batch_000_pilot.csv` (expected 26 / 27 / 3 / 986 / 661) |
| `harness/lint.py` | planted faults, as `assemble.py` does today, plus If/End If and For/Next balance, a ban list of Microsoft 365 functions, and a check that every loop has a bound | - |
| `batch_builder.py` | end to end: specs in, batch and expected file out | CI `--check` keeps committed batches in sync |

CI runs all of it on every push. Excel only sees a batch once everything is green.

## 7. Build order (one small PR each)

1. **`harness/lint.py`**: move `assemble.py`'s paste checks into it and add the new checks.
2. **`harness/simulate.py`**: enough for task_01 to task_04. It must reproduce the pilot's numbers.
3. **`harness/testgen.py`**: it must reproduce the pilot's 4 test procedures.
4. **`compiler/codegen.py` + templates for task_01's ops**: it must reproduce the pilot's procedure.
5. **`harness/batch_builder.py`**: the command above. The hand-written `batch_000` stays as the archived pilot.
6. **Templates for FORMAT, FOR_EACH and SORT**, then build batch_001. You run it in Excel, and
   we expect a fix round or two.
7. **`harness/results_import.py`**: your pasted Results go in and are saved to `data/results/`.
   They are compared with the expected file, and passing specs are written to `data/verified/`.
8. **Remaining templates and value functions**, then batch_002 (task_05 to task_10).
   Freeze the op vocabulary after reviewing the NOTE list.

## 8. First batch: batch_001

I agree with task_02 to task_04. Between them they exercise:
- row fills with FILL_COUNT (task_02)
- dictionaries, FOR_EACH over KEYS, SORT and number formats (task_03)
- five sheet arguments and Variant values (task_04)

I'd add **task_01, compiled rather than hand-written, as a control**. If task_01 passes and the
others fail, the fault is in the new templates, not the pipeline. That makes 15 tests (4+3+4+4),
a run of about the same length as the pilot.

batch_002 then takes task_05 to task_10: COPY_ROW, COUNTIFS/SUMIFS and TODAY.

## 9. Decisions I need from you

1. **batch_001 contents:** task_01 (compiled, as control) plus task_02 to task_04?
2. **TODAY in tests:** generated code calls `ForgeToday()`. It returns the real date normally, and
   a fixed batch date inside `RunBatch`, so Python and Excel agree. OK? (Needed for batch_002.)
3. **`[Header]` ranges:** keep "the whole column, header row included", which the specs'
   `= RESULT + 1` checks need? I'd fix spec.md's wording in VCL revision 4, together with the
   Orchestrator prompt.
4. **Dataset count:** does the pilot count as verified pair #1? I'd say no. Verified pairs should
   come from the pipeline, so the count starts with batch_001.
5. **Verified pair format:** `data/verified/<spec>.json`, holding the spec text, the generated VBA,
   the batch id and the Excel results for each TEST?

## Not in Phase 1

BUILD specs and UserForms, the local model, and automating Excel are all later work.

VBA emulators such as LibreOffice Basic are out too. They can't run Scripting.Dictionary or
`Application.Run` the way Excel does, so they would test the emulator rather than our code.
