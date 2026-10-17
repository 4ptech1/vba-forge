# VCL 0.2 proposal: from 10 tasks to diverse VBA

Status: **proposal, not adopted** (14 Oct 2026). VCL 0.1 revision 3 stays frozen for the pilot
and the Phase 1 compiler core. Adoption is staged (section 11); nothing here changes before the
pilot passes in Excel.

## 1. Is VCL 0.1 ready?

| Ready for | Not ready for |
| --- | --- |
| Register-style tasks: read rows, match IDs, compare quantities, copy, flag, summarise | Text work: splitting, searching, padding, cleaning copy-pasted text |
| Header-based columns, dirty-data guards, safe re-runs | Dates typed as text (dd/mm/yyyy), date arithmetic, working days |
| The 10 KGC specs, the validator, the pilot | Excel Tables (ListObjects), conditional formatting, data validation |
| | Creating or checking sheets (specs currently assume output sheets exist) |
| | Recovering from one bad row without aborting the whole macro |
| | Bounded `while` loops, skip-to-next-row, user-defined worksheet functions |
| | Anything outside the op list: `NOTE` is free text with no declared inputs or outputs |
| | Short tests: every spec hand-writes its own "nothing else changed" checks |

## 2. Design principles

1. **Closed core, not a closed world.** VCL states intent, guards and tests. It does not mirror
   all of VBA; that would make it a second programming language for the model to learn.
2. **One structured escape hatch.** Anything outside the vocabulary is a `DO` step with declared
   inputs and outputs. This is the local model's free zone, and it is still tested.
3. **Packs.** New ops arrive in themed packs. A pack is available only when it has spec text,
   validator rules, compiler templates, ForgeLib functions, harness support and one Excel batch
   that passed.
4. **Semantics live in one place.** A shared VBA runtime module, ForgeLib, implements every VCL
   meaning (empty, same, parse date...). Generated code calls it instead of re-inventing edge
   cases.
5. **One house style.** Generated VBA always looks the same, so the model learns one way to
   write each thing.
6. **Safety is declared, not hoped for.** `MODIFIES` says which sheets may change, and the
   harness checks every other sheet automatically. `ROLE` separates UI from logic.
7. **Dirty data by default.** Fixture modifiers plant spaces, wrong case, numbers stored as text,
   text dates and duplicates, so code without guards fails its tests and never enters the
   dataset.
8. **Excel 2016 desktop is the target.** Every op maps to Excel 2016 VBA. Excel 2019 and
   Microsoft 365 functions (TEXTJOIN, CONCAT, IFS, MAXIFS, MINIFS, SWITCH, XLOOKUP, FILTER, LET,
   LAMBDA, dynamic arrays) are never used.

## 3. Core language changes

### 3.1 Version line

The first line becomes `VCL 0.2`. The validator accepts 0.1 specs until the 10 specs are
migrated, then rejects them.

### 3.2 New keys

Key order becomes:
`PROC, KIND, ROLE, MODULE, ARGS, RET, CONST, VAR, PRE, MODIFIES, OP, PERF, ERR, ALLOW, FIXTURE, TEST`

| Key | Required | Meaning | Example |
| --- | --- | --- | --- |
| `ROLE` | No (default `logic`) | `logic`: testable, no UI ops. `ui`: may ask the user, no data logic, calls logic procedures. `udf`: a Function used in worksheet cells, no side effects | `ROLE     udf` |
| `CONST` | No | Named constant, one per line, usable in expressions and options | `CONST    dueWindow = 15` |
| `MODIFIES` | Yes, if any value is written to a sheet argument | Sheet arguments whose values may change. Formatting does not count | `MODIFIES $out` |

ALLOW changes: add `sheet_delete`; remove `shell` and `http` (no op uses them, and the hard
rules forbid them).

### 3.3 The DO step (replaces NOTE)

```
DO("<one technical sentence, no VBA>", reads=[<names>], sets=[<names>])
```

- `reads`: arguments, variables and constants the step may read.
- `sets`: variables the step may change, plus `$sheet` if it writes values to that sheet (the
  sheet must be in `MODIFIES`).
- At most 3 DO steps per TASK: a warning from 4, an error from 7. More than that means the spec
  is English, not VCL.
- Every DO step must be covered by a TEST that would fail if the step were skipped (author rule,
  in FINAL CHECK).
- `NOTE` still parses in 0.2 with a warning: "NOTE is deprecated; use DO".
- Example: `DO("Split heatNo at the first '-' into prefix and serial; both empty if there is no '-'", reads=[heatNo], sets=[prefix, serial])`

### 3.4 Loop control

| Syntax | Meaning |
| --- | --- |
| `SKIP_IF <condition>` | Go to the next iteration of the innermost loop |
| `EXIT_LOOP_IF <condition>` | Leave the innermost loop |
| `LOOP_WHILE cond=<expr> max=<n>` | Bounded while loop. Reaching `max` writes one ForgeLog entry and leaves the loop. `max` is required |
| `LOOP_ROWS ... on_error=skip` | If one iteration raises an error, log "row r: description" to ForgeLog and continue with the next row. Only for loops whose iterations are independent |

`SKIP_IF` replaces most nested `IF NOT IS_EMPTY(...)` blocks, which keeps specs flat and readable.

### 3.5 Expressions

| Syntax | Meaning |
| --- | --- |
| `SAME(a, b)` | Equal after CLEAN_TEXT, ignoring case. Use instead of `UPPER(TRIM(x)) = "Y"` |
| `IIF(cond, a, b)` | `a` if cond is true, else `b`. Both sides are evaluated, so only use safe values |
| `MOD` | Remainder operator: `r MOD 2 = 0` |

Plain `=` keeps VBA meaning (exact, case-sensitive) and should be used for numbers and dates.

Numeric text that passes IS_NUMBER ("25") is converted to a number wherever a number is needed:
arithmetic, comparisons, and `DICT_ADD` with `mode=sum`.

`[Header]` ranges in assertions include the header row (decision from the pilot), so counts over
them are one more than the number of data rows.

### 3.6 WRITE_ROW

```
WRITE_ROW(<ws>, <row>, <v1>, <v2>, ...)
```

Writes values to consecutive cells from column 1. `""` leaves a cell empty. One line replaces a
block of SET_VALUE lines for headers and output rows.

## 4. Packs

The "VBA basis" column is a note for the compiler and ForgeLib, not part of the language.

### 4.1 Text

| Syntax | Meaning | VBA basis |
| --- | --- | --- |
| `CLEAN_TEXT(v)` | Text with non-breaking spaces turned into spaces, line breaks and tabs removed, trimmed, inner spaces collapsed | Replace, Trim |
| `CONTAINS(s, part)`, `STARTS_WITH(s, part)`, `ENDS_WITH(s, part)` | Case-insensitive tests | InStr with vbTextCompare |
| `MATCHES(s, "<pattern>")` | Like-pattern test, case-insensitive: `*` `?` `#` `[A-Z]` | `Like` on UCase |
| `MID(s, start, n)`, `INSTR(s, find)`, `REPLACE(s, find, with)` | As in VBA; INSTR returns 0 if missing; matching ignores case | Mid, InStr, Replace |
| `PART(s, sep, n)` | The n-th piece (1-based), or `""` | Split |
| `PARTS(s, sep)` | All pieces, for `FOR_EACH var=p in=PARTS(s, "-")` | Split |
| `PAD_LEFT(s, n, "0")` | Left-pad to n characters; unchanged if already longer | String, Right |
| `PROPER(s)` | Title case | StrConv vbProperCase |
| `TEXT(v, "<format>")` | Formatted text | Format |
| `NUM(v, default)` | The number, or `default` if v is not a plain number | IsNumeric, CDbl |

Regular expressions are not included: VBScript.RegExp is being deprecated on Windows, so
nothing is built on it.

### 4.2 Dates

| Syntax | Meaning | VBA basis |
| --- | --- | --- |
| `PARSE_DATE(v, "dd/mm/yyyy")` | A real date stays as it is. Text is parsed strictly in the given format (1 or 2 digits for day and month; `/`, `-` or `.` as separator). Impossible dates like 31/02 give empty. Never guesses the order | Split, DateSerial, round-trip check (never CDate) |
| `DATE(y, m, d)` | A date | DateSerial |
| `YEAR(d)`, `MONTH(d)`, `DAY(d)`, `WEEKDAY(d)` | Parts; WEEKDAY has Monday = 1 | Year, Month, Day, Weekday(vbMonday) |
| `ADD_DAYS(d, n)`, `ADD_MONTHS(d, n)` | Date arithmetic | `+`, DateAdd("m") |
| `DAYS_BETWEEN(a, b)` | Whole days from a to b (negative if b is earlier) | Int(b) - Int(a) |
| `MONTH_END(d)` | Last day of d's month | DateSerial(y, m + 1, 0) |
| `WORKDAYS_BETWEEN(a, b)` | Monday to Friday days, both ends included | WorksheetFunction.NetworkDays |

### 4.3 Numbers

| Syntax | Meaning | VBA basis |
| --- | --- | --- |
| `ROUND(x, n)` | Rounds like Excel (2.5 becomes 3). VBA's own Round uses banker's rounding (2.5 becomes 2), which surprises users | WorksheetFunction.Round |
| `ABS(x)`, `INT(x)`, `MIN(a, b)`, `MAX(a, b)` | As usual | Abs, Int, comparisons |

### 4.4 Ranges and formatting

| Syntax | Meaning | VBA basis |
| --- | --- | --- |
| `BLOCK(ws, r1, c1, r2, c2)` | A rectangular range | Range(Cells, Cells) |
| `CURRENT_REGION(cell)` | The block around a cell | CurrentRegion |
| `NAMED(book, "<name>")` | A named range | Names(...).RefersToRange |
| `COPY_VALUES(from, toCell)` | Values only, resized to fit | `.Value = .Value` |
| `COPY_RANGE(from, toCell)` | Values and formats, without the clipboard | Copy Destination |
| `TO_VALUES(range)` | Replace formulas with their values | `.Value = .Value` |
| `REPLACE_IN(range, find, with, match=whole\|part)` | Find and replace in cells | Range.Replace |
| `AUTOFIT(range)` | Fit column widths | AutoFit |
| `HIDE_ROWS(range, yes\|no)`, `HIDE_COLS(range, yes\|no)` | Hide or show | Hidden |
| `COND_FORMAT(range, when="<formula>", fill=<colour>)` | Formula-based conditional format; the formula is written for the top-left cell | FormatConditions.Add xlExpression |
| `VALIDATION_LIST(range, "A,B,C")` | Drop-down list (replaces any existing validation) | Validation.Add xlValidateList |
| `CALCULATE()` | Recalculate | Application.Calculate |

FORMAT gains these options: `font_color=`, `italic=yes|no`, `border=thin|none`,
`align=left|center|right` and `wrap=yes|no`.

### 4.5 Sheets and workbooks

| Syntax | Meaning | VBA basis |
| --- | --- | --- |
| `SHEET_EXISTS(book, "<name>")` | True or false; never raises an error | loop over sheets |
| `ENSURE_SHEET(book, "<name>", into=<var>)` | The existing sheet, or a new one added at the end. Removes the "output sheet already exists" assumption | Sheets.Add |
| `RENAME_SHEET(sheet, <expr>)` | An invalid name (more than 31 characters, `[]:*?/\`) writes one ForgeLog entry and changes nothing | Name |
| `HIDE_SHEET(sheet, yes\|no)` | Hide or show | Visible |
| `DELETE_SHEET(sheet)` | Needs `ALLOW sheet_delete` | Delete with DisplayAlerts off |
| `SHEET_COUNT(book)` | Number of worksheets | Worksheets.Count |

### 4.6 Excel Tables (ListObjects)

| Syntax | Meaning | VBA basis |
| --- | --- | --- |
| `TABLE(ws, "<name>")` | The table, or NOTHING | ListObjects |
| `TABLE_ROWS(t)` | Number of data rows | ListRows.Count |
| `TABLE_COL(t, "<header>")` | The data cells of a column; header matched like COL | ListColumns(...).DataBodyRange |
| `TABLE_CELL(t, row, "<header>")` | One data cell, readable and writable | DataBodyRange.Cells |
| `ADD_TABLE_ROW(t, into=<var>)` | Append a row; var gets its row number | ListRows.Add |

### 4.7 Lookup and arrays

| Syntax | Meaning | VBA basis |
| --- | --- | --- |
| `MATCH_ROW(ws, col, value)` | First data row whose cell is SAME as value, or 0 | loop (no Match error handling needed) |
| `NEW_ARRAY(rows, cols, into=<var>)` | Empty 1-based 2D array | ReDim |
| `ADD_ITEM(collection, value)` | Append | Collection.Add |

### 4.8 Files (testable from March 2027)

`FILE_EXISTS(path)`, `FOLDER_EXISTS(path)`, `JOIN_PATH(a, b)`, `FILE_EXT(path)`,
`BOOK_FOLDER(book)`, `MAKE_FOLDER(path)` (does nothing if the folder exists), and
`COPY_FILE(from, to)`, which fails if `to` exists unless `ALLOW file_overwrite` is given.
Batches never include these before March.

### 4.9 UI (ROLE ui only)

`ASK(<prompt>, into=<var>)` (InputBox), `PICK_FILE("<filter>", into=<var>)`,
`PICK_FOLDER(into=<var>)` (FileDialog), and the existing `MSG`. `STATUS(<expr>)` (status bar
text) is allowed in any role; the harness ignores it.

The pattern is a thin `ROLE ui` Sub that collects inputs and CALLs a `ROLE logic` procedure that
takes them as ARGS. Only the logic procedure is tested.

### 4.10 UDF rules (ROLE udf)

- `KIND Function` with `RET`. ARGS are plain values or ranges.
- No write, format, sheet, UI or PERF ops, and no changes to Excel state.
- TEST calls use literal arguments: `call=PadSeam("7", 3) assert=RESULT = "007"`.

## 5. Tests and fixtures

### 5.1 New assertions

| Syntax | Meaning |
| --- | --- |
| `UNCHANGED(<Sheet>)` | Values and formulas in the sheet's used range are identical to before the first run. Formatting is ignored |
| `SORTED(<Sheet>![Header], asc\|desc)` | Data cells from row 2 down are in order; blanks must be at the bottom |
| `DISTINCT(<range>)` | Number of distinct non-blank values, compared with SAME |
| `COUNTBLANK(<range>)` | As in Excel |
| `HAS_HEADER(<Sheet>, "<header>")` | True if the row-1 header exists |

### 5.2 Automatic collateral checks

For every TEST, the harness adds `UNCHANGED(<Sheet>)` for each sheet passed in the call that is
not listed in `MODIFIES`. Spec writers no longer hand-write PRE_SUM checks for input sheets.
Failures are reported as "(auto) UNCHANGED(GRN)".

### 5.3 Fixture dirt modifiers

Add these after any rule. They are applied in the order written, each with its own per-column
seed:

| Modifier | Effect on n random cells of the column |
| --- | --- |
| `+blank(n)` | Empty (as in 0.1) |
| `+spaces(n)` | Leading or trailing spaces, sometimes a non-breaking space (common in copy-paste) |
| `+case(n)` | Random upper, lower or mixed case |
| `+astext(n)` | Numbers stored as text; dates stored as dd/mm/yyyy text |
| `+dupes(n)` | Replaced by another value from the same column |

New rule `datetext(<start>,<end>,"<format>")` writes dates as text in that format. New FIXTURE
option `table=<Name>` turns the sheet's data into an Excel Table with that name.

The default text-date format is `dd/mm/yyyy` (India). Text dates are only ever read with
PARSE_DATE and an explicit format.

## 6. VBA house style (additions to the output contract)

1. One procedure per spec, between the `'@VCL` and `'@END` markers.
2. Before each top-level OP line, a comment `' [OP n]` (n is the line number in the spec). DO
   steps also get `' [DO] <instruction text>`. This gives line-by-line traceability between spec
   and code, and helps the model keep its place.
3. Every Range and Cells call is qualified with a worksheet variable. No ActiveSheet, Selection or
   ActiveCell, except in ROLE ui when reading the user's choice.
4. Types: Long for rows, columns and counters; Double for quantities; Variant for raw cell values
   and dates that may be empty; String for cleaned IDs; Boolean for flags. Never Integer.
5. Every VCL meaning goes through a ForgeLib helper (ForgeIsEmpty, ForgeSame, ForgeParseDate and
   so on), never inline code.
6. Dictionaries come from `ForgeNewDict()`: late-bound, text compare.
7. Read a cell once into a Variant when it is used more than once in an iteration.
8. More than about 1,000 rows: arrays (rule 11 of 0.1).
9. Fixed error-handler shape. `on_error=skip` uses a fixed per-row label pattern. No
   `On Error Resume Next` in generated code; helpers that need it are inside ForgeLib and
   tested.
10. CONST lines become `Const` declarations at the top of the procedure.
11. CleanExit restores ScreenUpdating, Calculation, EnableEvents and DisplayAlerts.
12. Lines at most 100 characters. No Select, Activate, Integer, SendKeys or clipboard use. GoTo
    only to the error and skip labels.

## 7. ForgeLib, the runtime library

ForgeLib.bas is a standalone module shipped with every finished tool and pasted with every batch.
The pilot's helpers move out of ForgeTest.bas into it, and ForgeTest.bas keeps only the test
runner. The assembler still produces one paste-ready module.

Helpers (one per VCL meaning): ForgeNewDict, ForgeIsEmpty, ForgeIsNumber, ForgeIsDate, ForgeNum,
ForgeCleanText, ForgeSame, ForgeParseDate, ForgeRound, ForgeCol, ForgeLastRow, ForgeLastCol,
ForgeRowRange, ForgeCopyRow, ForgeWriteRow, ForgeSheetExists, ForgeEnsureSheet, ForgeMatchRow,
ForgeTable, ForgeTableCol, ForgeLog, ForgeErrorCount.

ForgeLib gets its own self-test batch (about 60 tests) that runs once in Excel before any pack
that depends on it is marked available.

## 8. New validator rules

- `VCL 0.2` header; the new key order.
- ROLE rules: logic has no UI ops and needs TESTs; ui has no loops and must CALL a logic
  procedure; udf has no write, format, sheet, UI or PERF ops, and its TESTs use literal
  arguments.
- MODIFIES: any value-writing op whose first argument is a sheet argument must name a sheet in
  MODIFIES; formatting ops are exempt. DO `sets=[$x]` requires `$x` in MODIFIES.
- DO: names in `reads` and `sets` must be declared; `sets` cannot contain plain arguments; count
  warnings at 4 and errors at 7.
- `LOOP_WHILE` without `max=` is an error. `SKIP_IF` and `EXIT_LOOP_IF` must be inside a loop.
- `on_error=skip` only on LOOP_ROWS, LOOP_COLS and FOR_EACH.
- File-pack ops: warning "testable from March 2027". `DELETE_SHEET` without `ALLOW sheet_delete`
  is an error.
- NOTE: warning "deprecated; use DO".
- CONST names must be unique and must not clash with VAR names.
- `PARSE_DATE` must have an explicit format string.

## 9. Deferred to VCL 0.3 (after March 2027)

| Area | Why not now |
| --- | --- |
| UserForms | Creating forms in code needs a Trust Center change; testing needs clicks |
| Worksheet and workbook events | The code lives in sheet modules; tests would need event simulation |
| Class modules | BUILD-level design; better once BUILD is proven |
| Pivot tables, charts | Large object model, and assertions can't check them yet |
| Outlook, Word, PowerPoint automation | Other applications, security prompts, untestable on the work PC |
| Regular expressions | VBScript.RegExp is being deprecated on Windows |
| ADO/SQL, API Declare | Risky, and not needed for the register tasks |
| SendKeys, clipboard | Never: fragile and unsafe |

## 10. Examples

### 10.1 task_01 in VCL 0.2

The same behaviour as 0.1, with 22 OP lines instead of 33 and 3 tests instead of 4. The
collateral test is automatic through MODIFIES, and the fixture now plants dirty part numbers.

```
VCL 0.2
TASK
# ASSUME: BOM headers (row 1) include "Part No" and "Required Qty"; GRN headers include "Part No" and "Received Qty"
# ASSUME: a part can appear on several GRN rows; its received quantities are added
# ASSUME: a blank or non-numeric Received Qty is ignored; a part with no valid GRN rows counts as 0 received
# ASSUME: a blank or non-numeric Required Qty is listed as "Required qty missing"
PROC     BuildShortageList
KIND     Function
ROLE     logic
MODULE   modMaterials
ARGS     bom As Worksheet, grn As Worksheet, out As Worksheet
RET      Long
VAR      received As Dictionary
VAR      bomPart As Long
VAR      bomReq As Long
VAR      grnPart As Long
VAR      grnQty As Long
VAR      partNo As String
VAR      got As Double
VAR      outRow As Long
PRE      COL($bom, "Part No") > 0; COL($bom, "Required Qty") > 0; COL($grn, "Part No") > 0; COL($grn, "Received Qty") > 0
MODIFIES $out
OP       LET(bomPart, COL($bom, "Part No"))
OP       LET(bomReq, COL($bom, "Required Qty"))
OP       LET(grnPart, COL($grn, "Part No"))
OP       LET(grnQty, COL($grn, "Received Qty"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Part No", "Required", "Received", "Shortage", "Remark")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($grn, grnPart) var=r
OP         SKIP_IF IS_EMPTY(CELL($grn, r, grnPart)) OR NOT IS_NUMBER(CELL($grn, r, grnQty))
OP         DICT_ADD(received, CLEAN_TEXT(CELL($grn, r, grnPart)), CELL($grn, r, grnQty), mode=sum)
OP       LET(outRow, 2)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($bom, bomPart) var=r
OP         SKIP_IF IS_EMPTY(CELL($bom, r, bomPart))
OP         LET(partNo, CLEAN_TEXT(CELL($bom, r, bomPart)))
OP         LET(got, 0)
OP         IF EXISTS(received, partNo) THEN DICT_GET(received, partNo, into=got)
OP         IF NOT IS_NUMBER(CELL($bom, r, bomReq))
OP           WRITE_ROW($out, outRow, partNo, "", got, "", "Required qty missing")
OP           LET(outRow, outRow + 1)
OP         SKIP_IF NOT IS_NUMBER(CELL($bom, r, bomReq))
OP         SKIP_IF got >= CELL($bom, r, bomReq)
OP         WRITE_ROW($out, outRow, partNo, CELL($bom, r, bomReq), got, CELL($bom, r, bomReq) - got, IIF(got = 0, "Not received", "Short"))
OP         LET(outRow, outRow + 1)
OP       RETURN(outRow - 2)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  bom_grn_01.xlsx sheet=BOM cols=A:Part No,B:Description,C:Required Qty rows=40 data=A:seq(1001,1)+astext(3);B:text(12);C:int(1,50)+blank(3)
FIXTURE  bom_grn_01.xlsx sheet=GRN cols=A:Part No,B:Received Qty rows=35 data=A:seq(1001,1)+spaces(4);B:int(1,50)+blank(4)+astext(2)
FIXTURE  bom_grn_01.xlsx sheet=Shortages cols=A:Part No,B:Required,C:Received,D:Shortage,E:Remark rows=0
TEST     fixture=bom_grn_01.xlsx call=BuildShortageList(BOM, GRN, Shortages) runs=2 assert=RESULT >= 5 AND COUNTA(Shortages![Part No]) = RESULT + 1 AND COUNTIF(Shortages![Shortage], "<0") = 0 AND ERROR_COUNT = 0
TEST     fixture=bom_grn_01.xlsx call=BuildShortageList(BOM, GRN, Shortages) assert=COUNTIF(Shortages![Remark], "Required qty missing") = 3 AND DISTINCT(Shortages![Part No]) = RESULT + 1
TEST     fixture=bom_grn_01.xlsx call=BuildShortageList(Shortages, GRN, Shortages) assert=RESULT = -1 AND ERROR_COUNT = 1
```

### 10.2 New capability: text dates and per-row recovery

```
VCL 0.2
TASK
# ASSUME: Weld Log headers (row 1) include "Weld Seam No", "Fit-up Date" and "Visual Date"
# ASSUME: dates may be real dates or dd/mm/yyyy text; anything else counts as missing
# ASSUME: a "Days to Visual" column is added after the last column if it does not exist
PROC     AddDaysToVisual
KIND     Function
ROLE     logic
MODULE   modWelding
ARGS     weldLog As Worksheet
RET      Long
CONST    dateFmt = "dd/mm/yyyy"
VAR      cSeam As Long
VAR      cFit As Long
VAR      cVis As Long
VAR      cOut As Long
VAR      fit As Variant
VAR      vis As Variant
VAR      done As Long
PRE      COL($weldLog, "Weld Seam No") > 0; COL($weldLog, "Fit-up Date") > 0; COL($weldLog, "Visual Date") > 0
MODIFIES $weldLog
OP       LET(cSeam, COL($weldLog, "Weld Seam No"))
OP       LET(cFit, COL($weldLog, "Fit-up Date"))
OP       LET(cVis, COL($weldLog, "Visual Date"))
OP       LET(cOut, COL($weldLog, "Days to Visual"))
OP       IF cOut = 0
OP         LET(cOut, LASTCOL($weldLog, 1) + 1)
OP         SET_VALUE(CELL($weldLog, 1, cOut), "Days to Visual")
OP       LET(done, 0)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($weldLog, cSeam) var=r on_error=skip
OP         SKIP_IF IS_EMPTY(CELL($weldLog, r, cSeam))
OP         CLEAR(CELL($weldLog, r, cOut), what=contents)
OP         LET(fit, PARSE_DATE(CELL($weldLog, r, cFit), dateFmt))
OP         LET(vis, PARSE_DATE(CELL($weldLog, r, cVis), dateFmt))
OP         SKIP_IF IS_EMPTY(fit) OR IS_EMPTY(vis)
OP         SET_VALUE(CELL($weldLog, r, cOut), DAYS_BETWEEN(fit, vis))
OP         LET(done, done + 1)
OP       RETURN(done)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  weld_dates_01.xlsx sheet=WeldLog cols=A:Weld Seam No,B:Welder ID,C:Fit-up Date,D:Visual Date rows=40 data=A:seq(1,1);B:values("W-01"x20,"W-02"x20)+spaces(4);C:datetext(2026-01-01,2026-06-30,"dd/mm/yyyy")+blank(3);D:dates(2026-07-01,2026-09-30)+blank(2)+astext(5)
FIXTURE  weld_dates_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(10)
TEST     fixture=weld_dates_01.xlsx call=AddDaysToVisual(WeldLog) runs=2 assert=RESULT >= 35 AND RESULT <= 37 AND COUNTA(WeldLog![Days to Visual]) = RESULT + 1 AND COUNTIF(WeldLog![Days to Visual], "<0") = 0 AND ERROR_COUNT = 0
TEST     fixture=weld_dates_01.xlsx call=AddDaysToVisual(WeldLog) assert=COUNTA(WeldLog![Fit-up Date]) = PRE_COUNTA(WeldLog![Fit-up Date]) AND COUNTA(WeldLog![Visual Date]) = PRE_COUNTA(WeldLog![Visual Date])
TEST     fixture=weld_dates_01.xlsx call=AddDaysToVisual(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
```

## 11. Adoption plan

A pack is "available" only when it has all of: spec.md text, validator rules, compiler
templates, ForgeLib functions, harness support, unit tests, and one Excel batch that passed.

| Stage | When | Contents | Why then |
| --- | --- | --- | --- |
| 0 | Now | Nothing. Run the pilot | The pilot proves the engine first |
| 1 | Phase 1 start, after the pilot passes | Core: `VCL 0.2` header, ROLE, CONST, MODIFIES, SKIP_IF, EXIT_LOOP_IF, LOOP_WHILE, on_error=skip, SAME, CLEAN_TEXT, IIF, MOD, WRITE_ROW, UNCHANGED with automatic collateral checks, dirt modifiers, ForgeLib split, house style. Migrate the 10 specs | The compiler's templates should be built once, on the final core |
| 2 | Phase 2 start (December) | Text, Dates, Numbers, Ranges and formatting, Sheets, Lookup and arrays, the DO step | These drive dataset diversity |
| 3 | Late Phase 2 | Excel Tables (needs `table=` fixtures) | Builds on stage 2 |
| 4 | March 2027 | Files and UI packs | They need the automated harness and sandbox folders |
| 0.3 | After March | Section 9 | Needs the new PC and a testable environment |

The Phase 2 plan's category table should be re-checked against stage 2: text cleanup, date
handling, formatting and sheet management become reachable categories.

## 12. Why this makes the local model a better VBA specialist

- **Each step is a small, precise decision.** SAME, SKIP_IF and WRITE_ROW remove boilerplate, so
  the model's attention goes to logic rather than ceremony.
- **One meaning, one implementation.** ForgeLib carries the edge cases, so the model learns "op
  to helper call" instead of re-deriving date parsing or text cleaning each time, and can't get
  them subtly wrong.
- **DO steps concentrate real VBA reasoning** where templates can't reach, with declared inputs
  and outputs so the step stays small and testable. They are the most valuable training pairs.
- **Dirty fixtures teach defence.** Code without guards fails its tests and never becomes
  training data, so the model only ever sees defensive VBA.
- **Traceable output.** `' [OP n]` comments align spec and code line by line, which makes
  training targets consistent and repairs easier to localise.
- **Coverage arithmetic.** About 130 ops and functions at roughly 15 verified examples each is
  about 2,000 pairs, which matches the Phase 2 dataset size. Staging packs keeps every op above
  that floor instead of spreading data thin.
