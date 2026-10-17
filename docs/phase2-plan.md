# Phase 2 plan: the verified dataset

**Status: proposal for Boss's approval. Phase 2 starts after Phase 1 is done (Dec 2026 to Feb 2027).**

Target, from the roadmap: 2,000 VCL-to-VBA TASK pairs, 500 repair pairs and a 300-task
held-out eval set (VBA-Bench). Every item is proven by running in Excel 2016. The data is used
for a QLoRA fine-tune of a 7B or 14B model in April.

## 0. Five assumptions I'd change first

1. **"A zero-hallucination model" is not something a model can be.** What we can build is zero
   *unverified output*. Nothing reaches you unless it passes three checks: lint, the simulator
   and the test harness. When the code fails, the repair loop fixes it, up to 3 tries. On top of
   that, the model is trained to say "can't compile" instead of inventing code. That makes it
   a property of the system, which is reachable.
2. **"Deep reasoning" is the wrong target for this model.** The reasoning happens when Claude
   writes the spec. The local model translates the spec, and its typical failure is skipping a
   detail (a TRIM, `dir=up`, `mode=sum`), not shallow reasoning. So the dataset should train
   *careful reading*. Section 2 explains how.
3. **Most of the 2,000 pairs are about things the Phase 1 compiler already does perfectly.**
   The 10 specs use 0 NOTE lines and only 13 ops, all of which become templates in Phase 1. So
   the model mainly earns its place on what templates can't do:
   - NOTE steps
   - new ops that don't have a template yet
   - repairs
   - knowing when to refuse

   The data mix follows from that: about half compiler-made pairs to teach the house style,
   and half Claude-written pairs for the hard parts.
4. **6 of the 14 categories don't fit v1 as written.** Some VCL 0.1 can't express, and some
   your locked-down PC can't verify. Training on code we can't verify would teach exactly the
   hallucinations we want to avoid. Section 1 re-allocates them.
5. **"Batches of 50" doesn't fit the calendar.** Counting repairs, eval and retries, it's roughly
   3,000 to 3,300 item runs. At 50 per paste that's 60 to 66 pastes, or 20 to 22 weeks at 3 a
   week. At 100 per paste it's 30 to 33 pastes, or 10 to 11 weeks, which fits Dec to Feb. The paste is the fixed cost, so each one
   should carry as much as is safe. Section 6 covers this.

## 1. What goes into v1

| Roadmap category | Roadmap target | Proposed v1 | Why |
|---|---|---|---|
| Dictionaries and collections | 150 | **320** | Reconciliation (BOM vs GRN, NDE vs weld log) is the paper trail's main job, and the hardest for a small model |
| Loops and logic | 150 | **240** | Nested conditions and row logic are where spec-reading slips happen |
| Ranges and cells | 200 | **220** | Core of every task |
| Tables, sorting, filtering | 150 | **220** | 120 for SORT, FILTER and DEDUPE now; 100 for ListObjects once VCL rev 4 adds ops for them |
| *new:* Dates and status | - | **180** | Due dates, ageing and expiry (3 of the 10 real specs) |
| Arrays and performance | 200 | **160** | READ_ARRAY/WRITE_ARRAY and PERF; a small model picks these up quickly |
| Sheets and workbooks | 200 | **160** | Rebuilding output sheets, multi-sheet work, ADD_SHEET/COPY_SHEET |
| Formatting | 100 | **160** | Highlighting rows is a daily need |
| Error handling | 100 | **140** stress items, plus all 2,000 | Not a separate topic: every item has PRE, guards and the standard handler. The 140 are dirty-data stress tests |
| Formulas | 100 | **120** | SET_FORMULA with Excel 2016 functions only |
| Files and folders | 150 | **80** | Sandbox folder only and no deleting; needs harness support |
| Pivot tables and charts | 150 | 0, in v2 | No VCL ops; rare in these registers |
| UserForms and events | 150 | 0, after March | A form can't be pasted as text, building one needs a Trust Center change, and events live in sheet modules |
| Class modules | 100 | 0, in VCL 0.2 | VCL 0.1 has no classes, and a class needs a second module paste |
| External automation | 100 | 0, out of v1 | Outlook, Word, ADO and HTTP have side effects and can't be tested safely on the work PC |
| **Total** | 2,000 | **2,000** | |

## 2. Distribution, difficulty and avoiding overfitting

**Difficulty tiers are measured by the validator, not judged:**

| Tier | Shape | Share |
|---|---|---|
| L1 | 1 sheet, 1 loop, up to 12 OP lines | 15% |
| L2 | 2 sheets or 1 dictionary, nesting up to 2, 13 to 24 OPs | 35% |
| L3 | 3+ sheets, several dictionaries or arrays, nesting 3, 25 to 34 OPs | 35% |
| L4 | 35 to 40 OPs, with dictionaries and arrays together, many guards | 15% |

Above 40 OPs a spec becomes a BUILD, which is v2; task_04 is the one accepted exception. Production rises through the tiers over
Phase 2, while training data is shuffled; a curriculum isn't worth it at this size.

**Teaching the model to read, not guess:**
- **Contrast clusters (20% of items):** 2 or 3 specs that differ in exactly one detail that changes
  the meaning, such as `dir=up` vs `dir=down`, `keep` vs `sum`, `<` vs `<=`, or a TRIM present
  or missing. The code must differ in exactly the right place. This is the strongest guard against
  "recognise the task, ignore the details".
- **Surface variety:** headers, sheet names, variable names and register types all vary. About
  70% are fabrication registers (BOM, weld log, NDE, NCR, MTC, ITP, calibration, hydrotest, PWHT,
  painting, transmittals) and 30% generic office data, so the model doesn't lock onto our vocabulary.
- **About 16 task archetypes:** reconcile, flag, exceptions list, summarise, merge, ageing or due
  date, cleanup, dedupe, lookup-fill, split by key, consolidate, validate entries, numbering,
  report layout, import/export, audit stamp. Two are held out for the eval set only (section 5).
- **No padding:**
  - At most 3 items share one *signature*: the op skeleton with the names removed.
  - No rename-augmentation; variety comes from generating different specs.
  - In March, train at 500, 1,000 and 2,000 items and check that more data still helps before
    chasing the last few hundred.

## 3. Making Excel 2016 and defensive code the default

These are built into the data, not left to examples:

1. **One house style in 100% of target code:** the pilot's skeleton, using the Forge helpers,
   PRE, PERF restore and the standard handler. Anything mechanical (line numbers, the order of
   Dim lines, indentation) is added by tools, never by the model, so it can't get those wrong.
2. **Dirty fixtures by default (VCL rev 4).** Every fixture column gets a small, seeded dose of
   hand-typed mess unless the spec opts out:
   - stray spaces and mixed case
   - numbers stored as text and dates stored as text
   - blanks and duplicate keys

   Code that isn't defensive fails its TESTs and never enters the dataset, so defensive code is
   the only kind the model ever sees. Hard rule 15 already asks for this; dirty fixtures make it
   impossible to skip.
3. **Standard negative TESTs**, required by the validator: wrong sheet (already required),
   missing header, empty input (`rows=0`) and re-run (`runs=2`).
4. **An Excel 2016 ban list in lint**, for both formulas and WorksheetFunction calls:
   - Microsoft 365 functions: XLOOKUP, XMATCH, FILTER(, SORT(, UNIQUE, SEQUENCE, LET, LAMBDA
     and `Range.Formula2`
   - the 2019 functions that Excel 2016 lacks: **TEXTJOIN, CONCAT, IFS, MAXIFS, MINIFS, SWITCH**
5. **Refusal examples (about 2%):** specs that ask for the impossible, such as a Microsoft 365
   function or external automation. The target is one fixed line,
   `' FORGE: CANNOT COMPILE: <reason>`, so the model learns to say no rather than invent.
6. **The same short system prompt on every training example:** Excel 2016, Forge helpers
   available, no Select. The model is always conditioned on it.

## 4. Repair pairs (500)

**Where they come from.** Every failure log is real, never imagined.
- **Natural failures (at least 40%):** Claude-written candidates (later, model-written ones) that
  failed in Excel. Claude fixes them and the fix is verified in the next batch.
- **Mutants (at most 60%):** verified code with one realistic bug added by a mutation tool, then
  run in Excel to get the real failure log. The fix is the original verified code, so no extra
  run is needed. The kinds of bug are copied from the natural failures, not invented.

| Kind | Share | Examples | Where the log comes from |
|---|---|---|---|
| Compile | 20% | undeclared variable, missing End If, a Microsoft 365 function | Our lint. A compile error would stop a whole Excel batch, so these never go to Excel. At inference the same lint produces the same messages |
| Runtime | 30% | 1004 bad range, 13 type mismatch on a text number, 91 object not set, 9 missing sheet | Excel Results: error number, line, ForgeLog |
| Wrong result | 50% | missing TRIM, deleting rows top-down, appending instead of rebuilding, wrong column, `<` vs `<=` | Excel Results: the failed check with the measured and expected values |

**Format: the same shape in training and in real use.**
```
INPUT                                    OUTPUT
REPAIR attempt 1 of 3                    CAUSE: MISSING_GUARD (line 33: text number treated as missing)
SPEC:   <the VCL spec>                   <the full corrected procedure>
CODE:   1| Public Function ...
        2|     Dim received As Object
        ...
FAILED: B012-T03 FAIL: COUNTIF(Shortages![Remark], "Required qty missing") = 3 [2 = 3]
LOG:    none
```
- **Line numbers on the code** let the model tie the failing line (Erl) to the code.
- **Only the failing tests and the log go in**, which keeps the context short enough for a 7B model.
- **CAUSE is one line from a fixed list of about 15 causes.** It makes the model locate the bug
  first, which is cheap reasoning, and lets us score its diagnosis separately. There are no
  free-text explanations, because that is where small models make things up.
- **The output is the full corrected procedure, not a diff.** It uses the same output grammar as
  normal generation, and there's nothing to apply wrongly. The data only keeps *minimal* fixes
  (a few changed lines), so the model learns small, precise repairs rather than rewrites.
- **15% are attempt 2 or 3:** the input includes the previous fix and its new failure, so the
  model learns not to repeat a failed fix.
- **They can be reused later:** broken versus fixed code can serve as preference pairs if we try
  DPO after v1.

## 5. VBA-Bench (300): build it first, then freeze it

- Build it in **weeks 1 to 3**, freeze it as v1, and only then generate training data. Training
  items are checked for duplicates against it.
- **200 in-distribution** items, stratified like the training set.
- **60 from the 2 held-out archetypes**, to test generalisation rather than memory.
- **40 from your real work tasks**, described in plain words with no names or figures.
  These are the most valuable items.
- Every eval item has a verified reference solution, which proves its TESTs can be passed.
- Claude's first-try pass rate on VBA-Bench is the ceiling to aim for (roadmap). Anything derived
  from an eval item (repairs, contrasts) never goes into training.

## 6. Workflow: two cloud factories and one human step

```
SPEC FACTORY (cloud)  ->  CODE FACTORY (cloud)  ->  EXCEL (you, about 20 min)  ->  IMPORT (cloud)
request -> spec          compiler, or Claude for     paste 1 module, RunBatch,    verified -> data/verified
validator, simulator,    NOTE and extended steps;    copy the Results back        failures -> repair queue
dedup                    lint + VBA grammar check                                  -> next batch
```
- **Specs and code in separate passes.** Specs are frozen before any code is written, so a spec
  can't be bent to fit the code. The simulator catches spec bugs in seconds, not in Excel.
  Specs can be made far ahead, about 200 per session.
- **Generation runs in Claude Code sessions straight into the repo:** no copying prompts in and
  out of claude.ai, and no paid API. The only manual step is your Excel run.
- **Batches of 100 items** (about 350 TESTs; my guess is 5 to 10 minutes of running, and the
  capacity batch will measure it) once a *capacity batch* in late Phase 1 shows a module that
  size pastes and runs smoothly. Until
  then, 50.
- **Keeping the rhythm going:**
  - three fixed slots a week (Monday, Wednesday, Friday)
  - the next batch always ready in `batches/`
  - the README shows the dataset count and the next batch
  - pasting back is one copy
- **If we fall behind:** once the simulator and Excel have agreed on every measure for 10 batches
  in a row, compiler-made items could be checked by sampling 20% instead of all of them. That's
  your call when the time comes.
- **Engine upgrades for big batches:**
  - After a hang, RunBatch resumes: a test left at RUNNING is marked HANG and skipped.
  - PASS rows get a shorter Description.
  - Fixture rows are packed several to a line.

## 7. Record format

- One JSON file per item in `data/verified/`, so git diffs stay readable. Fields:
  - id, split, source (compiler, Claude, model or mutant), category, archetype, tier
  - **request** (the plain-words task), spec, VBA, batch, Excel results, simulator results
  - VCL and engine versions
- Training files are exported from these records by a script and never edited by hand. Each one is:
  - the fixed system prompt
  - the spec, as the user turn
  - the procedure without line numbers, as the assistant turn
- Keeping the request text costs nothing. It leaves room for a later option: teaching the local
  model to write the *spec* from your request, which is where real reasoning (and offline value)
  would be.

## 8. "Verified" means all five of these

1. The spec validates with 0 errors.
2. The simulator passes every TEST, and the TESTs catch at least 80% of spec mutants. Weak tests
   (like task_07's old goal test) are rejected.
3. Lint and the VBA grammar check are clean.
4. Excel: every TEST passes, and every measure equals the simulator's.
5. It isn't a near-duplicate of anything, including the eval set.

## 9. What this would change in Phase 1, if approved

- The compiler emits code without line numbers; a separate step adds them, the same way the
  model's output will be handled.
- `harness/lint.py` gets the 2016/2019 ban list and a real VBA syntax check: an open-source ANTLR
  VBA grammar, after checking its licence and its accuracy on our pilot. That would be the first
  pinned dependency.
- A capacity batch of 100 items before Phase 1 ends.
- RunBatch resumes after a hang.

## 10. Decisions for Boss

1. **v1 scope:** the reallocated categories in section 1, with Pivot tables, UserForms, Classes
   and external automation moved later?
2. **Dirty fixtures by default** (VCL rev 4, together with the ListObject ops)?
3. **100 items per paste** after the capacity test?
4. **Repair format:** a CAUSE line plus the full procedure with a minimal fix?
5. **Real tasks:** can you note one real task from work in plain words (no names, no figures)
   whenever one comes up, aiming for 40 by December?
6. **One pinned dependency** (the ANTLR runtime) for the VBA syntax check?
