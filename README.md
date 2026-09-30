# Advanced Grading Console

**BITS Digital CodeForge V1.0 — Debug. Reimagine. Deploy.**

A single-file web application that turns a spreadsheet of examination marks into a
finalized grade sheet. Faculty upload a marks workbook, choose a course, adjust the
eight grade cutoffs while watching the distribution respond, review students sitting
near a boundary, and export an auditable CSV.

The starting point was a deliberately broken 500-line application. This repository
contains the repaired and extended version.

| | |
|---|---|
| **Repository** | https://github.com/cybersd12321/bits-codeforge-grading-console |
| **Live demo** | **https://cybersd12321.github.io/bits-codeforge-grading-console/** |
| **Mirror** | `https://<project>.vercel.app` — *optional; see [Vercel](#vercel)* |
| **Stack** | One `index.html`. No build step, no framework, no package manager. |
| **Dependency** | [SheetJS](https://sheetjs.com) `xlsx.full.min.js` via CDN — the only external script, unchanged from the original brief. |

---

## Contents

| File | Purpose |
|---|---|
| `index.html` | The application. Markup, styles and logic in one file. |
| `BITS_Digital_CodeForge_Challenge.html` | The original buggy file, kept for before/after comparison. |
| `sample-marks-clean.xlsx` | 40 students across 4 courses, no anomalies. Exercises the green "file verified" path. |
| `sample-marks-anomalies.xlsx` | 46 rows containing fractional marks, out-of-range scores, a duplicate ID, a missing ID and two unreadable rows. Exercises the anomaly inspector. |

Both sample workbooks also contain a **single-student course** (`BIO F111`) and a
**zero-variance course** (`PHY F110`, three students all on 55) so the statistical
guard paths can be exercised without hand-crafting a file.

## Running it locally

No install, no server required.

```bash
git clone https://github.com/cybersd12321/bits-codeforge-grading-console.git
cd bits-codeforge-grading-console
start index.html          # Windows
# open index.html         # macOS
# xdg-open index.html     # Linux
```

Opening it straight from the filesystem works, because the page loads no local
assets. If you prefer a server: `python -m http.server 8000`, then visit
`http://localhost:8000`.

---

## Stage 1: Bug Fix Log

Ten defects from the challenge brief, each reproduced before being fixed and
re-tested afterwards.

| # | Bug / Issue Identified | How You Reproduced It | Root Cause | Fix Implemented | How You Tested the Fix |
|---|---|---|---|---|---|
| 1 | **Modern Excel files could not be opened, and parsing corrupted cells** | Clicked the file picker with a `.xlsx` workbook — the file was greyed out and unselectable. Forcing it through "All files" produced garbled cell values on some browsers. | Two faults in one path. The input carried `accept=".xls"`, excluding the OOXML extension that every current Excel writes. Parsing used `FileReader.readAsBinaryString()` with `XLSX.read(data, {type:"binary"})`, which depends on a deprecated byte-to-character mapping that silently corrupts bytes. | `accept=".xlsx,.xls"`. Switched to `readAsArrayBuffer()` and `XLSX.read(new Uint8Array(buffer), {type:"array"})`. Added a `typeof XLSX === "undefined"` guard plus a `try/catch` that surfaces a readable message instead of dying inside the reader callback. | Generated `.xlsx` fixtures and pushed them through the real change handler in a DOM harness: 40 records across 4 courses parsed, headers resolved to `BITS ID` / `Course` / `Total Marks`. A 5-byte garbage file produced a user-facing explanation, not a thrown error. |
| 2 | **Course dropdown kept old courses and repeated each one** | Uploaded file A (3 courses), then file B (2 courses) — the dropdown listed all five. With 24 students in `CS F213`, that course name appeared 24 times. | `data.map(d => d.Course).forEach(c => course.add(new Option(c, c)))` iterated **every row** with no de-duplication, and nothing cleared the existing `<option>` elements before repopulating. | `resetCourseOptions()` now runs at the top of the change handler, *before* the asynchronous read, so a cancelled or failed second upload cannot leave the previous file's courses behind. Population uses `[...new Set(records.map(r => r.course))]`, sorted with `localeCompare`. | Uploaded A then B and asserted `options.length === courses + 1` (placeholder included). Ran 10 consecutive course switches and re-uploads and asserted no accumulation. |
| 3 | **Min and Max statistics displayed each other's value** | Loaded a course with marks spanning 18–92. The card labelled Min read 92; the card labelled Max read 18. | The markup was cross-wired: `<div class="stat">Min<br><b id="max">` sat above `Max<br><b id="min">`. The code wrote `min.textContent = m[0]` relying on implicit `id`-to-global binding, so the lowest value was written into the node visually labelled **Max**. | Renamed the value nodes to `statMin`, `statMax`, `statAvg`, `statMed` and replaced every implicit global with an explicit `document.getElementById` reference. | Asserted `statMin === "18"` and `statMax === "92"` on a real cohort. Added a static check that every `getElementById` string literal in the script resolves to an `id` that exists in the markup — the class of mistake that let this hide. |
| 4 | **Empty and single-student courses produced `NaN` and `undefined`** | Selected a course with no matching rows: Avg showed `NaN`, Min showed `undefined`. Uploaded an empty sheet: the histogram routine ran its binning path on an empty array. | `m.reduce((a,b)=>a+b,0) / m.length` evaluates `0/0 → NaN` when the array is empty; `m[0]` and `m[m.length-1]` return `undefined`. Neither `computeStats()` nor `drawHistogram()` checked length before doing arithmetic. | `computeStats()` returns `null` for `n === 0` and the renderer falls back to an em dash. `drawHistogram()` paints a labelled empty state and returns before any maths. `drawBellCurve()` returns early when `n < 2`, since a single point has no distribution. | `computeStats([]) === null`. A real 1-student course (`BIO F111`, 66 marks) asserted min = max = avg = median = 66. Drove an empty upload, a header-only upload and a 1-student course end to end with zero console errors. |
| 5 | **Median concatenated strings instead of averaging them** | A four-student course whose marks arrived as strings `"79","81","60","40"` reported a median in the thousands rather than 69.5. | `(m[m.length/2-1] + m[m.length/2]) / 2` — with string operands, `+` is concatenation. `"60" + "79"` becomes `"6079"`, which is then divided by 2. | `const median = n % 2 ? Number(sorted[mid]) : (Number(sorted[mid-1]) + Number(sorted[mid])) / 2`, with the array numerically coerced and sorted before the split. | Asserted the median of `["79","81","60","40"]` is exactly `69.5`. Cross-checked against real data: the 24-student `CS F213` cohort returns `63`, and the borderline pair (78, 79) returns `78.5`. |
| 6 | **Fractional marks were graded as fractions** | A sheet containing `80.2` carried that value into binning, cutoff comparison and the exported CSV, so a mark could fall inconsistently around an integer boundary. | No normalisation existed between `sheet_to_json` and the grading logic. Whatever Excel held was used verbatim. | `Math.round()` is applied **once**, at the parse boundary, together with clamping into 0–100. Every downstream consumer — histogram, statistics, cutoff comparison, CSV — therefore sees whole marks only. Per the BITS grading instruction, rounding is to the nearest whole integer. | Asserted `80.2 → 80`, `79.6 → 80`, `49.5 → 50`, and that every stored mark satisfies `Number.isInteger(m) && m >= 0 && m <= 100`. Verified through a real workbook, including a number-formatted cell. |
| 7 | **Normal curve sat left of the bars, and identical marks produced `Infinity`** | With three students all scoring 55, `std` evaluated to 0 and the curve height became `1/(0 × √2π)` — `Infinity`. Separately, on any normal cohort the curve's peak visibly sat to the left of the tallest bar. | No zero-variance guard existed. The x mapping was `px = 30 + (x/10) * 30`, a hardcoded 30px step, while the bars were laid out at `30 + i*32` with width 24 — placing true bin centres at `42 + i*32`. The two drifted apart by 2px per bin, 18px by the last one. | Added a guard: non-finite or zero `std` returns before the division and instead draws a dashed marker at the mean, so identical marks still communicate something. Introduced a single `markToX()` derived from the shared chart constants, so the bars and the curve read their geometry from the same source and cannot drift again. | Asserted `markToX(i*10 + 5)` equals the bar centre `30 + i*32 + 12` for all ten bins, to within floating-point tolerance. Confirmed the zero-variance course genuinely yields `std === 0` and renders with no `Infinity` or `NaN`. |
| 8 | **Reset raised two dialogs; hover and pulse animations never appeared** | Clicking "Reset Range" showed a confirmation, and dismissing it immediately raised a second, near-identical one. Changing a grade dropdown produced no visible lift, and changed summary pills never pulsed. | Two sequential `confirm()` guards sat in the same handler. The animation classes were removed by `setTimeout(..., 30)` — 30ms — while the corresponding CSS transitions run for `.25s`, so the class was gone long before the transition could render. | Reduced to a single, clearly worded confirmation. Introduced `ANIM_MS = 260` so the class always outlives the 250ms transition, and applied it to the card lift and the pill pulse alike. | Asserted exactly one `confirm(` call inside the reset handler, and that no `setTimeout(..., 30)` remains anywhere. Drove both the accept and the decline path: declining leaves edited cutoffs untouched. |
| 9 | **The grading timer measured how long the tab had been open** | Opened the page, left it for five minutes, then entered a name and selected a course. The elapsed time already read 05:00, and the finalize message reported that as grading time. | `const gradingStartTime = Date.now()` executed at module scope, and `startTimer()` was invoked from a `DOMContentLoaded` listener. The clock therefore began at page load regardless of whether grading had started. The display also sat blank for a full second waiting for the first interval tick. | `gradingStartTime` is initialised to `null`. `startGradingTimer()` is called from exactly one place — the course-selection handler, after the instructor-name check passes — so the clock starts only when an instructor and a valid course both exist. It paints immediately, then ticks. Later calls resume the display without rewinding the total. | Asserted that no interval is scheduled at boot and the clock reads `00:00`; that `startGradingTimer()` has exactly one call site and it lies inside `courseEl.onchange`; and that no `DOMContentLoaded` handler starts it. Simulated 7 min 23 s and confirmed the reported figure matched. |
| 10 | **Exported CSV was malformed and opened incorrectly in Excel** | An instructor named `Dr. Rao, Meera` split across two columns. A metadata row read `Course, CS F213` with a leading space inside the value. Non-ASCII characters appeared as mojibake when the file was opened in Excel. | The CSV was assembled by string concatenation — `` `Course, ${course.value}` `` — with a stray space baked into the template, no RFC 4180 quoting of any field, and no UTF-8 byte-order mark for Excel to detect the encoding. | A `csvField()` helper quotes any field containing a comma, double quote, CR, LF or edge whitespace, and escapes inner quotes by doubling them. Rows are joined with CRLF per RFC 4180, `\uFEFF` is prepended, and the blob is typed `text/csv;charset=utf-8`. | Captured the generated blob and asserted: the BOM is present; no bare `\n` exists; `Instructor,"Dr. Rao, Meera"` is quoted correctly; `Course,CS F213` has no stray space; and IDs containing commas and quotes round-trip byte-exactly. |

### Additional defects found during a follow-up audit

Three further faults surfaced when the finished application was executed
scenario-by-scenario rather than only read. They are listed separately because they
were not part of the original brief.

| # | Bug / Issue Identified | How You Reproduced It | Root Cause | Fix Implemented | How You Tested the Fix |
|---|---|---|---|---|---|
| 11 | **An empty workbook reported success** | Uploaded an empty sheet. The health banner showed a calm green *"File verified: 0 valid student records across 0 courses."* while the red line directly beneath it said nothing was gradable — two contradictory statements about the same file. | The banner's tone was derived only from the presence of *anomalies*. An empty file has none, so it fell through to the "all clear" branch. Zero records was never itself treated as a condition worth reporting. | Zero kept records now forces the warning tone regardless of anomaly count, with a message that says plainly what happened: *"No student rows found on that workbook's first sheet."* | Uploaded an empty sheet and a header-only sheet; asserted the banner carries the `warn` class, states that nothing was found, and does not repeat itself in a chip. |
| 12 | **The CSV carried no record of how long grading took** | Finalized a course, then opened the CSV. It held the instructor and course but no timing information — the elapsed figure existed only in an on-screen message that disappears with the page. | `minT` and `secT` were computed in the download handler but interpolated solely into the confirmation text. Nothing wrote them to the file. | The export now opens with a self-describing metadata block: instructor, course, students graded, a sortable finalize timestamp, grading time, and the attempt number. | Drove two finalizes against a controlled clock: the first recorded `7 min 23 sec`, the second a cumulative `11 min 5 sec` with the attempt counter advancing, and the on-screen message agreed with the file in both cases. |
| 13 | **Numeric BITS IDs lost their leading zeros** | A sheet where `0012345` was stored as a *number* carrying a `0000000` display format exported as `12345`. | `sheet_to_json` with raw values returns the underlying number, `12345`. The leading zeros existed only in the cell's display format, which the raw read discards. | The sheet is read a second time with `raw:false`, and the formatted text is preferred for the BITS ID — what the instructor sees in the cell is what gets exported. Guarded on row-count alignment between the two passes, falling back to raw values if they ever diverge. | Built a workbook with an explicitly number-formatted ID cell and confirmed `12345` now exports as `0012345`, while alphanumeric IDs, large numeric IDs, and IDs containing commas or quotes remain byte-exact. |

---

## Stage 2: Reimagined Product Enhancements

The brief is explicit that **more features ≠ better product**. Grade finalization is
a narrow, high-stakes task performed a few times a year under deadline, and a
cluttered tool costs more than a sparse one. So the test each addition had to pass
was not "is this nice to have" but **"does an instructor currently do this by hand,
or get it wrong?"** Three did. Everything else was left out.

### 1. Academic Data Health & Anomaly Inspector

The moment a workbook is parsed, a single line appears beneath the controls
reporting what was read and what had to be corrected: rows parsed, fractional scores
rounded, scores clamped into 0–100, duplicate BITS IDs, rows missing an ID, and
unreadable rows skipped. Repeated IDs expand into a collapsed list naming exactly
which ID in which course.

**Why this exists.** Marks files are assembled by hand, often by several people, and
an instructor has no way to audit one before grading from it. Previously the
application silently absorbed every defect: a mark of `105` simply fell outside all
eight bands and vanished from the export, and a student entered twice was graded
twice with nobody informed. A grade sheet is a statutory record, so a defect noticed
after finalization is a correction to someone's transcript.

**Restraint applied.**

- **Three tones, not two.** Green means nothing to report. Violet means the file was
  corrected automatically and nothing is wrong — rounding a fraction is expected
  behaviour, not a fault. Amber is reserved for things that genuinely need a human.
  A banner that shouts at every file gets ignored by the third upload.
- **Flag, never drop.** Duplicates are reported but kept. Silently removing a row
  would quietly change a transcript, which is worse than the duplicate.
- **Detail on demand.** The anomaly list is collapsed by default, so the clean case
  costs exactly one line of vertical space.

### 2. Boundary Auto-Cascader with Live Movement Delta

The eight bands are really seven cutoffs: the floor of A down to the floor of D. Each
band's ceiling is "the next cutoff minus one", which makes a continuous scale the
*only* state the interface can represent. Editing one boundary pins it and pushes the
others out of its way. Alongside the edit, the affected band reports its consequence
in plain language — *"+3 students moved into B"*.

**Why this exists.** The original application let an instructor create gaps and
overlaps, then scolded them with *"Grade ranges must be continuous with no gaps or
overlaps"* and disabled the export until they solved the arithmetic themselves. That
is the tool refusing to do its own job. Worse, nothing prevented A's ceiling drifting
below 100, which left top marks with no grade at all.

The delta answers the only question that actually matters when moving a cutoff: *who
does this move?* Faculty were previously reading eight counts before and after each
change and diffing them mentally.

**Restraint applied.**

- **Fewer controls, not more.** Because the ends of the scale are structural rather
  than preferences, A's maximum and E's minimum are locked and labelled instead of
  offered as editable fields. An input you are not allowed to use is worse than a
  plain number.
- **One delta, not eight.** A cascade can shift several bands. Only the band you
  edited reports a sentence; the rest use the existing pulse. Eight simultaneous
  messages would be noise.
- **No layout jump.** The delta slot holds its height whether or not it contains
  text, and fades after a few seconds.

Verified exhaustively: all 707 possible minimum edits and all 707 maximum edits keep
the scale continuous, 4000 chained random edits never break it, and across 200 random
configurations every mark from 0 to 100 maps to exactly one band.

### 3. Interactive Roster & Cutoff Inspection

The grade summary pills became toggle filters, and a compact roster below the
workspace lists BITS ID, course, total marks and assigned grade, sorted by marks
descending. Clicking a pill narrows the roster to that band; "Clear Filter" restores
it. Students within two marks of the next grade up are marked with a small chip
reading, for example, *"1 from A"*.

**Why this exists.** Setting a cutoff is a judgement about specific people, not a
number in the abstract. The application showed distributions and counts but never the
students, so an instructor deciding whether to drop the A boundary by one mark had to
reopen the spreadsheet and sort it by hand. The borderline chip answers the question
that motivates the whole exercise — who is one mark short — without asking faculty to
subtract in their heads.

**Restraint applied.**

- **No new columns.** The borderline indicator lives inside the Grade cell rather
  than adding a fifth column.
- **No second control.** Filtering reuses the pills that were already displaying
  those counts, instead of adding a filter bar that duplicates them.
- **Sorted, not sortable.** Marks-descending is the order cutoff review actually
  uses. Clickable column headers would have been four more controls serving a rarer
  need.
- **Focus preserved.** Toggling a filter updates the pressed state in place rather
  than rebuilding the pills, so keyboard focus stays where the user put it.

### What was deliberately left out

Turned down as failing the same test: per-column sorting, CSV import of previous
grade sheets, multi-course batch export, a printable summary view, dark mode,
per-student grade overrides, and a second chart of the grade distribution (the
histogram and the band counts already cover it from two useful angles).

---

## Verification

The application was not only read but **executed**. The script was run inside a
minimal DOM harness so each scenario could be driven and observed, and a final pass
recorded **136 checks passing with zero failures**, covering:

- **Scoping** — strict mode throughout, zero `var`, zero accidental globals (verified
  by diffing the execution context against the declared top-level functions), no
  implicit loop variables.
- **Zero uncaught errors** across an empty sheet, a header-only sheet, a 5-byte
  garbage file, a single-student course, a zero-variance course, ten consecutive
  course switches, re-uploading mid-session, resetting cutoffs while a filter is
  active, and filtering to a band then emptying it.
- **Keyboard access** — visible focus rings on inputs, dropdowns, pills, the summary
  disclosure and the scrollable roster, with a `:focus` fallback for browsers lacking
  `:focus-visible`. No positive `tabindex` values; locked controls are `disabled` so
  Tab skips them.
- **Export fidelity** — verified byte-for-byte from the generated blob.
- **Cutoff integrity** — the exhaustive cascader sweep described above.

Accessibility note: these checks cover keyboard operation, focus visibility,
contrast, semantic table structure and live-region announcements. Full WCAG
conformance additionally requires manual testing with assistive technology and expert
review, which has not been performed here.

---

## Deployment

`index.html` sits at the repository root, which is all either platform needs. There
is no build step and no `package.json`, so both hosts serve the file directly.

### GitHub Pages

1. Confirm `index.html` is at the repository root on the default branch:
   ```bash
   git ls-tree --name-only HEAD | findstr index.html
   ```
2. In the repository, open **Settings → Pages**.
3. Under **Build and deployment**, set **Source** to `Deploy from a branch`.
4. Set **Branch** to `main` and the folder to `/ (root)`. Click **Save**.
5. Open the **Actions** tab and wait for the `pages-build-deployment` run to finish
   green. First publish usually takes under two minutes.
6. The URL appears back on the Pages settings screen. For this repository it will
   be:

   ```
   https://cybersd12321.github.io/bits-codeforge-grading-console/
   ```

### Vercel

**From the dashboard**

1. **Add New → Project**, then import this GitHub repository.
2. Set **Framework Preset** to `Other`.
3. Leave **Build Command** and **Install Command** empty.
4. Set **Output Directory** to `./`.
5. **Deploy**.

**From the CLI**

```bash
npm i -g vercel
vercel          # preview deployment
vercel --prod   # production deployment
```

Because the repository contains no `package.json`, Vercel classifies it as a static
site and skips the build stage. If it ever reports a missing build command, the
Output Directory setting above is what to correct.

### Deployment verification checklist

Run this against the live URL, not just localhost. Both platforms serve over HTTPS,
so the SheetJS CDN script loads without mixed-content blocking.

The GitHub Pages deployment has been checked at the transport level already: the site
returns **200** over HTTPS, HTTP requests are **301**-upgraded, HSTS is set, the served
bytes are identical to the committed `index.html`, the SheetJS CDN resolves, there are
no `http://` subresources, both sample workbooks download intact, and the two
uncommitted files correctly return **404**. The rows below are the interactive checks,
which need a real browser.

| # | Check | Expected result |
|---|---|---|
| 1 | Open the URL with DevTools console visible | Page renders; console is clean |
| 2 | Network tab, filter `xlsx` | `xlsx.full.min.js` returns **200** over `https://` |
| 3 | Before touching anything, read the clock | Shows `00:00` and is not counting |
| 4 | Chart area before upload | Reads *"No marks to display yet"*, no blank canvas |
| 5 | Upload `sample-marks-clean.xlsx` | Green banner: *"File verified: 40 valid student records across 4 courses."* |
| 6 | Open the course dropdown | Exactly 4 courses plus the placeholder, each listed once |
| 7 | Select a course without entering a name first | Prompts for the instructor name and clears the selection |
| 8 | Enter a name, then select `CS F213` | Clock starts from `00:00`; 24 students; histogram and curve animate in |
| 9 | Read the statistics row | Min **18**, Max **92**, Median **63** — Min must be the lower number |
| 10 | Raise B's minimum from 60 to 64 | Adjacent bands follow automatically, **no** continuity error appears, and B reports *"-2 students moved out of B"* |
| 11 | Click the `A` summary pill | Roster narrows to grade A; "Clear Filter" becomes available |
| 12 | Upload `sample-marks-anomalies.xlsx` | Amber banner listing rounded, clamped, duplicate, missing-ID and skipped rows; the dropdown rebuilds cleanly with no leftovers |
| 13 | Select `BIO F111` (one student) | Statistics populate with no `NaN`; no bell curve is drawn |
| 14 | Select `PHY F110` (identical marks) | Renders without `Infinity`; a dashed line marks the mean |
| 15 | Click **Finalize & Download** | `grades-CS_F213.csv` downloads |
| 16 | Open the CSV in Excel | No mojibake; metadata block shows instructor, course, headcount, timestamp, grading time and attempt; BITS IDs are intact |
| 17 | Tab through the page | Every control shows a visible focus ring, in a sensible order |
| 18 | Narrow the window to phone width | Columns stack; the roster scrolls rather than overflowing |

---

## Known considerations

- **SheetJS is loaded from a CDN.** On a network that blocks jsDelivr the application
  detects this and shows *"The spreadsheet library could not be loaded"* rather than
  failing silently. For fully offline use, download `xlsx.full.min.js` and point the
  `<script>` tag at a local copy.
- **Excel re-types CSV columns on import.** The exported file holds exact
  characters — verified byte-for-byte — but Excel's importer will still re-parse an
  unquoted numeric field as a number when opening a `.csv`. This is a property of
  Excel's import heuristics, not of the file. Use *Data → From Text/CSV* and mark the
  BITS ID column as Text to preserve it visually.
- **CSV formula injection is not neutralised.** A BITS ID beginning with `=`, `+`, `-`
  or `@` would be interpreted as a formula by Excel. Every mitigation alters the
  stored value, which would contradict the requirement that IDs export exactly, so
  the data is kept faithful and the risk is documented instead.
- **The challenge PDF and `Info.txt` are intentionally not committed.** They are BITS
  Pilani's own material, and `Info.txt` contains Google Docs and Forms links
  restricted to BITS accounts.

---

## Credits

Built for **BITS Digital CodeForge V1.0**. The original challenge application and the
grading rules are the property of BITS Pilani Digital.
