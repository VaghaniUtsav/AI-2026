# MyBI — Like-for-Like Comparison Measure Set (toggle-aware)

> Fixes the "4 days of June vs full May = -86%" problem. The engine decides, per the toggle,
> whether the selected period is **LIVE** (still in progress → compare to-date) or **FINISHED**
> (a past period → compare the full period), then builds matching Previous and Prior-Year windows.
> One engine drives all five toggle options (YDAY / WTD / MTD / QTD / YTD) for both **Previous
> (MoM-style)** and **Prior Year (PY)**, for **Users** and **Sessions**. PBIRS-safe: no Calculation
> Groups, no standard time-intel, boolean predicates + `REMOVEFILTERS`.

## How it works (the idea in one paragraph)
There are two "today" dates. The **anchor** = `MAX('Date Table'[Date])` follows whatever you pick
in the date slicer. The **latest data date** = the newest date in the whole model, ignoring every
slicer — the platform's real "today". The engine compares the anchor's period to the latest-data
period: the selected date only picks **which bucket** (week/month/quarter/year). If that bucket is
the **same one** that holds the latest data, the period is **live** and runs to **today** (the
latest data date) — so clicking any day in Q2 gives April + May + June-to-date, not "up to the day
you clicked". If the selected bucket is **earlier and finished**, the period ends at the bucket's
**natural last day** (full period). Then the comparison switches on whether the current window is
**full** (its end = the bucket's last day) or **partial**: a **full** period compares to the
**whole** previous / prior-year bucket (full Feb → full Jan, even though 28 ≠ 31 days; full Q1 →
full Q4; leap years handled); a **partial** period compares to the **same number of days** into the
prior bucket, clamped. Result: live month → 1–7 Jun vs 1–7 May vs 1–7 Jun 2025; any finished month
(pick the 14th or the 28th of Feb, doesn't matter) → full Feb vs full Jan vs full Feb LY. Nine small
"boundary" measures compute the dates once; every value and label measure just reads them.

---

## STEP 0 — Prerequisite: the toggle must be a slicer table

These measures react to a single-column disconnected table read via `SELECTEDVALUE`. If your
YDAY/WTD/MTD/QTD/YTD pills are already a **slicer on a table**, just match the table/column name
below. If they are **bookmarks**, convert them to a slicer (these measures can't read a bookmark).

```dax
'Period' =
DATATABLE(
    "Period", STRING, "Sort", INTEGER,
    {
        { "YDAY", 1 },
        { "WTD",  2 },
        { "MTD",  3 },
        { "QTD",  4 },
        { "YTD",  5 }
    }
)
```
- Keep it **disconnected** (no relationship).
- Sort `Period` by `Sort` (Column tools → Sort by column) so the pills read YDAY→YTD.
- Put a slicer on `'Period'[Period]`, style **Tile**, single-select, default **MTD**.
- If your existing toggle table already exists, skip this and replace `'Period'[Period]` everywhere
  below with your real column name.

---

## STEP 1 — Boundary engine (9 measures, hide them all)

Put these in a `_Engine` display folder and **hide** them (right-click → Hide). They return dates,
not numbers, so they never go on a visual directly. Create them in this order (each one references
the ones above it).

```dax
-- (1) The newest date that actually has data, ignoring ALL slicers.
--     This is the platform's real "today"; it decides whether the selected
--     period is still LIVE (compare to-date) or already FINISHED (compare full).
_Latest Data Date = CALCULATE( MAX('SESSION USAGE'[Date]), REMOVEFILTERS() )

-- (2) First day of the period bucket that contains the anchor (the selected date).
_Period Start =
VAR A   = [_Anchor Date]
VAR P   = SELECTEDVALUE('Period'[Period], "MTD")
VAR Qtr = ROUNDUP( MONTH(A) / 3, 0 )
RETURN
    SWITCH( P,
        "YDAY", A,
        "WTD",  A - ( WEEKDAY(A, 2) - 1 ),               -- Monday of the anchor's week
        "MTD",  DATE( YEAR(A), MONTH(A), 1 ),            -- 1st of the anchor's month
        "QTD",  DATE( YEAR(A), (Qtr - 1) * 3 + 1, 1 ),   -- 1st day of the anchor's quarter
        "YTD",  DATE( YEAR(A), 1, 1 ),                   -- 1 Jan of the anchor's year
        DATE( YEAR(A), MONTH(A), 1 )                      -- fallback = MTD
    )

-- (3) The natural LAST day of the bucket that contains the anchor.
_Bucket End =
VAR A   = [_Anchor Date]
VAR P   = SELECTEDVALUE('Period'[Period], "MTD")
VAR S   = [_Period Start]
VAR Qtr = ROUNDUP( MONTH(A) / 3, 0 )
RETURN
    SWITCH( P,
        "YDAY", A,                                        -- a day is atomic
        "WTD",  S + 6,                                    -- Sunday of the week
        "MTD",  EOMONTH( A, 0 ),                          -- last day of the month
        "QTD",  EOMONTH( DATE( YEAR(A), Qtr * 3, 1 ), 0 ),-- last day of the quarter
        "YTD",  DATE( YEAR(A), 12, 31 ),                  -- 31 Dec
        EOMONTH( A, 0 )
    )

-- (4) End of the CURRENT period.
--     The selected date only picks the BUCKET (which week/month/quarter/year).
--     If that bucket is the one holding the latest data, the period is LIVE and
--     runs to TODAY (the latest data date) -> e.g. select any day in Q2 and you
--     get April + May + June-to-date, not "up to the day you clicked".
--     If the bucket is FINISHED, the period ends at its natural last day.
--     (YDAY is exempt: IsLiveBucket is always FALSE, so it ends on the chosen day.)
_Period End =
VAR A    = [_Anchor Date]
VAR L    = [_Latest Data Date]
VAR P    = SELECTEDVALUE('Period'[Period], "MTD")
VAR S    = [_Period Start]
VAR BEnd = [_Bucket End]
VAR Qtr  = ROUNDUP( MONTH(A) / 3, 0 )
VAR IsLiveBucket =
    SWITCH( P,
        "YDAY", FALSE(),                                  -- a single day is never "partial"
        "WTD",  L >= S && L <= BEnd,
        "MTD",  YEAR(A) = YEAR(L) && MONTH(A) = MONTH(L),
        "QTD",  YEAR(A) = YEAR(L) && Qtr = ROUNDUP( MONTH(L) / 3, 0 ),
        "YTD",  YEAR(A) = YEAR(L),
        FALSE()
    )
RETURN
    IF( IsLiveBucket, L, BEnd )                           -- LIVE -> today (L); FINISHED -> bucket end

-- (5) First day of the immediately-previous equivalent period.
_Prev Start =
VAR A = [_Anchor Date]
VAR P = SELECTEDVALUE('Period'[Period], "MTD")
VAR S = [_Period Start]
RETURN
    SWITCH( P,
        "YDAY", A - 1,                                    -- the day before
        "WTD",  S - 7,                                    -- previous week's Monday
        "MTD",  EDATE( S, -1 ),                           -- 1st of previous month
        "QTD",  EDATE( S, -3 ),                           -- 1st of previous quarter
        "YTD",  DATE( YEAR(A) - 1, 1, 1 ),                -- 1 Jan last year
        EDATE( S, -1 )
    )

-- (6) End of the previous period. The key fix:
--     FULL current period (E = bucket end)  -> the WHOLE previous bucket
--        (full Feb -> full Jan; 28 vs 31 days no longer matters).
--     PARTIAL current period (in progress) -> same days in, clamped to bucket end.
_Prev End =
VAR P      = SELECTEDVALUE('Period'[Period], "MTD")
VAR S      = [_Period Start]
VAR E      = [_Period End]
VAR IsFull = ( E = [_Bucket End] )                        -- did the current window cover the whole bucket?
VAR Offset = INT( E - S )                                 -- 0-based length of the current window
VAR PStart = [_Prev Start]
VAR PBucketEnd =
    SWITCH( P,
        "YDAY", PStart,
        "WTD",  PStart + 6,
        "MTD",  EOMONTH( PStart, 0 ),
        "QTD",  EOMONTH( PStart, 2 ),                     -- PStart is the quarter's 1st month
        "YTD",  DATE( YEAR(PStart), 12, 31 ),
        EOMONTH( PStart, 0 )
    )
RETURN
    IF( IsFull, PBucketEnd, MIN( PStart + Offset, PBucketEnd ) )

-- (7) HELPER (WTD only): Monday of the SAME ISO week, one ISO year earlier.
--     ISO rules: weeks start Monday; week 1 is the week containing 4 Jan; the
--     week's ISO year is the year of its Thursday. ISO week NUMBER is computed
--     MANUALLY (this PBIRS build has no ISOWEEKNUM) as the count of whole weeks
--     between the week's Monday and the Monday of ISO week 1. We map that week
--     number back one ISO year and rebuild the Monday, clamping 53 -> 52 when the
--     prior ISO year is shorter. Uses only WEEKDAY / DATE — both PBIRS-safe.
_PY Week Start (ISO) =
VAR S         = [_Period Start]                           -- Monday of the anchor's ISO week
VAR Thu       = S + 3                                     -- Thursday fixes the ISO year
VAR ISOYear   = YEAR( Thu )
-- ISO week 1 Monday of THIS ISO year (week containing 4 Jan):
VAR Jan4      = DATE( ISOYear, 1, 4 )
VAR Week1Mon  = Jan4 - ( WEEKDAY( Jan4, 2 ) - 1 )
VAR ISOWeek   = INT( ( S - Week1Mon ) / 7 ) + 1           -- manual ISO week number
-- Prior ISO year: its week-1 Monday and its total week count (28 Dec is always last week):
VAR PrevYear  = ISOYear - 1
VAR PJan4     = DATE( PrevYear, 1, 4 )
VAR PWeek1Mon = PJan4 - ( WEEKDAY( PJan4, 2 ) - 1 )
VAR PDec28    = DATE( PrevYear, 12, 28 )
VAR PDec28Mon = PDec28 - ( WEEKDAY( PDec28, 2 ) - 1 )
VAR MaxWeek   = INT( ( PDec28Mon - PWeek1Mon ) / 7 ) + 1  -- 52 or 53
VAR TgtWeek   = MIN( ISOWeek, MaxWeek )                   -- clamp 53 -> 52 if no W53 last year
RETURN
    PWeek1Mon + ( TgtWeek - 1 ) * 7                       -- Monday of the same ISO week last year

-- (8) First day of the same period one year earlier.
_PY Start =
VAR A = [_Anchor Date]
VAR P = SELECTEDVALUE('Period'[Period], "MTD")
VAR S = [_Period Start]
RETURN
    SWITCH( P,
        "YDAY", IF( MONTH(A) = 2 && DAY(A) = 29, DATE( YEAR(A) - 1, 2, 28 ), DATE( YEAR(A) - 1, MONTH(A), DAY(A) ) ),
        "WTD",  [_PY Week Start (ISO)],                   -- same ISO week number, prior ISO year
        "MTD",  DATE( YEAR(S) - 1, MONTH(S), 1 ),
        "QTD",  DATE( YEAR(S) - 1, MONTH(S), 1 ),
        "YTD",  DATE( YEAR(A) - 1, 1, 1 ),
        DATE( YEAR(S) - 1, MONTH(S), 1 )
    )

-- (9) End of the PY period: same FULL-vs-PARTIAL rule as _Prev End.
_PY End =
VAR P       = SELECTEDVALUE('Period'[Period], "MTD")
VAR S       = [_Period Start]
VAR E       = [_Period End]
VAR IsFull  = ( E = [_Bucket End] )
VAR Offset  = INT( E - S )
VAR PYStart = [_PY Start]
VAR PYBucketEnd =
    SWITCH( P,
        "YDAY", PYStart,
        "WTD",  PYStart + 6,
        "MTD",  EOMONTH( PYStart, 0 ),
        "QTD",  EOMONTH( PYStart, 2 ),
        "YTD",  DATE( YEAR(PYStart), 12, 31 ),
        EOMONTH( PYStart, 0 )
    )
RETURN
    IF( IsFull, PYBucketEnd, MIN( PYStart + Offset, PYBucketEnd ) )
```

---

## STEP 2 — Value measures (6) — Users & Sessions

Each applies the boundary dates over the right base measure. **Important:** the boundary measures
must be captured into `VAR`s first, then the VARs used in the filter. You **cannot** put a measure
reference directly inside a `CALCULATE` boolean filter (`'Date Table'[Date] >= [_Period Start]`) —
PBIRS throws *"A function 'PLACEHOLDER' has been used in a True/False expression..."*. Capturing to
a VAR also evaluates the date in the current (slicer) context **before** `REMOVEFILTERS` clears it,
so the anchor stays correct.

```dax
-- ===== USERS =====
Users (Current) =
VAR S = [_Period Start]
VAR E = [_Period End]
RETURN
CALCULATE(
    [_Base User Count],
    REMOVEFILTERS('Date Table'),
    'Date Table'[Date] >= S,
    'Date Table'[Date] <= E
)

Users (Prev LfL) =
VAR S = [_Prev Start]
VAR E = [_Prev End]
RETURN
CALCULATE(
    [_Base User Count],
    REMOVEFILTERS('Date Table'),
    'Date Table'[Date] >= S,
    'Date Table'[Date] <= E
)

Users (PY LfL) =
VAR S = [_PY Start]
VAR E = [_PY End]
RETURN
CALCULATE(
    [_Base User Count],
    REMOVEFILTERS('Date Table'),
    'Date Table'[Date] >= S,
    'Date Table'[Date] <= E
)

-- ===== SESSIONS =====
Sessions (Current) =
VAR S = [_Period Start]
VAR E = [_Period End]
RETURN
CALCULATE(
    [_Base Session Count],
    REMOVEFILTERS('Date Table'),
    'Date Table'[Date] >= S,
    'Date Table'[Date] <= E
)

Sessions (Prev LfL) =
VAR S = [_Prev Start]
VAR E = [_Prev End]
RETURN
CALCULATE(
    [_Base Session Count],
    REMOVEFILTERS('Date Table'),
    'Date Table'[Date] >= S,
    'Date Table'[Date] <= E
)

Sessions (PY LfL) =
VAR S = [_PY Start]
VAR E = [_PY End]
RETURN
CALCULATE(
    [_Base Session Count],
    REMOVEFILTERS('Date Table'),
    'Date Table'[Date] >= S,
    'Date Table'[Date] <= E
)
```

---

## STEP 3 — Percent-change measures (4)

```dax
Users MoM % (LfL)    = DIVIDE( [Users (Current)]    - [Users (Prev LfL)],    [Users (Prev LfL)] )
Users YoY % (LfL)    = DIVIDE( [Users (Current)]    - [Users (PY LfL)],      [Users (PY LfL)] )
Sessions MoM % (LfL) = DIVIDE( [Sessions (Current)] - [Sessions (Prev LfL)], [Sessions (Prev LfL)] )
Sessions YoY % (LfL) = DIVIDE( [Sessions (Current)] - [Sessions (PY LfL)],   [Sessions (PY LfL)] )
```

> `DIVIDE` returns BLANK (not an error) if the baseline is 0 — safe.

---

## STEP 4 — Honest subtitle labels (so the card SHOWS the windows it compares)

This is what kills the "is this comparing 4 days to a month?" confusion — the card literally states
the ranges.

```dax
Label Current Range =
"Current: " & FORMAT( [_Period Start], "dd mmm yyyy" ) &
    IF( [_Period Start] = [_Period End], "", " – " & FORMAT( [_Period End], "dd mmm yyyy" ) )

Label Prev Range =
"vs " & FORMAT( [_Prev Start], "dd mmm" ) &
    IF( [_Prev Start] = [_Prev End], "", " – " & FORMAT( [_Prev End], "dd mmm yyyy" ) )

Label PY Range =
"vs " & FORMAT( [_PY Start], "dd mmm" ) &
    IF( [_PY Start] = [_PY End], "", " – " & FORMAT( [_PY End], "dd mmm yyyy" ) )
```

---

## STEP 5 — Arrow + auto color for the % cards (all 8, fully written out)

```dax
-- ===== USERS MoM =====
Users MoM Arrow (LfL) =
VAR v = [Users MoM % (LfL)]
RETURN SWITCH( TRUE(), v > 0, "▲ " & FORMAT(v,"0.0%"), v < 0, "▼ " & FORMAT(v,"0.0%"), "▬ 0.0%" )

Users MoM Color (LfL) =
VAR v = [Users MoM % (LfL)]
RETURN SWITCH( TRUE(), v > 0, "#1E7D32", v < 0, "#C62828", "#9E9E9E" )

-- ===== USERS YoY =====
Users YoY Arrow (LfL) =
VAR v = [Users YoY % (LfL)]
RETURN SWITCH( TRUE(), v > 0, "▲ " & FORMAT(v,"0.0%"), v < 0, "▼ " & FORMAT(v,"0.0%"), "▬ 0.0%" )

Users YoY Color (LfL) =
VAR v = [Users YoY % (LfL)]
RETURN SWITCH( TRUE(), v > 0, "#1E7D32", v < 0, "#C62828", "#9E9E9E" )

-- ===== SESSIONS MoM =====
Sessions MoM Arrow (LfL) =
VAR v = [Sessions MoM % (LfL)]
RETURN SWITCH( TRUE(), v > 0, "▲ " & FORMAT(v,"0.0%"), v < 0, "▼ " & FORMAT(v,"0.0%"), "▬ 0.0%" )

Sessions MoM Color (LfL) =
VAR v = [Sessions MoM % (LfL)]
RETURN SWITCH( TRUE(), v > 0, "#1E7D32", v < 0, "#C62828", "#9E9E9E" )

-- ===== SESSIONS YoY =====
Sessions YoY Arrow (LfL) =
VAR v = [Sessions YoY % (LfL)]
RETURN SWITCH( TRUE(), v > 0, "▲ " & FORMAT(v,"0.0%"), v < 0, "▼ " & FORMAT(v,"0.0%"), "▬ 0.0%" )

Sessions YoY Color (LfL) =
VAR v = [Sessions YoY % (LfL)]
RETURN SWITCH( TRUE(), v > 0, "#1E7D32", v < 0, "#C62828", "#9E9E9E" )
```
For each arrow card: use the `... Arrow (LfL)` measure as the card field, then paint-roller →
**Visual → Callout value → Color → fx → Field value →** the matching `... Color (LfL)` measure.

---

## STEP 6 — Wire it into your existing cards

Repoint your left-panel cards to the new measures:

| Card on your page | New measure |
|---|---|
| Session Count (big number) | `Sessions (Current)` |
| Unique User (big number) | `Users (Current)` |
| Session Count change % ▼ | `Sessions MoM Arrow (LfL)` (+ color) |
| Unique User change % ▼ | `Users MoM Arrow (LfL)` (+ color) |
| Session Count "Previous Month" value | `Sessions (Prev LfL)` |
| Unique User "Previous Month" value | `Users (Prev LfL)` |
| vs PY Session Count | `Sessions (PY LfL)` |
| vs PY Unique User | `Users (PY LfL)` |
| Card subtitles (date ranges) | `Label Current Range` / `Label Prev Range` / `Label PY Range` |

After wiring, on **4 Jun with MTD selected** your Session card compares **1–4 Jun vs 1–4 May** — the
scary -86% becomes the real number, and the subtitle proves what it's comparing.

---

## STEP 7 — Dynamic DoD / WoW / MoM / QoQ / YoY change (+ label + color)

The "Current vs Previous" comparison already adapts to the toggle (`_Prev Start/_Prev End` shift by
the right unit), so **one** change measure becomes DoD when YDAY, WoW when WTD, MoM when MTD, QoQ
when QTD, YoY when YTD. These supersede the fixed-name `... MoM % (LfL)` measures from Step 3 — use
these instead (you can delete the MoM-named ones to avoid confusion). The `vs LY` pair below is the
separate "same period one year ago" comparison (your card's "vs PY" section).

```dax
-- (a) Dynamic acronym that names the current comparison
Change Label =
SWITCH( SELECTEDVALUE('Period'[Period], "MTD"),
    "YDAY", "DoD",
    "WTD",  "WoW",
    "MTD",  "MoM",
    "QTD",  "QoQ",
    "YTD",  "YoY",
    "MoM"
)

-- (b) Period-over-period change % (becomes DoD/WoW/MoM/QoQ/YoY automatically)
Users Change %    = DIVIDE( [Users (Current)]    - [Users (Prev LfL)],    [Users (Prev LfL)] )
Sessions Change % = DIVIDE( [Sessions (Current)] - [Sessions (Prev LfL)], [Sessions (Prev LfL)] )

-- (c) Display text: e.g. "MoM ▲ 4.2%" / "QoQ ▼ 3.1%" — label + arrow + value in one card
Users Change Display =
VAR v = [Users Change %]
RETURN
    [Change Label] & "  " &
    SWITCH( TRUE(), v > 0, "▲ ", v < 0, "▼ ", "▬ " ) &
    FORMAT( v, "0.0%" )

Sessions Change Display =
VAR v = [Sessions Change %]
RETURN
    [Change Label] & "  " &
    SWITCH( TRUE(), v > 0, "▲ ", v < 0, "▼ ", "▬ " ) &
    FORMAT( v, "0.0%" )

-- (d) Auto color (green up / red down / grey flat) for the change cards
Users Change Color =
VAR v = [Users Change %]
RETURN SWITCH( TRUE(), v > 0, "#1E7D32", v < 0, "#C62828", "#9E9E9E" )

Sessions Change Color =
VAR v = [Sessions Change %]
RETURN SWITCH( TRUE(), v > 0, "#1E7D32", v < 0, "#C62828", "#9E9E9E" )

-- (e) Same-period-LAST-YEAR comparison ("vs LY") + its color
Users vs LY %    = DIVIDE( [Users (Current)]    - [Users (PY LfL)],    [Users (PY LfL)] )
Sessions vs LY % = DIVIDE( [Sessions (Current)] - [Sessions (PY LfL)], [Sessions (PY LfL)] )

Users vs LY Display =
VAR v = [Users vs LY %]
RETURN "vs LY  " & SWITCH( TRUE(), v > 0, "▲ ", v < 0, "▼ ", "▬ " ) & FORMAT( v, "0.0%" )

Sessions vs LY Display =
VAR v = [Sessions vs LY %]
RETURN "vs LY  " & SWITCH( TRUE(), v > 0, "▲ ", v < 0, "▼ ", "▬ " ) & FORMAT( v, "0.0%" )

Users vs LY Color =
VAR v = [Users vs LY %]
RETURN SWITCH( TRUE(), v > 0, "#1E7D32", v < 0, "#C62828", "#9E9E9E" )

Sessions vs LY Color =
VAR v = [Sessions vs LY %]
RETURN SWITCH( TRUE(), v > 0, "#1E7D32", v < 0, "#C62828", "#9E9E9E" )
```

**Wire it on the card:**
- Put `Users Change Display` (or `Sessions Change Display`) as the card field — it already shows
  "MoM ▲ 4.2%" and renames itself to DoD/WoW/QoQ/YoY when you change the toggle.
- Color it: paint-roller → **Visual → Callout value → Color → fx → Field value → `Users Change
  Color`**.
- For a separate "vs LY" tile, use `Users vs LY Display` + `Users vs LY Color` the same way.
- Prefer label and value in **separate** card elements? Use `Change Label` as a text/card and
  `Users Change %` (formatted as %) as the value, both colored by `Users Change Color`.

> Note on **YoY vs vs-LY**: when **YTD** is selected, the period-over-period change *is* the
> year-over-year change, so `Users Change %` (labelled "YoY") and `Users vs LY %` return the **same**
> number. That is expected — for the other toggles they differ (e.g. MTD: "MoM" = vs last month,
> "vs LY" = vs the same month last year).

---

## Assumptions & notes (read these)
- **The selected date only picks the bucket; a live period runs to today.** Decided by
  `_Latest Data Date` (newest date in the model, ignoring slicers):
  - Latest data = 7 Jun. Pick **26 Apr** with **QTD** → it's the live Q2 → **1 Apr–7 Jun** (April +
    May + June-to-date) vs **1 Jan–~7 Mar** (same point of Q1) vs **1 Apr–7 Jun 2025**. The day 26
    Apr does not cap the window.
  - Pick **any date in 2026** with **YTD** → **1 Jan–7 Jun** vs prior-year-to-same-point.
  - Pick **any date in a finished period** (e.g. 25 May with MTD, or a Q1 date with QTD) → **full**
    May / full Q1, etc. The exact day does not matter.
  - **YDAY** is the exception: it always shows the single day you selected.
- **Full periods compare to the WHOLE previous/PY bucket** — not "the same number of days in". This
  is the fix for the 28-vs-31 bug: a finished **28-day February compares to the full 31-day
  January** (and full Feb last year), a finished **90-day quarter compares to the full 92-day prior
  quarter**, and **leap years** compare full-to-full. Offset alignment is used **only** while a
  period is still in progress (partial).
- **This depends on `_Latest Data Date` being correct.** It reads `MAX('SESSION USAGE'[Date])` with
  all filters removed. If your latest rows ever lag (e.g. a slow refresh), "live" shifts with the
  data, which is what you want.
- **Anchor = `MAX('Date Table'[Date])`** from the date slicer. Make sure the **Year tile slicer
  (2023/2024/2025) does NOT filter these KPI cards**, or it will move the anchor. Fix via
  **Format → Edit interactions** → set the Year slicer to **None** on the KPI cards (keep it on the
  month line charts only).
- **YDAY = the latest selected day** (the anchor). If you want *literally yesterday*, change
  `_Period End` to `[_Anchor Date] - 1` and `_Period Start`'s `"YDAY"` branch to `A - 1`.
- **WTD is ISO-week based.** Weeks start **Monday**; "same week last year" maps to the **same ISO
  week number** in the prior ISO year (not just 364 days back), so it stays correct across 52/53-week
  year boundaries. Week 53 clamps to 52 when the prior year has no week 53. The ISO week number is
  computed **manually** with `WEEKDAY` + `DATE` — this PBIRS build does **not** have `ISOWEEKNUM`,
  so never reintroduce that function here.
- **YTD "Previous"** naturally equals **YTD PY** (prior year to-date) — expected; for YTD the
  meaningful comparison is YoY anyway.
- All measures reuse your existing foundations `_Anchor Date`, `_Base User Count`,
  `_Base Session Count` — nothing in your 15+15 suite changes.
