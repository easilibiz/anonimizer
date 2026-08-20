# Anonymizer

Turns your real bookkeeping data into a shareable, anonymized copy. It reads
your master Excel workbook **and** bank-download CSVs, replaces the sensitive
text (names, towns, LLC/entity names, account names/numbers, addresses, long
codes, letter+digit reference codes) with consistent fakes, keeps every amount
and date exactly as-is, and writes files that open cleanly in Excel.

Everything is driven by a **seed**: the same seed always produces the same
fakes, so the workbook tabs and every bank CSV line up with each other. Change
the seed to make a different anonymized copy.

There are two generations of this tool in this folder:

- **Draft 9 (current, recommended)** — point at your Excel workbook + a
  folder of bank CSVs, get back one clean anonymized `.xlsx` + anonymized bank
  CSVs, with a built-in PASS/FAIL safety report. See **"Draft 9 — folder-based
  run"** below. Use `RUN_ANONYMIZER9.bat`.
- **Legacy (tab-delimited TSV workflow)** — the original tool, built around
  manually-exported `Transactions.txt` / `RuleN.txt`. Still here for
  compatibility; documented in the rest of this file. Use `RUN_ANONYMIZER.bat`.

---

## Draft 9 — folder-based run (recommended)

**One button, one folder.** Put your master workbook (`.xlsx`/`.xlsm`) and all
your bank statement CSVs in one folder, point `run9.config.txt` at that folder,
and double-click `RUN_ANONYMIZER9.bat`. You never list individual files.

### Step by step

1. **Put everything in one input folder** (e.g. `PersonalData\`): the master
   workbook, plus every bank statement CSV you want anonymized alongside it.
2. **Copy `run9.config.example.txt` to `run9.config.txt`** (if you don't
   already have one) and set:
   ```
   input_dir = PersonalData
   config    = anonimizer config.csv
   seed      = seed1
   output_dir = output
   ```
3. **Double-click `RUN_ANONYMIZER9.bat`.**
4. **Read the report.** All three checks should say **PASS**.
5. **Collect the results** from `output\<seed>\`.

### What "take everything from the folder" means

With `input_dir` set, the tool scans that folder (top level only) and
automatically picks up:
- **the workbook** — the one `.xlsx`/`.xlsm` file in the folder,
- **every bank CSV** — every `.csv` file in the folder.

It automatically **skips**: the `config` CSV itself, Excel lock files
(`~$...`), and any file already produced by a previous run (`*_anon.*`) — so
re-running on the same folder doesn't try to anonymize its own output.

If the folder has **no** workbook, or **more than one**, the run stops and
tells you — add an explicit `workbook = <file>` line to `run9.config.txt` to
pick one instead of relying on discovery.

> Prefer to list files by hand instead of pointing at a folder? Comment out
> `input_dir` and use `workbook =` + `bank =` lines — see
> `run9.config.example.txt` for both modes side by side.

### Draft 9 outputs

In `output\<seed>\`:
- **One anonymized workbook**: `<original name>_anon.xlsx` — a plain workbook
  with exactly the same tabs as the source (no formatting, no styling).
  Amounts and dates are kept as real Excel numbers/dates, only text is faked.
- **One anonymized CSV per bank statement**: `<bankfile>_anon.csv` — same
  columns, only the description text replaced.

The mapping (re-identification key) is written to
`mappings\<seed>\mapping.json` — **keep it private**, same rule as the legacy
tool (see "The mapping — keep it private" below).

### Amount factor (optional)

Set `amount_factor` in `run9.config.txt` to uniformly scale every dollar
amount — `Transactions`/`RulesN`.Amount, and each bank CSV's Amount, Running
Bal., and preamble summary lines (Beginning/Ending balance, Total
credits/debits) — by one constant factor. Must be between `0.80` and `1.20`;
blank or `1.0` = off (the default — amounts pass through exactly). A constant
factor preserves every total and running-balance calculation, so the books
still tie to each other — just no longer to the original real cent amounts,
which is the point.

### Draft 9 seeds

Same rules as the legacy tool: same seed → identical output; blank seed → a
fresh random one each run (printed as `seed: xxxxxx (auto-generated)`); a new
seed name → a brand-new, independent set of fakes.

### Draft 9 report

| Check | Meaning |
|---|---|
| Blacklist scan (whole-word) | None of your `Blacklist` entries (from `anonimizer config.csv`) survived anywhere in the output, workbook or bank CSVs. |
| Determinism | Re-running the same seed produces byte-identical mapping. |
| Bijection | Every distinct real value maps to its own distinct fake (no accidental collisions). |

If any check **FAIL**s, the reason and the offending value are printed above
the report — do not share the output until it's clean.

### One entity, one fake name

An LLC/owner that appears in several forms — bare (`idanamir`), checking
(`idanamir CK`), the LLC itself (`idanamir LLC`) — always maps to the **same**
fake root, e.g. `Denise` / `Denise CK` / `Denise LLC`. This keeps joins across
tabs intact after anonymization.

---

## Legacy tool — the two buttons (batch files)

You run the tool by **double-clicking a `.bat` file**. There are two:

| Double-click… | What it does | When to use |
|---|---|---|
| **`RUN_ANONYMIZER.bat`** | Anonymizes the **transactions + rules + all bank CSVs** together, rebuilding the mapping from the current config. | Your normal run. Always use this after changing anything in the config. |
| **`RUN_BANKS_ONLY.bat`** | Anonymizes **only the bank CSVs**, reusing the mapping already built. Faster; doesn't touch the transactions. | When you've just downloaded more bank CSVs and nothing else changed. |

Each opens a window, does the work, prints a **PASS/FAIL report**, and waits so
you can read it. Close the window when done.

> If double-clicking a `.bat` seems to do nothing, right-click it → **Run as
> administrator**, or open a terminal in this folder and run the same command
> the `.bat` contains. The `.bat` uses the project's own Python (`.venv\`), so
> you don't need Python set up separately.

---

## Normal run — step by step

1. **Export your data to tab-delimited text.** In Excel, for each sheet:
   *File → Save As → "Text (Tab delimited) (\*.txt)"*, saved into this folder —
   e.g. `Transactions.txt` and `RuleN.txt`.
2. **Download your bank CSVs** into this folder (or a subfolder like
   `publicData\`).
3. **List the bank files** in `run.config.txt` (one `bank =` line each — see below).
4. **Pick a seed** in `run.config.txt` (e.g. `seed1`).
5. **Double-click `RUN_ANONYMIZER.bat`.**
6. **Read the report.** All checks should say **PASS**.
7. **Collect the results** from the `output\` folder.

### Adding more bank CSVs later
Drop the new CSVs in, add their `bank =` lines to `run.config.txt`, and
**double-click `RUN_BANKS_ONLY.bat`.** It reuses the same seed's mapping, so the
new files line up with everything you already produced.

### If you change the config (`anonimizer config.csv`)
Adding/removing names, towns, or blacklist entries changes the mapping, so run
the **full** `RUN_ANONYMIZER.bat` (not banks-only). Banks-only reuses the old
mapping and won't pick up new names/towns.

---

## The config files

Just **two** plain-text files control everything. Edit them in Notepad.

### 1. `run.config.txt` — the ONE settings file (seed + everything)
```
seed         = seed1                # any label; blank = a new random one each run
transactions = Transactions.txt     # tab-delimited Transactions export
rules        = RuleN.txt            # tab-delimited rules-sheet export
config       = anonimizer config.csv
output_dir   =                      # blank = the .\output folder

# Bank CSVs (optional) — one 'bank =' line per file:
#     bank = <file> | <account last-four, optional> | <description column>
bank = Personaldata\stmt (3).csv                         |      | Description
bank = Personaldata\Chase7352_Activity_20260725.CSV      | 7352 | Description
```
- **bank lines** — add one `bank =` line per statement. A bare file name is
  looked for in this folder; or give a full path. The account last-four and
  description column are optional (column defaults to `Description`).
- **Both buttons read this same file** — `RUN_BANKS_ONLY.bat` just skips the
  transactions and processes the `bank =` lines.

### 2. `anonimizer config.csv` — what to replace
Three columns: **Names**, **Towns**, **Blacklist**.
- **Names / Towns** — the real values to replace with fakes.
- **Blacklist** — a safety list: these strings must NEVER appear in the output.
  Blacklist entries are *not* replaced on their own — put anything that must be
  scrubbed in **Names** or **Towns** too, or the run will fail the leakage check.
  *(Tip: bank text may spell a town differently, e.g. `Fair Lawn` vs `Fairlawn`
  — add both spellings to Towns.)*

---

## Outputs

In the **`output\`** folder, one anonymized file per input, seed-stamped:
- `Transactions_anon.<seed>.txt`, `RuleN_anon.<seed>.txt`
- `<bankfile>_anon.<seed>.csv` (keeps the bank's own columns)

Share these once the report says PASS.

### The mapping — keep it private
Each run writes/updates `Anon mapping\mapping.<seed>.json`. This is the **key
that maps fakes back to the real values** (the re-identification key). **Never
share it** and don't put it in any repo or upload.

---

## Seeds — one copy or many variants

- **Same seed → identical output** every time (reuse `seed1` to reproduce it).
- **New seed → all-new fakes** (`seed2`, `batchA`, … make different copies of
  the same data).
- **Blank `seed`** → the tool generates a random one and prints it.

Each seed writes its own files and its own `mapping.<seed>.json`, so nothing is
overwritten.

---

## What gets anonymized

| Thing | How |
|---|---|
| Names & towns | Replaced with picked fakes — but only the ones you list in `anonimizer config.csv`. Keep that list comprehensive. |
| Account name + number | Replaced together, keyed on the **last-four digits**, so `xxxx0585`, `0585`, and `585` are the same account and get the same fake. Numbers come out as `xxxx####`. |
| Institution | NOT faked — made **consistent** (each account keeps one real bank, the one the Transactions sheet uses). |
| Bare last-four in descriptions | Same fake last-four as the account, so transfers still reconcile. |
| 5+ digit codes | Randomized (loan/membership/gov numbers). |
| Letter+digit codes | Reference/confirmation codes (e.g. `6P038A5O05B`, `99clacfxo`) scrubbed to same-shape fakes. |
| Amounts, dates, other columns | Never touched. |

---

## How bank CSVs are handled

- The tool finds the real transaction **header row** automatically (skips
  summary/preamble rows like Bank of America's), and drops empty padding rows.
- It anonymizes **only the description column**; every other column (amount,
  date, balance, type…) is passed through unchanged.
- Each bank file is treated as **one account**; long codes new to the bank files
  are added to the mapping so they stay consistent.

---

## The report (what PASS means)

| Check | Meaning |
|---|---|
| Non-target integrity | Amounts / dates / other columns unchanged. |
| Leakage scan | No real name, town, account name/number, or code survived. |
| Hit-count | Account/transfer references still fire (rules stay in sync). |
| Determinism | Same seed reproduces identical output. |

If it says **FAIL — do NOT share**; the reason is printed above the report
(often a config-coverage gap — a real value you haven't added to Names/Towns).

---

## Troubleshooting

- **A real word survived** (e.g. a town spelled differently in bank text): add
  that exact spelling to Names/Towns in `anonimizer config.csv`, then run the
  **full** `RUN_ANONYMIZER.bat`.
- **Banks-only run flags a name you just added to the config:** that's expected
  — run the full `RUN_ANONYMIZER.bat` to rebuild the mapping first.
- **Excel shows odd characters:** open via *Data → From Text/CSV* and choose
  encoding **Windows (1252)**.
- **`.venv` missing:** open a terminal here and run
  `python -m venv .venv` then `.\.venv\Scripts\python.exe -m pip install openpyxl`.

---

## Command line (optional)

You normally don't need it. To run the same thing without the `.bat`:
```powershell
C:\ClaudeCodeOM\anonimizer\anonymize.ps1 fromconfig                       # full run
C:\ClaudeCodeOM\anonimizer\anonymize.ps1 fromconfig --banks-only          # banks only
```
Per-step commands (`roundtrip`, `inspect`, `extract`, `map`, `apply`, `scan`,
`hitcount`, `run`) also exist for debugging; run `anonymize.ps1 <command> -h`.
