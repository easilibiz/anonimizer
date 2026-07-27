# Anonymizer

Turns your real bookkeeping data into a shareable, anonymized copy. It reads
tab-delimited exports (Transactions + rules) **and** bank-download CSVs,
replaces the sensitive text (names, towns, account names/numbers, long codes,
letter+digit reference codes) with consistent fakes, keeps every amount and date
exactly as-is, and writes files that open cleanly in Excel.

Everything is driven by a **seed**: the same seed always produces the same
fakes, so the transactions, the rules, and every bank CSV line up with each
other. Change the seed to make a different anonymized copy.

---

## The two buttons (batch files)

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
