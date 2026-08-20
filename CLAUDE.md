# Anonymizer — build context

Turns the founder's real bookkeeping data into a shareable, anonymized twin.
Full requirements: see `Anonymization_Spec_Draft_10.docx` in this repo. Read it before writing code.

## Settled decisions — do NOT relitigate
- **Language:** Python 3. Read Excel with **openpyxl** (read-only; never rewrite the .xlsx). Read/write CSVs and TSVs with the `csv` module.
- **Input = one folder** (e.g. `PersonalData/`): the master Excel workbook + ~a dozen raw bank CSVs. Everything in the folder is anonymized in ONE run under ONE seed and ONE shared mapping, so all files stay mutually synchronized.
- **Output = per-seed batch:** `output/<seedname>/` (e.g. `output/John.Seed/`) — ONE plain `.xlsx` workbook (one sheet per Excel data tab, no styling; a tab-delimited `.txt` per tab was tried first but broke Excel's own parser on some real exports), plus one anonymized file per bank CSV. Inputs are read-only. Amount/date cells are written as native Excel numbers/dates, not text. Output file NAMES are anonymized too (the same known real values — not the blanket digit-run/reference-code scrub — so an ordinary filename with no real identifier in it is left alone), since export filenames sometimes carry the real account number (e.g. `Chase7352_Activity_....CSV`).
- **Data tabs only.** Process the identifier-bearing tabs (`Transactions`, `RulesN`, `accounts`, `LLCs`); skip hidden/`veryHidden` and formula/report tabs. **Clamp each tab to its real used column range** — `RulesN` reports 16,371 columns but has 17. Ignore phantom columns or the run explodes.
- **Only text cells are replaced.** Numeric and date cells are identified by type and NEVER read for replacement or modified. Amounts and dates pass through exactly; books must tie to the cent.
- **Case-insensitive matching is mandatory.** The data mixes forms (`danmir`/`Danmir`, `IdanAmir`/`idanamir`).
- **Determinism:** same seed → byte-identical output; a different seed → a fully independent, non-overlapping batch. Sort distinct reals before assigning fakes; never depend on dict/set order; use one `random.Random(seed)`.
- **No AI.** Fakes are *picked* from bundled dictionaries (first names, surnames, towns, streets), never generated.

## Two identifier sources
1. **Config CSV** (free-text identifiers with no registry) — columns `Names`, `Towns`, `Blacklist`:
   - `Names` (tenants) → replaced with realistic fakes.
   - `Towns` → replaced with realistic fakes.
   - `Blacklist` → **sanity check only**, NOT used for replacement. Owner/entity roots (e.g. `Michael`, `Danmir`). After the run, a case-insensitive, whitespace-trimmed substring scan must find NONE of these in any output → else fail loudly.
2. **Harvested from the workbook** (structural identifiers, read directly — not in config):
   - account names + account numbers (from `accounts` and `Transactions`)
   - LLC/entity names (from `LLCs`, `accounts`) — including compound strings like `Michael Team Ck - xxxx7705 (56CA)` and their 4-char codes
   - property codes (from `LLCs.Prop` and `Transactions.Prop`)
   - addresses (from `LLCs.Address`)

## Replacement rules (one combined pass, no chaining)
- **Names & towns:** matched anywhere in a string, **whole-word**, **case-insensitive**; consistent (same real → same fake everywhere).
- **Account name + Account #:** a bound pair; the fake last-four also replaces the bare last-four as a **whole 4-digit token** inside descriptions/notes. Handles both `xxxx7705` and bare `7705`.
- **Property codes:** fake = up to 3 digits from the property's fake street number, a hyphen, then 3 letters from the fake street name (e.g. `142 Maple Ave` → `142-MAP`; fewer than 3 digits uses what's present). Each real code maps to one such fake, **unique** across properties (collision-redraw), and that fake is the persistent replacement everywhere the real code appears (Transactions, notes, the `LLCs` key).
- **Addresses:** street-and-number → fake street; city → replaced via the `Towns` list; state and ZIP left unchanged.
- **5+ digit runs** in descriptions/notes → blanket-randomized. Known-value mappings win over the blanket scrub.
- **Bijection:** distinct reals → distinct fakes (collision-redraw), so the `LLCs` key joins (address ↔ prop ↔ account ↔ LLC) stay valid.
- **Institutions** (e.g. Chase) left as-is.
- **Output file names:** anonymized with the SAME known real values as above (account numbers/last-four, names, towns, property/entity codes) — but NOT the blanket 5+ digit run scrub or the generic reference-code scramble, since those would wrongly nuke an ordinary filename that has no real identifier in it at all. Export filenames sometimes glue the real account number directly onto other text (e.g. `Chase7352_Activity_20260725.CSV`); the last-four match uses a digit-only boundary (not the usual alnum boundary) so it still catches that.

## The key tab
The relationship-key tab (Address | Prop | LLC | bank1–3) is found BY ITS COLUMNS, not a fixed name — the founder has already renamed it once (`LLCs` → `Properties`); hardcoding the name silently broke harvesting (empty addresses, ZIPs caught by the digit scrub, property codes decoupled from their real address) with nothing in the report to say why. It is anonymized in the SAME run with the SAME mapping, so the anonymized key correctly describes the anonymized data.

## Amount factor (optional)
`amount_factor` (config setting, default 1.0/off, range 0.80–1.20) uniformly scales every dollar amount — `Transactions`/`RulesN`.Amount, and each bank CSV's Amount, Running Bal., and preamble summary lines (Beginning/Ending balance, Total credits/debits) — by one constant, founder-chosen factor. This is an OPT-IN departure from "amounts pass through exactly": a constant multiplier preserves every total/running-balance relationship (the books still tie to each other), just no longer to the original real cent amounts. Values outside 0.80–1.20 are rejected.

## Outputs
- `output/<seedname>/` — ONE anonymized `.xlsx` workbook (one sheet per data tab) + anonymized bank CSVs; output file names are anonymized too
- mapping table, local-only, namespaced per seed (`mappings/<seed>...`), never committed/shared
- pass/fail report

## Checks that must pass
- Blacklist leakage scan (case-insensitive, trimmed) across all output files → no owner/entity/tenant root survives.
- Bijection holds → `LLCs` joins still resolve after anonymization.
- Rule fidelity: `RulesN` rules fire on the identical row set; account/transfer hit-counts hold before vs after.
- Determinism: two runs, same seed, identical output.
- Books reconcile: regenerated per-LLC P&L / balance sheet tie to the original to the cent.

## Build order (one increment at a time, test each on the mini workbook + a tiny config CSV)
1. **Load + no-op round-trip.** Open the workbook read-only with openpyxl; list data tabs (skip hidden/formula); clamp used range; load the config CSV and any bank CSVs; print columns/counts. Confirm nothing is misread. THIS GATES EVERYTHING.
2. Harvest structural identifiers from `accounts`/`LLCs`/`Transactions`; load config Names/Towns/Blacklist. Print the distinct real-value inventory (no rows).
3. Build ONE deterministic, seed-driven, bijective mapping (names, towns, accounts+last-four, LLCs, property codes, addresses). Write the per-seed mapping table.
4. Value-driven, case-insensitive replacement pass across all data tabs + all bank CSVs; 5+ digit scrub; write `output/<seed>/`.
5. Blacklist leakage scan (case-insensitive) across all outputs.
6. Hit-count / rule-fidelity check (stub if rules not wired yet).
7. One-command runner: input folder + seed → full batch + pass/fail report. Loop over multiple seeds for multiple customers.

## Test discipline
Build and test against the mini workbook (`Transactions`, `RulesN`, `accounts`, `LLCs`) plus a tiny config CSV FIRST. Only run on the full Excel + real bank CSVs once every increment passes on the mini.
