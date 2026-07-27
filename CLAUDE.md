# Anonymizer — build context

Turns the founder's real bookkeeping data into a shareable, anonymized twin.
Full requirements: see `Anonymization_Spec_Draft_9.docx` in this repo. Read it before writing code.

## Settled decisions — do NOT relitigate
- **Language:** Python 3. Read Excel with **openpyxl** (read-only; never rewrite the .xlsx). Read/write CSVs and TSVs with the `csv` module.
- **Input = one folder** (e.g. `PersonalData/`): the master Excel workbook + ~a dozen raw bank CSVs. Everything in the folder is anonymized in ONE run under ONE seed and ONE shared mapping, so all files stay mutually synchronized.
- **Output = per-seed batch:** `output/<seedname>/` (e.g. `output/John.Seed/`) — one tab-delimited `.txt` per Excel data tab, plus one anonymized file per bank CSV. Inputs are read-only.
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
- **Property codes:** fake = 2 digits + 3 letters derived from the property's fake street (matches the real convention, e.g. `G42VW` = 42 Van Winkle). Must be **unique** across properties (collision-redraw) so the key never merges two properties.
- **Addresses:** street-and-number → fake street; city → replaced via the `Towns` list; state and ZIP left unchanged.
- **5+ digit runs** in descriptions/notes → blanket-randomized. Known-value mappings win over the blanket scrub.
- **Bijection:** distinct reals → distinct fakes (collision-redraw), so the `LLCs` key joins (address ↔ prop ↔ account ↔ LLC) stay valid.
- **Institutions** (e.g. Chase) left as-is.

## The key tab
`LLCs` (Address | LLC | Prop | bank1–3) is the relationship key. It is anonymized in the SAME run with the SAME mapping, so the anonymized key correctly describes the anonymized data.

## Outputs
- `output/<seedname>/` — anonymized `.txt` per tab + anonymized bank CSVs
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
