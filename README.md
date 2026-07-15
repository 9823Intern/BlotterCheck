# BlotterCheck

Beginning-of-day trade safety checks and end-of-day reconciliation in one
project.

## Beginning-of-day safety check

```bash
python bod.py
```

Open http://127.0.0.1:5051. Drop or paste the blotter to check trade direction
against the prior market day's Goldman positions, crossing-trade balances, and
fund/broker/account combinations.

The positions report is loaded from the dated Dropbox path configured in
`new_main.py`.

## End-of-day reconciliation

Reconciles four trade files with a four-step workflow. Only EMSX rows with
Status `Filled` (column D) are loaded. Trades are aggregated per
`(ticker, side)` — total quantity and quantity-weighted average price — and
every comparison is bidirectional: a bucket present on one side but not the
other is a discrepancy.

1. **EMSX vs Goldman tradefile** — EMSX filtered to brokers `GSPT`, `PREX`,
   `GSOP` (column G). Every Goldman EMSX trade must be in the tradefile and
   vice versa.
2. **EMSX vs PCMBlotter** — EMSX excluding broker `TDAI`. Every remaining EMSX
   trade must be in PCM and vice versa.
3. **PCM vs Goldman tradefile** — PCM filtered to brokers `GSPT`, `PREX`,
   `GSOP` (its broker column). The two Goldman views must match.
4. **Blotter EOD adjudication** — discrepancies from steps 1-3 that are pure
   side-classification differences (net signed quantity per ticker still
   matches) are checked against the Blotter EOD. When the blotter's per-ticker
   profile corroborates one source, the finding is downgraded to `RESOLVED`.
   Missing or extra trades that change net quantity always remain errors, even
   if the blotter confirms the trade exists elsewhere.

## Inputs

| Source | Format | Keyed on |
|---|---|---|
| EMSX grid export | `.xlsx` (named columns: Side, Security, SEDOL, FillQty, AvgPx) | ticker + SEDOL |
| Vantage Blotter EOD | `.xlsx`/`.csv`, S/P/C rows, `bl/sl/ss/cs` type codes | ticker |
| PCMBlotter | `.csv`, headered, accounting-style negatives `(21)` | ticker |
| Goldman tradefile | `tradefile.nine82.*.csv`, pipe-delimited ORDER/ALLOCAT | SEDOL |

### Web frontend (paste or drop files)

```bash
python eod.py
```

Open http://127.0.0.1:5052. Each of the four sources has its own box: drop the
file onto it, or copy the contents (Excel cells, raw CSV, or the pipe-delimited
tradefile) and paste into the textarea. Include header rows when pasting the
EMSX grid and PCMBlotter. Hit "Run reconciliation" to see every finding,
color-coded by severity and filterable by workflow step.

### CLI

```bash
python blotter_check.py \
    --emsx grid.xlsx \
    --blotter BlotterEOD07.10.26.xlsx \
    --pcm "PCMBlotter 2026-07-10.csv" \
    --tradefile tradefile.nine82.20260710045714.csv
```

Optional: `--out report.csv`, `--qty-tol 0` (shares), `--price-tol 0.01` (dollars).

Exit code is 1 if any ERROR-severity finding exists, 0 otherwise, so it can be
used in an automated pipeline.

## What it checks

Every trade is normalized to `(ticker, side, quantity, price)` with sides
BUY/SELL/SHORT/COVER, then the four-step comparisons above run. Finding types:

- **Missing / extra buckets** — a (ticker, side) present on one side of a
  comparison but not the other. One underlying discrepancy is reported once:
  the net-quantity check per ticker only fires when no bucket-level error
  already covers it.
- **Quantity mismatches** — per (ticker, side) aggregate totals differ beyond
  the share tolerance.
- **Price mismatches** — quantity-weighted average price beyond tolerance
  (warning only).
- **EMSX partial fills** — Filled rows where FillQty < Qty (warning).
- **Blotter unfilled orders** — blotter rows with zero fills (warning only).
- **PCM sign consistency** — quantity sign vs its own transaction type.
- **Tradefile internal integrity** — allocations sum to order quantity, trailer
  row count matches, allocation accounts are recognized.
- **Internal crosses** — fund-to-fund crosses (broker `CROS`) never route
  through EMSX, so they are excluded from every comparison. Instead they are
  checked to net to zero per ticker, and cross volume is compared between the
  blotter and PCM.

## Comparison scope

Goldman comparisons (steps 1 and 3) are limited to brokers `GSPT,GSOP,PREX`
(override with `--goldman-brokers`). The PCM comparison (step 2) covers every
EMSX broker except `TDAI` — TDAI executions live in EMSX only and are expected
nowhere else. The Blotter EOD is used only for step 4 adjudication, scoped the
same way as the comparison that produced each discrepancy.

## SEDOL -> ticker resolution (tradefile)

1. The EMSX file itself (it carries both ticker and SEDOL) — no network needed.
2. Backtester security master (`~/GitHub/Backtester`, local parquet).
3. `xbbg` / Bloomberg terminal, if installed and running.

Unresolvable SEDOLs are flagged as errors.

## Output

A findings CSV (`blottercheck_report_<timestamp>.csv` by default) with columns:
severity (`ERROR`/`WARN`/`RESOLVED`), step (`Step 1`-`Step 3`, `Crosses`,
`Load`), source (the comparison, e.g. `EMSX vs Tradefile`), check, ticker,
side, left_value, right_value, detail. The same table is printed to the
console.
