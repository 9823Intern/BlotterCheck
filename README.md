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

Reconciles four trade files against each other and flags any differences. The
EMSX export is treated as authoritative.

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
EMSX grid and PCMBlotter. Hit "Run reconciliation" to see every diff, color-coded
by severity and filterable by source.

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
BUY/SELL/SHORT/COVER, then each source is compared to EMSX:

- **Missing / extra trades** — a (ticker, side) present in one file but not the other.
- **Quantity mismatches** — per (ticker, side) and per net signed quantity per ticker.
  Side-level differences that still net out (e.g. a buy-to-open recorded as a
  buy-to-cover) are downgraded to warnings.
- **Price mismatches** — quantity-weighted average price beyond tolerance.
- **EMSX partial fills** — orders where FillQty < Qty.
- **Blotter unfilled orders** — blotter rows with zero fills (warning only).
- **PCM sign consistency** — quantity sign vs its own transaction type.
- **Tradefile internal integrity** — allocations sum to order quantity, trailer
  row count matches, allocation accounts are recognized.
- **Internal crosses** — fund-to-fund crosses (broker `CROS`) never route
  through EMSX, so they are excluded from the EMSX comparison. Instead they are
  checked to net to zero per ticker, and cross volume is compared between the
  blotter and PCM.

## Comparison scope

The Goldman tradefile only carries Goldman-custodied executions, so it is
compared against EMSX filtered to brokers `GSPT,GSCO,PREX` (override with
`--goldman-brokers`). CFR/Cantor executions are checked via the blotter and PCM
comparisons, which cover the full broker universe.

## SEDOL -> ticker resolution (tradefile)

1. The EMSX file itself (it carries both ticker and SEDOL) — no network needed.
2. Backtester security master (`~/GitHub/Backtester`, local parquet).
3. `xbbg` / Bloomberg terminal, if installed and running.

Unresolvable SEDOLs are flagged as errors.

## Output

A findings CSV (`blottercheck_report_<timestamp>.csv` by default) with columns:
severity, source, check, ticker, side, emsx_value, source_value, detail.
The same table is printed to the console.
