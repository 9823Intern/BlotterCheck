"""Unit tests for the four-step reconciliation workflow (synthetic DataFrames)."""

import sys

import pandas as pd

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import blotter_check as bc

# Real Bloomberg tickers / SEDOLs so the tradefile resolves via EMSX alone.
SEDOLS = {"AAPL": "2046251", "MSFT": "2588173", "NVDA": "2379504", "ABNB": "BMGYYH4"}


def emsx_df(rows):
    return pd.DataFrame(
        [{"ticker": t, "sedol": SEDOLS[t], "side": s, "qty": q, "price": p, "broker": b}
         for (t, s, q, p, b) in rows],
        columns=["ticker", "sedol", "side", "qty", "price", "broker"],
    )


def trade_df(rows):
    return pd.DataFrame(
        [{"sedol": SEDOLS[t], "side": s, "qty": q, "price": p, "broker": b,
          "order_id": f"O{i}", "trade_date": "20260715"}
         for i, (t, s, q, p, b) in enumerate(rows)],
        columns=["sedol", "side", "qty", "price", "broker", "order_id", "trade_date"],
    )


def src_df(rows):
    return pd.DataFrame(
        [{"ticker": t, "side": s, "qty": q, "price": p, "broker": b, "fund": "F1"}
         for (t, s, q, p, b) in rows],
        columns=["ticker", "side", "qty", "price", "broker", "fund"],
    )


def run(emsx, blotter, pcm, tradefile, **kw):
    findings = bc.Findings()
    return bc.reconcile(emsx, blotter, pcm, tradefile, findings, **kw)


def errors(report):
    return report[report["severity"] == "ERROR"]


def resolved(report):
    return report[report["severity"] == "RESOLVED"]


# --- EMSX loads only Filled rows --------------------------------------------
raw = pd.DataFrame([
    ["Security", "SEDOL", "Side", "Status", "FillQty", "AvgPx", "Def Brkr Code", "Qty"],
    ["AAPL", "2046251", "Buy", "Filled", 10, 100, "GSPT", 10],
    ["MSFT", "2588173", "Sell", "Working", 5, 200, "GSPT", 5],
    ["NVDA", "2379504", "Buy", " filled ", 7, 150, "GSOP", 7],
])
f = bc.Findings()
out = bc._emsx_from_table(raw, f)
assert out["ticker"].tolist() == ["AAPL", "NVDA"], out.to_dict("records")
assert not f.rows, f.rows
print("filled-only EMSX ingestion passed")

# --- Clean run: everything matches, including aggregate wavg price ----------
# AAPL is split across two EMSX fills (60@100 + 40@110 -> wavg 104) that must
# aggregate to match a single 100@104 row in every other source.
emsx = emsx_df([
    ("AAPL", "BUY", 60, 100.0, "GSPT"),
    ("AAPL", "BUY", 40, 110.0, "GSPT"),
    ("NVDA", "SELL", 25, 500.0, "PREX"),
    ("MSFT", "SELL", 50, 300.0, "TDAI"),   # TDAI: EMSX only, everywhere excluded
])
tradefile = trade_df([
    ("AAPL", "BUY", 100, 104.0, "GSPT"),
    ("NVDA", "SELL", 25, 500.0, "PREX"),
])
pcm = src_df([
    ("AAPL", "BUY", 100, 104.0, "GSPT"),
    ("NVDA", "SELL", 25, 500.0, "PREX"),
])
blotter = src_df([
    ("AAPL", "BUY", 100, 104.0, "GSPT"),
    ("NVDA", "SELL", 25, 500.0, "PREX"),
])
report = run(emsx, blotter, pcm, tradefile)
assert errors(report).empty, errors(report).to_dict("records")
assert (report["severity"] != "WARN").all(), report.to_dict("records")
print("clean four-step run passed")

# --- Step 1: Goldman scope, bidirectional ------------------------------------
# CFR trade in EMSX is out of tradefile scope (no finding). A Goldman EMSX trade
# missing from the tradefile and an extra tradefile trade are both errors.
emsx = emsx_df([
    ("AAPL", "BUY", 100, 104.0, "GSPT"),
    ("MSFT", "SELL", 50, 300.0, "CFR"),
])
tradefile = trade_df([("NVDA", "SELL", 25, 500.0, "PREX")])
pcm = src_df([
    ("AAPL", "BUY", 100, 104.0, "GSPT"),
    ("MSFT", "SELL", 50, 300.0, "CFR"),
    ("NVDA", "SELL", 25, 500.0, "PREX"),
])
blotter = src_df([])
# NVDA sedol must still resolve: it's absent from this EMSX, so put it in PCM
# scope only via EMSX... instead resolve by including an NVDA EMSX row at a
# non-Goldman broker so step 1 exclusion still applies.
emsx = pd.concat([emsx, emsx_df([("NVDA", "SELL", 25, 500.0, "CFR")])], ignore_index=True)
report = run(emsx, blotter, pcm, tradefile)
step1_errors = errors(report)[errors(report)["step"] == "Step 1"]
assert set(step1_errors["check"]) == {"missing_bucket", "extra_bucket"}, step1_errors.to_dict("records")
assert set(step1_errors["ticker"]) == {"AAPL", "NVDA"}, step1_errors.to_dict("records")
assert "MSFT" not in set(step1_errors["ticker"])
print("step 1 Goldman scope and bidirectional checks passed")

# --- Step 2: TDAI excluded from the EMSX side --------------------------------
emsx = emsx_df([
    ("MSFT", "SELL", 50, 300.0, "TDAI"),   # not expected in PCM
    ("AAPL", "BUY", 100, 104.0, "CFR"),    # expected in PCM
])
tradefile = trade_df([])
pcm = src_df([])                            # AAPL missing -> error; MSFT absent -> fine
blotter = src_df([])
report = run(emsx, blotter, pcm, tradefile)
step2_errors = errors(report)[errors(report)["step"] == "Step 2"]
assert set(step2_errors["ticker"]) == {"AAPL"}, step2_errors.to_dict("records")
print("step 2 TDAI exclusion passed")

# --- Step 3: PCM Goldman scope vs tradefile ----------------------------------
emsx = emsx_df([
    ("AAPL", "BUY", 100, 104.0, "GSPT"),
    ("MSFT", "SELL", 50, 300.0, "CFR"),
])
tradefile = trade_df([("AAPL", "BUY", 100, 104.0, "GSPT")])
pcm = src_df([
    ("AAPL", "BUY", 100, 104.0, "GSPT"),
    ("MSFT", "SELL", 50, 300.0, "CFR"),    # non-Goldman: out of step 3 scope
    ("NVDA", "SELL", 25, 500.0, "GSOP"),   # Goldman PCM trade not in tradefile
])
blotter = src_df([])
report = run(emsx, blotter, pcm, tradefile)
step3_errors = errors(report)[errors(report)["step"] == "Step 3"]
assert set(step3_errors["ticker"]) == {"NVDA"}, step3_errors.to_dict("records")
# NVDA (GSOP) also legitimately fails step 2 since it's not in EMSX ex-TDAI.
print("step 3 PCM Goldman scope passed")

# --- Step 4: blotter resolves a side-classification difference ---------------
# EMSX and tradefile say COVER 82; PCM classified it as BUY 82 (net matches).
# Blotter shows COVER 82, corroborating EMSX/tradefile -> RESOLVED, no errors.
emsx = emsx_df([("AAPL", "COVER", 82, 50.0, "GSPT")])
tradefile = trade_df([("AAPL", "COVER", 82, 50.0, "GSPT")])
pcm = src_df([("AAPL", "BUY", 82, 50.0, "GSPT")])
blotter = src_df([("AAPL", "COVER", 82, 50.0, "GSPT")])
report = run(emsx, blotter, pcm, tradefile)
assert errors(report).empty, errors(report).to_dict("records")
res = resolved(report)
assert set(res["step"]) == {"Step 2", "Step 3"}, res.to_dict("records")
assert all("corroborating" in d for d in res["detail"]), res.to_dict("records")
print("step 4 blotter resolution passed")

# --- Step 4: ambiguous side difference stays an error ------------------------
# Same disagreement, but the blotter matches neither classification.
blotter = src_df([("AAPL", "SELL", 82, 50.0, "GSPT")])
report = run(emsx, blotter, pcm, tradefile)
amb = errors(report)
assert not amb.empty and all("does not corroborate" in d for d in amb["detail"]), amb.to_dict("records")
print("step 4 ambiguous discrepancy stays error passed")

# --- Step 4: a required missing trade is never suppressed --------------------
# EMSX Goldman SELL absent from the tradefile changes net qty; the blotter
# corroborating EMSX must NOT hide that the tradefile is missing the trade.
emsx = emsx_df([("ABNB", "SELL", 6, 120.0, "GSPT")])
tradefile = trade_df([])
pcm = src_df([("ABNB", "SELL", 6, 120.0, "GSPT")])
blotter = src_df([("ABNB", "SELL", 6, 120.0, "GSPT")])
report = run(emsx, blotter, pcm, tradefile)
missing = errors(report)
assert set(missing["step"]) == {"Step 1", "Step 3"}, missing.to_dict("records")
assert set(missing["check"]) == {"missing_bucket"}, missing.to_dict("records")
print("required missing trade stays error passed")

# --- Duplicate suppression: one discrepancy, one finding per step ------------
emsx = emsx_df([("AAPL", "BUY", 100, 104.0, "GSPT")])
tradefile = trade_df([("AAPL", "BUY", 90, 104.0, "GSPT")])
pcm = src_df([("AAPL", "BUY", 100, 104.0, "GSPT")])
blotter = src_df([("AAPL", "BUY", 100, 104.0, "GSPT")])
report = run(emsx, blotter, pcm, tradefile)
step1 = report[(report["step"] == "Step 1") & (report["severity"] == "ERROR")]
assert len(step1) == 1 and step1.iloc[0]["check"] == "qty_mismatch", step1.to_dict("records")
assert "net_qty_mismatch" not in set(step1["check"])
print("duplicate suppression passed")

# --- Price mismatch is a warning ---------------------------------------------
emsx = emsx_df([("AAPL", "BUY", 100, 104.0, "GSPT")])
tradefile = trade_df([("AAPL", "BUY", 100, 110.0, "GSPT")])
pcm = src_df([("AAPL", "BUY", 100, 104.0, "GSPT")])
blotter = src_df([("AAPL", "BUY", 100, 104.0, "GSPT")])
report = run(emsx, blotter, pcm, tradefile)
warns = report[report["severity"] == "WARN"]
assert "price_mismatch" in set(warns["check"]), report.to_dict("records")
assert errors(report).empty, errors(report).to_dict("records")
print("price mismatch warning passed")

# --- CROS cross checks are retained ------------------------------------------
emsx = emsx_df([("AAPL", "BUY", 100, 104.0, "GSPT")])
tradefile = trade_df([("AAPL", "BUY", 100, 104.0, "GSPT")])
pcm = src_df([
    ("AAPL", "BUY", 100, 104.0, "GSPT"),
    ("MSFT", "BUY", 10, 300.0, "CROS"),    # unbalanced cross
])
blotter = src_df([
    ("AAPL", "BUY", 100, 104.0, "GSPT"),
    ("MSFT", "BUY", 20, 300.0, "CROS"),
    ("MSFT", "SELL", 20, 300.0, "CROS"),   # balanced, but volume differs vs PCM
])
report = run(emsx, blotter, pcm, tradefile)
cross = report[report["step"] == "Crosses"]
assert set(cross["check"]) == {"cross_not_flat", "cross_qty_mismatch"}, cross.to_dict("records")
print("CROS checks passed")

print("\nALL RECONCILIATION TESTS PASSED")
