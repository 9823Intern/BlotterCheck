"""Exercise the /check endpoint with synthetic pasted inputs for all four sources."""

import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from eod import app

client = app.test_client()


def post(data):
    resp = client.post("/check", data=data, content_type="multipart/form-data")
    body = resp.get_json()
    assert body is not None, resp.data[:500]
    return resp.status_code, body


# --- Synthetic pasted grids ---------------------------------------------------
# AAPL BUY 100 @ 104 (GSPT) appears in all four sources; MSFT SELL 50 (TDAI)
# is in EMSX only and is excluded from every comparison scope.
EMSX_HEADER = ["Security", "SEDOL", "Side", "Status", "FillQty", "AvgPx", "Def Brkr Code", "Qty"]
emsx_text = "\n".join("\t".join(str(c) for c in row) for row in [
    EMSX_HEADER,
    ["AAPL", "2046251", "Buy", "Filled", 100, 104.0, "GSPT", 100],
    ["MSFT", "2588173", "Sell", "Filled", 50, 300.0, "TDAI", 50],
    ["NVDA", "2379504", "Buy", "Working", 7, 150.0, "GSPT", 7],  # not Filled: ignored
])


def blotter_row(side_code, ticker, qty, price, broker, fund="F1"):
    cells = [""] * 19
    cells[0] = "S"
    cells[2] = side_code
    cells[3] = ticker
    cells[4] = str(qty)
    cells[8] = str(qty)
    cells[9] = str(price)
    cells[11] = broker
    cells[18] = fund
    return "\t".join(cells)


blotter_text = blotter_row("bl", "AAPL", 100, 104.0, "GSPT")

pcm_text = "\n".join([
    "Fund,Ticker,Transaction Type,Quantity,Price,Broker",
    "F1,AAPL,Buy,100,104.00,GSPT",
])

tradefile_text = "\n".join([
    "HEADER|1",
    "ORDER||O1|20260715|||B|||2046251|||||GSPT|||104.0000|100|",
    "ALLOCAT||065465783|100",
    "TRAILER|2",
])

# --- Case 1: consistent pastes reconcile cleanly ------------------------------
code, body = post({
    "emsx_text": emsx_text,
    "blotter_text": blotter_text,
    "pcm_text": pcm_text,
    "tradefile_text": tradefile_text,
})
print("clean:", code, body.get("summary"), body.get("error", ""))
assert code == 200 and body["ok"], body
assert body["summary"] == {"errors": 0, "warnings": 0, "resolved": 0}, body["summary"]
assert body["sources"]["emsx"]["rows"] == 2, body["sources"]  # Working row dropped

# --- Case 2: tampered tradefile shows step findings ---------------------------
tampered = tradefile_text.replace("|104.0000|100|", "|104.0000|90|").replace(
    "ALLOCAT||065465783|100", "ALLOCAT||065465783|90")
code, body = post({
    "emsx_text": emsx_text,
    "blotter_text": blotter_text,
    "pcm_text": pcm_text,
    "tradefile_text": tampered,
})
print("tampered:", code, body.get("summary"))
assert code == 200 and body["ok"] and body["summary"]["errors"] >= 2, body["summary"]
steps = {f["step"] for f in body["findings"] if f["severity"] == "ERROR"}
assert {"Step 1", "Step 3"} <= steps, steps
assert all("left_value" in f and "right_value" in f for f in body["findings"])

# --- Case 3: blotter-resolved side classification ------------------------------
pcm_cover_as_buy = pcm_text  # PCM says BUY
emsx_cover = emsx_text.replace("AAPL\t2046251\tBuy\tFilled", "AAPL\t2046251\tBuy to Cover\tFilled")
trade_cover = tradefile_text.replace("|B|", "|BC|")
blotter_cover = blotter_row("cs", "AAPL", 100, 104.0, "GSPT")  # corroborates COVER
code, body = post({
    "emsx_text": emsx_cover,
    "blotter_text": blotter_cover,
    "pcm_text": pcm_cover_as_buy,
    "tradefile_text": trade_cover,
})
print("resolved:", code, body.get("summary"))
assert code == 200 and body["ok"], body
assert body["summary"]["errors"] == 0, body["findings"]
assert body["summary"]["resolved"] >= 1, body["summary"]

# --- Case 4: missing source rejected -------------------------------------------
code, body = post({"emsx_text": emsx_text})
print("missing source:", code, body.get("error"))
assert code == 400 and not body["ok"]

print("\nALL WEB TESTS PASSED")
