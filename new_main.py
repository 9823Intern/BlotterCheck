from collections import Counter
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd
import pandas_market_calendars as mcal
import pytz


ALLOWED_TRADES = {
    "Long": ["sl", "bl"],
    "Short": ["ss", "cs"],
    "None": ["bl", "ss"],
}

_NYSE = mcal.get_calendar("XNYS")


def previous_market_day(reference: datetime) -> datetime:
    """Return the most recent NYSE trading day strictly before ``reference``."""
    ref_date = reference.date()
    trading_days = _NYSE.valid_days(
        start_date=ref_date - timedelta(days=10),
        end_date=ref_date - timedelta(days=1),
    )
    if len(trading_days) == 0:
        raise RuntimeError(
            f"No NYSE trading day found in the 10 days before {ref_date}"
        )
    return trading_days[-1].to_pydatetime()


# Get the previous market day relative to the current CST date.
cst = pytz.timezone("US/Central")
now_cst = datetime.now(cst)
prev_day = previous_market_day(now_cst)
date_str = prev_day.strftime("%Y%m%d")

POSITIONS_FILE = Path(
    fr"C:\Users\EnnTurn\Precept Dropbox\PCM Share\Quant Models\Nick GS Position Reports\SRPB_198764_1200680426_Custody_Position_301508_409969_{date_str}.xls"
)
BLOTTER_FILE = Path(__file__).parent / "BlotterEOD06.03.26.xlsx"
BLOTTER_SKIPROWS = 26
RED = "\033[91m"
YELLOW = "\033[93m"
RESET = "\033[0m"


def load_positions_df(positions_file: Path = POSITIONS_FILE) -> pd.DataFrame:
    """Read the GS custody positions report, raising if it is missing."""
    positions_file = Path(positions_file)
    if not positions_file.exists():
        raise FileNotFoundError(
            f"Positions file not found for {date_str}: {positions_file}"
        )
    return pd.read_excel(
        positions_file, skiprows=8, header=None, engine="calamine"
    )


def analyze(
    blotter_df: pd.DataFrame,
    positions_df: pd.DataFrame,
    skiprows: int = BLOTTER_SKIPROWS,
) -> list[dict]:
    """Run the full blotter check against an already-loaded blotter dataframe."""
    positions_dict = generate_positions_dict(positions_df)
    trades_df = generate_trades_df(blotter_df, skiprows=skiprows)
    return find_trade_errors(positions_dict, trades_df)


def analyze_crosses(
    blotter_df: pd.DataFrame,
    skiprows: int = BLOTTER_SKIPROWS,
) -> list[dict]:
    """Run the crossing-trade balance check against a loaded blotter dataframe."""
    crosses_df = generate_crosses_df(blotter_df, skiprows=skiprows)
    return find_cross_imbalances(crosses_df)


def generate_positions_dict(positions_df) -> dict:
    """Return {Ticker: [L/S code, trade-date quantity]}."""
    positions_dict = {}
    for _, row in positions_df.iterrows():
        last_ticker = row.iloc[-1]
        if pd.isna(last_ticker) or (
            isinstance(last_ticker, str) and last_ticker.strip() == ""
        ):
            ticker = row.iloc[4]
        else:
            ticker = last_ticker

        if pd.isna(ticker) or (
            isinstance(ticker, str) and str(ticker).strip() == ""
        ):
            continue

        ticker = str(ticker).strip()
        position = row.iloc[5]
        shares = row.iloc[6]
        if pd.isna(position) or pd.isna(shares):
            continue

        shares = float(shares)
        if ticker not in positions_dict:
            positions_dict[ticker] = [position, shares]
        else:
            positions_dict[ticker][1] += shares
    return positions_dict


def position_state(positions_dict: dict, ticker: str) -> tuple[str, float]:
    """Map net share quantity to (Long|Short|None, absolute share count)."""
    entry = positions_dict.get(ticker)
    if not entry:
        return "None", 0.0

    shares = float(entry[1])
    if shares > 0:
        return "Long", shares
    if shares < 0:
        return "Short", abs(shares)
    return "None", 0.0


def generate_trades_df(
    blotter_df: pd.DataFrame, skiprows: int = BLOTTER_SKIPROWS
) -> pd.DataFrame:
    filtered_df = blotter_df[blotter_df.iloc[:, 0].astype(str) != "C"].copy()
    filtered_df = filtered_df[
        (filtered_df.iloc[:, 19].astype(str) == "gs-wsf")
        & (
            ~filtered_df.iloc[:, 11]
            .astype(str)
            .str.lower()
            .str.contains("cros")
        )
    ].copy()
    trades_df = filtered_df
    excel_line = trades_df.index + skiprows + 1
    # Column 11 ("L" in Excel) == "gsop" marks an options trade.
    is_option = (
        trades_df.iloc[:, 11].astype(str).str.strip().str.lower() == "gsop"
    )
    trades_df = trades_df.iloc[:, [3, 2, 4]].copy()
    trades_df.columns = ["Ticker", "Requested_Trade", "share count"]
    trades_df["line"] = excel_line.values
    trades_df["is_option"] = is_option.values
    trades_df["Requested_Trade"] = (
        trades_df["Requested_Trade"].astype(str).str.strip().str.lower()
    )
    trades_df["Ticker"] = trades_df["Ticker"].astype(str).str.strip()
    trades_df["share count"] = pd.to_numeric(
        trades_df["share count"], errors="coerce"
    )
    return trades_df.reset_index(drop=True)


# ---------------------------------------------------------------------------
# Fund / broker / account check (columns S, T vs broker column L)
# ---------------------------------------------------------------------------

TICKER_COLUMN = 3  # Excel column D
BROKER_COLUMN = 11  # Excel column L
FUND_COLUMN = 18  # Excel column S
ACCOUNT_COLUMN = 19  # Excel column T

# Column S value -> allowed column L values
ALLOWED_L_FOR_S = {
    "wsf": {"gspt", "cros", "cfr", "prex", "gsop"},
    "top": {"gspt", "cros"},
    "multiple": {"gspt", "cfr"},
    "ral2": {"tdai"},
    "sche": {"tdai"},
    "sthw": {"tdai"},
    "dbbs": {"tdai"},
    "jk": {"tdai"},
    "ral1": {"tdai"},
}

# Column T value -> allowed column L values
ALLOWED_L_FOR_T = {
    "gs-wsf": {"gspt", "cros", "prex", "gsop"},
    "gs-top": {"gspt", "cros"},
    "cros-wsf": {"cros"},
    "cros-top": {"cros"},
    "can-wsf": {"cros", "cfr"},
    "sch-ral-060": {"tdai"},
    "sch-sch-666": {"tdai"},
    "sch-sou-851": {"tdai"},
    "sch-dbb-129": {"tdai"},
    "sch-jk-606": {"tdai"},
    "sch-ral-551": {"tdai"},
    "multiple": {"prex", "gspt"},
}

# Column S value -> allowed column T values
ALLOWED_T_FOR_S = {
    "wsf": {"can-wsf", "gs-wsf", "cros-wsf"},
    "top": {"cros-top", "gs-top"},
    "multiple": {"multiple"},
    "ral2": {"sch-ral-060"},
    "sche": {"sch-sch-666"},
    "sthw": {"sch-sou-851"},
    "dbbs": {"sch-dbb-129"},
    "jk": {"sch-jk-606"},
    "ral1": {"sch-ral-551"},
}


def _norm_cell(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip().lower()


def analyze_fund_broker_account(
    blotter_df: pd.DataFrame, skiprows: int = BLOTTER_SKIPROWS
) -> list[dict]:
    """Check fund (S), broker (L), and account (T) combinations."""
    issues = []
    for _, row in blotter_df.iterrows():
        ticker = (
            str(row.iloc[TICKER_COLUMN]).strip()
            if blotter_df.shape[1] > TICKER_COLUMN
            and not pd.isna(row.iloc[TICKER_COLUMN])
            else ""
        )

        l_val = (
            _norm_cell(row.iloc[BROKER_COLUMN])
            if blotter_df.shape[1] > BROKER_COLUMN
            else ""
        )
        s_val = (
            _norm_cell(row.iloc[FUND_COLUMN])
            if blotter_df.shape[1] > FUND_COLUMN
            else ""
        )
        t_val = (
            _norm_cell(row.iloc[ACCOUNT_COLUMN])
            if blotter_df.shape[1] > ACCOUNT_COLUMN
            else ""
        )

        # Skip fully blank rows (e.g. trailing empty rows in the sheet).
        if not l_val and not s_val and not t_val:
            continue

        row_issues = []

        if s_val not in ALLOWED_L_FOR_S:
            row_issues.append(f"Unrecognized fund (S) value '{s_val}'")
        elif l_val not in ALLOWED_L_FOR_S[s_val]:
            row_issues.append(
                f"Broker (L) '{l_val}' not allowed for fund (S) '{s_val}' "
                f"(allowed: {', '.join(sorted(ALLOWED_L_FOR_S[s_val]))})"
            )

        if t_val not in ALLOWED_L_FOR_T:
            row_issues.append(f"Unrecognized account (T) value '{t_val}'")
        elif l_val not in ALLOWED_L_FOR_T[t_val]:
            row_issues.append(
                f"Broker (L) '{l_val}' not allowed for account (T) '{t_val}' "
                f"(allowed: {', '.join(sorted(ALLOWED_L_FOR_T[t_val]))})"
            )

        if s_val in ALLOWED_T_FOR_S and t_val not in ALLOWED_T_FOR_S[s_val]:
            row_issues.append(
                f"Account (T) '{t_val}' not allowed for fund (S) '{s_val}' "
                f"(allowed: {', '.join(sorted(ALLOWED_T_FOR_S[s_val]))})"
            )

        for issue in row_issues:
            issues.append({
                "ticker": ticker,
                "broker": l_val,
                "fund": s_val,
                "account": t_val,
                "issue": issue,
            })

    return issues


# Trade types that reduce/sell inventory; their ordered shares are signed
# negative so a balanced cross (buys vs. sells) nets to zero per ticker.
_SELL_TRADES = {"sl", "ss"}


def generate_crosses_df(
    blotter_df: pd.DataFrame, skiprows: int = BLOTTER_SKIPROWS
) -> pd.DataFrame:
    """Build a crossing-trades-only blotter."""
    filtered_df = blotter_df[
        blotter_df.iloc[:, 11].astype(str).str.lower().str.contains("cros")
    ].copy()
    excel_line = filtered_df.index + skiprows + 1
    crosses_df = filtered_df.iloc[:, [3, 2, 4]].copy()
    crosses_df.columns = ["Ticker", "Requested_Trade", "share count"]
    crosses_df["line"] = excel_line.values
    crosses_df["Requested_Trade"] = (
        crosses_df["Requested_Trade"].astype(str).str.strip().str.lower()
    )
    crosses_df["Ticker"] = crosses_df["Ticker"].astype(str).str.strip()
    crosses_df["share count"] = pd.to_numeric(
        crosses_df["share count"], errors="coerce"
    )
    return crosses_df.reset_index(drop=True)


def find_cross_imbalances(crosses_df: pd.DataFrame) -> list[dict]:
    """Return tickers whose signed crossing-share total is non-zero."""
    net_shares: Counter = Counter()
    trade_counts: Counter = Counter()
    for _, row in crosses_df.iterrows():
        size = row["share count"]
        if pd.isna(size):
            continue
        size = float(size)
        if row["Requested_Trade"] in _SELL_TRADES:
            size = -size
        net_shares[row["Ticker"]] += size
        trade_counts[row["Ticker"]] += 1

    imbalances = []
    for ticker, net in net_shares.items():
        if round(net, 6) != 0.0:
            imbalances.append({
                "Ticker": ticker,
                "net_shares": net,
                "trade_count": trade_counts[ticker],
            })
    imbalances.sort(key=lambda entry: abs(entry["net_shares"]), reverse=True)
    return imbalances


def _trade_error(
    line: int,
    ticker: str,
    trade_type: str,
    share_count,
    reason: str,
    position: str,
    has_position: bool,
    holdings_side: str,
    holdings_qty: float,
    is_option: bool = False,
) -> dict:
    return {
        "line": line,
        "Ticker": ticker,
        "Requested_Trade": trade_type,
        "share count": share_count,
        "Position": position,
        "reason": reason,
        "has_position": has_position,
        "holdings_side": holdings_side,
        "holdings_qty": holdings_qty,
        "is_option": bool(is_option),
    }


def _signed_net(side: str, qty: float) -> float:
    """Convert a (side, absolute qty) pair into a signed share count."""
    if side == "Long":
        return float(qty)
    if side == "Short":
        return -float(qty)
    return 0.0


def _classify_trade(
    net: float, trade_type: str, size: float
) -> tuple[bool, str | None, float]:
    """Validate one trade against the running, share-level position."""
    if pd.isna(size):
        size = 0.0

    if trade_type == "bl":
        if net < 0:
            return (
                True,
                "buy long while still short (cover the short first)",
                net,
            )
        return False, None, net + size

    if trade_type == "sl":
        if net <= 0:
            return True, "sell long with no long inventory", net
        if size > net:
            return True, f"sell long exceeds long inventory ({net:g} held)", net
        return False, None, net - size

    if trade_type == "ss":
        if net > 0:
            return (
                True,
                "short sale while still long (sell the long first)",
                net,
            )
        return False, None, net - size

    if trade_type == "cs":
        if net >= 0:
            return True, "cover short with no short position", net
        if size > -net:
            return True, f"cover exceeds short position ({-net:g} short)", net
        return False, None, net + size

    return True, f"unknown trade type '{trade_type}'", net


_ORDER_FROM_LONG = {"bl": 0, "sl": 1, "ss": 2, "cs": 3}
_ORDER_FROM_SHORT = {"ss": 0, "cs": 1, "bl": 2, "sl": 3}


def _execution_order_key(trade_type: str, start_net: float) -> int:
    table = _ORDER_FROM_SHORT if start_net < 0 else _ORDER_FROM_LONG
    return table.get(trade_type, 99)


def check_ticker_trades(
    trades: pd.DataFrame,
    start_side: str,
    start_qty: float,
) -> list[dict]:
    """Validate the full set of a ticker's trades, ignoring blotter row order."""
    errors = []
    has_position = start_side != "None"
    net = _signed_net(start_side, start_qty)

    ordered_rows = sorted(
        (row for _, row in trades.iterrows()),
        key=lambda row: _execution_order_key(row["Requested_Trade"], net),
    )

    for row in ordered_rows:
        trade_type = row["Requested_Trade"]
        is_error, reason, net = _classify_trade(
            net, trade_type, row["share count"]
        )
        if is_error:
            errors.append(
                _trade_error(
                    int(row["line"]),
                    row["Ticker"],
                    trade_type,
                    row["share count"],
                    reason,
                    start_side,
                    has_position,
                    start_side,
                    start_qty,
                    bool(row.get("is_option", False)),
                )
            )

    errors.sort(key=lambda entry: entry["line"])
    return errors


def find_trade_errors(
    positions_dict: dict, trades_df: pd.DataFrame
) -> list[dict]:
    errors = []
    for ticker in trades_df["Ticker"].unique():
        ticker_trades = trades_df[trades_df["Ticker"] == ticker]
        start_side, start_qty = position_state(positions_dict, ticker)
        errors.extend(check_ticker_trades(ticker_trades, start_side, start_qty))
    return errors


def _format_error_block(err: dict, number: int) -> list[str]:
    option_tag = f" {YELLOW}[OPTION]{RESET}" if err.get("is_option") else ""
    trade_line = (
        f"  #{number}  LINE {err['line']} | {err['Ticker']} | "
        f"{err['Requested_Trade']} | {err['share count']} shares{option_tag}"
    )
    if err["has_position"]:
        detail = (
            f"       Current holdings: {err['holdings_side']} "
            f"({err['holdings_qty']:g} shares)  Reason: {err['reason']}"
        )
    else:
        detail = (
            f"       Position: {err['Position']}  Reason: {err['reason']}"
        )
    block = [trade_line, detail, ""]
    if err["has_position"]:
        block = [
            f"{RED}{line}{RESET}" if line else line for line in block
        ]
    return block


def format_errors(errors: list[dict]) -> str:
    with_position = [entry for entry in errors if entry["has_position"]]
    without_position = [entry for entry in errors if not entry["has_position"]]

    lines = [
        "Blotter Check Report",
        "=" * 60,
        f"Total ERROR trades: {len(errors)}",
        f"  With current holdings: {len(with_position)}",
        f"  Without current holdings: {len(without_position)}",
        "",
    ]
    if not errors:
        lines.append("No problematic trades detected.")
        return "\n".join(lines)

    if with_position:
        lines.append(
            f"{RED}--- Trades on tickers WITH a current position ---{RESET}"
        )
        lines.append("")
        for index, err in enumerate(with_position, 1):
            lines.extend(_format_error_block(err, index))

    if without_position:
        lines.append("--- Trades on tickers WITHOUT a current position ---")
        lines.append("")
        offset = len(with_position)
        for index, err in enumerate(without_position, 1):
            lines.extend(_format_error_block(err, offset + index))

    return "\n".join(lines)


def main() -> list[dict]:
    try:
        positions_df = load_positions_df()
    except FileNotFoundError:
        print("Error: Positions File not updated for today")
        raise SystemExit(1)

    blotter_df = pd.read_excel(
        BLOTTER_FILE, skiprows=BLOTTER_SKIPROWS, header=None, engine="calamine"
    )
    return analyze(blotter_df, positions_df, skiprows=BLOTTER_SKIPROWS)


if __name__ == "__main__":
    flagged = main()
    print(format_errors(flagged))
