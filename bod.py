"""Web frontend for the beginning-of-day blotter safety check.

Drag/drop a blotter file or paste an Excel grid; the contents are used as the
blotter input for the trade-error analysis in ``new_main``.

Run with::

    python bod.py

then open http://127.0.0.1:5051 in a browser.
"""

from __future__ import annotations

import inspect
import sys
import threading
import traceback
from io import BytesIO, StringIO
from pathlib import Path

import pandas as pd
from flask import Flask, jsonify, render_template, request

import new_main

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB upload cap

EXCEL_EXTS = {".xlsx", ".xls", ".xlsm", ".xlsb"}
TICKER_COLUMN_INDEX = 3  # Excel column D
TICKER_VALIDATION_FIELDS = (
    "MARKET_STATUS",
    "EQY_PRIM_SECURITY_TICKER",
)
SECURITY_MASTER_UPDATE_LOCK = threading.Lock()


def _read_excel_bytes(data: bytes, skiprows: int) -> pd.DataFrame:
    return pd.read_excel(
        BytesIO(data), skiprows=skiprows, header=None, engine="calamine"
    )


def read_blotter_from_upload(
    data: bytes, filename: str, skiprows: int
) -> pd.DataFrame:
    """Parse an uploaded blotter file into the expected dataframe shape."""
    ext = Path(filename or "").suffix.lower()
    if ext in EXCEL_EXTS:
        return _read_excel_bytes(data, skiprows)
    if ext == ".csv":
        return pd.read_csv(BytesIO(data), skiprows=skiprows, header=None)
    # Unknown extension: try Excel first, then CSV as a fallback.
    try:
        return _read_excel_bytes(data, skiprows)
    except Exception:
        return pd.read_csv(BytesIO(data), skiprows=skiprows, header=None)


def read_blotter_from_paste(text: str, skiprows: int) -> pd.DataFrame:
    """Parse an Excel/CSV grid pasted as text (tab- or comma-separated)."""
    sep = "\t" if "\t" in text else ","
    return pd.read_csv(
        StringIO(text),
        sep=sep,
        header=None,
        skiprows=skiprows,
        engine="python",
        on_bad_lines="skip",
    )


@app.route("/")
def index():
    return render_template(
        "bod/index.html", default_skiprows=new_main.BLOTTER_SKIPROWS
    )


def _read_request_blotter(skiprows: int):
    """Load the uploaded file or pasted grid from the active request."""
    uploaded = request.files.get("file")
    pasted = request.form.get("pasted", "").strip()

    if uploaded and uploaded.filename:
        return (
            read_blotter_from_upload(uploaded.read(), uploaded.filename, skiprows),
            uploaded.filename,
        )
    if pasted:
        return read_blotter_from_paste(pasted, skiprows), "pasted grid"
    raise ValueError("No blotter file or pasted grid provided.")


def _requested_skiprows() -> int:
    try:
        return int(request.form.get("skiprows", new_main.BLOTTER_SKIPROWS))
    except (TypeError, ValueError):
        return new_main.BLOTTER_SKIPROWS


def unique_tickers(blotter_df: pd.DataFrame) -> list[str]:
    """Return non-blank column-D tickers in first-seen order."""
    if blotter_df.shape[1] <= TICKER_COLUMN_INDEX:
        raise ValueError("The blotter does not contain column D.")

    tickers = (
        str(value).strip()
        for value in blotter_df.iloc[:, TICKER_COLUMN_INDEX]
        if not pd.isna(value)
    )
    return list(dict.fromkeys(ticker for ticker in tickers if ticker))


def invalid_bloomberg_tickers(tickers: list[str]) -> list[dict]:
    """Return inactive tickers with Bloomberg status and replacement ticker."""
    if not tickers:
        return []

    try:
        from xbbg import blp
    except ImportError as exc:
        raise RuntimeError(
            "xbbg is not installed; install it to run ticker validation."
        ) from exc

    # Bloomberg requires a yellow key for bare equity symbols (for example,
    # "AAPL US Equity", not "AAPL"). Preserve identifiers that already include
    # one, such as options or non-equity securities.
    requested_securities = [
        ticker if " " in ticker else f"{ticker} US Equity"
        for ticker in tickers
    ]

    # xbbg 0.x omits rejected securities from BDP results. xbbg 1.x can return
    # explicit security-error rows; force pandas/long output there so global
    # backend settings cannot change the response schema.
    request_options = {}
    if "include_security_errors" in inspect.signature(blp.bdp).parameters:
        request_options.update({
            "include_security_errors": True,
            "backend": "pandas",
            "format": "long",
        })

    response = blp.bdp(
        tickers=requested_securities,
        flds=list(TICKER_VALIDATION_FIELDS),
        **request_options,
    )

    if not isinstance(response, pd.DataFrame):
        if hasattr(response, "to_pandas"):
            response = response.to_pandas()
        else:
            response = pd.DataFrame(response)

    security_values: dict[str, dict[str, str]] = {}
    if {"ticker", "field", "value"}.issubset(response.columns):
        for _, row in response.iterrows():
            field = str(row["field"]).upper()
            if field in TICKER_VALIDATION_FIELDS:
                security = str(row["ticker"]).strip().casefold()
                security_values.setdefault(security, {})[field] = str(
                    row["value"]
                ).strip()
    else:
        columns_by_field = {
            str(column).upper(): column
            for column in response.columns
            if str(column).upper() in TICKER_VALIDATION_FIELDS
        }
        for security, row in response.iterrows():
            security_values[str(security).strip().casefold()] = {
                field: str(row[column]).strip()
                for field, column in columns_by_field.items()
                if not pd.isna(row[column])
            }

    invalid_tickers = []
    for ticker, security in zip(tickers, requested_securities):
        values = security_values.get(security.casefold(), {})
        market_status = values.get("MARKET_STATUS", "INVALID").upper()
        if market_status != "ACTV":
            invalid_tickers.append({
                "ticker": ticker,
                "market_status": market_status,
                "new_ticker": values.get("EQY_PRIM_SECURITY_TICKER", ""),
            })
    return invalid_tickers


def _security_master_identifier_values(value) -> list[str]:
    """Normalize a Security Master list cell into clean identifier strings."""
    if value is None:
        return []
    if isinstance(value, str):
        values = value.split(",")
    elif hasattr(value, "tolist"):
        values = value.tolist()
    else:
        try:
            values = list(value)
        except TypeError:
            values = [value]
    return [
        str(item).strip()
        for item in values
        if not pd.isna(item) and str(item).strip()
    ]


def _security_master_ticker_keys(sm_df: pd.DataFrame, ticker: str) -> list[str]:
    """Find every Security Master row containing an exact ticker alias."""
    normalized_ticker = ticker.casefold()
    return [
        str(security_key)
        for security_key, value in sm_df["bloomberg_ticker"].items()
        if normalized_ticker
        in {
            identifier.casefold()
            for identifier in _security_master_identifier_values(value)
        }
    ]


def _load_security_master():
    """Load Backtester's Security Master without requiring package installation."""
    backtester_root = Path(__file__).resolve().parent.parent / "Backtester"
    if not backtester_root.exists():
        raise RuntimeError(f"Backtester repository not found: {backtester_root}")
    root_string = str(backtester_root)
    if root_string not in sys.path:
        sys.path.insert(0, root_string)

    from backtester.security_master.security_master import SecurityMaster

    return SecurityMaster()


def resolve_security_master_tickers(
    invalid_tickers: list[dict],
    security_master=None,
) -> tuple[list[dict], list[dict]]:
    """Promote simple Bloomberg ticker changes in the Security Master."""
    sm = security_master or _load_security_master()
    resolved = []
    skipped = []

    with SECURITY_MASTER_UPDATE_LOCK:
        sm_df = sm.df_disk_image
        pending_updates: dict[str, list[str]] = {}

        for entry in invalid_tickers:
            old_ticker = str(entry.get("ticker", "")).strip()
            new_ticker = str(entry.get("new_ticker", "")).strip()
            market_status = str(entry.get("market_status", "")).upper()

            # Only automate straightforward ticker changes. Acquisitions,
            # delistings, bad/replaced CUSIPs, FIGI changes, and any status
            # other than TKCH can represent a different security and require
            # permanent-ID review before changing the Security Master.
            if market_status != "TKCH":
                skipped.append({
                    "ticker": old_ticker,
                    "reason": f"Market status {market_status or 'UNKNOWN'} "
                    "requires manual review.",
                })
                continue
            if not old_ticker or not new_ticker:
                skipped.append({
                    "ticker": old_ticker,
                    "reason": "Bloomberg did not provide a replacement ticker.",
                })
                continue

            matching_keys = _security_master_ticker_keys(sm_df, old_ticker)
            if len(matching_keys) != 1:
                skipped.append({
                    "ticker": old_ticker,
                    "reason": (
                        "Ticker was not found in the Security Master."
                        if not matching_keys
                        else "Ticker appears in multiple Security Master rows."
                    ),
                })
                continue

            security_key = matching_keys[0]
            conflicting_keys = [
                key
                for key in _security_master_ticker_keys(sm_df, new_ticker)
                if key != security_key
            ]
            if conflicting_keys:
                skipped.append({
                    "ticker": old_ticker,
                    "reason": (
                        f"Replacement ticker {new_ticker} already belongs to "
                        "another Security Master row."
                    ),
                })
                continue

            existing = pending_updates.get(
                security_key,
                _security_master_identifier_values(
                    sm_df.at[security_key, "bloomberg_ticker"]
                ),
            )
            promoted = []
            for ticker in [new_ticker, *existing]:
                if ticker.casefold() not in {
                    value.casefold() for value in promoted
                }:
                    promoted.append(ticker)
            promoted = promoted[:3]

            already_current = (
                bool(existing)
                and existing[0].casefold() == new_ticker.casefold()
            )
            if not already_current:
                pending_updates[security_key] = promoted

            resolved.append({
                "ticker": old_ticker,
                "new_ticker": new_ticker,
                "security_key": security_key,
                "bloomberg_tickers": promoted,
                "status": "already_current" if already_current else "updated",
            })

        if pending_updates:
            updated_count = sm.batch_update_securities(
                pending_updates,
                "bloomberg_ticker",
            )
            if updated_count != len(pending_updates):
                raise RuntimeError(
                    "Security Master updated fewer rows than expected "
                    f"({updated_count}/{len(pending_updates)})."
                )

    return resolved, skipped


@app.route("/check", methods=["POST"])
def check():
    skiprows = _requested_skiprows()

    try:
        blotter_df, source = _read_request_blotter(skiprows)
    except ValueError as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400
    except Exception as exc:  # noqa: BLE001 - surface parse errors to the UI
        return jsonify({
            "ok": False,
            "error": f"Could not read the blotter input: {exc}",
            "detail": traceback.format_exc(),
        }), 400

    try:
        positions_df = new_main.load_positions_df()
    except FileNotFoundError as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400
    except Exception as exc:  # noqa: BLE001
        return jsonify({
            "ok": False,
            "error": f"Could not read the positions file: {exc}",
            "detail": traceback.format_exc(),
        }), 500

    try:
        errors = new_main.analyze(blotter_df, positions_df, skiprows=skiprows)
    except Exception as exc:  # noqa: BLE001
        return jsonify({
            "ok": False,
            "error": f"Analysis failed: {exc}",
            "detail": traceback.format_exc(),
        }), 500

    with_position = [e for e in errors if e.get("has_position")]
    without_position = [e for e in errors if not e.get("has_position")]

    return jsonify({
        "ok": True,
        "source": source,
        "skiprows": skiprows,
        "summary": {
            "total": len(errors),
            "with_position": len(with_position),
            "without_position": len(without_position),
        },
        "with_position": with_position,
        "without_position": without_position,
    })


@app.route("/crosses", methods=["POST"])
def crosses():
    skiprows = _requested_skiprows()

    try:
        blotter_df, source = _read_request_blotter(skiprows)
    except ValueError as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400
    except Exception as exc:  # noqa: BLE001 - surface parse errors to the UI
        return jsonify({
            "ok": False,
            "error": f"Could not read the blotter input: {exc}",
            "detail": traceback.format_exc(),
        }), 400

    try:
        imbalances = new_main.analyze_crosses(blotter_df, skiprows=skiprows)
    except Exception as exc:  # noqa: BLE001
        return jsonify({
            "ok": False,
            "error": f"Crossing-trade analysis failed: {exc}",
            "detail": traceback.format_exc(),
        }), 500

    return jsonify({
        "ok": True,
        "source": source,
        "skiprows": skiprows,
        "summary": {"total": len(imbalances)},
        "imbalances": imbalances,
    })


@app.route("/accounts", methods=["POST"])
def accounts():
    skiprows = _requested_skiprows()

    try:
        blotter_df, source = _read_request_blotter(skiprows)
    except ValueError as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400
    except Exception as exc:  # noqa: BLE001 - surface parse errors to the UI
        return jsonify({
            "ok": False,
            "error": f"Could not read the blotter input: {exc}",
            "detail": traceback.format_exc(),
        }), 400

    try:
        issues = new_main.analyze_fund_broker_account(
            blotter_df, skiprows=skiprows
        )
    except Exception as exc:  # noqa: BLE001
        return jsonify({
            "ok": False,
            "error": f"Fund/broker/account analysis failed: {exc}",
            "detail": traceback.format_exc(),
        }), 500

    return jsonify({
        "ok": True,
        "source": source,
        "skiprows": skiprows,
        "summary": {"total": len(issues)},
        "issues": issues,
    })


@app.route("/tickers", methods=["POST"])
def tickers():
    skiprows = _requested_skiprows()

    try:
        blotter_df, source = _read_request_blotter(skiprows)
        requested_tickers = unique_tickers(blotter_df)
    except ValueError as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400
    except Exception as exc:  # noqa: BLE001 - surface parse errors to the UI
        return jsonify({
            "ok": False,
            "error": f"Could not read the blotter input: {exc}",
            "detail": traceback.format_exc(),
        }), 400

    try:
        invalid_tickers = invalid_bloomberg_tickers(requested_tickers)
    except Exception as exc:  # noqa: BLE001
        return jsonify({
            "ok": False,
            "error": f"Bloomberg ticker validation failed: {exc}",
            "detail": traceback.format_exc(),
        }), 500

    return jsonify({
        "ok": True,
        "source": source,
        "skiprows": skiprows,
        "summary": {
            "unique": len(requested_tickers),
            "valid": len(requested_tickers) - len(invalid_tickers),
            "invalid": len(invalid_tickers),
        },
        "invalid_tickers": invalid_tickers,
    })


@app.route("/resolve-security-master", methods=["POST"])
def resolve_security_master():
    payload = request.get_json(silent=True) or {}
    requested_tickers = payload.get("tickers", [])
    if not isinstance(requested_tickers, list):
        return jsonify({
            "ok": False,
            "error": "tickers must be a JSON list.",
        }), 400

    requested_tickers = list(dict.fromkeys(
        str(ticker).strip()
        for ticker in requested_tickers
        if str(ticker).strip()
    ))
    if not requested_tickers:
        return jsonify({
            "ok": False,
            "error": "No invalid tickers were provided.",
        }), 400
    if len(requested_tickers) > 5_000:
        return jsonify({
            "ok": False,
            "error": "At most 5,000 tickers can be resolved at once.",
        }), 400

    try:
        # Never trust the browser's status/replacement values. Recheck Bloomberg
        # immediately before making a shared Security Master change.
        invalid_tickers = invalid_bloomberg_tickers(requested_tickers)
        invalid_by_ticker = {
            entry["ticker"].casefold(): entry
            for entry in invalid_tickers
        }
        skipped = [
            {
                "ticker": ticker,
                "reason": "Bloomberg now reports this ticker as active.",
            }
            for ticker in requested_tickers
            if ticker.casefold() not in invalid_by_ticker
        ]
        resolved, resolution_skips = resolve_security_master_tickers(
            invalid_tickers
        )
        skipped.extend(resolution_skips)
    except Exception as exc:  # noqa: BLE001
        return jsonify({
            "ok": False,
            "error": f"Security Master resolution failed: {exc}",
            "detail": traceback.format_exc(),
        }), 500

    return jsonify({
        "ok": True,
        "summary": {
            "resolved": len(resolved),
            "updated": sum(
                entry["status"] == "updated" for entry in resolved
            ),
            "already_current": sum(
                entry["status"] == "already_current" for entry in resolved
            ),
            "skipped": len(skipped),
        },
        "resolved": resolved,
        "skipped": skipped,
    })


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5051, debug=True)
