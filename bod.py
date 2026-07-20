"""Web frontend for the beginning-of-day blotter safety check.

Drag/drop a blotter file or paste an Excel grid; the contents are used as the
blotter input for the trade-error analysis in ``new_main``.

Run with::

    python bod.py

then open http://127.0.0.1:5051 in a browser.
"""

from __future__ import annotations

import inspect
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


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5051, debug=True)
