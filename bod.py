"""Web frontend for the beginning-of-day blotter safety check.

Drag/drop a blotter file or paste an Excel grid; the contents are used as the
blotter input for the trade-error analysis in ``new_main``.

Run with::

    python bod.py

then open http://127.0.0.1:5051 in a browser.
"""

from __future__ import annotations

import traceback
from io import BytesIO, StringIO
from pathlib import Path

import pandas as pd
from flask import Flask, jsonify, render_template, request

import new_main

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB upload cap

EXCEL_EXTS = {".xlsx", ".xls", ".xlsm", ".xlsb"}


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


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5051, debug=True)
