"""Web frontend for the four-step end-of-day blotter reconciliation.

Each of the four sources (EMSX grid, Blotter EOD, PCMBlotter, Goldman tradefile)
can be provided as an uploaded file or pasted text. The workflow from
``blotter_check`` runs and all findings are returned to the browser:

    Step 1: EMSX (GSPT/PREX/GSOP) vs Goldman tradefile
    Step 2: EMSX (excluding TDAI) vs PCMBlotter
    Step 3: PCM (GSPT/PREX/GSOP)  vs Goldman tradefile
    Step 4: Blotter EOD adjudication of remaining discrepancies

Run with::

    python eod.py

then open http://127.0.0.1:5052 in a browser.
"""

from __future__ import annotations

import traceback

from flask import Flask, jsonify, render_template, request

import blotter_check as bc

app = Flask(__name__)
# This is a localhost-only tool whose inputs can be large pasted exports.
# Disable Flask/Werkzeug's request and multipart parser limits.
app.config.update(
    MAX_CONTENT_LENGTH=None,
    MAX_FORM_MEMORY_SIZE=None,
    MAX_FORM_PARTS=None,
)

SOURCES = {
    "emsx": (bc.load_emsx_bytes, bc.load_emsx_text, "EMSX grid"),
    "blotter": (bc.load_blotter_bytes, bc.load_blotter_text, "Blotter EOD"),
    "pcm": (bc.load_pcm_bytes, bc.load_pcm_text, "PCMBlotter"),
    "tradefile": (bc.load_tradefile_bytes, bc.load_tradefile_text, "Tradefile"),
}


def _load_source(key, findings):
    """Load one source from the request (file upload wins over pasted text)."""
    from_bytes, from_text, label = SOURCES[key]
    uploaded = request.files.get(f"{key}_file")
    # Do NOT strip the pasted text: a leading tab means an empty first cell
    # (e.g. the EMSX header row) and stripping it would shift every column.
    pasted = request.form.get(f"{key}_text") or ""

    if uploaded and uploaded.filename:
        return (
            from_bytes(uploaded.read(), uploaded.filename, findings),
            uploaded.filename,
        )
    if pasted.strip():
        return from_text(pasted, findings), "pasted"
    raise ValueError(f"No {label} file or pasted text provided.")


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/check", methods=["POST"])
def check():
    try:
        qty_tol = int(request.form.get("qty_tol", 0))
    except (TypeError, ValueError):
        qty_tol = 0
    try:
        price_tol = float(request.form.get("price_tol", 0.01))
    except (TypeError, ValueError):
        price_tol = 0.01

    findings = bc.Findings()
    loaded = {}
    sources_meta = {}
    for key in SOURCES:
        try:
            df, origin = _load_source(key, findings)
            loaded[key] = df
            sources_meta[key] = {"origin": origin, "rows": int(len(df))}
        except ValueError as exc:
            return jsonify({"ok": False, "error": str(exc)}), 400
        except Exception as exc:  # noqa: BLE001 - surface parse errors to the UI
            return jsonify({
                "ok": False,
                "error": f"Could not read the {SOURCES[key][2]} input: {exc}",
                "detail": traceback.format_exc(),
            }), 400

    try:
        report = bc.reconcile(
            loaded["emsx"],
            loaded["blotter"],
            loaded["pcm"],
            loaded["tradefile"],
            findings,
            qty_tol=qty_tol,
            price_tol=price_tol,
        )
    except Exception as exc:  # noqa: BLE001
        return jsonify({
            "ok": False,
            "error": f"Reconciliation failed: {exc}",
            "detail": traceback.format_exc(),
        }), 500

    return jsonify({
        "ok": True,
        "sources": sources_meta,
        "summary": {
            "errors": int((report["severity"] == "ERROR").sum()),
            "warnings": int((report["severity"] == "WARN").sum()),
            "resolved": int((report["severity"] == "RESOLVED").sum()),
        },
        "findings": report.to_dict(orient="records"),
    })


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5052, debug=True)
