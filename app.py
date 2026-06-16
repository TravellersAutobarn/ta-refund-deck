from flask import Flask, request, jsonify, send_file
import json
import os
import subprocess
import tempfile
import sys

from generate_praise_deck import generate_region as generate_praise_region
from generate_day1_deck import generate_region as generate_day1_region
from generate_wwyd_deck import generate_region as generate_wwyd_region
from generate_qotw_deck import generate_region as generate_qotw_region
from generate_negotiation_deck import generate_region as generate_negotiation_region

app = Flask(__name__)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def run_deck_generator(script_name, data, output_filename, date_label="This Period", api_key=None):
    """
    Shared helper for routes that run a Python deck generator script.
    Saves request JSON to a temporary file, runs the selected generator,
    then returns the generated PowerPoint file.
    """
    with tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False) as f:
        json.dump(data, f)
        data_path = f.name

    output_path = tempfile.mktemp(suffix=".pptx")

    command = [
        sys.executable,
        script_name,
        "--data",
        data_path,
        "--output",
        output_path,
        "--date-label",
        date_label,
    ]

    if api_key:
        command.extend(["--api-key", api_key])

    result = subprocess.run(
        command,
        capture_output=True,
        text=True,
        cwd=BASE_DIR,
        env={**os.environ, "NODE_PATH": os.path.join(BASE_DIR, "node_modules")},
    )

    os.unlink(data_path)

    if result.returncode != 0:
        return jsonify({"error": result.stderr + result.stdout}), 500

    return send_file(
        output_path,
        as_attachment=True,
        download_name=output_filename,
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )


@app.route("/generate", methods=["POST"])
def generate():
    data = request.get_json()

    if not data:
        return jsonify({"error": "No JSON body received"}), 400

    rows = data.get("rows", [])
    date_label = data.get("date_label", "This Period")
    api_key = data.get("api_key", os.environ.get("ANTHROPIC_API_KEY"))

    return run_deck_generator(
        script_name="generate_aunz_deck.py",
        data=rows,
        output_filename="aunz_refund_report.pptx",
        date_label=date_label,
        api_key=api_key,
    )


@app.route("/generate-nz", methods=["POST"])
def generate_nz():
    data = request.get_json()

    if not data:
        return jsonify({"error": "No JSON body received"}), 400

    rows = data.get("rows", [])
    date_label = data.get("date_label", "This Period")
    api_key = data.get("api_key", os.environ.get("ANTHROPIC_API_KEY"))

    return run_deck_generator(
        script_name="generate_nz_deck.py",
        data=rows,
        output_filename="nz_refund_report.pptx",
        date_label=date_label,
        api_key=api_key,
    )


@app.route("/generate-america", methods=["POST"])
def generate_america():
    data = request.get_json()

    if not data:
        return jsonify({"error": "No JSON body received"}), 400

    rows = data.get("rows", [])
    date_label = data.get("date_label", "This Period")
    api_key = data.get("api_key", os.environ.get("ANTHROPIC_API_KEY"))

    return run_deck_generator(
        script_name="generate_america_deck.py",
        data=rows,
        output_filename="america_refund_report.pptx",
        date_label=date_label,
        api_key=api_key,
    )


@app.route("/generate-info-requests", methods=["POST"])
def generate_info_requests():
    """
    Existing information-request trends deck.

    Expected body:
    {
        "rows": [
            {
                "branch": "...",
                "issue": "Category :: question",
                "subject": "..."
            }
        ],
        "date_label": ""
    }
    """
    data = request.get_json()

    if not data:
        return jsonify({"error": "No JSON body received"}), 400

    rows = data.get("rows", [])
    date_label = data.get("date_label", "")
    api_key = data.get("api_key", os.environ.get("ANTHROPIC_API_KEY"))

    return run_deck_generator(
        script_name="generate_info_request_deck.py",
        data=rows,
        output_filename="info_request_report.pptx",
        date_label=date_label,
        api_key=api_key,
    )


@app.route("/generate-accidents", methods=["POST"])
def generate_accidents():
    """
    Accident report deck.

    Expected body:
    {
        "rows": [
            {
                "license_plate": "...",
                "cause_summary": "...",
                "location": "...",
                "date": "...",
                "photo_url": "..."
            }
        ],
        "region": "America",
        "date_label": ""
    }
    """
    data = request.get_json()

    if not data:
        return jsonify({"error": "No JSON body received"}), 400

    region = data.get("region", "All Regions")
    date_label = data.get("date_label", "")
    rows = data.get("rows", [])
    api_key = data.get("api_key", os.environ.get("ANTHROPIC_API_KEY"))

    return run_deck_generator(
        script_name="generate_accident_deck.py",
        data={"rows": rows, "region": region, "date_label": date_label},
        output_filename="accident_report_" + region.replace(" ", "_").replace("&", "and").lower() + ".pptx",
        date_label=date_label,
        api_key=api_key,
    )


@app.route("/generate-us-info-trends", methods=["POST"])
def generate_us_info_trends():
    """
    US More Information Trends deck.

    Expected body:
    {
        "date_label": "Previous Month",
        "slides": [
            {
                "slide_number": 1,
                "title": "Las Vegas top 3 trends",
                "bullets": [
                    "Trend 1: ...",
                    "Trend 2: ...",
                    "Trend 3: ..."
                ]
            }
        ]
    }
    """
    data = request.get_json()

    if not data:
        return jsonify({"error": "No JSON body received"}), 400

    date_label = data.get("date_label", "Previous Month")

    return run_deck_generator(
        script_name="generate_us_info_trends_deck.py",
        data=data,
        output_filename="us_more_information_trends_deck.pptx",
        date_label=date_label,
        api_key=None,
    )


@app.route("/generate-aunz-info-trends", methods=["POST"])
def generate_aunz_info_trends():
    """
    Australia and New Zealand More Information Trends deck.

    Expected body:
    {
        "date_label": "Previous Month",
        "slides": [
            {
                "slide_number": 1,
                "title": "Auckland top 3 trends",
                "bullets": [
                    "Trend 1: ...",
                    "Trend 2: ...",
                    "Trend 3: ..."
                ]
            }
        ]
    }

    This uses the same slide JSON to PPTX generator as the US deck.
    The generator is location-agnostic; it builds slides from the JSON supplied.
    """
    data = request.get_json()

    if not data:
        return jsonify({"error": "No JSON body received"}), 400

    date_label = data.get("date_label", "Previous Month")

    return run_deck_generator(
        script_name="generate_us_info_trends_deck.py",
        data=data,
        output_filename="aunz_more_information_trends_deck.pptx",
        date_label=date_label,
        api_key=None,
    )


def _run_praise_region(region, filename):
    """
    Build one region's Wall of Praise deck and return it as a PowerPoint file.
    """
    import traceback

    try:
        payload = request.get_json(force=True)

        if not payload:
            return jsonify({"error": "No JSON body received"}), 400

        outdir = tempfile.mkdtemp(prefix="praise_")
        path = generate_praise_region(payload, outdir, region)

        if not path:
            return jsonify({
                "error": "No praise items matched region '%s'" % region,
                "received_praise_count": len(payload.get("praise", []) or [])
                    if isinstance(payload.get("praise"), (list, tuple)) else 1,
            }), 400

        return send_file(
            path,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )

    except Exception as e:
        return jsonify({"error": str(e), "traceback": traceback.format_exc()}), 500


@app.route("/generate-praise-deck/aunz", methods=["POST"])
def praise_deck_aunz():
    return _run_praise_region("aunz", "praise_aunz.pptx")


@app.route("/generate-praise-deck/america", methods=["POST"])
def praise_deck_america():
    return _run_praise_region("america", "praise_america.pptx")


def _run_day1_region(region, filename):
    """
    Build one region's Day 1 issues deck and return it as a PowerPoint file.
    """
    import traceback

    try:
        payload = request.get_json(force=True)

        if not payload:
            return jsonify({"error": "No JSON body received"}), 400

        outdir = tempfile.mkdtemp(prefix="day1_")
        path = generate_day1_region(payload, outdir, region)

        if not path:
            return jsonify({
                "error": "No Day 1 issue items matched region '%s'" % region,
                "received_count": len(payload.get("praise", []) or [])
                    if isinstance(payload.get("praise"), (list, tuple)) else 1,
            }), 400

        return send_file(
            path,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )

    except Exception as e:
        return jsonify({"error": str(e), "traceback": traceback.format_exc()}), 500


@app.route("/generate-day1-deck/aunz", methods=["POST"])
def day1_deck_aunz():
    return _run_day1_region("aunz", "day1_aunz.pptx")


@app.route("/generate-day1-deck/america", methods=["POST"])
def day1_deck_america():
    return _run_day1_region("america", "day1_america.pptx")


def _run_wwyd_region(region, filename):
    """
    Build one region's What Would You Do deck and return it as a PowerPoint file.
    Accepts Claude raw JSON output and tolerates markdown fences.
    """
    import traceback
    import json as _json
    import re as _re

    try:
        raw = request.get_data(as_text=True) or ""
        cleaned = raw.strip()

        cleaned = _re.sub(r"^```[a-zA-Z]*", "", cleaned).strip()
        cleaned = _re.sub(r"```$", "", cleaned).strip()

        try:
            payload = _json.loads(cleaned)
        except Exception:
            payload = request.get_json(force=True, silent=True) or {}

        if not payload:
            return jsonify({"error": "No JSON body received"}), 400

        outdir = tempfile.mkdtemp(prefix="wwyd_")
        path = generate_wwyd_region(payload, outdir, region)

        if not path:
            return jsonify({"error": "No scenario matched region '%s'" % region}), 400

        return send_file(
            path,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )

    except Exception as e:
        return jsonify({"error": str(e), "traceback": traceback.format_exc()}), 500


@app.route("/generate-wwyd-deck/aunz", methods=["POST"])
def wwyd_deck_aunz():
    return _run_wwyd_region("aunz", "wwyd_aunz.pptx")


@app.route("/generate-wwyd-deck/america", methods=["POST"])
def wwyd_deck_america():
    return _run_wwyd_region("america", "wwyd_america.pptx")


def _run_qotw_region(region, filename):
    """
    Build one region's Question of the Week deck and return it as a PowerPoint file.
    Accepts Claude raw JSON output and tolerates markdown fences.
    """
    import traceback
    import json as _json
    import re as _re

    try:
        raw = request.get_data(as_text=True) or ""
        cleaned = raw.strip()

        cleaned = _re.sub(r"^```[a-zA-Z]*", "", cleaned).strip()
        cleaned = _re.sub(r"```$", "", cleaned).strip()

        try:
            payload = _json.loads(cleaned)
        except Exception:
            payload = request.get_json(force=True, silent=True) or {}

        if not payload:
            return jsonify({"error": "No JSON body received"}), 400

        outdir = tempfile.mkdtemp(prefix="qotw_")
        path = generate_qotw_region(payload, outdir, region)

        if not path:
            return jsonify({"error": "No question matched region '%s'" % region}), 400

        return send_file(
            path,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )

    except Exception as e:
        return jsonify({"error": str(e), "traceback": traceback.format_exc()}), 500


@app.route("/generate-qotw-deck/aunz", methods=["POST"])
def qotw_deck_aunz():
    return _run_qotw_region("aunz", "qotw_aunz.pptx")


@app.route("/generate-qotw-deck/america", methods=["POST"])
def qotw_deck_america():
    return _run_qotw_region("america", "qotw_america.pptx")


def _run_negotiation_region(region, filename):
    """
    Build one region's Difficult Conversations curriculum deck and return it
    as a PowerPoint file. Accepts Claude raw JSON output and tolerates
    markdown fences (same handling as wwyd/qotw).
    """
    import traceback
    import json as _json
    import re as _re

    try:
        raw = request.get_data(as_text=True) or ""
        cleaned = raw.strip()

        cleaned = _re.sub(r"^```[a-zA-Z]*", "", cleaned).strip()
        cleaned = _re.sub(r"```$", "", cleaned).strip()

        try:
            payload = _json.loads(cleaned)
        except Exception:
            payload = request.get_json(force=True, silent=True) or {}

        if not payload:
            return jsonify({"error": "No JSON body received"}), 400

        outdir = tempfile.mkdtemp(prefix="negotiation_")
        path = generate_negotiation_region(payload, outdir, region)

        if not path:
            return jsonify({"error": "No lesson payload received for region '%s'" % region}), 400

        return send_file(
            path,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )

    except Exception as e:
        return jsonify({"error": str(e), "traceback": traceback.format_exc()}), 500


@app.route("/generate-negotiation-deck/aunz", methods=["POST"])
def negotiation_deck_aunz():
    return _run_negotiation_region("aunz", "negotiation_aunz.pptx")


@app.route("/generate-negotiation-deck/america", methods=["POST"])
def negotiation_deck_america():
    return _run_negotiation_region("america", "negotiation_america.pptx")


@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
