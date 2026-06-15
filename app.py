from flask import Flask, request, jsonify, send_file
import json
import os
import subprocess
import tempfile
import sys

from generate_praise_deck import generate_region as generate_praise_region
from generate_day1_deck import generate_region as generate_day1_region

app = Flask(__name__)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def run_deck_generator(script_name, data, output_filename, date_label="This Period", api_key=None):
    """
    Shared helper for all deck routes.
    Saves request JSON to a temp file, runs the selected generator script,
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
    Information-request trends deck.
    Expects: { "rows": [ {"branch": "...", "issue": "Category :: question", "subject": "..."} ], "date_label": "" }
    Empty date_label = date-less deck.
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
    Expects: { "rows": [ {"license_plate": "...", "cause_summary": "...", "location": "...", "date": "...", "photo_url": "..."} ], "region": "America", "date_label": "" }
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


def _run_praise_region(region, filename):
    """Build one region's praise deck and return it as a single .pptx, or 400 if empty."""
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
    """Wall of Praise deck for Australia & New Zealand. Returns praise_aunz.pptx."""
    return _run_praise_region("aunz", "praise_aunz.pptx")


@app.route("/generate-praise-deck/america", methods=["POST"])
def praise_deck_america():
    """Wall of Praise deck for the USA. Returns praise_america.pptx."""
    return _run_praise_region("america", "praise_america.pptx")


def _run_day1_region(region, filename):
    """Build one region's Day 1 issues deck and return it as a single .pptx, or 400 if empty."""
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
    """Day 1 issues deck for Australia & New Zealand. Returns day1_aunz.pptx."""
    return _run_day1_region("aunz", "day1_aunz.pptx")


@app.route("/generate-day1-deck/america", methods=["POST"])
def day1_deck_america():
    """Day 1 issues deck for the USA. Returns day1_america.pptx."""
    return _run_day1_region("america", "day1_america.pptx")


@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
