"""
generate_negotiation_deck.py
Flask routes for the "Difficult Conversations" curriculum deck.
Add this file to the ta-refund-deck repo, then in app.py:
    from generate_negotiation_deck import register_negotiation_routes
    register_negotiation_routes(app)

Slide building runs via Node (generate_negotiation_deck.js).
"""

import subprocess
import json
import os
import tempfile
from flask import request, send_file, jsonify


def _extract_json_payload(text: str) -> dict:
    """
    Claude (esp. Haiku) sometimes wraps JSON in preamble or ```json fences.
    Grab everything between the first { and the last } and parse that.
    """
    start = text.find("{")
    end = text.rfind("}")
    if start == -1 or end == -1:
        raise ValueError("No JSON object found in Claude response")
    return json.loads(text[start:end + 1])


def register_negotiation_routes(app):
    @app.route("/generate-negotiation-deck/aunz", methods=["POST"])
    def generate_negotiation_deck_aunz():
        return _run(request, "aunz")

    @app.route("/generate-negotiation-deck/america", methods=["POST"])
    def generate_negotiation_deck_america():
        return _run(request, "america")


def _run(request, region: str):
    try:
        raw = request.get_data(as_text=True)

        # Body may be clean JSON or a Claude text blob with fences/preamble
        try:
            payload = json.loads(raw)
        except json.JSONDecodeError:
            payload = _extract_json_payload(raw)

        payload["region"] = region  # trust the route over the body

        with tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False) as tmp:
            json.dump(payload, tmp)
            tmp_path = tmp.name

        out_path = tempfile.mktemp(suffix=f"_neg_{region}.pptx")

        result = subprocess.run(
            ["node", "generate_negotiation_deck.js", "--json", tmp_path, "--out", out_path],
            capture_output=True, text=True, timeout=60,
        )
        os.unlink(tmp_path)

        if result.returncode != 0:
            return jsonify({"error": result.stderr}), 500

        lesson = payload.get("lesson_number", "X")
        skill = payload.get("skill_name", "lesson").replace(" ", "_").replace("'", "")
        date_label = payload.get("date_label", "wk").replace(" ", "_")
        filename = f"DifficultConversations_L{lesson}_{skill}_{date_label}_{region}.pptx"

        return send_file(
            out_path,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500
