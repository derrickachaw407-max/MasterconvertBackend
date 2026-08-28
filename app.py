import os
import tempfile
import shutil
import uuid
import requests
from flask import Flask, request, send_file, jsonify
from converters import convert, text_to_pptx, extract_text, ConversionError

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024  # 25MB upload cap

MIME_TYPES = {
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "pdf": "application/pdf",
}

# CORS: allow the frontend (any origin, since this is a public conversion API).
# Tighten this to your real domain once it's live, e.g. "https://masterconvert.app"
ALLOWED_ORIGIN = "https://masterconvert-tau.vercel.app"

# AI features (Smart Summarize / drafting) call Claude directly.
# Set ANTHROPIC_API_KEY as an environment variable on Render — never hardcode it here.
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
ANTHROPIC_MODEL = "claude-sonnet-5"


def call_claude(system_prompt, user_message, max_tokens=600):
    if not ANTHROPIC_API_KEY:
        raise ConversionError("AI features need ANTHROPIC_API_KEY set on the server")
    resp = requests.post(
        "https://api.anthropic.com/v1/messages",
        headers={
            "x-api-key": ANTHROPIC_API_KEY,
            "anthropic-version": "2023-06-01",
            "content-type": "application/json",
        },
        json={
            "model": ANTHROPIC_MODEL,
            "max_tokens": max_tokens,
            "system": system_prompt,
            "messages": [{"role": "user", "content": user_message}],
        },
        timeout=45,
    )
    if resp.status_code != 200:
        raise ConversionError(f"AI request failed ({resp.status_code}): {resp.text[:200]}")
    data = resp.json()
    return "".join(b.get("text", "") for b in data.get("content", []) if b.get("type") == "text")


@app.after_request
def add_cors_headers(resp):
    resp.headers["Access-Control-Allow-Origin"] = ALLOWED_ORIGIN
    resp.headers["Access-Control-Allow-Methods"] = "POST, GET, OPTIONS"
    resp.headers["Access-Control-Allow-Headers"] = "Content-Type"
    return resp


@app.route("/api/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})


@app.route("/api/convert", methods=["POST", "OPTIONS"])
def convert_endpoint():
    if request.method == "OPTIONS":
        return "", 204

    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files["file"]
    from_fmt = (request.form.get("from") or "").lower().strip()
    to_fmt = (request.form.get("to") or "").lower().strip()
    style = (request.form.get("style") or "").lower().strip() or None

    if not from_fmt or not to_fmt:
        return jsonify({"error": "Missing 'from' or 'to' format"}), 400
    if file.filename == "":
        return jsonify({"error": "Empty filename"}), 400

    ext = file.filename.rsplit(".", 1)[-1].lower() if "." in file.filename else ""
    if ext != from_fmt:
        return jsonify({"error": f"File extension .{ext} doesn't match declared source format '{from_fmt}'"}), 400

    work_dir = tempfile.mkdtemp(prefix=f"mc_{uuid.uuid4().hex[:8]}_")
    try:
        src_path = os.path.join(work_dir, file.filename)
        file.save(src_path)

        try:
            result_path = convert(src_path, from_fmt, to_fmt, work_dir, style=style)
        except ConversionError as e:
            return jsonify({"error": str(e)}), 422
        except FileNotFoundError as e:
            return jsonify({"error": f"Required conversion tool missing on server: {e}"}), 500

        base_name = file.filename.rsplit(".", 1)[0]
        download_name = f"{base_name}.{to_fmt}"

        return send_file(
            result_path,
            mimetype=MIME_TYPES.get(to_fmt, "application/octet-stream"),
            as_attachment=True,
            download_name=download_name,
        )
    except Exception as e:
        return jsonify({"error": f"Conversion failed: {e}"}), 500
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


@app.route("/api/text-to-pptx", methods=["POST", "OPTIONS"])
def text_to_pptx_endpoint():
    if request.method == "OPTIONS":
        return "", 204

    data = request.get_json(silent=True) or {}
    text = (data.get("text") or "").strip()
    if not text:
        return jsonify({"error": "No text provided"}), 400
    if len(text) > 50000:
        return jsonify({"error": "Text too long (50,000 character limit)"}), 400

    work_dir = tempfile.mkdtemp(prefix=f"mc_txt_{uuid.uuid4().hex[:8]}_")
    try:
        result_path = text_to_pptx(text, work_dir)
        return send_file(
            result_path,
            mimetype=MIME_TYPES["pptx"],
            as_attachment=True,
            download_name="Presentation.pptx",
        )
    except ConversionError as e:
        return jsonify({"error": str(e)}), 422
    except Exception as e:
        return jsonify({"error": f"Conversion failed: {e}"}), 500
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


@app.route("/api/extract-text", methods=["POST", "OPTIONS"])
def extract_text_endpoint():
    if request.method == "OPTIONS":
        return "", 204
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "Empty filename"}), 400
    ext = file.filename.rsplit(".", 1)[-1].lower() if "." in file.filename else ""

    work_dir = tempfile.mkdtemp(prefix=f"mc_ext_{uuid.uuid4().hex[:8]}_")
    try:
        src_path = os.path.join(work_dir, file.filename)
        file.save(src_path)
        text = extract_text(src_path, ext)
        if len(text) > 20000:
            text = text[:20000]
        return jsonify({"text": text, "filename": file.filename})
    except ConversionError as e:
        return jsonify({"error": str(e)}), 422
    except Exception as e:
        return jsonify({"error": f"Couldn't read file: {e}"}), 500
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


@app.route("/api/summarize", methods=["POST", "OPTIONS"])
def summarize_endpoint():
    if request.method == "OPTIONS":
        return "", 204
    data = request.get_json(silent=True) or {}
    text = (data.get("text") or "").strip()
    if not text:
        return jsonify({"error": "No text provided"}), 400
    if len(text) > 20000:
        return jsonify({"error": "Text too long (20,000 character limit)"}), 400
    try:
        result = call_claude(
            system_prompt=(
                "You condense study notes into slide-ready bullet points for a tutoring app. "
                "Return 4 to 7 short bullet points capturing the material, one per line, no "
                "numbering, no markdown, no preamble or closing remarks — plain lines only."
            ),
            user_message=text,
            max_tokens=500,
        )
        bullets = [ln.strip(" -•\t") for ln in result.strip().split("\n") if ln.strip()]
        return jsonify({"bullets": bullets})
    except ConversionError as e:
        return jsonify({"error": str(e)}), 502


@app.route("/api/write", methods=["POST", "OPTIONS"])
def write_endpoint():
    if request.method == "OPTIONS":
        return "", 204
    data = request.get_json(silent=True) or {}
    topic = (data.get("topic") or "").strip()
    if not topic:
        return jsonify({"error": "No topic provided"}), 400
    if len(topic) > 2000:
        return jsonify({"error": "Topic too long (2,000 character limit)"}), 400
    try:
        result = call_claude(
            system_prompt=(
                "Write a clear explanatory paragraph, 120-180 words, on the given academic topic "
                "for a student. Do not invent citations, author names, journal names, or source "
                "titles under any circumstance — you cannot verify real sources exist, and a "
                "fabricated citation in a student's work is a serious problem. Write in plain "
                "prose with no citations or references at all."
            ),
            user_message=topic,
            max_tokens=450,
        )
        return jsonify({"text": result.strip()})
    except ConversionError as e:
        return jsonify({"error": str(e)}), 502


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    app.run(host="0.0.0.0", port=port)
