"""
ECONOMIC EVALUATION GROUP 2 (FIELD DEVELOPMENT) Web Application
Flask backend serving API endpoints for uploading Excel datasets,
running evaluations, streaming figures/results, and saving ZIP archives.
"""

import os
import sys
import shutil
import tempfile
import traceback
from pathlib import Path
from flask import Flask, request, jsonify, send_file, render_template, url_for

import economic_engine

app = Flask(__name__, static_folder="static", template_folder="templates")
app.config["MAX_CONTENT_LENGTH"] = 64 * 1024 * 1024  # 64 MB max upload size

# Store active runs in memory: run_id -> run_data
ACTIVE_RUNS = {}

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SAMPLE_TEMPLATE_PATH = os.path.join(BASE_DIR, "sample_template.xlsx")


def get_user_preset_paths():
    """Return common user directories on Windows for quick folder selection."""
    home = Path.home()
    paths = {
        "Desktop": str(home / "Desktop"),
        "Downloads": str(home / "Downloads"),
        "Documents": str(home / "Documents"),
        "Project Folder": BASE_DIR
    }
    # Filter to existing
    return {k: v for k, v in paths.items() if os.path.exists(v)}


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/preset-paths", methods=["GET"])
def preset_paths():
    return jsonify(get_user_preset_paths())


@app.route("/api/download-template", methods=["GET"])
def download_template():
    if not os.path.exists(SAMPLE_TEMPLATE_PATH):
        return jsonify({"error": "Sample template not found"}), 404
    return send_file(
        SAMPLE_TEMPLATE_PATH,
        as_attachment=True,
        download_name="Economics_Indicators_Template.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.route("/api/run-sample", methods=["POST"])
def run_sample():
    try:
        if not os.path.exists(SAMPLE_TEMPLATE_PATH):
            return jsonify({"error": "Sample template file not found on server"}), 404
            
        data = request.get_json(silent=True) or {}
        user_config = {}
        if "discount_rate" in data and data["discount_rate"] is not None:
            user_config["DISCOUNT_RATE_OVERRIDE"] = float(data["discount_rate"]) / 100.0
        if "capital_budget" in data and data["capital_budget"]:
            user_config["CAPITAL_BUDGET"] = float(data["capital_budget"])
        if "economic_cutoff" in data:
            user_config["ECONOMIC_LIMIT_CUTOFF"] = bool(data["economic_cutoff"])
        if "price_escalation" in data:
            user_config["APPLY_PRICE_ESCALATION"] = bool(data["price_escalation"])
            
        result = economic_engine.run_full_analysis(
            SAMPLE_TEMPLATE_PATH,
            user_config=user_config
        )
        ACTIVE_RUNS[result["run_id"]] = result
        return jsonify(result)
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500


@app.route("/api/upload-and-run", methods=["POST"])
def upload_and_run():
    try:
        if "file" not in request.files:
            return jsonify({"error": "No file uploaded in request."}), 400
            
        file = request.files["file"]
        if file.filename == "":
            return jsonify({"error": "Selected file has an empty filename."}), 400
            
        filename = file.filename
        ext = os.path.splitext(filename)[1].lower()
        if ext not in [".xlsx", ".xls", ".xlsm"]:
            return jsonify({"error": f"Invalid file format '{ext}'. Please upload an Excel workbook (.xlsx, .xls, .xlsm)."}), 400
            
        # Save temporary file
        temp_dir = tempfile.mkdtemp(prefix="econ_upload_")
        temp_file_path = os.path.join(temp_dir, filename)
        file.save(temp_file_path)
        
        # Parse optional user configuration overrides
        user_config = {}
        if "discount_rate" in request.form and request.form["discount_rate"]:
            try:
                user_config["DISCOUNT_RATE_OVERRIDE"] = float(request.form["discount_rate"]) / 100.0
            except ValueError:
                pass
        if "capital_budget" in request.form and request.form["capital_budget"]:
            try:
                user_config["CAPITAL_BUDGET"] = float(request.form["capital_budget"])
            except ValueError:
                pass
        if "economic_cutoff" in request.form:
            user_config["ECONOMIC_LIMIT_CUTOFF"] = request.form["economic_cutoff"].lower() in ["true", "1", "yes"]
        if "price_escalation" in request.form:
            user_config["APPLY_PRICE_ESCALATION"] = request.form["price_escalation"].lower() in ["true", "1", "yes"]
            
        result = economic_engine.run_full_analysis(
            temp_file_path,
            user_config=user_config
        )
        ACTIVE_RUNS[result["run_id"]] = result
        return jsonify(result)
        
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500


@app.route("/api/download-zip/<run_id>", methods=["GET"])
def download_zip(run_id):
    run_data = ACTIVE_RUNS.get(run_id)
    if not run_data or not os.path.exists(run_data["zip_filepath"]):
        return jsonify({"error": f"Run ID '{run_id}' not found or zip file missing."}), 404
        
    return send_file(
        run_data["zip_filepath"],
        as_attachment=True,
        download_name=run_data["zip_filename"],
        mimetype="application/zip"
    )


@app.route("/api/save-zip-to-folder", methods=["POST"])
def save_zip_to_folder():
    """Saves the generated ZIP file directly to a designated directory on the user's machine."""
    try:
        data = request.get_json(force=True)
        run_id = data.get("run_id")
        target_dir = data.get("target_dir", "").strip().strip('"').strip("'")
        custom_name = data.get("custom_name", "").strip()
        
        if not run_id or run_id not in ACTIVE_RUNS:
            return jsonify({"error": "Invalid or expired run session."}), 400
            
        if not target_dir:
            return jsonify({"error": "Please specify a target destination folder."}), 400
            
        run_data = ACTIVE_RUNS[run_id]
        src_zip = run_data["zip_filepath"]
        
        if not os.path.exists(src_zip):
            return jsonify({"error": "Source zip file not found on server."}), 404
            
        target_dir_expanded = os.path.expanduser(target_dir)
        os.makedirs(target_dir_expanded, exist_ok=True)
        
        filename = custom_name if custom_name else run_data["zip_filename"]
        if not filename.lower().endswith(".zip"):
            filename += ".zip"
            
        dest_path = os.path.join(target_dir_expanded, filename)
        shutil.copy2(src_zip, dest_path)
        
        return jsonify({
            "success": True,
            "message": f"Successfully saved ZIP archive to selected location.",
            "saved_path": os.path.abspath(dest_path),
            "size_bytes": os.path.getsize(dest_path)
        })
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/api/figures/<run_id>/<fig_name>", methods=["GET"])
def get_figure(run_id, fig_name):
    run_data = ACTIVE_RUNS.get(run_id)
    if not run_data:
        return "Run not found", 404
        
    fig_path = os.path.join(run_data["output_dir"], "figures", fig_name)
    if not os.path.exists(fig_path):
        return "Figure not found", 404
        
    return send_file(fig_path, mimetype="image/png")


@app.route("/api/download-results-excel/<run_id>", methods=["GET"])
def download_results_excel(run_id):
    run_data = ACTIVE_RUNS.get(run_id)
    if not run_data:
        return "Run not found", 404
        
    excel_path = os.path.join(run_data["output_dir"], "Economic_Indicators_Results.xlsx")
    if not os.path.exists(excel_path):
        return "Results Excel file not found", 404
        
    return send_file(
        excel_path,
        as_attachment=True,
        download_name="Economic_Indicators_Results.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print(f"Starting ECONOMIC EVALUATION GROUP 2 (FIELD DEVELOPMENT) Web App on http://localhost:{port}")
    app.run(host="0.0.0.0", port=port, debug=False)
