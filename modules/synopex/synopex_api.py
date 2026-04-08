import io
import os
import re
import shutil
import tempfile
import traceback
import zipfile
from pathlib import Path, PurePosixPath

from flask import Blueprint, jsonify, request, send_file

synopex_bp = Blueprint("synopex_bp", __name__)


def _safe_relative_path(name):
    raw = PurePosixPath((name or "").replace("\\", "/"))
    parts = [part for part in raw.parts if part not in ("", ".")]
    if raw.is_absolute() or any(part == ".." for part in parts):
        raise ValueError(f"Đường dẫn không hợp lệ: {name}")
    return Path(*parts)


def _extract_zip_safe(archive_path, destination):
    with zipfile.ZipFile(archive_path, "r") as archive:
        for member in archive.infolist():
            if member.is_dir():
                continue
            relative_path = _safe_relative_path(member.filename)
            if not relative_path.parts:
                continue
            target = Path(destination) / relative_path
            target.parent.mkdir(parents=True, exist_ok=True)
            with archive.open(member, "r") as source, open(target, "wb") as output:
                shutil.copyfileobj(source, output)


def _find_machine_root(root_dir):
    root_dir = Path(root_dir)
    direct_matches = [
        child for child in root_dir.iterdir()
        if child.is_dir() and re.match(r"^[Ss]\d+", child.name)
    ]
    if direct_matches:
        return root_dir

    candidates = []
    for current_root, dirnames, _ in os.walk(root_dir):
        matches = [name for name in dirnames if re.match(r"^[Ss]\d+", name)]
        if not matches:
            continue
        current_path = Path(current_root)
        depth = len(current_path.relative_to(root_dir).parts)
        candidates.append((depth, -len(matches), current_path))

    if not candidates:
        if re.match(r"^[Ss]\d+", root_dir.name):
            return root_dir
        return None

    candidates.sort()
    return candidates[0][2]


def _normalize_output_name(name):
    cleaned = re.sub(r'[<>:"/\\\\|?*]+', "_", (name or "").strip())
    if not cleaned:
        cleaned = "KEW_Synopex_Report.docx"
    if not cleaned.lower().endswith(".docx"):
        cleaned += ".docx"
    return cleaned


@synopex_bp.route("/generate", methods=["POST"])
def generate_report():
    data_zip = request.files.get("data_zip")
    if not data_zip or data_zip.filename == "":
        return jsonify({"error": "Cần upload file ZIP dữ liệu."}), 400

    work_dir = tempfile.mkdtemp(prefix="synopex_web_")
    try:
        upload_zip_path = os.path.join(work_dir, "upload.zip")
        input_dir = os.path.join(work_dir, "input")
        output_dir = os.path.join(work_dir, "output")
        os.makedirs(input_dir, exist_ok=True)
        os.makedirs(output_dir, exist_ok=True)
        data_zip.save(upload_zip_path)
        _extract_zip_safe(upload_zip_path, input_dir)

        machine_root = _find_machine_root(input_dir)
        if machine_root is None:
            return jsonify({"error": "Không tìm thấy thư mục dữ liệu dạng Sxxxx trong nguồn upload."}), 400

        output_name = _normalize_output_name(request.form.get("output_name"))
        output_path = os.path.join(output_dir, output_name)

        from generate_kew_synopex import build_synopex_report

        generated_path = build_synopex_report(
            base_dir=str(machine_root),
            output_file=output_path,
        )

        if not generated_path or not os.path.exists(generated_path):
            return jsonify({"error": "Không thể tạo báo cáo Synopex."}), 500

        with open(generated_path, "rb") as report_file:
            data = io.BytesIO(report_file.read())
        data.seek(0)
        return send_file(
            data,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            as_attachment=True,
            download_name=output_name,
        )
    except zipfile.BadZipFile:
        return jsonify({"error": "File ZIP dữ liệu không hợp lệ."}), 400
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception as exc:
        traceback.print_exc()
        return jsonify({"error": f"Lỗi tạo báo cáo: {exc}"}), 500
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)
