import os
import time
from uuid import uuid4
from flask import (
    Blueprint, current_app, render_template, request,
    redirect, url_for, send_from_directory, flash, jsonify
)
from werkzeug.utils import secure_filename
from .services import DocumentComparator

bp = Blueprint("main", __name__)

ALLOWED_EXTENSIONS = {".xlsx", ".xls", ".docx", ".doc", ".pdf"}


def allowed_file(filename):
    _, ext = os.path.splitext(filename.lower())
    return ext in ALLOWED_EXTENSIONS


@bp.route("/", methods=["GET"])
def home():
    return render_template("home.html")


@bp.route("/compare", methods=["GET"])
def compare_page():
    return render_template("compare.html")


@bp.route("/compare/process", methods=["POST"])
def process_comparison():
    """문서 비교 처리"""
    original_file = request.files.get("original_file")
    revised_file = request.files.get("revised_file")

    # 파일 검증
    if not original_file or original_file.filename == "":
        flash("원본 파일을 업로드하세요.")
        return redirect(url_for("main.compare_page"))

    if not revised_file or revised_file.filename == "":
        flash("수정본 파일을 업로드하세요.")
        return redirect(url_for("main.compare_page"))

    if not allowed_file(original_file.filename):
        flash("지원하지 않는 원본 파일 형식입니다.")
        return redirect(url_for("main.compare_page"))

    if not allowed_file(revised_file.filename):
        flash("지원하지 않는 수정본 파일 형식입니다.")
        return redirect(url_for("main.compare_page"))

    # 작업 ID 생성 및 파일 저장
    job_id = f"{int(time.time())}_{uuid4().hex[:8]}"

    original_name = secure_filename(f"{job_id}_original_" + original_file.filename)
    revised_name = secure_filename(f"{job_id}_revised_" + revised_file.filename)

    original_path = os.path.join(current_app.config["UPLOAD_ORIGINAL_DIR"], original_name)
    revised_path = os.path.join(current_app.config["UPLOAD_REVISED_DIR"], revised_name)

    original_file.save(original_path)
    revised_file.save(revised_path)

    # 비교 옵션 가져오기
    options = {
        "ignore_case": request.form.get("ignore_case") == "on",
        "ignore_whitespace": request.form.get("ignore_whitespace") == "on",
        "show_context": request.form.get("show_context") == "on",
        "context_lines": int(request.form.get("context_lines", 2))
    }

    try:
        # 문서 비교 실행
        comparator = DocumentComparator()
        results = comparator.compare(original_path, revised_path, options)

        # 결과 저장
        output_name = f"comparison_{job_id}.html"
        output_path = os.path.join(current_app.config["OUTPUT_DIR"], output_name)

        # HTML 리포트 생성
        comparator.save_html_report(results, output_path)

        # JSON 형식 결과도 저장
        json_name = f"comparison_{job_id}.json"
        json_path = os.path.join(current_app.config["OUTPUT_DIR"], json_name)
        comparator.save_json_results(results, json_path)

        current_app.logger.info(f"Comparison completed: {job_id}")

        return render_template(
            "results.html",
            results=results,
            output_html=output_name,
            output_json=json_name,
            job_id=job_id
        )

    except Exception as e:
        current_app.logger.exception(e)
        flash(f"비교 작업 중 오류가 발생했습니다: {e}")
        return redirect(url_for("main.compare_page"))


@bp.route("/api/compare", methods=["POST"])
def api_compare():
    """API 엔드포인트 - JSON 결과 반환"""
    original_file = request.files.get("original")
    revised_file = request.files.get("revised")

    if not original_file or not revised_file:
        return jsonify({"error": "Both files are required"}), 400

    job_id = f"{int(time.time())}_{uuid4().hex[:8]}"

    # 임시 파일 저장
    original_path = f"/tmp/{job_id}_original"
    revised_path = f"/tmp/{job_id}_revised"

    original_file.save(original_path)
    revised_file.save(revised_path)

    try:
        comparator = DocumentComparator()
        results = comparator.compare(original_path, revised_path)

        # 임시 파일 정리
        os.remove(original_path)
        os.remove(revised_path)

        return jsonify(results)

    except Exception as e:
        return jsonify({"error": str(e)}), 500


@bp.route("/download/<path:filename>")
def download_file(filename):
    return send_from_directory(current_app.config["OUTPUT_DIR"], filename, as_attachment=True)


@bp.route('/shutdown')
def shutdown():
    """서버 종료"""
    shutdown_func = request.environ.get('werkzeug.server.shutdown')
    if shutdown_func:
        print("서버 종료 중...")
        shutdown_func()
        return "서버 종료 중..."
    print("강제 종료 실행")
    os._exit(0)