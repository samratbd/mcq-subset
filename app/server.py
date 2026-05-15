"""Flask application — endpoints + the index page.

Endpoints:
    POST   /upload                multipart upload -> {paper_id, name, n_questions, has_katex}
    POST   /generate              multipart: paper_id + optional header_image -> ZIP of sets
    GET    /papers                list saved papers (persistence only)
    GET    /papers/<id>/sets      list previously-generated sets
    DELETE /papers/<id>           delete a saved paper
    GET    /samples/<filename>    download a sample/template file
    GET    /samples               JSON manifest of available samples
    GET    /health                liveness + feature flags

The server holds an in-memory cache so that one-shot uploads (persistence off)
don't have to be stored to disk just to be downloaded a moment later.
"""

from __future__ import annotations
import io
import os
import re
import time
import uuid
import zipfile
from typing import Optional

from flask import Flask, jsonify, render_template, request, send_file, abort

from .parsers import parse_upload
from .writers import write_set
from .writers.pdf_writer import libreoffice_available
from .shuffler import make_set, verify_set
from .math_utils import has_katex, pandoc_available
from .db import Store
from .samples import build_sample_paper, sample_manifest, write_sample_to_bytes


# --- app & storage -----------------------------------------------------------

def _available_ram_mb() -> int:
    """Return approximate total system RAM in MB (cross-platform)."""
    try:
        # Linux: read /proc/meminfo
        with open("/proc/meminfo") as f:
            for line in f:
                if line.startswith("MemTotal:"):
                    return int(line.split()[1]) // 1024
    except Exception:
        pass
    try:
        import os
        return os.sysconf("SC_PHYS_PAGES") * os.sysconf("SC_PAGE_SIZE") // (1024 * 1024)
    except Exception:
        return 2048  # assume 2 GB if unknown


def create_app() -> Flask:
    app = Flask(
        __name__,
        template_folder=os.path.join(os.path.dirname(__file__), "..", "templates"),
        static_folder=os.path.join(os.path.dirname(__file__), "..", "static"),
    )

    # 200 MB — generous enough for batch OMR uploads (each scan ~200-400 KB,
    # so this comfortably fits 500+ sheets in one request).
    app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024

    @app.errorhandler(413)
    def _too_large(_e):
        return jsonify(
            error=(
                "Upload exceeded the 200 MB size limit. "
                "Try uploading fewer sheets at a time, or contact the admin "
                "to raise MAX_CONTENT_LENGTH in app/server.py."
            )
        ), 413

    @app.errorhandler(500)
    def _server_error(e):
        # Without this, Flask returns a generic HTML page. With it, the
        # frontend gets a JSON body it can show inline.
        import traceback
        app.logger.error("500 error: %s\n%s", e, traceback.format_exc())
        return jsonify(
            error=f"Server error: {type(e).__name__}: {e}",
        ), 500

    db_path = os.path.join(os.path.dirname(__file__), "..", "data", "store.sqlite3")
    store = Store(db_path)

    # In-memory cache for one-shot uploads (persist=false): paper_id -> (name, filename, blob)
    one_shot: dict[str, tuple[str, str, bytes]] = {}

    # --- helpers -------------------------------------------------------------

    def _resolve_paper(paper_id: str):
        """Look up a paper in either the in-memory cache or the DB. Returns
        (name, filename, blob) or None."""
        if paper_id in one_shot:
            return one_shot[paper_id]
        row = store.get_paper(paper_id)
        return row

    def _safe_filename(name: str) -> str:
        # Replace anything that's not alnum/dash/underscore/dot with underscore.
        s = re.sub(r"[^A-Za-z0-9._-]+", "_", (name or "paper").strip())
        return s.strip("._-") or "paper"

    # --- routes --------------------------------------------------------------

    @app.get("/")
    def index():
        return render_template(
            "index.html",
            pandoc_available=pandoc_available(),
            libreoffice_available=libreoffice_available(),
            has_default_header=os.path.exists(
                os.path.join(os.path.dirname(__file__), "..",
                             "static", "assets", "default_header.jpg")
            ),
        )

    @app.get("/health")
    def health():
        return jsonify(
            ok=True,
            pandoc=pandoc_available(),
            libreoffice=libreoffice_available(),
        )

    @app.post("/upload")
    def upload():
        f = request.files.get("file")
        if not f:
            return jsonify(error="No file uploaded."), 400
        persist = (request.form.get("persist", "false").lower() == "true")
        display_name = (request.form.get("name", "") or "").strip() \
            or os.path.splitext(f.filename or "paper")[0]

        try:
            data = f.read()
            questions = parse_upload(f.filename or "", data)
        except ValueError as e:
            # Expected: layout couldn't be recognised, structural issues, etc.
            return jsonify(error=str(e)), 400
        except Exception as e:
            # Unexpected: log the full traceback so we can debug from logs.
            import traceback
            tb = traceback.format_exc()
            app.logger.error("upload parse failed:\n%s", tb)
            return jsonify(
                error=f"{type(e).__name__}: {e}",
                detail="See server logs for the full traceback.",
            ), 400

        any_katex = any(
            has_katex(q.question) or has_katex(q.explanation)
            or any(has_katex(o) for o in q.options)
            for q in questions
        )

        if persist:
            paper_id = store.add_paper(display_name, f.filename or "upload", data)
        else:
            paper_id = uuid.uuid4().hex
            one_shot[paper_id] = (display_name, f.filename or "upload", data)

        return jsonify(
            paper_id=paper_id,
            name=display_name,
            n_questions=len(questions),
            has_katex=any_katex,
            persisted=persist,
        )

    @app.post("/generate")
    def generate():
        # We accept BOTH JSON (legacy) and multipart/form-data (new — required
        # for uploading the header image). For multipart, all params come from
        # request.form and the image from request.files["header_image"].
        if request.content_type and request.content_type.startswith(
                "multipart/form-data"):
            body = request.form
            get = body.get
            header_file = request.files.get("header_image")
        else:
            body = request.get_json(silent=True) or {}
            get = body.get
            header_file = None

        def _bool(v, default=False):
            if v is None:
                return default
            return str(v).lower() in ("true", "1", "yes", "on")

        paper_id = get("paper_id")
        n_sets = int(get("n_sets", 1))
        shuffle_q = _bool(get("shuffle_questions"), True)
        shuffle_o = _bool(get("shuffle_options"), True)
        fmt = (get("format") or "csv").lower()
        persist = _bool(get("persist"), False)
        math_in_docx = (get("math_in_docx") or "equation").lower()
        math_in_data = (get("math_in_data") or "katex").lower()

        # Header image options:
        #   header_mode = "none"    → no header
        #   header_mode = "default" → use static/assets/default_header.jpg
        #   header_mode = "custom"  → use uploaded bytes from header_image file
        header_mode = (get("header_mode") or "none").lower()
        header_bytes: Optional[bytes] = None
        if header_mode == "default":
            default_path = os.path.join(
                os.path.dirname(__file__), "..",
                "static", "assets", "default_header.jpg",
            )
            if os.path.exists(default_path):
                with open(default_path, "rb") as f:
                    header_bytes = f.read()
        elif header_mode == "custom":
            if header_file:
                header_bytes = header_file.read()
                if not header_bytes:
                    header_bytes = None

        if not paper_id:
            return jsonify(error="paper_id is required"), 400
        if not (1 <= n_sets <= 20):
            return jsonify(error="n_sets must be between 1 and 20"), 400
        if math_in_docx not in ("equation", "text", "unicode"):
            return jsonify(error=f"bad math_in_docx: {math_in_docx!r}"), 400
        if math_in_data not in ("katex", "unicode"):
            return jsonify(error=f"bad math_in_data: {math_in_data!r}"), 400
        if fmt.startswith("pdf_") and not libreoffice_available():
            return jsonify(
                error="PDF output requires LibreOffice on the server. "
                      "It isn't installed in this environment. Use the "
                      "Word format instead, or install LibreOffice "
                      "(`apt install libreoffice`) and restart."
            ), 400

        record = _resolve_paper(paper_id)
        if not record:
            return jsonify(error="Paper not found. Re-upload or check the ID."), 404
        display_name, source_filename, blob = record

        try:
            questions = parse_upload(source_filename or "", blob)
        except Exception as e:
            return jsonify(error=f"Re-parse failed: {e}"), 500

        safe_base = _safe_filename(display_name)
        buf = io.BytesIO()
        integrity_lines = []

        manifest_lines = [
            f"Paper: {display_name}",
            f"Source: {source_filename}",
            f"Questions: {len(questions)}",
            f"Sets generated: {n_sets}",
            f"Shuffle questions: {shuffle_q}",
            f"Shuffle options: {shuffle_o}",
            f"Output format: {fmt}",
            f"Math in Word/PDF (KaTeX → ...): {math_in_docx}",
            f"Math in CSV/XLSX: {math_in_data}",
            f"Header image: {header_mode}",
            f"Generated at: {time.strftime('%Y-%m-%d %H:%M:%S')}",
            "",
            "Reproducibility: shuffles are seeded by (paper_id, set_number, mode);",
            "the same source + set number always produces the same paper.",
            "",
            "Integrity:",
            "  Each set is verified BEFORE writing — the option text at the",
            "  shuffled answer position must equal the original correct option",
            "  text, all questions must appear exactly once with their original",
            "  options intact (as a multiset), and SLs must be 1..N.",
            "",
        ]

        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for n in range(1, n_sets + 1):
                shuffled = make_set(
                    questions,
                    paper_id=paper_id,
                    set_number=n,
                    shuffle_questions=shuffle_q,
                    shuffle_options=shuffle_o,
                )
                try:
                    verify_set(questions, shuffled)
                    integrity_lines.append(f"  Set {n:02d}: OK ({len(shuffled)} Qs)")
                except AssertionError as e:
                    return jsonify(
                        error=f"Integrity check failed for set {n}: {e}"
                    ), 500

                set_title = f"{display_name} — Set {n}"
                try:
                    data, ext = write_set(
                        shuffled, fmt, title=set_title,
                        math_in_docx=math_in_docx,
                        math_in_data=math_in_data,
                        header_image=header_bytes,
                    )
                except Exception as e:
                    import traceback
                    app.logger.error("set %d writer error:\n%s",
                                     n, traceback.format_exc())
                    return jsonify(error=f"Set {n} writer error: {e}"), 500

                fname = f"{safe_base}_Set{n:02d}.{ext}"
                zf.writestr(fname, data)

                if persist:
                    store.record_set(paper_id, n, shuffle_q, shuffle_o)

            zf.writestr(
                "MANIFEST.txt",
                "\n".join(manifest_lines + integrity_lines)
            )

        buf.seek(0)
        download_name = f"{safe_base}_sets.zip"
        return send_file(
            buf,
            mimetype="application/zip",
            as_attachment=True,
            download_name=download_name,
        )

    @app.get("/papers")
    def list_papers():
        return jsonify(papers=store.list_papers())

    @app.get("/papers/<paper_id>/sets")
    def list_sets(paper_id: str):
        return jsonify(sets=store.list_sets(paper_id))

    @app.delete("/papers/<paper_id>")
    def delete_paper(paper_id: str):
        ok = store.delete_paper(paper_id)
        if not ok:
            return jsonify(error="Paper not found."), 404
        return jsonify(deleted=True)

    @app.get("/samples")
    def list_samples():
        """JSON manifest of available template/sample files."""
        return jsonify(samples=sample_manifest())

    @app.get("/samples/<filename>")
    def download_sample(filename: str):
        """Serve a generated sample file with the right MIME type."""
        # Sanitise to avoid path traversal
        safe = re.sub(r"[^A-Za-z0-9._-]+", "", filename)
        if safe != filename:
            return jsonify(error="Invalid filename"), 400
        try:
            data, mimetype = write_sample_to_bytes(filename)
        except KeyError:
            return jsonify(error=f"Unknown sample: {filename!r}"), 404
        except Exception as e:
            return jsonify(error=f"Sample generation failed: {e}"), 500
        return send_file(
            io.BytesIO(data),
            mimetype=mimetype,
            as_attachment=True,
            download_name=filename,
        )

    # --- OMR scanner ---------------------------------------------------------

    @app.post("/omr/scan")
    def omr_scan():
        """Scan one or more uploaded OMR sheet images.

        Multipart form fields:
            sheet_type   : 'auto' (default), 'omr_50', 'omr_100'
            output_format: 'xlsx' (default), 'csv', 'json'
            include_review_images: 'true' to include per-sheet annotated PNGs
            files        : one or more image files (BMP/PNG/JPEG)

        Returns a ZIP containing the results file + (optionally) review images.
        """
        import traceback

        try:
            from .omr import (
                scan_omr, render_review_image,
                write_csv as omr_write_csv,
                write_xlsx as omr_write_xlsx,
                write_json as omr_write_json,
                TEMPLATES as OMR_TEMPLATES,
            )
        except ImportError as e:
            # OpenCV or similar missing — give the user an actionable message
            app.logger.error("OMR module import failed: %s", e)
            return jsonify(
                error=(
                    "OMR scanner is not available because a required library "
                    f"is missing: {e}. On Linux: "
                    "'sudo apt install libgl1 libglib2.0-0' then "
                    "'pip install -r requirements.txt'."
                ),
            ), 500

        files = request.files.getlist("files")
        if not files:
            return jsonify(error="No image files uploaded."), 400

        sheet_type = (request.form.get("sheet_type") or "auto").lower()
        output_format = (request.form.get("output_format") or "xlsx").lower()
        include_review = (
            request.form.get("include_review_images", "false").lower()
            in ("true", "1", "yes", "on")
        )

        if sheet_type not in ("auto", "omr_50", "omr_100"):
            return jsonify(
                error=f"bad sheet_type: {sheet_type!r} "
                      "(must be 'auto', 'omr_50', or 'omr_100')"
            ), 400
        if output_format not in ("csv", "xlsx", "json"):
            return jsonify(
                error=f"bad output_format: {output_format!r}"
            ), 400

        try:
            # Read all file bytes first (must do this in the request thread).
            files_data: list = []
            for f in files:
                fname = f.filename or "unknown"
                try:
                    data = f.read()
                except Exception as e:
                    app.logger.warning("Read failed for %s: %s", fname, e)
                    from .omr.scanner import OmrResult
                    files_data.append((
                        fname, None,
                        OmrResult(
                            sheet_type=("omr_50" if sheet_type == "auto"
                                        else sheet_type),
                            roll_number="?", set_letter="?",
                            answers=[], confidence=0.0,
                            needs_review=True,
                            review_items=["read_failed"],
                            fill_fractions=[],
                            error=f"Could not read uploaded file: {e}",
                        ),
                    ))
                    continue
                if not data:
                    app.logger.warning("Empty file: %s", fname)
                    continue
                files_data.append((fname, data, None))

            # Scan all files. On low-memory servers (≤1 GB) we process ONE
            # sheet at a time and immediately discard the image from memory.
            # On multi-CPU servers we use a small thread pool.
            import os, gc
            total_ram_mb = _available_ram_mb()
            cpus = os.cpu_count() or 1
            # Use parallel only when we have both multiple cores AND enough RAM
            use_parallel = cpus >= 2 and total_ram_mb >= 1800
            max_workers = min(4, cpus) if use_parallel else 1

            def _scan_one(item):
                fname, data, pre_error = item
                if pre_error is not None:
                    return (pre_error, fname, None)
                try:
                    res = scan_omr(data, sheet_type=sheet_type)
                except Exception as e:
                    app.logger.error(
                        "Scan failed for %s: %s\n%s",
                        fname, e, traceback.format_exc(),
                    )
                    from .omr.scanner import OmrResult
                    res = OmrResult(
                        sheet_type=("omr_50" if sheet_type == "auto"
                                    else sheet_type),
                        roll_number="?", set_letter="?",
                        answers=[], confidence=0.0,
                        needs_review=True,
                        review_items=["scan_exception"],
                        fill_fractions=[],
                        error=f"{type(e).__name__}: {e}",
                    )
                review_png = None
                if include_review and not res.error:
                    try:
                        review_png = render_review_image(data, res)
                    except Exception as e:
                        app.logger.warning(
                            "Review image failed for %s: %s", fname, e
                        )
                return (res, fname, review_png)

            results: list = []
            per_sheet_review_imgs: list = []

            if use_parallel:
                from concurrent.futures import ThreadPoolExecutor
                with ThreadPoolExecutor(max_workers=max_workers) as pool:
                    for res, fname, review_png in pool.map(_scan_one, files_data):
                        results.append((res, fname))
                        if review_png:
                            per_sheet_review_imgs.append(
                                ((fname or "sheet") + "_review.png", review_png)
                            )
            else:
                # Low-memory serial path: process one sheet, free memory, repeat
                for item in files_data:
                    res, fname, review_png = _scan_one(item)
                    results.append((res, fname))
                    if review_png:
                        per_sheet_review_imgs.append(
                            ((fname or "sheet") + "_review.png", review_png)
                        )
                    gc.collect()  # release image data promptly

            if not results:
                return jsonify(
                    error="No sheets could be read from the upload."
                ), 400

            # All output rows share a column count — use the max question
            # count across all sheets (so a 100-mark sheet mixed in with
            # 50-mark sheets doesn't truncate).
            n_questions = max(
                OMR_TEMPLATES[r.sheet_type].n_questions
                for r, _ in results
                if r.sheet_type in OMR_TEMPLATES
            )

            if output_format == "csv":
                data_bytes = omr_write_csv(results, n_questions)
                data_name = "omr_results.csv"
            elif output_format == "json":
                data_bytes = omr_write_json(results, n_questions)
                data_name = "omr_results.json"
            else:
                data_bytes = omr_write_xlsx(results, n_questions)
                data_name = "omr_results.xlsx"

            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
                zf.writestr(data_name, data_bytes)

                ok_count = sum(1 for r, _ in results if not r.error)
                review_count = sum(1 for r, _ in results if r.needs_review)
                avg_conf = (
                    sum(r.confidence for r, _ in results) / len(results) * 100
                ) if results else 0.0
                summary = [
                    f"Sheets uploaded: {len(files)}",
                    f"Sheets scanned:  {len(results)} "
                    f"({ok_count} OK, {len(results) - ok_count} failed)",
                    f"Sheets needing review: {review_count}",
                    f"Average confidence: {avg_conf:.1f}%",
                    f"Sheet type setting: {sheet_type}",
                    "",
                    "Per-sheet summary:",
                ]
                for serial, (res, src) in enumerate(results, start=1):
                    safe_src = src.encode("ascii", "replace").decode()
                    line = (
                        f"  {serial:3d}. {safe_src:40.40s} "
                        f"roll={res.roll_number}  set={res.set_letter}  "
                        f"conf={res.confidence * 100:5.1f}%  "
                        f"review={'YES' if res.needs_review else 'no'}"
                    )
                    if res.error:
                        line += f"  ERROR: {res.error}"
                    summary.append(line)
                zf.writestr("SUMMARY.txt", "\n".join(summary))

                for fname, img_bytes in per_sheet_review_imgs:
                    if img_bytes:
                        safe_fname = re.sub(
                            r"[^A-Za-z0-9._-]+", "_", fname
                        )
                        zf.writestr(f"review/{safe_fname}", img_bytes)

            buf.seek(0)
            return send_file(
                buf,
                mimetype="application/zip",
                as_attachment=True,
                download_name="omr_results.zip",
            )

        except Exception as e:
            # Last-resort catch — never return a 500 without explanation
            app.logger.error(
                "Unhandled OMR scan error: %s\n%s", e, traceback.format_exc()
            )
            return jsonify(
                error=(
                    f"OMR scan failed unexpectedly: {type(e).__name__}: {e}. "
                    "Please check the server log for details."
                ),
                traceback=traceback.format_exc().splitlines()[-3:],
            ), 500

    @app.get("/omr/health")
    def omr_health():
        from .omr import TEMPLATES as OMR_TEMPLATES
        try:
            import cv2  # noqa
            cv_ok = True
        except Exception:
            cv_ok = False
        return jsonify(
            ok=cv_ok,
            opencv=cv_ok,
            sheet_types=sorted(OMR_TEMPLATES.keys()),
        )

    return app
