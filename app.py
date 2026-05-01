
import io
import json
import os
import re
import time
import queue

import threading
from pathlib import Path
from collections import OrderedDict

import fitz
from flask import (
    Flask, render_template, request, Response, jsonify, send_file,
)
from db import get_setting, set_setting

app = Flask(__name__)
app.secret_key = os.urandom(32)
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024

_progress_queues: dict[str, queue.Queue] = {}
_progress_lock = threading.Lock()


def _new_progress_queue(task_id: str) -> queue.Queue:
    q: queue.Queue = queue.Queue()
    with _progress_lock:
        _progress_queues[task_id] = q
    return q


def _get_progress_queue(task_id: str) -> queue.Queue | None:
    with _progress_lock:
        return _progress_queues.get(task_id)


def _remove_progress_queue(task_id: str):
    with _progress_lock:
        _progress_queues.pop(task_id, None)


def _emit(q: queue.Queue, *, stage: str, pct: int, detail: str,
          eta_sec: float | None = None, done: bool = False,
          error: str | None = None, result: dict | None = None):
    payload = {
        "stage": stage,
        "pct": pct,
        "detail": detail,
        "done": done,
    }
    if eta_sec is not None:
        payload["eta_sec"] = round(eta_sec, 1)
    if error:
        payload["error"] = error
    if result:
        payload["result"] = result
    q.put(payload)


VARIANT_ORDER = {"4x3": 0, "3x2": 1, "3x3": 2, "4x6": 3}
VARIANT_RE = re.compile(r"(3x2|3x3|4x6)", re.IGNORECASE)
FOUR_DIGIT_RE = re.compile(r"-(\d{4})")
COMPLEX_INDICATORS = re.compile(r"[&]|1_2|2_2|\bset\b", re.IGNORECASE)

COMBINED_RE = re.compile(r"combined", re.IGNORECASE)

PREFIX_RE = re.compile(r"^\d{2,}_")


def extract_four_digit_id(filename: str) -> str | None:
    m = FOUR_DIGIT_RE.search(filename)
    return m.group(1) if m else None


def extract_all_four_digit_ids(filename: str) -> list[str]:
    return FOUR_DIGIT_RE.findall(filename)


def detect_variant(filename: str) -> str:
    m = VARIANT_RE.search(filename)
    return m.group(1).lower() if m else "4x3"


def is_complex(filename: str) -> bool:
    return bool(COMPLEX_INDICATORS.search(Path(filename).stem))


def sort_key(filename: str):
    variant = detect_variant(filename)
    complexity = 1 if is_complex(filename) else 0
    alpha = Path(filename).stem.lower()
    return (VARIANT_ORDER.get(variant, 99), complexity, alpha)


def _strip_prefix(filename: str) -> str:
    return PREFIX_RE.sub("", filename)


def _is_combined(filename: str) -> bool:
    return bool(COMBINED_RE.search(filename))


def _run_module1(folder_path: str, task_id: str):
    q = _get_progress_queue(task_id)
    if q is None:
        return

    t0 = time.perf_counter()

    try:
        folder = Path(folder_path)
        if not folder.is_dir():
            _emit(q, stage="error", pct=0, detail="", done=True,
                  error=f"Folder not found: {folder_path}")
            return

        _emit(q, stage="cleanup", pct=0,
              detail="Deleting old combined PDFs and ZIP files...")

        for f in folder.iterdir():
            if not f.is_file():
                continue
            if f.suffix.lower() == ".pdf" and _is_combined(f.stem):
                f.unlink()
            

        _emit(q, stage="cleanup", pct=5, detail="Old files cleaned up.")

        _emit(q, stage="scan", pct=5, detail="Scanning folder for PDFs...")

        pdf_files = []
        for f in folder.iterdir():
            if not f.is_file():
                continue
            if f.suffix.lower() != ".pdf":
                continue
            if _is_combined(f.stem):
                continue
            clean_name = _strip_prefix(f.name)
            if clean_name != f.name:
                new_path = folder / clean_name
                if new_path.exists() and new_path != f:
                    pdf_files.append(f.name)
                else:
                    f.rename(new_path)
                    pdf_files.append(clean_name)
            else:
                pdf_files.append(f.name)

        if not pdf_files:
            _emit(q, stage="error", pct=0, detail="", done=True,
                  error="No PDF files found (after excluding 'combined').")
            return

        _emit(q, stage="scan", pct=10,
              detail=f"Found {len(pdf_files)} PDFs (excluded 'combined').")

        _emit(q, stage="sort", pct=10, detail="Sorting PDFs...")
        pdf_files.sort(key=sort_key)

        sorted_list = []
        for idx, fname in enumerate(pdf_files, start=1):
            prefix = f"{idx:02d}_"
            sorted_list.append({
                "original": fname,
                "new_name": f"{prefix}{fname}",
                "four_digit_id": extract_four_digit_id(fname),
                "variant": detect_variant(fname),
            })

        _emit(q, stage="sort", pct=15,
              detail=f"Sorted {len(sorted_list)} PDFs.")

        _emit(q, stage="rename", pct=15, detail="Renaming files...")
        total = len(sorted_list)

        temp_map: dict[str, str] = {}
        for item in sorted_list:
            src = folder / item["original"]
            if src.exists():
                tmp_name = f"__tmp_{item['new_name']}"
                src.rename(folder / tmp_name)
                temp_map[tmp_name] = item["new_name"]

        renamed_count = 0
        for tmp_name, final_name in temp_map.items():
            (folder / tmp_name).rename(folder / final_name)
            renamed_count += 1
            pct = 15 + int((renamed_count / total) * 15)
            _emit(q, stage="rename", pct=pct,
                  detail=f"Renamed {renamed_count}/{total}")

        _emit(q, stage="rename", pct=30,
              detail=f"All {total} files renamed.")

        _emit(q, stage="combine", pct=30,
              detail="Combining PDFs (lossless)...")

        non_4x3_count = sum(1 for item in sorted_list if item["variant"] != "4x3")
        combined_name = f"{total}-combined-{non_4x3_count}.pdf"
        combined_path = folder / combined_name
        combined_doc = fitz.open()

        for i, item in enumerate(sorted_list):
            src_path = folder / item["new_name"]
            if not src_path.exists():
                continue
            src_doc = fitz.open(str(src_path))
            combined_doc.insert_pdf(src_doc)
            src_doc.close()

            done_count = i + 1
            pct = 30 + int((done_count / total) * 50)
            elapsed = time.perf_counter() - t0
            rate = done_count / elapsed if elapsed > 0 else 1
            remaining = (total - done_count) / rate if rate > 0 else 0
            _emit(q, stage="combine", pct=pct,
                  detail=f"Combined {done_count}/{total} PDFs",
                  eta_sec=remaining)

        combined_doc.save(str(combined_path), deflate=False, garbage=0)
        combined_doc.close()

        _emit(q, stage="combine", pct=80,
              detail=f"Created {combined_name}")

        _emit(q, stage="report", pct=85,
              detail="Generating report.txt...")
        id_info: OrderedDict[str, dict] = OrderedDict()
        for item in sorted_list:
            ids = extract_all_four_digit_ids(item["new_name"])
            variant = item["variant"]
            n = len(ids) if ids else 1
            for oid in ids:
                if oid not in id_info:
                    id_info[oid] = {"variant": variant, "count": 0.0}
                id_info[oid]["count"] += 1.0 / n

        report_lines: list[str] = []
        for oid, info in id_info.items():
            c = info["count"]
            count_str = str(int(c)) if c == int(c) else f"{c:.1f}"
            report_lines.append(f"{oid}\t{info['variant']}\t{count_str}")

            num_orders = len(id_info)
            total_boards = total
            boards_4x3 = sum(1 for item in sorted_list if item["variant"] == "4x3")
            boards_3x2 = sum(1 for item in sorted_list if item["variant"] == "3x2")
            boards_3x3 = sum(1 for item in sorted_list if item["variant"] == "3x3")

            summary_header = (
                    f"Orders: {num_orders}\n"
                 f"Total Boards: {total_boards}\n"
                    f"4x3 Boards: {boards_4x3}\n"
                     f"3x2 Boards: {boards_3x2}\n"
                        f"3x3 Boards: {boards_3x3}\n"
                    f"{'=' * 40}\n"
                        )
            
            

        report_path = folder / "report.txt"
        report_path.write_text(
         summary_header + "Order ID\tVariant\tCount\n" + "\n".join(report_lines),
    encoding="utf-8",
        )

        _emit(q, stage="report", pct=98,
             detail=f"report.txt written ({num_orders} orders, {total_boards} boards)")

        elapsed_total = round(time.perf_counter() - t0, 1)
        _emit(q, stage="done", pct=100, done=True,
              detail=f"Completed in {elapsed_total}s",
              result={
                  "sorted_list": sorted_list,
                  "combined_pdf": combined_name,
                "unique_ids": num_orders,
        "total_boards": total_boards,
        "boards_4x3": boards_4x3,
        "boards_3x2": boards_3x2,
        "boards_3x3": boards_3x3,
                  "elapsed": elapsed_total,
              })

    except Exception as exc:
        _emit(q, stage="error", pct=0, detail="", done=True,
              error=str(exc))


_preview_cache: dict[str, dict] = {}
_preview_cache_lock = threading.Lock()


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/module1/start", methods=["POST"])
def module1_start():
    data = request.get_json(silent=True) or {}
    folder_path = data.get("folder_path", "").strip()
    if not folder_path:
        return jsonify({"error": "folder_path is required"}), 400

    folder_path = os.path.normpath(folder_path)

    task_id = f"m1_{int(time.time()*1000)}"
    _new_progress_queue(task_id)

    thread = threading.Thread(
        target=_run_module1, args=(folder_path, task_id), daemon=True,
    )
    thread.start()

    return jsonify({"task_id": task_id})


@app.route("/module1/progress/<task_id>")
def module1_progress(task_id: str):
    def generate():
        q = _get_progress_queue(task_id)
        if q is None:
            yield f"data: {json.dumps({'error': 'Unknown task_id', 'done': True})}\n\n"
            return
        while True:
            try:
                msg = q.get(timeout=30)
            except queue.Empty:
                yield ": keepalive\n\n"
                continue
            yield f"data: {json.dumps(msg)}\n\n"
            if msg.get("done"):
                _remove_progress_queue(task_id)
                break
    return Response(generate(), mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache",
                             "X-Accel-Buffering": "no"})

_M1_SETTING_BASE = "m1_base_folder"
_OUTPUT_FOLDER_RE = re.compile(r"ship.*output", re.IGNORECASE)

@app.route("/module1/base-folder", methods=["GET"])
def m1_get_base_folder():
    folder = get_setting(_M1_SETTING_BASE, "")
    return jsonify({"folder": folder})

@app.route("/module1/base-folder", methods=["POST"])
def m1_set_base_folder():
    data = request.get_json(silent=True) or {}
    folder = data.get("folder", "").strip()
    if not folder:
        return jsonify({"error": "Folder path is required."}), 400
    p = Path(folder)
    if not p.is_dir():
        return jsonify({"error": f"Folder not found: {folder}"}), 400
    
    set_setting(_M1_SETTING_BASE, str(p))
    return jsonify({"ok": True, "folder": str(p)})

@app.route("/module1/output-folders", methods=["GET"])
def m1_list_output_folders():
    base = get_setting(_M1_SETTING_BASE, "")
    if not base or not Path(base).is_dir():
        return jsonify({"error": "Base folder not set or not found."}), 400
    
    folders = []
    for d in sorted(Path(base).iterdir(), key=lambda x: x.name.lower()):
        if d.is_dir() and _OUTPUT_FOLDER_RE.search(d.name):
            pdf_count = sum(1 for f in d.iterdir() if f.is_file() and f.suffix.lower() == ".pdf")
            folders.append({"name": d.name, "path": str(d), "pdf_count": pdf_count})
            
    return jsonify({"folders": folders})
A4_WIDTH = 595.28
A4_HEIGHT = 841.89
HALF_WIDTH = A4_WIDTH / 2
HALF_HEIGHT = A4_HEIGHT / 2

LEFT_HALF_RECT = fitz.Rect(0, 0, HALF_WIDTH, A4_HEIGHT)


def _process_shipping_labels(file_streams: list) -> io.BytesIO:
    merged = fitz.open()
    for stream in file_streams:
        src = fitz.open(stream=stream.read(), filetype="pdf")
        merged.insert_pdf(src)
        src.close()

    left_halves = fitz.open()
    for page_num in range(len(merged)):
        page = merged[page_num]
        new_page = left_halves.new_page(width=HALF_WIDTH, height=A4_HEIGHT)
        new_page.show_pdf_page(fitz.Rect(0, 0, HALF_WIDTH, A4_HEIGHT),
                               merged, page_num,
                               clip=LEFT_HALF_RECT)

    merged.close()

    output = fitz.open()
    total_halves = len(left_halves)

    for i in range(0, total_halves, 2):
        new_page = output.new_page(width=A4_WIDTH, height=A4_HEIGHT)

        dest_left = fitz.Rect(0, 0, HALF_WIDTH, A4_HEIGHT)
        new_page.show_pdf_page(dest_left, left_halves, i)

        if i + 1 < total_halves:
            dest_right = fitz.Rect(HALF_WIDTH, 0, A4_WIDTH, A4_HEIGHT)
            new_page.show_pdf_page(dest_right, left_halves, i + 1)

    left_halves.close()

    buf = io.BytesIO()
    output.save(buf)
    output.close()
    buf.seek(0)
    return buf


@app.route("/shipping-labels", methods=["POST"])
def shipping_labels():
    files = request.files.getlist("pdfs")
    if not files or all(f.filename == "" for f in files):
        return jsonify({"error": "No PDF files uploaded."}), 400

    pdf_files = [f for f in files if f.filename and f.filename.lower().endswith(".pdf")]
    if not pdf_files:
        return jsonify({"error": "No valid PDF files found."}), 400

    buf = _process_shipping_labels(pdf_files)

    return send_file(
        buf,
        mimetype="application/pdf",
        as_attachment=True,
        download_name="optimized_shipping_labels.pdf",
    )

_M3_SETTING_BASE = "m3_base_folder"

@app.route("/module3/base-folder", methods=["GET"])
def m3_get_base_folder():
    folder = get_setting(_M3_SETTING_BASE, "")
    return jsonify({"folder": folder})

@app.route("/module3/base-folder", methods=["POST"])
def m3_set_base_folder():
    data = request.get_json(silent=True) or {}
    folder = data.get("folder", "").strip()
    if not folder:
        return jsonify({"error": "Folder path is required."}), 400
    p = Path(folder)
    if not p.is_dir():
        return jsonify({"error": f"Folder not found: {folder}"}), 400
    set_setting(_M3_SETTING_BASE, str(p))
    return jsonify({"ok": True, "folder": str(p)})
@app.route("/module3/search", methods=["POST"])
def module3_search():
    data = request.get_json(silent=True) or {}
    order_id = data.get("order_id", "").strip()

    if not re.fullmatch(r"\d{4}", order_id):
        return jsonify({"error": "Order ID must be exactly 4 digits."}), 400

    base = get_setting(_M3_SETTING_BASE, "")
    if not base or not Path(base).is_dir():
        return jsonify({"error": "Base folder not set or not found."}), 400

    base_path = Path(base)
    matches = []
    match_folders = [] # track which folder each match came from

    for d in sorted(base_path.iterdir(), key=lambda x: x.name.lower()):
        if not d.is_dir() or not _OUTPUT_FOLDER_RE.search(d.name):
            continue
            
        for f in sorted(d.iterdir(), key=lambda x: x.name.lower()):
            if not f.is_file() or f.suffix.lower() != ".pdf":
                continue

            ids_in_name = FOUR_DIGIT_RE.findall(f.name)
            if order_id in ids_in_name:
                try:
                    doc = fitz.open(str(f))
                    page_count = len(doc)
                    doc.close()
                    matches.append({
                        "filename": f.name,
                        "page_count": page_count,
                        "folder_name": d.name,
                    })
                    match_folders.append(str(d))
                except Exception:
                    continue

    if not matches:
        return jsonify({"error": f"No PDFs found with order ID {order_id} in any output folder."}), 404

    search_id = f"pv_{int(time.time() * 1000)}"
    with _preview_cache_lock:
        _preview_cache[search_id] = {
            "folders": match_folders,
            "matches": matches,
            "created": time.time(),
        }
        
        # Cleanup old cache entries
        cutoff = time.time() - 1800
        for k in [k for k, v in _preview_cache.items() if v["created"] < cutoff]:
            del _preview_cache[k]

    return jsonify({"search_id": search_id, "matches": matches})


@app.route("/module3/page/<search_id>/<int:match_idx>/<int:page_num>")
def module3_page(search_id: str, match_idx: int, page_num: int):
    with _preview_cache_lock:
        entry = _preview_cache.get(search_id)

    if entry is None:
        return jsonify({"error": "Search expired. Please search again."}), 404
    if match_idx < 0 or match_idx >= len(entry["matches"]):
        return jsonify({"error": "Invalid PDF index."}), 400

    info = entry["matches"][match_idx]
    pdf_path = Path(entry["folders"][match_idx]) / info["filename"]

    if not pdf_path.is_file():
        return jsonify({"error": "PDF file no longer exists."}), 404

    try:
        doc = fitz.open(str(pdf_path))
        if page_num < 0 or page_num >= len(doc):
            doc.close()
            return jsonify({"error": "Invalid page number."}), 400
        pix = doc[page_num].get_pixmap(matrix=fitz.Matrix(1, 1), alpha=False)
        img_data = pix.tobytes("jpeg", jpg_quality=55)
        doc.close()

        return Response(img_data, mimetype="image/jpeg",
                        headers={"Cache-Control": "public, max-age=300"})
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500

from PIL import Image as PILImage


def _compute_dhash(pil_img, hash_size=16):
    img = pil_img.convert("L").resize((hash_size + 1, hash_size), PILImage.LANCZOS)
    pixels = list(img.getdata())
    bits = []
    for row in range(hash_size):
        offset = row * (hash_size + 1)
        for col in range(hash_size):
            bits.append(1 if pixels[offset + col] > pixels[offset + col + 1] else 0)
    return bits


def _hamming(h1, h2):
    return sum(a != b for a, b in zip(h1, h2))


@app.route("/module4/search", methods=["POST"])
def module4_search():
    folder_path = request.form.get("folder_path", "").strip()
    if not folder_path:
        return jsonify({"error": "Folder path is required."}), 400

    folder = Path(os.path.normpath(folder_path))
    if not folder.is_dir():
        return jsonify({"error": f"Folder not found: {folder_path}"}), 400

    if "image" not in request.files:
        return jsonify({"error": "No image uploaded."}), 400
    img_file = request.files["image"]
    if not img_file.filename:
        return jsonify({"error": "No image selected."}), 400

    try:
        query_img = PILImage.open(img_file.stream)
        query_hash = _compute_dhash(query_img)
    except Exception:
        return jsonify({"error": "Could not read the uploaded image."}), 400

    MATCH_THRESHOLD = 64
    results = []

    pdf_files = sorted(
        [f for f in folder.iterdir() if f.is_file() and f.suffix.lower() == ".pdf"],
        key=lambda x: x.name.lower(),
    )

    for pdf_file in pdf_files:
        try:
            doc = fitz.open(str(pdf_file))
            best_dist = 256
            seen_xrefs = set()

            for page in doc:
                for img_info in page.get_images(full=True):
                    xref = img_info[0]
                    if xref in seen_xrefs:
                        continue
                    seen_xrefs.add(xref)
                    try:
                        raw = doc.extract_image(xref)
                        pil = PILImage.open(io.BytesIO(raw["image"]))
                        dist = _hamming(query_hash, _compute_dhash(pil))
                        if dist < best_dist:
                            best_dist = dist
                    except Exception:
                        continue

            if not seen_xrefs:
                for page_num in range(min(len(doc), 5)):
                    try:
                        pix = doc[page_num].get_pixmap(
                            matrix=fitz.Matrix(0.5, 0.5), alpha=False,
                        )
                        pil = PILImage.open(io.BytesIO(pix.tobytes("png")))
                        dist = _hamming(query_hash, _compute_dhash(pil))
                        if dist < best_dist:
                            best_dist = dist
                    except Exception:
                        continue

            page_count = len(doc)
            doc.close()

            if best_dist <= MATCH_THRESHOLD:
                results.append({
                    "filename": pdf_file.name,
                    "page_count": page_count,
                    "distance": best_dist,
                    "similarity": round((1 - best_dist / 256) * 100, 1),
                })
        except Exception:
            continue

    results.sort(key=lambda x: x["distance"])

    if not results:
        return jsonify({"error": "No matching PDFs found for the uploaded image."}), 404

    search_id = f"m4_{int(time.time() * 1000)}"
    with _preview_cache_lock:
        _preview_cache[search_id] = {
            "folder": str(folder),
            "matches": [
                {"filename": m["filename"], "page_count": m["page_count"]}
                for m in results
            ],
            "created": time.time(),
        }

        cutoff = time.time() - 1800
        for k in [k for k, v in _preview_cache.items() if v["created"] < cutoff]:
            del _preview_cache[k]

    return jsonify({"search_id": search_id, "matches": results})


# --- Register Polaroid bulk-processing blueprint ---
from polaroid_tool import polaroid_bp        # noqa: E402
app.register_blueprint(polaroid_bp)

# --- Module 6: Final Label ---

_M6_SETTING_BASE = "m6_base_folder"

# Regex to find Amazon-style order IDs like 407-6518891-4216357
_AMAZON_ORDER_RE = re.compile(r"(\d{3})\s*[-—]\s*(\d{7})\s*[-—]\s*(\d{4})\d{3}")

def _extract_label_order_ids_from_page(doc, page_num: int) -> list[dict]:
    """Extract order IDs from a page that may have 1 or 2 labels (left+right halves).
    Returns list of dicts: {'order_id_4': str, 'half': 'left'|'right'|'full'}"""
    page = doc[page_num]
    pw = page.rect.width
    ph = page.rect.height
    results = []

    # Try left half
    left_rect = fitz.Rect(0, 0, pw / 2, ph)
    left_text = page.get_text("text", clip=left_rect)
    left_ids = _AMAZON_ORDER_RE.findall(left_text)
    if left_ids:
        results.append({"order_id_4": left_ids[0], "half": "left"})

    # Try right half
    right_rect = fitz.Rect(pw / 2, 0, pw, ph)
    right_text = page.get_text("text", clip=right_rect)
    right_ids = _AMAZON_ORDER_RE.findall(right_text)
    if right_ids:
        results.append({"order_id_4": right_ids[0], "half": "right"})

    # If nothing found in halves, try full page
    if not results:
        full_text = page.get_text("text")
        full_ids = _AMAZON_ORDER_RE.findall(full_text)
        if full_ids:
            for fid in full_ids:
                results.append({"order_id_4": fid, "half": "full"})

    return results

@app.route("/final-label/base-folder", methods=["GET"])
def fl_get_base_folder():
    folder = get_setting(_M6_SETTING_BASE, "")
    return jsonify({"folder": folder})

@app.route("/final-label/base-folder", methods=["POST"])
def fl_set_base_folder():
    data = request.get_json(silent=True) or {}
    folder = data.get("folder", "").strip()
    if not folder:
        return jsonify({"error": "Folder path is required."}), 400
    p = Path(folder)
    if not p.is_dir():
        return jsonify({"error": f"Folder not found: {folder}"}), 400
    set_setting(_M6_SETTING_BASE, str(p))
    return jsonify({"ok": True, "folder": str(p)})

@app.route("/final-label/process", methods=["POST"])
def fl_process():
    """Process uploaded label PDFs against all output folders in the base folder."""
    base = get_setting(_M6_SETTING_BASE, "")
    if not base or not Path(base).is_dir():
        return jsonify({"error": "Base folder not set or not found."}), 400

    files = request.files.getlist("pdfs")
    if not files or all(f.filename == "" for f in files):
        return jsonify({"error": "No PDF files uploaded."}), 400

    pdf_files = [f for f in files if f.filename and f.filename.lower().endswith(".pdf")]
    if not pdf_files:
        return jsonify({"error": "No valid PDF files found."}), 400

    base_path = Path(base)

    # --- Step 1: Parse all labels from uploaded PDFs ---
    label_docs = []
    label_map: dict[str, list[dict]] = {} # oid4 -> [{ doc_idx, page_num, half }]

    for f in pdf_files:
        doc = fitz.open(stream=f.read(), filetype="pdf")
        doc_idx = len(label_docs)
        label_docs.append(doc)

        for pn in range(len(doc)):
            entries = _extract_label_order_ids_from_page(doc, pn)
            for entry in entries:
                oid4 = entry["order_id_4"]
                if oid4 not in label_map:
                    label_map[oid4] = []
                label_map[oid4].append({
                    "doc_idx": doc_idx,
                    "page_num": pn,
                    "half": entry["half"],
                })

    # --- Step 2: Find all output folders ---
    output_dirs = []
    for d in sorted(base_path.iterdir(), key=lambda x: x.name.lower()):
        if d.is_dir() and _OUTPUT_FOLDER_RE.search(d.name):
            output_dirs.append(d)

    if not output_dirs:
        for doc in label_docs:
            doc.close()
        return jsonify({"error": "No output folders found in base folder."}), 400

    # --- Step 3: For each output folder, match labels ---
    overall_results = []
    total_matched = 0
    total_cancelled = 0
    total_label_pdfs = 0

    for out_dir in output_dirs:
        folder_order_ids: set[str] = set()
        for f in out_dir.iterdir():
            if f.is_file() and f.suffix.lower() == ".pdf":
                ids = FOUR_DIGIT_RE.findall(f.name)
                folder_order_ids.update(ids)

        if not folder_order_ids:
            continue

        matched_ids = []
        cancelled_ids = []
        matched_labels: list[dict] = []

        for oid in sorted(folder_order_ids):
            if oid in label_map:
                matched_ids.append(oid)
                matched_labels.extend(label_map[oid])
            else:
                cancelled_ids.append(oid)

        matched_labels.sort(key=lambda x: (x["doc_idx"], x["page_num"]))

        if matched_labels:
            filtered_doc = fitz.open()

            for lbl in matched_labels:
                src_doc = label_docs[lbl["doc_idx"]]
                src_page = src_doc[lbl["page_num"]]
                pw = src_page.rect.width
                ph = src_page.rect.height

                if lbl["half"] == "left":
                    clip = fitz.Rect(0, 0, pw / 2, ph)
                    new_page = filtered_doc.new_page(width=pw / 2, height=ph)
                    new_page.show_pdf_page(
                        fitz.Rect(0, 0, pw / 2, ph), src_doc, lbl["page_num"],
                        clip=clip
                    )
                elif lbl["half"] == "right":
                    clip = fitz.Rect(pw / 2, 0, pw, ph)
                    new_page = filtered_doc.new_page(width=pw / 2, height=ph)
                    new_page.show_pdf_page(
                        fitz.Rect(0, 0, pw / 2, ph), src_doc, lbl["page_num"],
                        clip=clip
                    )
                else:
                    filtered_doc.insert_pdf(src_doc, from_page=lbl["page_num"],
                                            to_page=lbl["page_num"])

            single_labels = fitz.open()
            for pn in range(len(filtered_doc)):
                pg = filtered_doc[pn]
                new_pg = single_labels.new_page(width=HALF_WIDTH, height=A4_HEIGHT)
                new_pg.show_pdf_page(
                    fitz.Rect(0, 0, HALF_WIDTH, A4_HEIGHT),
                    filtered_doc, pn
                )
            filtered_doc.close()

            final_doc = fitz.open()
            total_singles = len(single_labels)
            for i in range(0, total_singles, 2):
                new_page = final_doc.new_page(width=A4_WIDTH, height=A4_HEIGHT)
                dest_left = fitz.Rect(0, 0, HALF_WIDTH, A4_HEIGHT)
                new_page.show_pdf_page(dest_left, single_labels, i)
                if i + 1 < total_singles:
                    dest_right = fitz.Rect(HALF_WIDTH, 0, A4_WIDTH, A4_HEIGHT)
                    new_page.show_pdf_page(dest_right, single_labels, i + 1)
            single_labels.close()

            label_pdf_name = f"labels-{out_dir.name}.pdf"
            label_pdf_path = out_dir / label_pdf_name
            final_doc.save(str(label_pdf_path))
            final_doc.close()
            total_label_pdfs += 1

        report_lines = []
        report_lines.append(f"Label Report for: {out_dir.name}")
        report_lines.append("=" * 50)
        report_lines.append(f"Total orders in folder: {len(folder_order_ids)}")
        report_lines.append(f"Labels matched: {len(matched_ids)}")
        report_lines.append(f"Cancelled (no label found): {len(cancelled_ids)}")
        report_lines.append("")

        if matched_ids:
            report_lines.append("MATCHED ORDERS:")
            for oid in matched_ids:
                report_lines.append(f" {oid} - Label found")

        if cancelled_ids:
            report_lines.append("")
            report_lines.append("CANCELLED ORDERS (no label found):")
            for oid in cancelled_ids:
                report_lines.append(f" {oid} - CANCELLED")

        report_path = out_dir / "label-report.txt"
        report_path.write_text("\n".join(report_lines) + "\n", encoding="utf-8")

        total_matched += len(matched_ids)
        total_cancelled += len(cancelled_ids)

        overall_results.append({
            "folder": out_dir.name,
            "total_orders": len(folder_order_ids),
            "matched": len(matched_ids),
            "cancelled": len(cancelled_ids),
            "cancelled_ids": cancelled_ids,
            "label_pdf": label_pdf_name if matched_labels else None,
        })

    for doc in label_docs:
        doc.close()

    return jsonify({
        "results": overall_results,
        "summary": {
            "output_folders_processed": len(overall_results),
            "total_matched": total_matched,
            "total_cancelled": total_cancelled,
            "label_pdfs_created": total_label_pdfs,
            "total_labels_parsed": sum(len(v) for v in label_map.values()),
            "unique_label_ids": len(label_map),
        }
    })



if __name__ == "__main__":
    app.run(debug=True, port=5200)
    
                










    
                        