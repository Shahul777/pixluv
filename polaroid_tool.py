import json
import logging
import shutil
import zipfile
from PIL import Image
import math
import os
import queue
import re
import threading
import time
from pathlib import Path

from natsort import natsorted
from flask import Blueprint, Response, jsonify, request

from db import get_setting, set_setting, log_activity

# --- Import image-placement engine from imageplacer.py ---
from imageplacer import LAYOUTS, SUPPORTED_EXTENSIONS, generate_pdf

polaroid_bp = Blueprint("polaroid", __name__)

# --- Progress infrastructure (same pattern as Module 1 in app.py) ---
_pq: dict[str, queue.Queue] = {}
_pq_lock = threading.Lock()

SETTING_BASE_FOLDER = "polaroid_base_folder"

# Variant regex: ends with -3x2 or -3x3 (case-insensitive)
_VARIANT_SUFFIX_RE = re.compile(r"-(3x[23])\s*$", re.IGNORECASE)

# Ship-folder pattern: contains "ship" and "image" (flexible date formats)
_SHIP_FOLDER_RE = re.compile(r"ship.*image", re.IGNORECASE)

# Special multi-set indicators in folder name
_SPECIAL_RE = re.compile(r"(q\d+|set\d+)", re.IGNORECASE)

# Four-digit order-id extractor from folder name
_ORDER_ID_RE = re.compile(r"-(\d{4})(?:-|$)")

# Layout keys mapped by variant shorthand
_VARIANT_TO_LAYOUT = {
    "4x3": "4x3_polaroid_18",
    "3x2": "3x2_polaroid_36",
    "3x3": "3x3_square_24",
}

_UPS = {
    "4x3": 18,
    "3x2": 36,
    "3x3": 24,
}
def _new_q(task_id: str) -> queue.Queue:
    q: queue.Queue = queue.Queue()
    with _pq_lock:
        _pq[task_id] = q
    return q

def _get_q(task_id: str) -> queue.Queue | None:
    with _pq_lock:
        return _pq.get(task_id)

def _del_q(task_id: str):
    with _pq_lock:
        _pq.pop(task_id, None)

def _emit(q: queue.Queue, **kw):
    q.put(kw)

def _detect_variant(folder_name: str) -> str:
    """Return '3x2', '3x3', or '4x3' based on folder name suffix."""
    m = _VARIANT_SUFFIX_RE.search(folder_name)
    if m:
        return m.group(1).lower()
    return "4x3"

def _extract_order_id(folder_name: str) -> str | None:
    m = _ORDER_ID_RE.search(folder_name)
    return m.group(1) if m else None

def _is_special(folder_name: str) -> bool:
    """True if folder name contains q1, q2, set1, set2 etc."""
    return bool(_SPECIAL_RE.search(folder_name))

def _collect_images(folder: Path) -> list[str]:
    """Return sorted list of image file paths inside a folder."""
    imgs = []
    for f in folder.iterdir():
        if f.is_file() and f.suffix.lower() in SUPPORTED_EXTENSIONS:
            imgs.append(str(f))
    imgs = natsorted(imgs, key=lambda p: Path(p).name.lower())
    return imgs

def _output_folder_for(ship_folder: Path) -> Path:
    """ship-30-4-image -> ship-30-4-output (sibling folder)."""
    name = ship_folder.name
    # Replace last occurrence of "image" with "output"
    if "image" in name.lower():
        idx = name.lower().rfind("image")
        out_name = name[:idx] + "output" + name[idx + 5:]
    else:
        out_name = name + "-output"
    return ship_folder.parent / out_name

def _pdf_already_exists(output_dir: Path, order_id: str | None, folder_name: str, is_special: bool) -> bool:
    """Check if output PDF for this order already exists.
    Skip only if NOT a special (q/set) order and a matching PDF is present."""
    if is_special or order_id is None:
        return False
    for f in output_dir.iterdir():
        if not f.is_file() or f.suffix.lower() != ".pdf":
            continue
        if order_id in f.stem:
            return True
    return False

# --- Core processing thread ---

def _run_polaroid(ship_folder_path: str, task_id: str):
    q = _get_q(task_id)
    if q is None:
        return

    t0 = time.perf_counter()

    try:
        ship_folder = Path(ship_folder_path)
        if not ship_folder.is_dir():
            _emit(q, stage="error", pct=0, detail="", done=True,
                  error=f"Folder not found: {ship_folder_path}")
            return

        # --- Stage: scan -----------------------------------------
        _emit(q, stage="scan", pct=0, detail="Scanning order folders...", done=False)

        order_dirs = []
        for d in sorted(ship_folder.iterdir(), key=lambda x: x.name.lower()):
            if d.is_dir():
                order_dirs.append(d)

        if not order_dirs:
            _emit(q, stage="error", pct=0, detail="", done=True,
                  error="No order folders found inside the selected ship folder.")
            return

        _emit(q, stage="scan", pct=5,
              detail=f"Found {len(order_dirs)} order folders.", done=False)

        # --- Stage: prepare output folder ------------------------
        output_dir = _output_folder_for(ship_folder)
        output_dir.mkdir(parents=True, exist_ok=True)

        _emit(q, stage="prepare", pct=8,
              detail=f"Output folder: {output_dir.name}", done=False)

        # --- Stage: process orders -------------------------------
        total_orders = len(order_dirs)
        processed_count = 0
        skipped_count = 0
        pdf_count = 0
        results: list[dict] = []

        for oi, order_dir in enumerate(order_dirs):
            folder_name = order_dir.name
            variant = _detect_variant(folder_name)
            order_id = _extract_order_id(folder_name)
            special = _is_special(folder_name)
            layout_key = _VARIANT_TO_LAYOUT[variant]
            layout = LAYOUTS[layout_key]
            ups = _UPS[variant]

            # Check if already processed
            if _pdf_already_exists(output_dir, order_id, folder_name, special):
                skipped_count += 1
                pct = 10 + int((oi + 1) / total_orders * 85)
                _emit(q, stage="process", pct=pct,
                      detail=f"Skipped (already exists): {folder_name}",
                      done=False, order_index=oi, order_total=total_orders,
                      order_name=folder_name, status="skipped")
                results.append({
                    "folder": folder_name, "variant": variant,
                    "order_id": order_id or "", "status": "skipped",
                    "pdfs": 0, "photos": 0,
                })
                continue

            # Collect images
            images = _collect_images(order_dir)
            if not images:
                skipped_count += 1
                pct = 10 + int((oi + 1) / total_orders * 85)
                _emit(q, stage="process", pct=pct,
                      detail=f"Skipped (no images): {folder_name}",
                      done=False, order_index=oi, order_total=total_orders,
                      order_name=folder_name, status="empty")
                results.append({
                    "folder": folder_name, "variant": variant,
                    "order_id": order_id or "", "status": "no images",
                })
                continue
            num_photos = len(images)
            num_pdfs = math.ceil(num_photos / ups)

            _emit(q, stage="process", pct=10 + int(oi / total_orders * 85),
                  detail=f"Processing: {folder_name} ({num_photos} photos, {num_pdfs} PDF{'s' if num_pdfs > 1 else ''})",
                  done=False, order_index=oi, order_total=total_orders,
                  order_name=folder_name, status="processing")

            for pi in range(num_pdfs):
                start_idx = pi * ups
                end_idx = min(start_idx + ups, num_photos)
                batch = images[start_idx:end_idx]

                # Build label & filename
                if num_pdfs == 1:
                    label = folder_name
                    pdf_filename = f"{folder_name}.pdf"
                else:
                    label = f"{folder_name}-({pi + 1}/{num_pdfs})"
                    pdf_filename = f"{folder_name}-({pi + 1}_{num_pdfs}).pdf"

                pdf_out = str(output_dir / pdf_filename)
                generate_pdf(batch, pdf_out, label, layout,layout_key)
                pdf_count += 1

                # Sub-progress within order
                sub_pct = 10 + int(((oi + (pi + 1) / num_pdfs) / total_orders) * 85)
                _emit(q, stage="process", pct=sub_pct,
                      detail=f"{folder_name}: PDF {pi + 1}/{num_pdfs} done",
                      done=False, order_index=oi, order_total=total_orders,
                      order_name=folder_name, status="processing")

            processed_count += 1
            results.append({
                "folder": folder_name, "variant": variant,
                "order_id": order_id or "", "status": "done",
                "pdfs": num_pdfs, "photos": num_photos,
            })

        # --- Stage: done -----------------------------------------
        elapsed = round(time.perf_counter() - t0, 1)

        log_activity("polaroid", "bulk_process", json.dumps({
            "ship_folder": ship_folder_path,
            "processed": processed_count,
            "skipped": skipped_count,
            "pdfs": pdf_count,
            "elapsed": elapsed,
        }))

        _emit(q, stage="done", pct=100, done=True,
              detail=f"Completed in {elapsed}s",
              result={
                  "orders": results,
                  "processed": processed_count,
                  "skipped": skipped_count,
                  "total_pdfs": pdf_count,
                  "output_folder": str(output_dir),
                  "elapsed": elapsed,
              })

    except Exception as exc:
        _emit(q, stage="error", pct=0, detail="", done=True,
              error=str(exc))

# --- Routes ------------------------------------------------------

# --- Routes ------------------------------------------------------

@polaroid_bp.route("/polaroid/base-folder", methods=["GET"])
def get_base_folder():
    folder = get_setting(SETTING_BASE_FOLDER, "")
    return jsonify({"folder": folder})


@polaroid_bp.route("/polaroid/base-folder", methods=["POST"])
def set_base_folder():
    data = request.get_json(silent=True) or {}
    folder = data.get("folder", "").strip()
    if not folder:
        return jsonify({"error": "Folder path is required."}), 400
    p = Path(folder)
    if not p.is_dir():
        return jsonify({"error": f"Folder not found: {folder}"}), 400
    set_setting(SETTING_BASE_FOLDER, str(p))
    log_activity("polaroid", "set_base_folder", str(p))
    return jsonify({"ok": True, "folder": str(p)})

@polaroid_bp.route("/polaroid/ship-folders",methods=["GET"])
def list_ship_folders():
    base = get_setting(SETTING_BASE_FOLDER, "")
    if not base or not Path(base).is_dir():
        return jsonify({"error": "Base folder not set or not found."}), 400

    folders = []
    for d in sorted(Path(base).iterdir(), key=lambda x: x.name.lower()):
        if d.is_dir() and _SHIP_FOLDER_RE.search(d.name):
            # Count order sub-folders
            order_count = sum(1 for sub in d.iterdir() if sub.is_dir())
            folders.append({"name": d.name, "path": str(d), "order_count": order_count})

    return jsonify({"folders": folders})


@polaroid_bp.route("/polaroid/start", methods=["POST"])
def start_processing():
    data = request.get_json(silent=True) or {}
    ship_folder = data.get("ship_folder", "").strip()
    if not ship_folder:
        return jsonify({"error": "ship_folder is required."}), 400

    ship_folder = os.path.normpath(ship_folder)
    if not Path(ship_folder).is_dir():
        return jsonify({"error": f"Folder not found: {ship_folder}"}), 400

    task_id = f"pol_{int(time.time() * 1000)}"
    _new_q(task_id)

    thread = threading.Thread(
        target=_run_polaroid, args=(ship_folder, task_id), daemon=True,
    )
    thread.start()

    return jsonify({"task_id": task_id})


@polaroid_bp.route("/polaroid/progress/<task_id>")
def polaroid_progress(task_id: str):
    def generate():
        q = _get_q(task_id)
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
                _del_q(task_id)
                break

    return Response(generate(), mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache",
                             "X-Accel-Buffering": "no"})

# --- Order Search (cancelled-order lookup) -----------------------

# Pattern to match output folders (sibling of ship-image folders)
_OUTPUT_FOLDER_RE = re.compile(r"ship.*output", re.IGNORECASE)

@polaroid_bp.route("/polaroid/order-search", methods=["POST"])
def order_search():
    """Search the entire base folder for a 4-digit order ID in both
    ship-image (order sub-folders) and ship-output (PDF filenames)."""
    data = request.get_json(silent=True) or {}
    order_id = data.get("order_id", "").strip()

    if not re.fullmatch(r"\d{4}", order_id):
        return jsonify({"error": "Please enter a valid 4-digit order ID."}), 400

    base = get_setting(SETTING_BASE_FOLDER, "")
    if not base or not Path(base).is_dir():
        return jsonify({"error": "Base folder not set or not found."}), 400

    base_path = Path(base)
    results: list[dict] = []

    for d in sorted(base_path.iterdir(), key=lambda x: x.name.lower()):
        if not d.is_dir():
            continue

        is_image_folder = bool(_SHIP_FOLDER_RE.search(d.name))
        is_output_folder = bool(_OUTPUT_FOLDER_RE.search(d.name))

        if not is_image_folder and not is_output_folder:
            continue

        if is_image_folder:
            # Search for order sub-folders containing the order ID
            for sub in sorted(d.iterdir(), key=lambda x: x.name.lower()):
                if not sub.is_dir():
                    continue
                if order_id in sub.name:
                    img_count = sum(
                        1 for f in sub.iterdir()
                        if f.is_file() and f.suffix.lower() in SUPPORTED_EXTENSIONS
                    )
                    variant = _detect_variant(sub.name)
                    results.append({
                        "ship_folder": d.name,
                        "type": "input",
                        "name": sub.name,
                        "variant": variant,
                        "photos": img_count,
                        "path": str(sub),
                    })

        if is_output_folder:
            # Search for PDFs containing the order ID in filename
            for f in sorted(d.iterdir(), key=lambda x: x.name.lower()):
                if not f.is_file() or f.suffix.lower() != ".pdf":
                    continue
                if order_id in f.stem:
                    results.append({
                        "ship_folder": d.name,
                        "type": "output",
                        "name": f.name,
                        "variant": _detect_variant(f.stem),
                        "photos": None,
                        "path": str(f),
                    })

    found_input = sum(1 for r in results if r["type"] == "input")
    found_output = sum(1 for r in results if r["type"] == "output")

    return jsonify({
        "order_id": order_id,
        "results": results,
        "summary": {
            "total": len(results),
            "input_folders": found_input,
            "output_pdfs": found_output,
        },
    })

_log = logging.getLogger("extract-image")
_IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif", ".heic", ".heif", ".webp"}

_OCR_KEYWORDS = re.compile(r"(Note\s*:|PIXLUV)", re.IGNORECASE)

_ocr_engine = None
_OCR_SKIP_SIZE = 100 * 1024  # skip files > 100KB (real photos, not promo cards)

def _get_ocr():
    global _ocr_engine
    if _ocr_engine is None:
        from rapidocr_onnxruntime import RapidOCR
        _ocr_engine = RapidOCR()
    return _ocr_engine

def _has_ocr_keyword(image_path: str) -> bool:
    """Fast OCR check: skip large files, then detect keywords."""
    try:
        # Large files are real photos - promo/note images are under 1MB
        if os.path.getsize(image_path) > _OCR_SKIP_SIZE:
            return False

        img = Image.open(image_path)
        img = img.convert("RGB")
        import numpy as np
        img_array = np.array(img)
        ocr = _get_ocr()
        result, _ = ocr(img_array)
        if result:
            text = " ".join(line[1] for line in result)
            return bool(_OCR_KEYWORDS.search(text))
        return False
    except Exception as exc:
        _log.warning("OCR failed for %s: %s", image_path, exc)
        return False

def _is_image_file(p: Path) -> bool:
    return p.is_file() and p.suffix.lower() in _IMAGE_EXTS

def _run_extract_images(ship_folder_path: str, task_id: str):
    """Worker thread: clean up order folders, extract zips, OCR-filter, write report."""
    q = _get_q(task_id)
    if q is None:
        return

    t0 = time.perf_counter()

    try:
        ship_folder = Path(ship_folder_path)
        if not ship_folder.is_dir():
            _emit(q, stage="error", pct=0, detail="", done=True,
                  error=f"Folder not found: {ship_folder_path}")
            return

        # --- Stage: scan ----------------------------------------------------
        _emit(q, stage="scan", pct=0, detail="Scanning order folders...", done=False)

        order_dirs = []
        for d in sorted(ship_folder.iterdir(), key=lambda x: x.name.lower()):
            if d.is_dir():
                order_dirs.append(d)

        if not order_dirs:
            _emit(q, stage="error", pct=0, detail="", done=True,
                  error="No order folders found inside the selected ship folder.")
            return

        total = len(order_dirs)
        _emit(q, stage="scan", pct=5,
              detail=f"Found {total} order folders.", done=False)

        # --- Stage: extract & cleanup ---------------------------------------
        results: list[dict] = []

        for oi, order_dir in enumerate(order_dirs):
            folder_name = order_dir.name
            order_id = _extract_order_id(folder_name) or ""
            pct = 5 + int((oi / total) * 60)

            _emit(q, stage="extract", pct=pct,
                  detail=f"Cleaning: {folder_name}",
                  done=False, order_name=folder_name,
                  order_index=oi, order_total=total)
            all_files = [f for f in order_dir.iterdir() if f.is_file()]
            zip_files = [f for f in all_files if f.suffix.lower() == ".zip"]
            image_files = [f for f in all_files if _is_image_file(f)]
            other_files = [f for f in all_files 
                           if not _is_image_file(f) and f.suffix.lower() != ".zip"]

            extracted_zip = False
            deleted_files: list[str] = []

            if len(zip_files) >= 1 and len(image_files) == 0:
                for zf_path in zip_files:
                    max_retries = 3
                    for attempt in range(max_retries):
                        try:
                            with zipfile.ZipFile(str(zf_path), 'r') as zf:
                    # Security: skip entries with path traversal
                                for info in zf.infolist():
                                    if info.is_dir():
                                        continue
                                    member_name = Path(info.filename).name
                                    if not member_name or '..' in info.filename:
                                        continue
                                    target = order_dir / member_name
                                    with zf.open(info) as src, open(str(target), 'wb') as dst:
                                        dst.write(src.read())
                            extracted_zip = True
                            _log.info("Extracted zip: %s", zf_path.name)
                            break  # success, no retry needed
                        except Exception as exc:
                            _log.warning("Attempt %d/%d failed to extract %s: %s",
                                attempt + 1, max_retries, zf_path, exc)
                            if attempt < max_retries - 1:
                                time.sleep(1)  # wait before retry (file lock release)

        # Delete the zip after extraction (or failed attempts)
                    try:
                        zf_path.unlink()
                        deleted_files.append(zf_path.name)
                    except OSError:
                        _log.warning("Could not delete zip: %s", zf_path.name)

            # Case 2 & 3: images exist (maybe with zips/other files)
            # -> keep only image files, delete everything else
            elif len(image_files) > 0:
                for zf in zip_files:
                    try:
                        zf.unlink()
                        deleted_files.append(zf.name)
                    except OSError:
                        pass
            
            # Delete all non-image files remaining in the folder
            for f in order_dir.iterdir():
                if f.is_file() and not _is_image_file(f):
                    try:
                        f.unlink()
                        deleted_files.append(f.name)
                    except OSError:
                        pass

            # Also remove any subdirectories that zip might have created
            for d in order_dir.iterdir():
                if d.is_dir():
                    # Move images from subdirs up to order folder
                    for sub_f in d.rglob("*"):
                        if sub_f.is_file() and _is_image_file(sub_f):
                            dest = order_dir / sub_f.name
                            if not dest.exists():
                                sub_f.rename(dest)
                    try:
                        shutil.rmtree(str(d))
                    except OSError:
                        pass
            
            # Count images after cleanup
            
                        # Count images after cleanup
            final_images = [f for f in order_dir.iterdir() if f.is_file() and _is_image_file(f)]

            results.append({
                "folder": folder_name,
                "order_id": order_id,
                "extracted_zip": extracted_zip,
                "deleted": deleted_files,
                "images_before_ocr": len(final_images),
                "ocr_removed": [],
                "final_count": len(final_images),
            })

        # --- Stage: OCR check (threaded) ------------------------------------
        _emit(q, stage="ocr", pct=65,
              detail="Running OCR check on images...", done=False)

        # Pre-init OCR engine before threads start
        _get_ocr()

        from concurrent.futures import ThreadPoolExecutor, as_completed

        # Collect all images across all orders for parallel OCR
        all_ocr_tasks: list[tuple[int, Path]] = []  # (result_index, file_path)
        for ri, res in enumerate(results):
            order_dir = ship_folder / res["folder"]
            for f in sorted(order_dir.iterdir(), key=lambda x: x.name.lower()):
                if f.is_file() and _is_image_file(f):
                    all_ocr_tasks.append((ri, f))

        total_ocr = len(all_ocr_tasks)
        ocr_done_count = 0
        ocr_to_remove: dict[int, list[Path]] = {}  # result_index -> files to delete

        num_workers = min(os.cpu_count() or 4, 8)

        with ThreadPoolExecutor(max_workers=num_workers) as pool:
            future_map = {
                pool.submit(_has_ocr_keyword, str(fp)): (ri, fp)
                for ri, fp in all_ocr_tasks
            }
            for future in as_completed(future_map):
                ri, fp = future_map[future]
                ocr_done_count += 1
                try:
                    if future.result():
                        ocr_to_remove.setdefault(ri, []).append(fp)
                except Exception:
                    pass

                # Progress update every few images
                if ocr_done_count % 5 == 0 or ocr_done_count == total_ocr:
                    pct = 65 + int((ocr_done_count / max(total_ocr, 1)) * 30)
                    _emit(q, stage="ocr", pct=pct,
                          detail=f"OCR checked {ocr_done_count}/{total_ocr} images",
                          done=False)

        # Delete flagged images and update results
        for ri, files in ocr_to_remove.items():
            removed_names = []
            for fp in files:
                try:
                    fp.unlink()
                    removed_names.append(fp.name)
                    _log.info("OCR removed: %s", fp)
                except OSError:
                    pass
            results[ri]["ocr_removed"] = removed_names

        # Recount final images for all orders
                # Recount final images for all orders
        for res in results:
            order_dir = ship_folder / res["folder"]
            final_count = sum(1 for f in order_dir.iterdir()
                              if f.is_file() and _is_image_file(f))
            res["final_count"] = final_count

        # --- Stage: report --------------------------------------------------
        _emit(q, stage="report", pct=96,
              detail="Generating report...", done=False)

        report_lines = ["Extract Image Report", "=" * 50, ""]
        report_lines.append(f"Ship Folder: {ship_folder.name}")
        report_lines.append(f"Date: {time.strftime('%Y-%m-%d %H:%M:%S')}")
        report_lines.append(f"Total Orders: {total}")
        report_lines.append("")
        report_lines.append(f"{'Order Folder':<45} {'Order ID':<12} {'Images':>8}")
        report_lines.append("-" * 70)

        total_images = 0
        for res in results:
            report_lines.append(
                f"{res['folder']:<45} {res['order_id']:<12} {res['final_count']:>8}"
            )
            total_images += res["final_count"]

        report_lines.append("-" * 70)
        report_lines.append(f"{'TOTAL':<45} {'':<12} {total_images:>8}")
        report_lines.append("")

        # Details section
        has_details = any(r["extracted_zip"] or r["deleted"] or r["ocr_removed"]
                          for r in results)
        if has_details:
            report_lines.append("Details:")
            report_lines.append("-" * 50)
            for res in results:
                if res["extracted_zip"] or res["deleted"] or res["ocr_removed"]:
                    report_lines.append(f"\n  {res['folder']}:")
                    if res["extracted_zip"]:
                        report_lines.append("    - Zip extracted")
                    if res["deleted"]:
                        report_lines.append(f"    - Deleted: {', '.join(res['deleted'])}")
                    if res["ocr_removed"]:
                        report_lines.append(f"    - OCR removed: {', '.join(res['ocr_removed'])}")

        report_path = ship_folder / "report.txt"
        report_path.write_text("\n".join(report_lines), encoding="utf-8")

        # --- Stage: done ----------------------------------------------------
        elapsed = round(time.perf_counter() - t0, 1)

        orders_with_zip = sum(1 for r in results if r["extracted_zip"])
        total_ocr_removed = sum(len(r["ocr_removed"]) for r in results)
        total_deleted = sum(len(r["deleted"]) for r in results)

        log_activity("polaroid", "extract_images", json.dumps({
            "ship_folder": ship_folder_path,
            "total_orders": total,
            "elapsed": elapsed,
        }))

        _emit(q, stage="done", pct=100, done=True,
              detail=f"Completed in {elapsed}s",
              result={
                  "orders": results,
                  "total_orders": total,
                  "total_images": total_images,
                  "zips_extracted": orders_with_zip,
                  "files_deleted": total_deleted,
                  "ocr_removed": total_ocr_removed,
                  "files_deleted": total_deleted,
                  "ocr_removed": total_ocr_removed,
                  "report_path": str(report_path),
                  "elapsed": elapsed,
              })

    except Exception as exc:
        _log.exception("Extract image failed")
        _emit(q, stage="error", pct=0, detail="", done=True,
              error=str(exc))

@polaroid_bp.route("/polaroid/extract-start", methods=["POST"])
def start_extract():
    data = request.get_json(silent=True) or {}
    ship_folder = data.get("ship_folder", "").strip()
    if not ship_folder:
        return jsonify({"error": "ship_folder is required."}), 400

    ship_folder = os.path.normpath(ship_folder)
    if not Path(ship_folder).is_dir():
        return jsonify({"error": f"Folder not found: {ship_folder}"}), 400

    task_id = f"ext_{int(time.time() * 1000)}"
    _new_q(task_id)

    thread = threading.Thread(
        target=_run_extract_images, args=(ship_folder, task_id), daemon=True,
    )
    thread.start()

    return jsonify({"task_id": task_id})

@polaroid_bp.route("/polaroid/extract-progress/<task_id>")
def extract_progress(task_id: str):
    def generate():
        q = _get_q(task_id)
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
                _del_q(task_id)
                break

    return Response(generate(), mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache",
                             "X-Accel-Buffering": "no"})
                             