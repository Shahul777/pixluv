import json
import math
import os
import queue
import re
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

from natsort import natsorted
from flask import Blueprint, Response, jsonify, request

from db import get_setting, set_setting, log_activity

# --- Import image-placement engine from imageplacer.py ---
from imageplacer import LAYOUTS, SUPPORTED_EXTENSIONS, generate_pdf, make_thumbnail

polaroid_bp = Blueprint("polaroid", __name__)
# --- OCR engine (lazy-loaded) ----------------------------------------------------
_ocr_engine = None
_OCR_AVAILABLE = False

def _get_ocr():
    """Lazy-load RapidOCR engine on first use."""
    global _ocr_engine, _OCR_AVAILABLE
    if _ocr_engine is not None:
        return _ocr_engine
    try:
        from rapidocr_onnxruntime import RapidOCR
        _ocr_engine = RapidOCR()
        _OCR_AVAILABLE = True
    except ImportError:
        _ocr_engine = False  # sentinel: attempted but not available
        _OCR_AVAILABLE = False
    return _ocr_engine

def _ocr_contains_blacklist(image_path: Path, words: list[str]) -> str | None:
    """Run OCR on an image and return the first blacklisted word found
    in the detected text, or None if clean.  Resizes to 400px for speed."""
    ocr = _get_ocr()
    if not ocr or ocr is False:
        return None
    try:
        import numpy as np
        img = Image.open(image_path)
        img.thumbnail((400, 400), Image.LANCZOS)
        img_array = np.array(img.convert("RGB"))
        result, _ = ocr(img_array)
        if not result:
            return None
        full_text = " ".join(line[1] for line in result).lower()
        for w in words:
            if w in full_text:
                return w
    except Exception:
        pass
    return None

# --- Progress infrastructure (same pattern as Module 1 in app.py) ----------
# --- Progress infrastructure (same pattern as Module 1 in app.py) ---
_pq: dict[str, queue.Queue] = {}
_pq_lock = threading.Lock()
from PIL import Image
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
                generate_pdf(batch, pdf_out, label, layout)
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
@polaroid_bp.route("/polaroid/list-orders", methods=["POST"])
def list_orders():
    """Return the list of all order folders inside a ship folder with
    metadata (variant, image count, whether already processed)."""
    data = request.get_json(silent=True) or {}
    ship_folder = data.get("ship_folder", "").strip()
    if not ship_folder or not Path(ship_folder).is_dir():
        return jsonify({"error": "Ship folder not found."}), 400

    ship_path = Path(ship_folder)
    output_dir = _output_folder_for(ship_path)

    orders = []
    for d in sorted(ship_path.iterdir(), key=lambda x: x.name.lower()):
        if not d.is_dir():
            continue
        folder_name = d.name
        variant = _detect_variant(folder_name)
        order_id = _extract_order_id(folder_name)
        special = _is_special(folder_name)
        images = _collect_images(d)
        already_done = (output_dir.is_dir() and 
                        _pdf_already_exists(output_dir, order_id, folder_name, special))
        orders.append({
            "folder": folder_name,
            "path": str(d),
            "variant": variant,
            "order_id": order_id or "",
            "photos": len(images),
            "already_done": already_done,
            "special": special,
        })

    return jsonify({
        "orders": orders,
        "output_dir": str(output_dir),
        "total": len(orders),
    })


@polaroid_bp.route("/polaroid/preview-order", methods=["POST"])
def preview_order():
    """Return low-res thumbnails + metadata for a single order folder,
    used to build an interactive preview in the browser."""
    import traceback
    data = request.get_json(silent=True) or {}
    order_path = data.get("order_path", "").strip()
    if not order_path or not Path(order_path).is_dir():
        return jsonify({"error": "Order folder not found."}), 400

    try:
        order_dir = Path(order_path)
        folder_name = order_dir.name
        variant = _detect_variant(folder_name)
        layout_key = _VARIANT_TO_LAYOUT[variant]
        layout = LAYOUTS[layout_key]
        ups = _UPS[variant]
        images = _collect_images(order_dir)
        num_photos = len(images)
        num_pdfs = math.ceil(num_photos / ups) if num_photos > 0 else 0

        thumbnails = []
        for i, img_path in enumerate(images):
            try:
                b64, auto_px, auto_py = make_thumbnail(img_path, layout)
                thumbnails.append({
                    "index": i,
                    "filename": Path(img_path).name,
                    "thumbnail": b64,
                    "mode": "fill",
                    "pan_x": auto_px,
                    "pan_y": auto_py,
                })
            except Exception:
                thumbnails.append({
                    "index": i,
                    "filename": Path(img_path).name,
                    "thumbnail": "",
                    "mode": "fill",
                    "pan_x": 0.5,
                    "pan_y": 0.5,
                })

        return jsonify({
            "folder": folder_name,
            "path": str(order_dir),
            "variant": variant,
            "layout_key": layout_key,
            "cols": layout["cols"],
            "rows": layout["rows"],
            "ups": ups,
            "rotate": layout["rotate"],
            "photos": num_photos,
            "num_pdfs": num_pdfs,
            "thumbnails": thumbnails,
        })
    except Exception as exc:
        traceback.print_exc()
        return jsonify({"error": f"Preview failed: {exc}"}), 500


@polaroid_bp.route("/polaroid/confirm-order", methods=["POST"])
def confirm_order():
    """Generate final PDF(s) for a single order with per-image adjustments."""
    data = request.get_json(silent=True) or {}
    order_path = data.get("order_path", "").strip()
    ship_folder = data.get("ship_folder", "").strip()
    adjustments = data.get("adjustments", [])  # [{mode, pan_x, pan_y}, ...]

    if not order_path or not Path(order_path).is_dir():
        return jsonify({"error": "Order folder not found."}), 400
    if not ship_folder or not Path(ship_folder).is_dir():
        return jsonify({"error": "Ship folder not found."}), 400

    order_dir = Path(order_path)
    ship_path = Path(ship_folder)
    folder_name = order_dir.name
    variant = _detect_variant(folder_name)
    layout_key = _VARIANT_TO_LAYOUT[variant]
    layout = LAYOUTS[layout_key]
    ups = _UPS[variant]

    images = _collect_images(order_dir)
    if not images:
        return jsonify({"error": "No images found in order folder."}), 400

    output_dir = _output_folder_for(ship_path)
    output_dir.mkdir(parents=True, exist_ok=True)

    num_photos = len(images)
    num_pdfs = math.ceil(num_photos / ups)
    pdf_names = []

    for pi in range(num_pdfs):
        start_idx = pi * ups
        end_idx = min(start_idx + ups, num_photos)
        batch = images[start_idx:end_idx]
        batch_adj = adjustments[start_idx:end_idx] if adjustments else None

        if num_pdfs == 1:
            label = folder_name
            pdf_filename = f"{folder_name}.pdf"
        else:
            label = f"{folder_name}-{pi + 1}/{num_pdfs}"
            pdf_filename = f"{folder_name}-{pi + 1}_{num_pdfs}.pdf"

        pdf_out = str(output_dir / pdf_filename)
        generate_pdf(batch, pdf_out, label, layout, adjustments=batch_adj)
        pdf_names.append(pdf_filename)

    log_activity("polaroid", "confirm_order", json.dumps({
        "folder": folder_name, "variant": variant,
        "photos": num_photos, "pdfs": num_pdfs,
    }))

    return jsonify({
        "ok": True,
        "folder": folder_name,
        "pdfs": pdf_names,
        "output_dir": str(output_dir),
    })


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
_CLEAN_BLACKLIST_WORDS = ["note:", "pixluv"]

@polaroid_bp.route("/polaroid/clean-orders", methods=["POST"])
def clean_orders():
    """Scan every order folder inside the selected ship folder and delete
    any file that is NOT an image, plus any image whose filename or
    visible text (via OCR) contains 'Note:' or 'PIXLUV' (case-insensitive).
    OCR runs in parallel across threads for speed."""
    data = request.get_json(silent=True) or {}
    ship_folder = data.get("ship_folder", "").strip()
    if not ship_folder or not Path(ship_folder).is_dir():
        return jsonify({"error": "Ship folder not found."}), 400

    ship_path = Path(ship_folder)
    deleted_files: list[dict] = []
    folders_with_deletes: set[str] = set()

    # --- Pass 1: fast checks (filename + extension) ---
    quick_deletes: list[tuple[Path, str, str]] = []  # (path, folder, reason)
    ocr_candidates: list[tuple[Path, str]] = []      # (path, folder)

    for d in sorted(ship_path.iterdir(), key=lambda x: x.name.lower()):
        if not d.is_dir():
            continue
        for f in list(d.iterdir()):
            if not f.is_file():
                continue
            fname_lower = f.name.lower()
            ext = f.suffix.lower()

            if ext not in SUPPORTED_EXTENSIONS:
                quick_deletes.append((f, d.name, f"non-image ({ext})"))
                continue

            # Check filename for blacklisted words
            hit = None
            for word in _CLEAN_BLACKLIST_WORDS:
                if word in fname_lower:
                    hit = word
                    break
            if hit:
                quick_deletes.append((f, d.name, f"filename contains '{hit}'"))
            else:
                ocr_candidates.append((f, d.name))

    # Delete quick matches immediately
    for fpath, folder, reason in quick_deletes:
        try:
            fpath.unlink()
            deleted_files.append({"folder": folder, "file": fpath.name, "reason": reason})
            folders_with_deletes.add(folder)
        except OSError:
            pass

    # --- Pass 2: parallel OCR on remaining images ---
    ocr_scanned = len(ocr_candidates)
    ocr_deletes: list[tuple[Path, str, str]] = []

    if ocr_candidates and _get_ocr() and _get_ocr() is not False:
        workers = min(os.cpu_count() or 4, 8, len(ocr_candidates))

        def _scan(item):
            fpath, folder = item
            hit = _ocr_contains_blacklist(fpath, _CLEAN_BLACKLIST_WORDS)
            if hit:
                return (fpath, folder, f"OCR detected '{hit}' in image")
            return None

        with ThreadPoolExecutor(max_workers=workers) as pool:
            for result in pool.map(_scan, ocr_candidates):
                if result:
                    ocr_deletes.append(result)

    # Delete OCR-flagged files
    for fpath, folder, reason in ocr_deletes:
        try:
            fpath.unlink()
            deleted_files.append({"folder": folder, "file": fpath.name, "reason": reason})
            folders_with_deletes.add(folder)
        except OSError:
            pass

    total_deleted = len(deleted_files)
    folders_cleaned = len(folders_with_deletes)

    log_activity("polaroid", "clean_orders", json.dumps({
        "ship_folder": ship_folder,
        "folders_cleaned": folders_cleaned,
        "total_deleted": total_deleted,
        "ocr_scanned": ocr_scanned,
    }))

    return jsonify({
        "folders_cleaned": folders_cleaned,
        "total_deleted": total_deleted,
        "ocr_scanned": ocr_scanned,
        "deleted_files": deleted_files,
    })