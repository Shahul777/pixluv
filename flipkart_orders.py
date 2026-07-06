import csv
import io
import os
import re
from datetime import datetime
from pathlib import Path

from flask import Blueprint, jsonify, request

from db import get_setting, set_setting

flipkart_bp = Blueprint("flipkart_orders", __name__)

SETTING_BASE_FOLDER = "flipkart_orders_base_folder"
REPORT_FILENAME = "FlipkartReport.txt"


def _sanitize_name(name: str) -> str:
    """Clean buyer name for folder naming: strip, replace spaces with _, remove invalid chars."""
    name = name.strip()
    name = re.sub(r"\s+", " ", name)
    name = re.sub(r'[<>:"/\\|?*]', "", name)
    return name.replace(" ", "_")

def _extract_order_digits(order_id: str) -> str:
    # Remove OD prefix and any quotes/spaces
    clean = order_id.strip().strip('"').strip()
    if clean.upper().startswith("OD"):
        clean = clean[2:]
    # Skip first 4, take next 4
    if len(clean) >= 8:
        return clean[4:8]
    return clean[:4] if len(clean) >= 4 else clean


def _parse_report(base_folder: Path) -> set:
    """Parse FlipkartReport.txt and return set of already processed order IDs."""
    report_path = base_folder / REPORT_FILENAME
    if not report_path.exists():
        return set()
        
    processed = set()
    try:
        for line in report_path.read_text(encoding="utf-8").splitlines():
            # Lines formatted as: "OrderID | BuyerName | Qty"
            parts = line.split("|")
            if len(parts) >= 1:
                oid = parts[0].strip()
                if oid.startswith("OD"):
                    processed.add(oid)
    except Exception:
        pass
    return processed


def _append_report(base_folder: Path, entries: list):
    """Append processed order entries to FlipkartReport.txt."""
    report_path = base_folder / REPORT_FILENAME
    with open(report_path, "a", encoding="utf-8") as f:
        if report_path.stat().st_size == 0 or not report_path.exists():
            f.write(f"# Flipkart Order Report\n")
            f.write(f"# Format: OrderID | BuyerName | Quantity | ProcessedDate\n")
            f.write(f"# {'='*60}\n")
        for entry in entries:
            f.write(f"{entry['order_id']} | {entry['buyer_name']} | {entry['quantity']} | {entry['date']}\n")


# --- Routes -----------------------------------------------------------

@flipkart_bp.route("/flipkart-orders/base-folder", methods=["GET"])
def get_base_folder():
    folder = get_setting(SETTING_BASE_FOLDER) or ""
    return jsonify({"folder": folder})


@flipkart_bp.route("/flipkart-orders/base-folder", methods=["POST"])
def set_base_folder():
    data = request.get_json(force=True)
    folder = data.get("folder", "").strip()
    if not folder:
        return jsonify({"error": "Folder path is required."}), 400
    if not os.path.isdir(folder):
        try:
            os.makedirs(folder, exist_ok=True)
        except Exception as e:
            return jsonify({"error": f"Cannot create folder: {e}"}), 400
    set_setting(SETTING_BASE_FOLDER, folder)
    return jsonify({"ok": True, "folder": folder})


@flipkart_bp.route("/flipkart-orders/process-csv", methods=["POST"])
def process_csv():
    """Upload and process Flipkart Order CSV."""
    base_folder_str = get_setting(SETTING_BASE_FOLDER)
    if not base_folder_str:
        return jsonify({"error": "Base folder not set. Please set it first."}), 400
    base_folder = Path(base_folder_str)
    if not base_folder.is_dir():
        return jsonify({"error": f"Base folder does not exist: {base_folder_str}"}), 400
        
    # Get uploaded CSV
    f = request.files.get("csv")
    if not f or f.filename == "":
        return jsonify({"error": "No CSV file uploaded."}), 400
        
    # Parse CSV
    try:
        content = f.read().decode("utf-8-sig")
        reader = csv.DictReader(io.StringIO(content))
        rows = list(reader)
    except Exception as e:
        return jsonify({"error": f"Failed to parse CSV: {e}"}), 400
        
    if not rows:
        return jsonify({"error": "CSV is empty."}), 400
        
    # Check required columns
    required = {"Order Id", "Buyer name", "Quantity"}
    if not required.issubset(set(rows[0].keys())):
        missing = required - set(rows[0].keys())
        return jsonify({"error": f"CSV missing columns: {', '.join(missing)}"}), 400
        
    # Load already processed orders
    processed_ids = _parse_report(base_folder)
    
    # Process each order
    results = []
    created = 0
    skipped = 0
    
    for row in rows:
        order_id = row["Order Id"].strip().strip('"')
        buyer_name = row["Buyer name"].strip()
        quantity = row["Quantity"].strip()
        
        # Skip if already processed
        if order_id in processed_ids:
            results.append({
                "order_id": order_id,
                "buyer_name": buyer_name,
                "status": "skipped",
                "folder": ""
            })
            skipped += 1
            continue
            
        # Build folder name
        digits = _extract_order_digits(order_id)
        safe_name = _sanitize_name(buyer_name)
        folder_name = f"{safe_name}-{digits}"
        
        # Create folder
        folder_path = base_folder / folder_name
        try:
            folder_path.mkdir(parents=True, exist_ok=True)
            created += 1
            results.append({
                "order_id": order_id,
                "buyer_name": buyer_name,
                "status": "created",
                "folder": folder_name
            })
        except Exception as e:
            results.append({
                "order_id": order_id,
                "buyer_name": buyer_name,
                "status": f"error: {e}",
                "folder": folder_name
            })
            
    # Update report with newly processed orders
    new_entries = [
        {
            "order_id": r["order_id"],
            "buyer_name": r["buyer_name"],
            "quantity": rows[i]["Quantity"].strip(),
            "date": datetime.now().strftime("%Y-%m-%d %H:%M")
        }
        for i, r in enumerate(results) if r["status"] == "created"
    ]
    if new_entries:
        _append_report(base_folder, new_entries)
        
    return jsonify({
        "ok": True,
        "total": len(rows),
        "created": created,
        "skipped": skipped,
        "results": results
    })