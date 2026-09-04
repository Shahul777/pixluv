"""PDF Handler - Split combined PDF by variant & Merge multiple PDFs."""

import os
import re
from pathlib import Path
from collections import OrderedDict

import fitz  # PyMuPDF

# Variant detection (same logic as app.py)
VARIANT_ORDER = {"4x3": 0, "3x2": 1, "3x3": 2, "4x6": 3, "magnet": 4}
VARIANT_RE = re.compile(r"(3x2|3x3|4x6|magnet)", re.IGNORECASE)
PREFIX_RE = re.compile(r"^\d{2,}_")
COMBINED_RE = re.compile(r"combined", re.IGNORECASE)


def _detect_variant(filename: str) -> str:
    m = VARIANT_RE.search(filename)
    return m.group(1).lower() if m else "4x3"


def _is_combined(filename: str) -> bool:
    return bool(COMBINED_RE.search(filename))

def split_combined_pdf(combined_pdf_path: str) -> dict:
    """Group individual order PDFs by variant (from filename) and combine each group."""

    input_path = Path(combined_pdf_path)

    # Accept either a folder or a file (use its parent folder)
    if input_path.is_dir():
        folder = input_path
    elif input_path.is_file():
        folder = input_path.parent
    else:
        return {"error": f"Path not found: {combined_pdf_path}"}

    # Collect all PDFs in the folder (skip subfolders and combined PDFs)
    individual_pdfs = []
    for f in folder.iterdir():
        if not f.is_file():
            continue
        if f.suffix.lower() != ".pdf":
            continue
        if _is_combined(f.stem):
            continue
        individual_pdfs.append(f)

    if not individual_pdfs:
        return {"error": "No individual order PDFs found in the folder."}

    # Sort by name ascending
    individual_pdfs.sort(key=lambda x: x.name.lower())

    # Group PDFs by variant based on filename only
    # 4x3 -> separate | 4x6 -> separate | 3x2, 3x3, magnet -> others
    groups = OrderedDict()
    groups["4x3"] = []
    groups["4x6"] = []
    groups["others"] = []

    other_page_counts = OrderedDict()  # variant -> list of pdf files for naming

    for pdf_file in individual_pdfs:
        variant = _detect_variant(pdf_file.name)
        if variant == "4x3":
            groups["4x3"].append(pdf_file)
        elif variant == "4x6":
            groups["4x6"].append(pdf_file)
        else:
            groups["others"].append(pdf_file)
            other_page_counts.setdefault(variant, 0)

    # Create output folder
    output_folder = folder / "final_pdf_outputs"
    output_folder.mkdir(exist_ok=True)

    output_files = []
    total_pages = 0

    for group_name, pdf_list in groups.items():
        if not pdf_list:
            continue

        # Merge all PDFs in this group (already in ascending name order)
        out_doc = fitz.open()
        group_pages = 0

        for pdf_file in pdf_list:
            try:
                src = fitz.open(str(pdf_file))
                pc = len(src)
                out_doc.insert_pdf(src)
                group_pages += pc
                src.close()
                # Track per-variant page counts for "others" naming
                if group_name == "others":
                    v = _detect_variant(pdf_file.name)
                    other_page_counts[v] = other_page_counts.get(v, 0) + pc
            except Exception as e:
                out_doc.close()
                return {"error": f"Failed to read {pdf_file.name}: {e}"}

        # Build output filename
        if group_name == "4x3":
            out_name = f"4x3-{group_pages}pages.pdf"
        elif group_name == "4x6":
            out_name = f"4x6-{group_pages}pages.pdf"
        else:
            # others: 3x3-10_3x2-15_magnet-2
            parts = []
            for v, pc in other_page_counts.items():
                if pc > 0:
                    parts.append(f"{v}-{pc}")
            out_name = "_".join(parts) + ".pdf" if parts else f"others-{group_pages}.pdf"

        out_path = output_folder / out_name
        out_doc.save(str(out_path), garbage=0, deflate=False)
        out_doc.close()

        total_pages += group_pages
        output_files.append({
            "filename": out_name,
            "pages": group_pages,
            "group": group_name,
            "pdf_count": len(pdf_list),
        })

    return {
        "status": "ok",
        "output_files": output_files,
        "total_pages": total_pages,
    }

def split_pdf_half(pdf_path: str) -> dict:
    """Split a PDF into two halves. Removes the original after splitting."""
    path = Path(pdf_path)
    if not path.exists():
        return {"error": f"File not found: {pdf_path}"}
    if path.suffix.lower() != ".pdf":
        return {"error": "Not a PDF file."}

    doc = fitz.open(str(path))
    total = len(doc)
    if total < 2:
        doc.close()
        return {"error": "PDF must have at least 2 pages to split."}

    # First half gets the extra page if odd
    half1 = (total + 1) // 2
    half2 = total - half1

    folder = path.parent

    # Part 1
    name1 = f"split-{half1}-part1.pdf"
    out1 = fitz.open()
    out1.insert_pdf(doc, from_page=0, to_page=half1 - 1)
    out1.save(str(folder / name1), garbage=0, deflate=False)
    out1.close()

    # Part 2
    name2 = f"split-{half2}-part2.pdf"
    out2 = fitz.open()
    out2.insert_pdf(doc, from_page=half1, to_page=total - 1)
    out2.save(str(folder / name2), garbage=0, deflate=False)
    out2.close()

    doc.close()

    # Remove original
    try:
        path.unlink()
    except Exception:
        pass

    return {
        "status": "ok",
        "total_pages": total,
        "files": [
            {"filename": name1, "pages": half1},
            {"filename": name2, "pages": half2},
        ],
    }

def merge_pdfs(pdf_paths: list[str]) -> dict:

    if not pdf_paths or len(pdf_paths) < 2:
        return {"error": "At least 2 PDFs are required for merging."}

    # Validate all paths exist
    paths = [Path(p) for p in pdf_paths]
    for p in paths:
        if not p.exists():
            return {"error": f"File not found: {p}"}
        if p.suffix.lower() != ".pdf":
            return {"error": f"Not a PDF file: {p}"}

    first_path = paths[0]

    # Open the first PDF
    try:
        first_doc = fitz.open(str(first_path))
    except Exception as e:
        return {"error": f"Cannot open first PDF: {e}"}

    original_pages = len(first_doc)
    added_pages = 0

    # Append pages from each subsequent PDF
    for p in paths[1:]:
        try:
            src_doc = fitz.open(str(p))
            page_count = len(src_doc)
            first_doc.insert_pdf(src_doc)
            src_doc.close()
            added_pages += page_count
        except Exception as e:
            first_doc.close()
            return {"error": f"Error reading {p.name}: {e}"}

    # Save to the renamed path: merged-{total_pages}.pdf
    total_pages = len(first_doc)
    final_name = f"merged-{total_pages}.pdf"
    final_path = first_path.parent / final_name

    try:
        first_doc.save(str(final_path), deflate=False, garbage=0)
    except Exception as e:
        first_doc.close()
        return {"error": f"Failed to save merged PDF: {e}"}

    first_doc.close()

    # Remove the original first PDF since we saved with new name
    if first_path != final_path:
        try:
            first_path.unlink()
        except Exception:
            pass

    return {
        "status": "ok",
        "first_pdf": final_name,
        "original_pages": original_pages,
        "added_pages": added_pages,
        "total_pages": total_pages,
        "merged_count": len(paths),
    }