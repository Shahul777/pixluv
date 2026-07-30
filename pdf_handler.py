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
    """Split a combined PDF into separate PDFs by variant."""
    combined_path = Path(combined_pdf_path)
    if not combined_path.exists():
        return {"error": f"File not found: {combined_pdf_path}"}

    folder = combined_path.parent

    # Find all individual order PDFs (non-combined) in the folder
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
        return {"error": "No individual order PDFs found in the folder to determine variant boundaries."}

    # Sort them the same way the combine process does (by numeric prefix)
    individual_pdfs.sort(key=lambda x: x.name.lower())

    # Build page mapping: for each individual PDF, get its page count and variant
    page_map = []  # list of (variant, page_count, filename)
    for pdf_file in individual_pdfs:
        try:
            doc = fitz.open(str(pdf_file))
            pc = len(doc)
            doc.close()
            variant = _detect_variant(pdf_file.name)
            page_map.append((variant, pc, pdf_file.name))
        except Exception:
            continue

    if not page_map:
        return {"error": "Could not read any individual PDFs in the folder."}

    # Calculate total pages from individual PDFs
    total_individual_pages = sum(pc for _, pc, _ in page_map)

    # Open the combined PDF
    combined_doc = fitz.open(str(combined_path))
    combined_pages = len(combined_doc)

    # Verify page count matches (allow slight mismatch due to edge cases)
    if abs(combined_pages - total_individual_pages) > 0:
        # If mismatch, we still proceed but use the combined PDF's actual page count
        # and map proportionally or just use what we have
        pass

    # Group consecutive pages by variant
    # Build: variant -> list of page indices in the combined PDF
    variant_pages = OrderedDict()  # variant -> [page_indices]
    current_page = 0

    for variant, pc, _ in page_map:
        if current_page >= combined_pages:
            break
        # Clamp page count to remaining pages
        actual_pc = min(pc, combined_pages - current_page)
        if variant not in variant_pages:
            variant_pages[variant] = []
        for i in range(actual_pc):
            variant_pages[variant].append(current_page + i)
        current_page += actual_pc

    # Group variants for output:
    # - 4x3 -> separate PDF
    # - 4x6 -> separate PDF
    # - 3x3, 3x2, magnet -> combined into one PDF
    groups = OrderedDict()
    groups["4x3"] = []
    groups["4x6"] = []
    groups["others"] = []  # 3x3, 3x2, magnet

    # Track counts for naming
    other_counts = OrderedDict()  # variant -> count for the "others" group

    for variant, pages in variant_pages.items():
        if variant == "4x3":
            groups["4x3"].extend(pages)
        elif variant == "4x6":
            groups["4x6"].extend(pages)
        else:
            groups["others"].extend(pages)
            other_counts[variant] = len(pages)

    output_folder = folder  / "final_pdf_outputs"
    output_folder.mkdir(exist_ok=True)



    # Generate output PDFs
    output_files = []

    for group_name, pages in groups.items():
        if not pages:
            continue



        if group_name == "4x3":
            out_name = f"4x3-{len(pages)}.pdf"
        elif group_name == "4x6":
            out_name = f"4x6-{len(pages)}.pdf"
        else:
            # others: 3x3-10_3x2-15_magnet-2
            parts = []
            for v, c in other_counts.items():
                parts.append(f"{v}-{c}")
            out_name = "_".join(parts) + ".pdf" if parts else f"others-{len(pages)}.pdf"

        out_path = output_folder / out_name

        # Create new PDF with selected pages
        out_doc = fitz.open()
        # Pages are contiguous per group, use range extraction for speed
        if pages:
            start = pages[0]
            end = pages[-1]
            # Verify contiguous
            if end - start + 1 == len(pages) and end < len(combined_doc):
                out_doc.insert_pdf(combined_doc, from_page=start, to_page=end)
            else:
                # Fallback: page by page (if non-contiguous)
                for page_idx in pages:
                    if page_idx < len(combined_doc):
                        out_doc.insert_pdf(combined_doc, from_page=page_idx, to_page=page_idx)

        out_doc.save(str(out_path), garbage=0, deflate=False)
        out_doc.close()

        output_files.append({
            "filename": out_name,
            "pages": len(pages),
            "group": group_name,
        })

    combined_doc.close()

    # Remove the original combined PDF since we've split it
    # try:
    #     combined_path.unlink()
    # except Exception:
    #     pass

    return {
        "status": "ok",
        "output_files": output_files,
        "total_pages": combined_pages,
        # "original_removed": combined_path.name,
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