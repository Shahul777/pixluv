import datetime
import json
import logging
import os
import queue
import re
import shutil
import threading
import time
from pathlib import Path


from flask import Blueprint, Response, jsonify, request

from db import get_setting, set_setting, log_activity

amazon_bp = Blueprint("amazon_download", __name__)
log = logging.getLogger("amazon_download")
log.setLevel(logging.INFO)
if not log.handlers:
    _h = logging.StreamHandler()
    _h.setFormatter(logging.Formatter("%(asctime)s %(name)s %(message)s"))
    log.addHandler(_h)

# --- Settings -----------------------------------------------------------------
SETTING_BASE_FOLDER = "amazon_download_base_folder"
BROWSER_DATA_DIR = Path(__file__).parent / ".amazon_browser_data"
SELLER_CENTRAL_URL = "https://sellercentral.amazon.in"
ORDERS_URL = f"{SELLER_CENTRAL_URL}/orders-v3/ref=xx_myo_favb_xx?page=1"

# --- Progress infrastructure --------------------------------------------------
_pq: dict[str, queue.Queue] = {}
_pq_lock = threading.Lock()
_stop_flags: dict[str, bool] = {}

# --- Browser state ------------------------------------------------------------
_browser_context = None
_browser_instance = None
_playwright_instance = None
_browser_lock = threading.Lock()

# --- Day-of-week mapping ------------------------------------------------------
DAY_ABBREVS = {
    "Mon": ["Mon", "Monday"],
    "Tue": ["Tue", "Tuesday"],
    "Wed": ["Wed", "Wednesday"],
    "Thu": ["Thu", "Thursday"],
    "Fri": ["Fri", "Friday"],
    "Sat": ["Sat", "Saturday"],
    "Sun": ["Sun", "Sunday"],
}

# --- Variant detection --------------------------------------------------------
# Matches "4 x 3", "4x3", "4 x 3", "3 x 2" etc. in product title
_VARIANT_RE = re.compile(r"(\d)\s*[xX]\s*(\d)")
_VARIANT_PRIORITY = ["4x6", "3x2", "3x3"]
_KNOWN_VARIANTS = {"4x3", "4x6", "3x2", "3x3"}

def _detect_variant_from_title(title: str) -> str:

    # Normalize: replace various separators to standard form for matching
    normalized = re.sub(r'(\d)\s*[xX×]\s*(\d)', r'\1x\2', title)
    normalized_lower = normalized.lower()

    # Check non-default variants first (priority)
    for v in _VARIANT_PRIORITY:
        if v in normalized_lower:
            return v

    # Check for 4x3 (default)
    if "4x3" in normalized_lower:
        return "4x3"
    # Fallback: regex scan
    matches = _VARIANT_RE.findall(title)
    for w, h in matches:
        key = f"{w}x{h}"
        if key in _KNOWN_VARIANTS and key != "4x3":
            return key

    return "4x3"  # default
REPORT_FILENAME = "amazonDownload_report.txt"

def _parse_report_order_ids(base_folder: Path) -> set:
    """Parse the existing amazonDownload_report.txt and extract all order IDs (full 3-digit+ IDs)
    that have been previously processed/downloaded. Returns a set of full order IDs."""
    report_path = base_folder / REPORT_FILENAME
    if not report_path.exists():
        return set()

    processed_ids = set()
    try:
        content = report_path.read_text(encoding="utf-8")
        # Look for lines in PROCESSED ORDERS section with format:
        # "   1. order_id | name | variants | status"
        # or lines containing order IDs (pattern: ###-#######-#######)
        for line in content.split("\n"):
            # Match full Amazon order IDs like "408-1234567-1234567"
            matches = re.findall(r"\d{3}-\d{7}-\d{7}", line)
            for m in matches:
                processed_ids.add(m)
    except Exception:
        pass

    return processed_ids
def _sanitize_name(name: str) -> str:
    """Clean up buyer name for use in folder naming."""
    # Remove special characters, keep alphanumeric and spaces
    name = name.strip()
    # Replace multiple spaces with single
    name = re.sub(r"\s+", " ", name)
    # Remove characters that are invalid in folder names
    name = re.sub(r'[<>:"/\\|?*]', "", name)
    # Convert to lowercase, replace spaces with nothing (or keep as-is)
    # Based on existing convention: shahul-1234, so use lowercase
    return name.strip().lower().replace(" ", "_")

def _build_folder_name(name: str, order_id_4: str, variant: str, photo_count: str = "") -> str:
    """Build folder name based on variant type.
    4x3: name-XXXX
    3x2/3x3: name-XXXX-(count)-variant
    Others: name-XXXX-variant
    """
    clean_name = _sanitize_name(name)
    if variant == "4x3":
        return f"{clean_name}-{order_id_4}"
    elif variant in ("3x2", "3x3") and photo_count:
        return f"{clean_name}-{order_id_4}-({photo_count})-{variant}"
    else:
        return f"{clean_name}-{order_id_4}-{variant}"

def _folder_exists(base_folder: Path, order_id_4: str, name: str, variant: str) -> bool:
    """Check if a folder for this order already exists in base folder or WhatsApp subfolder.
    Matches on name-orderID_last4 pattern only (ignores variant suffix).
    """
    # safe_name = re.sub(r'[<>:"/\\|?*]', '', name).strip().lower()
    # pattern = f"{safe_name}-{order_id_4}"

    safe_name = _sanitize_name(name)
    tag = f"{safe_name}-{order_id_4}"
    def _scan(folder: Path) -> bool:
        if not folder.exists():
            return False
        try:
            for d in folder.iterdir():
                if not d.is_dir():
                    continue
                low = d.name.lower()
                # Matches: name-1234, name-1234-4x6, name-1234-(12)-3x2, etc.
                if low.startswith(tag):
                    return True
        except Exception:
            pass
        return False

    if _scan(base_folder):
        return True
    if _scan(base_folder / "WhatsApp"):
        return True
    return False

  

# --- Queue helpers ------------------------------------------------------------

def _new_q(task_id: str) -> queue.Queue:
    q: queue.Queue = queue.Queue()
    with _pq_lock:
        _pq[task_id] = q
        _stop_flags[task_id] = False
    return q

def _get_q(task_id: str) -> queue.Queue | None:
    with _pq_lock:
        return _pq.get(task_id)

def _del_q(task_id: str):
    with _pq_lock:
        _pq.pop(task_id, None)
        _stop_flags.pop(task_id, None)

def _is_stopped(task_id: str) -> bool:
    with _pq_lock:
        return _stop_flags.get(task_id, False)

def _emit(q: queue.Queue, **kw):
    q.put(kw)


# --- Browser Management -------------------------------------------------------

def _ensure_browser_dir():
    """Ensure the browser data directory exists."""
    BROWSER_DATA_DIR.mkdir(parents=True, exist_ok=True)

def _launch_browser():
    """Launch a persistent Playwright browser context.
    Returns (playwright, browser, context) tuple.
    """
    global _playwright_instance, _browser_instance, _browser_context
    
    with _browser_lock:
        if _browser_context is not None:
            try:
                # Test if context is still alive
                _browser_context.pages
                return _playwright_instance, _browser_instance, _browser_context
            except Exception:
                # Context is dead, clean up
                _browser_context = None
                _browser_instance = None
                _playwright_instance = None

        from playwright.sync_api import sync_playwright

        _ensure_browser_dir()

        pw = sync_playwright().start()
        _playwright_instance = pw

        # Launch with persistent context (stores cookies, localStorage, etc.)
        context = pw.chromium.launch_persistent_context(
            user_data_dir=str(BROWSER_DATA_DIR),
            headless=False,
            viewport={"width": 1400, "height": 900},
            args=[
                "--disable-blink-features=AutomationControlled",
                "--no-first-run",
                "--no-default-browser-check",
            ],
            ignore_default_args=["--enable-automation"],
            channel="chromium",
        )
        _browser_context = context
        _browser_instance = None  # persistent context doesn't have separate browser
        
        log.info("Browser launched with persistent context at %s", BROWSER_DATA_DIR)
        return pw, None, context

def _close_browser():
    """Close the browser and playwright instance."""
    global _playwright_instance, _browser_instance, _browser_context
    
    with _browser_lock:
        if _browser_context:
            try:
                _browser_context.close()
            except Exception:
                pass
            _browser_context = None
        if _playwright_instance:
            try:
                _playwright_instance.stop()
            except Exception:
                pass
            _playwright_instance = None
            _browser_instance = None

def _is_logged_in(page) -> bool:
    """Check if the current page shows Seller Central logged-in state."""
    try:
        # Check for common logged-in indicators
        # The navbar should have seller name or account menu
        content = page.content()
        if "Sign in" in content and "sellercentral" not in page.url:
            return False
        if "sellercentral.amazon.in" in page.url:
            # Check if we're on a login/auth page
            if "/ap/signin" in page.url or "/ap/oa" in page.url:
                return False
            return True
        return False
    except Exception:
        return False

def _wait_for_login(page, q: queue.Queue, timeout: int = 300) -> bool:
    """Navigate to Seller Central and wait for user to log in if needed."""
    try:
        page.goto(ORDERS_URL, wait_until="domcontentloaded", timeout=20000)
        time.sleep(1)
    except Exception as e:
        log.warning("Navigation timeout, checking state: %s", e)

    # Check if we're already logged in
    if _is_logged_in(page):
        return True

    # We're on a login page - wait for user to manually log in
    _emit(q, stage="login", pct=0, 
          detail="Please log in to Amazon Seller Central in the browser window...", 
          done=False, needs_login=True)

    start_time = time.time()
    while time.time() - start_time < timeout:
        time.sleep(2)
        try:
            if _is_logged_in(page):
                _emit(q, stage="login", pct=5, 
                      detail="Login successful! Proceeding...", done=False)
                return True
            # Also check if URL changed to orders page
            if "orders-v3" in page.url or "sellercentral.amazon.in" in page.url:
                if "/ap/" not in page.url:
                    return True
        except Exception:
            pass
            
    return False


# --- Scraping Logic -----------------------------------------------------------

def _navigate_to_unshipped(page, q: queue.Queue) -> bool:
    """Navigate to the Unshipped orders tab."""
    try:
        # Go to orders page if not already there
        if "orders-v3" not in page.url:
            page.goto(ORDERS_URL, wait_until="domcontentloaded", timeout=20000)
            time.sleep(2)

        # Wait for page to be usable
        try:
            page.wait_for_load_state("networkidle", timeout=8000)
        except Exception:
            time.sleep(1)

        # Click on "Unshipped" tab/link if available
        try:
            unshipped_tab = page.locator(
                "a:has-text('Unshipped'), "
                "button:has-text('Unshipped'), "
                "span:has-text('Unshipped'), "
                "[data-test-id*='unshipped'], "
                "li:has-text('Unshipped')"
            )
            if unshipped_tab.count() > 0:
                unshipped_tab.first.click()
                time.sleep(2)
        except Exception:
            pass

        # Verify we can see orders (page loaded successfully)
        time.sleep(1)

        return True
    except Exception as e:
        log.error("Failed to navigate to unshipped: %s", e)
        return False

def _extract_ship_day(ship_text: str) -> str | None:
    # Strategy 1: Look for "Ship by" followed by a day abbreviation
    m = re.search(
    r'Ship\s*by[^:]*:\s*(Mon|Tue|Wed|Thu|Fri|Sat|Sun)',
    ship_text, re.IGNORECASE
            )
    if m:
        raw = m.group(1).strip().capitalize()
        for abbrev in DAY_ABBREVS:
            if raw.startswith(abbrev):
                return abbrev

# Strategy 2: Parse date like "1 Jul 2026" or "1 Jul, 2026" and get weekday
    date_m = re.search(r'(\d{1,2})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*[,.]?\s*(\d{4})', ship_text, re.IGNORECASE)
    if date_m:
        try:
            from datetime import datetime
            date_str = f"{date_m.group(1)} {date_m.group(2)[:3]} {date_m.group(3)}"
            dt = datetime.strptime(date_str, "%d %b %Y")
            day_name = dt.strftime("%a") # Mon, Tue, Wed, ...
            return day_name
        except Exception:
            pass

    # Strategy 3: Direct day abbreviation search (fallback)
    for abbrev in DAY_ABBREVS:
        if re.search(rf'\b{abbrev}\b', ship_text, re.IGNORECASE):
            return abbrev
        for variant in DAY_ABBREVS[abbrev]:
            if variant.lower() in ship_text.lower():
                return abbrev
    return None

def _scrape_order_list(page, ship_day: str, q: queue.Queue) -> list[dict]:
    """Scrape all orders from the unshipped orders page and filter by ship day.
    
    Returns list of dicts: {order_id, order_id_4, ship_day, items: [{title, variant, order_item_id}]}
    """
    orders = []
    
    _emit(q, stage="scan", pct=5, detail="Scanning unshipped orders...", done=False)
    
    # Wait for the orders table to load
    try:
        page.wait_for_selector(
            "a[href*='orders-v3/order/']",
            timeout=10000
        )
 
    except Exception:
        pass
            
    time.sleep(1)
    
    # Strategy: Find all order rows by looking for order ID links
    # Amazon order IDs are in format: XXX-XXXXXXX-XXXXXXX
    order_links = page.locator("a[href*='orders-v3/order/']")
    link_count = order_links.count()
    seen_ids = set()

    for li in range(link_count):
        try:
            link_el = order_links.nth(li)
            href = link_el.get_attribute("href") or ""
            # Extract order ID from href
            m = re.search(r"(\d{3}-\d{7}-\d{7})", href)
            if not m:
                continue
            order_id = m.group(1)

            # Skip duplicates (same order ID may appear multiple times on page)
            if order_id in seen_ids:
                continue
            seen_ids.add(order_id)

            order_id_4 = order_id[-4:]

            # Navigate up to the closest <tr> ancestor for THIS order only
            row = link_el.locator("xpath=ancestor::tr[1]")
            if row.count() == 0:
                # Fallback: try a smaller parent container
                row = link_el.locator("xpath=ancestor::div[1]")
                if row.count() == 0:
                    continue

            row_text = row.first.inner_text()

            # Extract ship date from THIS row's text only
            detected_day = _extract_ship_day(row_text)

            # Extract product title from THIS row
            # Look for the product title link in this row
            product_link = row.first.locator(
                "a[href*='catalog'], a[href*='product'], "
                "td:nth-child(4) a, td:nth-child(5) a"
            )
            title = ""
            if product_link.count() > 0:
                title = product_link.first.inner_text().strip()

            # If no product link found, try to extract title from row text
            if not title:
                # Look for PixLuv/PIXLUV pattern in row text
                title_m = re.search(
                    r'((?:PixLuv|Pixluv|PIXLUV)[^\|]{10,})',
                    row_text, re.IGNORECASE
                )
                if title_m:
                    title = title_m.group(1).strip()

            # Detect variant from this order's product title
            variant = _detect_variant_from_title(title) if title else "4x3"
            items = [{"title": title[:150], "variant": variant}] if title else []

            if detected_day and detected_day.lower() != ship_day.lower():
                log.info("  SKIP %s (ship: %s != %s)", order_id, detected_day, ship_day)
                continue

            # If no day detected and no product info, skip
            if not detected_day and not items:
                continue

            orders.append({
                "order_id": order_id,
                "order_id_4": order_id_4,
                "ship_day": detected_day,
                "items": items,
                "ship_text": row_text[:200],
            })

        except Exception as e:
            log.warning("Error processing order link %d: %s", li, e)
            continue
    log.info("=" * 60)
    log.info("SCAN RESULTS: %d orders matched ship day '%s'", len(orders), ship_day)
    log.info("-" * 60)
    for i, o in enumerate(orders, 1):
        variants = ", ".join(it.get("variant", "?") for it in o.get("items", [])) or "?"
        log.info("  %d. %s  (ship: %s, variant: %s)",
             i, o["order_id"], o.get("ship_day", "?"), variants)
    log.info("=" * 60)
        
    _emit(q, stage="scan", pct=10, 
          detail=f"Found {len(orders)} orders matching ship day '{ship_day}'", 
          done=False)
          
    return orders
def _get_order_details(page, order_id: str, q: queue.Queue) -> dict | None:
    """Navigate to order detail page and extract:
    - Ship-to name
    - All order items with their orderItemId
    - Ship-by date (for verification)
    
    Returns dict: {name, ship_day, items: [{title, variant, order_item_id}]}
    """
    detail_url = f"{SELLER_CENTRAL_URL}/orders-v3/order/{order_id}"
    
    try:
        page.goto(detail_url, wait_until="domcontentloaded", timeout=20000)
        time.sleep(1)
        try:

            page.wait_for_load_state("networkidle", timeout=8000)
        except Exception:
            time.sleep(1)
    except Exception:
        time.sleep(2)
        
    time.sleep(2)
    
    # Extract "Ship to" name
    ship_to_name = None
    try:
        # The "Ship to" section contains the recipient name
        # Look for the Ship to heading and get the name below it
        ship_to_section = page.locator(
            "div:has(> h2:has-text('Ship to')), "
            "div:has(> h3:has-text('Ship to')), "
            "div:has(> *:has-text('Ship to'))"
        )
        if ship_to_section.count() > 0:
            section_text = ship_to_section.first.inner_text()
            # Name is typically the first line after "Ship to"
            lines = [l.strip() for l in section_text.split("\n") if l.strip()]
            for i, line in enumerate(lines):
                if "ship to" in line.lower():
                    # Next non-empty line should be the name
                    if i + 1 < len(lines):
                        candidate = lines[i + 1].strip()
                        # Skip if it looks like an address part
                        if candidate and not candidate.startswith("#") and \
                           not re.match(r"^\d", candidate) and \
                           len(candidate) < 50:
                            ship_to_name = candidate
                            break
    except Exception as e:
        log.warning("Could not extract Ship To name: %s", e)
        
    # Fallback: try to find name directly
    if not ship_to_name:
        try:
            # Look for bold text in the ship-to section
            name_el = page.locator(
                ".ship-to-name, "
                "[data-test-id='ship-to-name'], "
                "div.a-row.a-spacing-none.a-size-base.a-text-bold"
            )
            if name_el.count() > 0:
                ship_to_name = name_el.first.inner_text().strip()
        except Exception:
            pass
            
    # Another fallback: regex search in page content
    if not ship_to_name:
        try:
            content = page.content()
            # Look for text after "Ship to" header
            m = re.search(
                r"Ship\s+to.*?[<[^>]*>]\s*([A-Z][a-z]+(?:\s+[A-Z][a-z]+)*)",
                content, re.DOTALL
            )
            if m:
                ship_to_name = m.group(1).strip()
        except Exception:
            pass
            
    if not ship_to_name:
        # Last resort: try the "Contact Buyer" name area
        try:
            buyer_el = page.locator("a:has-text('See all'), span:has-text('Contact Buyer')")
            if buyer_el.count() > 0:
                parent_text = buyer_el.first.locator("xpath=..").inner_text()
                # Extract name near "Contact Buyer"
                lines = parent_text.split("\n")
                for line in lines:
                    line = line.strip()
                    if line and "Contact" not in line and "See all" not in line \
                       and "order" not in line.lower() and len(line) < 30:
                        ship_to_name = line
                        break
        except Exception:
            pass

    # Extract ship-by date
    ship_day = None
    try:
        ship_by_el = page.locator("text=/Ship by/i")
        if ship_by_el.count() > 0:
            ship_text = ship_by_el.first.locator("xpath=..").inner_text()
            ship_day = _extract_ship_day(ship_text)
    except Exception:
        pass

    # Extract order items with their Order Item IDs
    items = []
    try:
        # Each item in "Order contents" section has product title and Order Item ID
        # Look for all product rows
        content = page.content()
        
        # Find orderItemId values - they appear in links or data attributes
        # Pattern: orderItemId=XXXXXXXXXX
        item_id_pattern = re.compile(r"orderItemId[=:]\s*(\d{10,})")
        found_item_ids = list(set(item_id_pattern.findall(content)))
        
        # Also look for "Order Item ID:" text patterns
        item_id_text_pattern = re.compile(r"Order\s*Item\s*ID:\s*(\d+)")
        found_item_ids.extend(item_id_text_pattern.findall(content))
        
        log.info("Order %s: found item IDs: %s", order_id, found_item_ids)
        
        # For each item, try to get its title/variant
        # Look for product title links on this page
        product_els = page.locator(
            "a[href*='catalog'], a[href*='sellercentral'][href*='product'], "
            "td.product-name-column a, a.product-title"
        )
        
        titles_found = []
        for i in range(min(product_els.count(), 10)):
            try:
                title = product_els.nth(i).inner_text().strip()
                if title and len(title) > 20 and "PIXLUV" in title.upper() or "photo" in title.lower():
                    titles_found.append(title)
            except Exception:
                continue
                
        # Also search for product titles via regex
        title_matches = re.findall(
            r"(?:PixLuv|PIXLUV)[^<]{20,200}?(?:photo|print|poster|frame)",
            content, re.IGNORECASE
        )
        for t in title_matches:
            clean_t = re.sub(r"<[^>]+>", "", t).strip()
            if clean_t and clean_t not in titles_found:
                titles_found.append(clean_t)
                
        # Remove duplicates while preserving order
        seen = set()
        unique_titles = []
        for t in titles_found:
            t_key = t[:50].lower()
            if t_key not in seen:
                seen.add(t_key)
                unique_titles.append(t)
                
        # Match items with IDs
        # If we have same count, zip them; otherwise, create items with what we have
        if len(unique_titles) == len(found_item_ids):
            for title, item_id in zip(unique_titles, found_item_ids):
                variant = _detect_variant_from_title(title)
                items.append({
                    "title": title,
                    "variant": variant,
                    "order_item_id": item_id,
                })
        elif unique_titles:
            # Try to associate items by finding item IDs near each title
            for title in unique_titles:
                variant = _detect_variant_from_title(title)
                # Try to find the closest orderItemId
                item_id = None
                
                # Look for "Customisation Information" links which contain orderItemId
                cust_links = page.locator(f"a[href*='orderItemId']")
                for ci in range(cust_links.count()):
                    try:
                        href = cust_links.nth(ci).get_attribute("href")
                        if href:
                            m = re.search(r"orderItemId=(\d+)", href)
                            if m:
                                cid = m.group(1)
                                if cid not in [it.get("order_item_id") for it in items]:
                                    item_id = cid
                                    break
                    except Exception:
                        continue
                        
                if not item_id and found_item_ids:
                    # Use remaining IDs
                    used = {it.get("order_item_id") for it in items}
                    remaining = [x for x in found_item_ids if x not in used]
                    if remaining:
                        item_id = remaining[0]
                        
                items.append({
                    "title": title,
                    "variant": variant,
                    "order_item_id": item_id,
                })
        elif found_item_ids:
            # We have IDs but no titles - create placeholder items
            for item_id in found_item_ids:
                items.append({
                    "title": "Unknown Product",
                    "variant": "4x3",
                    "order_item_id": item_id,
                })
    except Exception as e:
        log.warning("Error extracting order items for %s: %s", order_id, e)

    # Extract quantities for each item from the page
    try:
        # Quantities appear in the "Quantity" column of the order contents table
        qty_matches = re.findall(
            r'(?:Quantity|Qty)[^<]*?</(?:th|td)>[^<]*<(?:td|td)[^>]*>\s*(\d+)',
            page.content(), re.IGNORECASE | re.DOTALL
        )
        
        # Fallback: look for Quantity column values in table rows
        if not qty_matches:
            rows = page.locator("table tr, div[class*='order-item']")
            for ri in range(rows.count()):
                try:
                    row_text = rows.nth(ri).inner_text()
                    # Look for standalone numbers that likely represent quantity
                    if "Quantity" in row_text:
                        continue  # header row
                    qty_m = re.search(r'\b(\d+)\s*₹', row_text)
                    if qty_m:
                        qty_matches.append(qty_m.group(1))
                except Exception:
                    continue
                    
        # Assign quantities to items (in order)
        for idx, item in enumerate(items):
            if idx < len(qty_matches):
                item["quantity"] = int(qty_matches[idx])
            else:
                item["quantity"] = 1
    except Exception:
        for item in items:
            item.setdefault("quantity", 1)

    # Always click "Show more" buttons to reveal Customisation text (WhatsApp info etc.)
    try:
        show_more_btns = page.locator(
            "button:has-text('Show more'), a:has-text('Show more'), "
            "span:has-text('Show more')"
        )
        for i in range(show_more_btns.count()):
            try:
                show_more_btns.nth(i).click()
                time.sleep(0.3)
            except Exception:
                continue
        if show_more_btns.count() > 0:
            time.sleep(1)  
    except Exception:
        pass

    # If we still have no items, look for Customisation Information links
    if not items:
        try:
            # Now look for Customisation Information links
            cust_links = page.locator("a:has-text('Customisation Information')")
            for i in range(cust_links.count()):
                try:
                    href = cust_links.nth(i).get_attribute("href") or ""
                    m = re.search(r"orderItemId=(\d+)", href)
                    item_id = m.group(1) if m else None

                    # Try to get the variant from nearby product title
                    parent = cust_links.nth(i).locator("xpath=ancestor::tr[1]|ancestor::div[contains(@class,'item')]")
                    title = ""
                    if parent.count() > 0:
                        title = parent.first.inner_text()[:200]
                    variant = _detect_variant_from_title(title) if title else "4x3"

                    items.append({
                        "title": title[:100] if title else "Product",
                        "variant": variant,
                        "order_item_id": item_id,
                    })
                except Exception:
                    continue
        except Exception as e:
            log.warning("Fallback item extraction failed: %s", e)

    log.info("Order %s: name=%s, ship_day=%s, items=%d",
             order_id, ship_to_name, ship_day, len(items))

    # --- WhatsApp order detection ---
    # After items are extracted (and "Show more" has been clicked),
    # check if the page contains the WhatsApp pattern
    whatsapp_info = None
    try:
        page_text = page.inner_text("body")
        # Pattern: "SEND PHOTOS WITHIN 1 HOUR: <number>"
        wa_match = re.search(
            r"SEND\s+PHOTOS\s+WITHIN\s+1\s+HOUR[:\s]*(\d{7,15})",
            page_text, re.IGNORECASE
        )
        if wa_match:
            wa_number = wa_match.group(1).strip()
            
            # Extract "Included Components:" value
            inc_match = re.search(
                r"Included\s+Components?:\s*(.+?)(?:\n|SHARE|$)",
                page_text, re.IGNORECASE
            )
            included_component = inc_match.group(1).strip() if inc_match else ""
            
            whatsapp_info = {
                "whatsapp_number": wa_number,
                "included_component": included_component,
            }
            log.info("Order %s is a WhatsApp order: number=%s, component=%s",
                     order_id, wa_number, included_component)
    except Exception as e:
        log.warning("WhatsApp detection error for %s: %s", order_id, e)

    return {
        "name": ship_to_name,
        "ship_day": ship_day,
        "items": items,
        "whatsapp": whatsapp_info,
    }
def _download_zip_for_item(page, order_id: str, order_item_id: str,
                           dest_folder: Path, q: queue.Queue) -> bool:
    """Navigate to the Customisation Information page and download the zip.
    
    Returns True on success.
    """
    if not order_item_id:
        log.warning("No order_item_id for order %s, cannot download", order_id)
        return False
        
    # Construct the customisation URL
    cust_url = (
        f"{SELLER_CENTRAL_URL}/gestalt/fulfillment/index.html"
        f"?orderId={order_id}&orderItemId={order_item_id}"
        f"&marketplaceId=A21TJRUUN4KGV"  # Amazon.in marketplace ID
    )
    
    try:
        page.goto(cust_url, wait_until="domcontentloaded", timeout=20000)
        time.sleep(1)
        try:

            page.wait_for_load_state("networkidle", timeout=8000)
        except Exception:
            time.sleep(1)
    except Exception:
        time.sleep(2)
        
   
    
    # Find and click the "Download zip file" button
    try:
        download_btn = page.locator(
            "button:has-text('Download zip'), a:has-text('Download zip'), "
            "input[value*='Download zip'], span:has-text('Download zip')"
        )
        
        if download_btn.count() == 0:
            # Try broader selectors
            download_btn = page.locator(
                "[data-action*='download'], .download-btn, "
                "a[href*='download'], button[class*='download']"
            )
            
        if download_btn.count() == 0:
            log.warning("Download button not found for order %s item %s", 
                        order_id, order_item_id)
            # Take a screenshot for debugging
            return False
            
        # Ensure destination folder exists
        dest_folder.mkdir(parents=True, exist_ok=True)
        
        # Set up download handling

        def _save_download(dl, folder, oid, oiid):
            try:
                filename = dl.suggested_filename or f"{oid}_{oiid}.zip"
                dl.save_as(str(folder / filename))
                log.info("Download saved: %s -> %s", filename, folder)
            except Exception as ex:
                log.error("Background save failed for %s: %s", oid, ex)


        with page.expect_download(timeout=30000) as download_info:
            download_btn.first.click()
            
        download = download_info.value
        save_thread = threading.Thread(
        target=_save_download,
        args=(download, dest_folder, order_id, order_item_id),
        daemon=True
    )
        save_thread.start()

        log.info("Download triggered for order %s, saving in background", order_id)
        return True
        
    except Exception as e:
        log.error("Download failed for order %s item %s: %s",
                  order_id, order_item_id, e)
                  
        # Fallback: try clicking any download-like element
        try:
            # Look for links with 'zip' in href
            zip_links = page.locator("a[href*='.zip'], a[href*='download']")
            if zip_links.count() > 0:
                dest_folder.mkdir(parents=True, exist_ok=True)
                with page.expect_download(timeout=30000) as download_info:
                    zip_links.first.click()
                download = download_info.value


                save_thread = threading.Thread(
                target=_save_download,
                args=(download, dest_folder, order_id, order_item_id),
                daemon=True
            )
                save_thread.start()
                log.info("Fallback download triggered: %s", order_id)
            
                return True
        except Exception as e2:
            log.error("Fallback download also failed: %s", e2)
            
        return False
        
def _generate_report(base_folder: Path, results: list[dict], ship_day: str,
                     total_orders: int, downloaded: int, skipped: int, errors: int,
                     elapsed: float,dup_warnings: list = None):
    """Generate a summary report file in the base folder. Overwrites each time."""
    from datetime import datetime

    if dup_warnings is None:
        dup_warnings = []


    report_path = base_folder / REPORT_FILENAME

# Load previously processed order IDs from existing report
    previously_processed = _parse_report_order_ids(base_folder)

# Add newly processed orders (downloaded or errored - not skipped-from-report ones)
    all_processed = set(previously_processed)
    for r in results:
        if r.get("status") != "skipped":
            all_processed.add(r.get("order_id", ""))
    all_processed.discard("")

    lines = []
    lines.append("=" * 70)
    lines.append("       AMAZON ORDER DOWNLOAD REPORT")
    lines.append("=" * 70)
    lines.append(f"  Last Updated: {datetime.now().strftime('%d %b %Y, %I:%M %p')}")
    lines.append(f"  Ship Day    : {ship_day}")
    lines.append(f"  Base Folder : {base_folder}")
    lines.append("")
    lines.append("-" * 70)
    lines.append("  LAST RUN SUMMARY")
    lines.append("-" * 70)
    lines.append(f"  Total orders this run        : {total_orders}")
    lines.append(f"  Downloaded (new)             : {downloaded}")
    lines.append(f"  Skipped (already in report)  : {skipped}")
    lines.append(f"  Errors                       : {errors}")
    lines.append(f"  Time Taken                   : {elapsed}s")
    lines.append("")

    # Regular orders list with detail
    lines.append("-" * 70)
    lines.append(f"  ALL PROCESSED ORDERS ({len(all_processed)})")
    lines.append("-" * 70)
    for i, oid in enumerate(sorted(all_processed), 1):
    # Find name and variant info from results
        detail = ""
        for r in results:
            if r.get("order_id") == oid:
                name = r.get("name", "")
                variants = r.get("variants", [])
                if name and name != "?":
                    detail = f" | {name}"
                    if variants:
                        detail += f" | {', '.join(variants)}"
                break
        lines.append(f"  {i:3}. {oid}{detail}")
    if not all_processed:
        lines.append("  (none)")
    lines.append("")


    # -- This run's newly processed orders --
    new_results = [r for r in results if r.get("status") != "skipped"]
    if new_results:
        lines.append("-" * 70)
        lines.append(f"  THIS RUN - NEWLY PROCESSED ({len(new_results)})")
        lines.append("-" * 70)
        for i, r in enumerate(new_results, 1):
            oid = r.get("order_id", "?")
            name = r.get("name", "?")
            variants = ", ".join(r.get("variants", []))
            status = r.get("status", "?")
            lines.append(f"  {i:3}. {oid} | {name} | {variants} | {status}")
        lines.append("")



    # Orders with multiple variants or quantity > 1
    multi_variant_orders = []
    high_qty_orders = []
    for r in results:
        variants = r.get("variants", [])
        if len(variants) > 1:
            multi_variant_orders.append(r)
        # Check quantity from items
        for item in r.get("items_detail", []):
            if item.get("quantity", 1) > 1:
                high_qty_orders.append(r)
                break

    if multi_variant_orders or high_qty_orders or dup_warnings:
        lines.append("-" * 70)
        lines.append("  SPECIAL ATTENTION")
        lines.append("-" * 70)
        if multi_variant_orders:
            lines.append("")
            lines.append("  Orders with MULTIPLE VARIANTS:")
            for r in multi_variant_orders:
                name = r.get("name", "?")
                oid = r.get("order_id", "?")[-4:]
                variants = r.get("variants", [])
                lines.append(f"    • {name}-{oid} -> Variants: {', '.join(variants)}")

        if high_qty_orders:
            lines.append("")
            lines.append("  Orders with QUANTITY > 1:")
            for r in high_qty_orders:
                name = r.get("name", "?")
                oid = r.get("order_id", "?")[-4:]
                for item in r.get("items_detail", []):
                    qty = item.get("quantity", 1)
                    if qty > 1:
                        variant = item.get("variant", "4x3")
                        lines.append(f"    • {name}-{oid} ({variant}) -> Qty: {qty}")
        if dup_warnings:
            lines.append("")
            lines.append("  ⚠ DUPLICATE LAST-4 ORDER IDs (verify manually):")
            for oid4, oid_list in dup_warnings:
                lines.append(f"    • Last-4 '{oid4}' shared by:")
                for oid in oid_list:
                    lines.append(f"        {oid}")

        lines.append("")
# WhatsApp orders (folder-based, still useful info)
    wa_folders = []
    wa_dir = base_folder / "WhatsApp"
    if wa_dir.exists():
        for d in sorted(wa_dir.iterdir()):
            if d.is_dir():
                wa_folders.append(d.name)

    if wa_folders:
        lines.append("-" * 70)
        lines.append(f"  WHATSAPP ORDERS ({len(wa_folders)})")
        lines.append("-" * 70)
        for i, fname in enumerate(wa_folders, 1):
            lines.append(f"  {i:3}. {fname}")
        lines.append("")
    lines.append("=" * 70)
    lines.append("  END OF REPORT")
    lines.append("=" * 70)

    report_text = "\n".join(lines)
    report_path.write_text(report_text, encoding="utf-8")
    log.info("Report written to %s (%d total processed orders)", report_path, len(all_processed))


def _get_order_detail_note(folder_name: str, results: list[dict]) -> str:
    """Check if a folder's order has multi-variant or high quantity."""
    # Extract order_id_4 from folder name (pattern: name-XXXX or name-XXXX-variant)
    parts = folder_name.split("-")
    if len(parts) < 2:
        return ""

    # Find 4-digit order ID part
    order_id_4 = None
    for p in parts:
        if re.match(r"^\d{4}$", p):
            order_id_4 = p
            break

    if not order_id_4:
        return ""

    # Find matching result
    for r in results:
        if r.get("order_id", "").endswith(order_id_4):
            notes = []
            variants = r.get("variants", [])
            if len(variants) > 1:
                notes.append(f"Multi-variant: {', '.join(variants)}")
            for item in r.get("items_detail", []):
                qty = item.get("quantity", 1)
                if qty > 1:
                    v = item.get("variant", "4x3")
                    notes.append(f"Qty {qty} ({v})")
            return "; ".join(notes)

    return ""


# --- Main Processing Thread ---------------------------------------------------
# --- Main Processing Thread ---------------------------------------------------

def _run_download(base_folder_path: str, ship_day: str, task_id: str):
    """Main worker thread: process all unshipped orders for the given ship day."""
    q = _get_q(task_id)
    if q is None:
        return

    t0 = time.perf_counter()
    base_folder = Path(base_folder_path)

    try:
        # --- Stage: Launch browser
        _emit(q, stage="browser", pct=0, 
              detail="Launching browser...", done=False)
        
        _close_browser()
        from playwright.sync_api import sync_playwright

        _ensure_browser_dir()
        try:
            pw = sync_playwright().start()
            context = pw.chromium.launch_persistent_context(
                user_data_dir=str(BROWSER_DATA_DIR),
                headless=False,
                viewport={"width": 1400, "height": 900},
                args=[
                    "--disable-blink-features=AutomationControlled",
                    "--no-first-run",
                    "--no-default-browser-check",
                ],
                ignore_default_args=["--enable-automation"],
                channel="chromium",
            )
        except Exception as e:
            _emit(q, stage="error", pct=0, detail="", done=True, 
                  error=f"Failed to launch browser: {e}")
            return
        global _playwright_instance, _browser_instance, _browser_context
        with _browser_lock:
            _playwright_instance = pw
            _browser_instance = None
            _browser_context = context
        # Get or create a page
        pages = context.pages
        if pages:
            page = pages[0]
        else:
            page = context.new_page()

        # --- Stage: Login check -----------------------------------------------
        _emit(q, stage="login", pct=2, 
              detail="Checking login status...", done=False)

        if not _wait_for_login(page, q, timeout=300):
            _emit(q, stage="error", pct=0, detail="", done=True, 
                  error="Login timeout. Please log in to Seller Central and try again.")
            return

        if _is_stopped(task_id):
            _emit(q, stage="stopped", pct=0, detail="Stopped by user.", done=True)
            return

        # --- Stage: Navigate to unshipped -------------------------------------
        _emit(q, stage="navigate", pct=5, 
              detail="Navigating to Unshipped orders...", done=False)

        if not _navigate_to_unshipped(page, q):
            _emit(q, stage="error", pct=0, detail="", done=True, 
                  error="Could not navigate to Unshipped orders page.")
            return

        if _is_stopped(task_id):
            _emit(q, stage="stopped", pct=0, detail="Stopped by user.", done=True)
            return

        # --- Stage: Scan orders -----------------------------------------------
        _emit(q, stage="scan", pct=8, 
              detail=f"Scanning orders for ship day: {ship_day}...", done=False)

        orders = _scrape_order_list(page, ship_day, q)

        if not orders:
            _emit(q, stage="done", pct=100, done=True, 
                  detail=f"No orders found for ship day '{ship_day}'.",
                  result={"processed": 0, "skipped": 0, "downloaded": 0, 
                          "errors": 0, "orders": []})
            return

        _emit(q, stage="scan", pct=12, 
              detail=f"Found {len(orders)} orders for {ship_day}. Processing...", 
              done=False)

 
        new_orders = []
        early_skipped = []

        # Parse previously processed order IDs from the report
        previously_processed = _parse_report_order_ids(base_folder)
        log.info("Report contains %d previously processed order IDs", len(previously_processed))
# Group orders by last-4 digits to detect duplicates (for report)
        duplicate_id4 = {}
        for order in orders:
            oid4 = order["order_id_4"]
            duplicate_id4.setdefault(oid4, []).append(order)

# Report duplicate last-4-digit orders
        dup_warnings = []
        for oid4, oid_list in duplicate_id4.items():
            if len(oid_list) > 1:
                dup_warnings.append((oid4, [o["order_id"] for o in oid_list]))
                log.warning(" ⚠ DUPLICATE LAST-4 '%s': %s", oid4,
                    [o["order_id"] for o in oid_list])
 

# Group orders by last-4 digits to detect duplicates
        for order in orders:
            full_order_id = order["order_id"]
            if full_order_id in previously_processed:
                early_skipped.append(order)
                log.info("  SKIP (in report): %s", full_order_id)
            else:
                new_orders.append(order)

        skipped = len(early_skipped)
        if early_skipped:
            log.info("SKIPPED %d orders (already in report). %d new to process.",
             len(early_skipped), len(new_orders))
            _emit(q, stage="process", pct=14,
          detail=f"Skipped {len(early_skipped)} existing orders. Processing {len(new_orders)} new...",
          done=False)

# --- Stage: Process each NEW order -------------------------------------
        total_orders = len(new_orders)
        processed = 0
    
        downloaded = 0
        errors = 0
        results: list[dict] = []


        # Add early-skipped to results
        for o in early_skipped:
            results.append({
        "order_id": o["order_id"],
        "name": "?",
        "variants": [it.get("variant", "?") for it in o.get("items", [])],
        "status": "skipped",
        "reason": "Already in report",
            })

        
        for oi, order in enumerate(new_orders):
            if _is_stopped(task_id):
                _emit(q, stage="stopped", pct=0, detail="Stopped by user.", done=True)
                return

            order_id = order["order_id"]
            order_id_4 = order["order_id_4"]
            order_t0 = time.perf_counter()
            pct = 12 + int((oi / total_orders) * 80)

            _emit(q, stage="process", pct=pct, 
                  detail=f"Processing order {oi+1}/{total_orders}: {order_id}", 
                  done=False, order_index=oi, order_total=total_orders, 
                  order_id=order_id)

            # Navigate to order detail to get name and items
            details = _get_order_details(page, order_id, q)
           
            if not details or not details.get("name"):
                log.warning("Could not get details for order %s", order_id)
                _emit(q, stage="process", pct=pct, 
                      detail=f"⚠️ Could not get details for {order_id}, skipping", 
                      done=False)
                errors += 1
                results.append({
                    "order_id": order_id,
                    "status": "error",
                    "reason": "Could not extract order details",
                })
                continue
                
            buyer_name = details["name"]
            items = details["items"]
            # If detail page items all default to 4x3 but the scan detected
# a specific variant from the orders list page, use the scan's variant.
# The scan reads the product title directly from the table row which is reliable.
            scan_variant = None
            scan_items = order.get("items", [])
            if scan_items:
                for si in scan_items:
                    sv = si.get("variant", "4x3")
                    if sv != "4x3":
                        scan_variant = sv
                        break

            if scan_variant and items:
    # Check if all detail items are 4x3 (likely failed to detect)
                all_default = all(it.get("variant", "4x3") == "4x3" for it in items)
                if all_default:
                    log.info("  Using scan variant '%s' (detail page defaulted to 4x3)", scan_variant)
                    for it in items:
                        it["variant"] = scan_variant
            # Also update title from scan if detail title is generic
                        if scan_items[0].get("title"):
                            it["title"] = scan_items[0]["title"]




            wa_info = details.get("whatsapp")
            variants_str = ", ".join(it.get("variant", "?") for it in items) or "?"
            wa_tag = " [WHATSAPP]" if wa_info else ""
            log.info("ORDER %d/%d: %s  (%s, %s, %s)%s",
                        oi + 1, total_orders, order_id, buyer_name, variants_str,
                            order.get("ship_day", "?"), wa_tag)
            # --- Handle WhatsApp orders ---
            wa_info = details.get("whatsapp")
            if wa_info:
                wa_number = wa_info["whatsapp_number"]
                inc_component = wa_info["included_component"]
                
                # Check if folder already exists (by name-orderID pattern)
                # Create WhatsApp folder inside base_folder
                wa_base = base_folder / "WhatsApp"
                wa_base.mkdir(parents=True, exist_ok=True)


                # Folder name: name-orderID4-whatsappNumber-includedComponent
                safe_name = re.sub(r'[<>:"/\\|?*]', '', buyer_name).strip()
                safe_component = re.sub(r'[<>:"/\\|?*]', '', inc_component).strip()
                wa_folder_name = f"{safe_name}-{order_id_4}-{wa_number}-{safe_component}"
                wa_folder = wa_base / wa_folder_name
                wa_folder.mkdir(parents=True, exist_ok=True)
                downloaded += 1
                log.info("Created WhatsApp folder: %s", wa_folder_name)
                results.append({
    "order_id": order_id,
    "name": buyer_name,
    "variant": f"WhatsApp ({inc_component})",
    "status": "whatsapp_created",
})
                _emit(q, stage="process", pct=pct,
      detail=f"📱 WhatsApp order: created folder {wa_folder_name}",
      done=False,
      order={"order_id": order_id_4, "name": buyer_name,
             "variant": f"WA: {inc_component}", "status": "downloaded"})



                # Skip zip download for WhatsApp orders
                continue
                
            if not items:
                log.warning("No items found for order %s", order_id)
                _emit(q, stage="process", pct=pct, 
                      detail=f"⚠️ No downloadable items for {order_id}", 
                      done=False)
                errors += 1
                results.append({
                    "order_id": order_id,
                    "name": buyer_name,
                    "status": "error",
                    "reason": "No items found",
                })
                continue
                
            # Process each item/variant
            order_downloaded = 0
            order_skipped = 0
            
            for item in items:
                if _is_stopped(task_id):
                    _emit(q, stage="stopped", pct=0, 
                          detail="Stopped by user.", done=True)
                    return
                    
                variant = item.get("variant", "4x3")
                order_item_id = item.get("order_item_id")
                
                # Extract photo count from title for 3x2/3x3 variants
                photo_count = ""
                if variant in ("3x2", "3x3"):
                    title = item.get("title", "")
                    pc_match = re.search(r'Customised\s+(\d+)', title, re.IGNORECASE)
                    if pc_match:
                        photo_count = pc_match.group(1)
                        
                folder_name = _build_folder_name(buyer_name, order_id_4, variant, photo_count)

                    
                # Create the folder
                target_folder = base_folder / folder_name
                target_folder.mkdir(parents=True, exist_ok=True)
                
                _emit(q, stage="download", pct=pct,
                      detail=f"Downloading zip for {buyer_name}-{order_id_4} ({variant})...",
                      done=False)
                      
                # Download the zip
                success = _download_zip_for_item(
                    page, order_id, order_item_id, target_folder, q
                )
                
                if success:
                    order_downloaded += 1
                    downloaded += 1
                    log.info("Successfully downloaded for %s", folder_name)
                else:
                    errors += 1
                    log.warning("Failed to download for %s", folder_name)
                    # Remove empty folder if download failed
                    if target_folder.exists() and not any(target_folder.iterdir()):
                        target_folder.rmdir()
                        
                # Small delay between downloads
         
                
            processed += 1
            status = "done" if order_downloaded > 0 else (
                "skipped" if order_skipped > 0 else "error"
            )
            results.append({
                "order_id": order_id,
                "name": buyer_name,
                "variants": [it.get("variant") for it in items],
                "items_detail": [{"variant": it.get("variant", "4x3"), "quantity": it.get("quantity", 1)} for it in items],
                "downloaded": order_downloaded,
                "skipped": order_skipped,
                "status": status,
            })
            
            _emit(q, stage="process", pct=pct,
                  detail=f"✔️ {buyer_name}-{order_id_4}: "
                         f"{order_downloaded} downloaded, {order_skipped} skipped",
                  done=False)
            
            order_elapsed = round(time.perf_counter() - order_t0, 1)
            log.info("DONE %d/%d: %s  (%s, %s) -> %s  [%.1fs]",
                        oi + 1, total_orders, order_id, buyer_name, variants_str,
                        status, order_elapsed)
                  
        # --- Stage: Done ------------------------------------------------------
        elapsed = round(time.perf_counter() - t0, 1)
        
        log_activity("amazon_download", "process_orders", json.dumps({
            "ship_day": ship_day,
            "total_orders": total_orders,
            "processed": processed,
            "downloaded": downloaded,
            "skipped": skipped,
            "errors": errors,
            "elapsed": elapsed,
        }))
        
        # Generate report
        try:
            _generate_report(
                base_folder, results, ship_day, 
                total_orders + len(early_skipped), downloaded, skipped, errors, elapsed,dup_warnings=dup_warnings
            )
        except Exception as re_err:
            log.warning("Failed to generate report: %s", re_err)
            
        _emit(q, stage="done", pct=100, done=True,
              detail=f"Completed in {elapsed}s",
              result={
                  "orders": results,
                  "total_orders": total_orders,
                  "processed": processed,
                  "downloaded": downloaded,
                  "skipped": skipped,
                  "errors": errors,
                  "elapsed": elapsed,
                  "base_folder": base_folder_path,
              })
              
    except Exception as exc:
        log.exception("Amazon download failed")
        _emit(q, stage="error", pct=0, detail="", done=True,
              error=str(exc))
              
# --- Routes -------------------------------------------------------------------

@amazon_bp.route("/amazon-download/base-folder", methods=["GET"])
def ad_get_base_folder():
    folder = get_setting(SETTING_BASE_FOLDER, "")
    return jsonify({"folder": folder})

@amazon_bp.route("/amazon-download/base-folder", methods=["POST"])
def ad_set_base_folder():
    data = request.get_json(silent=True) or {}
    folder = (data.get("base_folder") or data.get("folder") or "").strip()
    if not folder:
        return jsonify({"error": "Folder path is required."}), 400
    p = Path(folder)
    if not p.is_dir():
        try:
            p.mkdir(parents=True, exist_ok=True)
        except Exception:
            return jsonify({"error": f"Folder not found and could not create: {folder}"}), 400
    set_setting(SETTING_BASE_FOLDER, str(p))
    log_activity("amazon_download", "set_base_folder", str(p))
    return jsonify({"status": "ok", "folder": str(p)})
    
@amazon_bp.route("/amazon-download/browser-status", methods=["GET"])
def ad_browser_status():
    """Check if the browser is running and logged in."""
    with _browser_lock:
        if _browser_context is None:
            return jsonify({"status": "not_running", "logged_in": False})
        try:
            pages = _browser_context.pages
            if not pages:
                return jsonify({"status": "running", "logged_in": False})
            page = pages[0]
            logged_in = _is_logged_in(page)
            return jsonify({
                "status": "running",
                "logged_in": logged_in,
                "url": page.url,
            })
        except Exception:
            return jsonify({"status": "error", "logged_in": False})
            
@amazon_bp.route("/amazon-download/launch-browser", methods=["POST"])
def ad_launch_browser():
    """Launch the browser for manual login."""
    try:
        # Close any existing browser first
        _close_browser()

        from playwright.sync_api import sync_playwright

        _ensure_browser_dir()
        pw = sync_playwright().start()
        context = pw.chromium.launch_persistent_context(
            user_data_dir=str(BROWSER_DATA_DIR),
            headless=False,
            viewport={"width": 1400, "height": 900},
            args=[
                "--disable-blink-features=AutomationControlled",
                "--no-first-run",
                "--no-default-browser-check",
            ],
            ignore_default_args=["--enable-automation"],
            channel="chromium",
        )
        pages = context.pages
        page = pages[0] if pages else context.new_page()

        # Navigate to Seller Central
        try:
            page.goto(SELLER_CENTRAL_URL, wait_until="domcontentloaded", timeout=20000)
        except Exception:
            pass
        # Store references so we can close later
        global _playwright_instance, _browser_instance, _browser_context
        with _browser_lock:
            _playwright_instance = pw
            _browser_instance = None
            _browser_context = context
        return jsonify({"status": "ok", "message": "Browser launched. Please log in if needed."})
    except Exception as e:
        return jsonify({"error": f"Failed to launch browser: {e}"}), 500
        
@amazon_bp.route("/amazon-download/close-browser", methods=["POST"])
def ad_close_browser():
    """Close the browser."""
    _close_browser()
    return jsonify({"ok": True})
    
@amazon_bp.route("/amazon-download/start", methods=["POST"])
def ad_start():
    """Start the order download process.
    
    Body JSON: {"ship_day": "Wed"}
    """
    data = request.get_json(silent=True) or {}
    ship_day = data.get("ship_day", "").strip()
    
    if not ship_day or ship_day not in DAY_ABBREVS:
        return jsonify({
            "error": f"ship_day must be one of: {', '.join(DAY_ABBREVS.keys())}"
        }), 400
        
    base_folder = get_setting(SETTING_BASE_FOLDER, "")
    if not base_folder or not Path(base_folder).is_dir():
        return jsonify({"error": "Base folder not set or not found."}), 400
        
    task_id = f"amz_{int(time.time() * 1000)}"
    _new_q(task_id)
    
    thread = threading.Thread(
        target=_run_download, args=(base_folder, ship_day, task_id), daemon=True
    )
    thread.start()
    
    return jsonify({"task_id": task_id, "ship_day": ship_day})
    
@amazon_bp.route("/amazon-download/progress/<task_id>")
def ad_progress(task_id: str):
    """SSE stream for download progress."""
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

@amazon_bp.route("/amazon-download/stop/<task_id>", methods=["POST"])
def ad_stop(task_id: str):
    """Stop a running download task."""
    with _pq_lock:
        if task_id in _stop_flags:
            _stop_flags[task_id] = True
            return jsonify({"ok": True, "message": "Stop signal sent."})
        return jsonify({"error": "Task not found."}), 404
        
WA_BROWSER_DATA_DIR = Path(__file__).parent / ".whatsapp_browser_data"
SETTING_WA_GROUP = "amazon_download_wa_group"

_wa_playwright = None
_wa_context = None
_wa_lock = threading.Lock()

def _ensure_wa_dir():
    WA_BROWSER_DATA_DIR.mkdir(parents=True, exist_ok=True)

def _close_wa_browser():
    global _wa_playwright, _wa_context
    with _wa_lock:
        if _wa_context:
            try:
                _wa_context.close()
            except Exception:
                pass
            _wa_context = None
        if _wa_playwright:
            try:
                _wa_playwright.stop()
            except Exception:
                pass
            _wa_playwright = None

@amazon_bp.route("/amazon-download/wa-launch", methods=["POST"])
def ad_wa_launch():
    """Launch WhatsApp Web browser for QR login."""
    try:
        _close_wa_browser()

        from playwright.sync_api import sync_playwright
        
        _ensure_wa_dir()
        pw = sync_playwright().start()
        context = pw.chromium.launch_persistent_context(
            user_data_dir=str(WA_BROWSER_DATA_DIR),
            headless=False,
            viewport={"width": 1300, "height": 850},
            args=[
                "--disable-blink-features=AutomationControlled",
                "--no-first-run",
                "--no-default-browser-check",
            ],
            ignore_default_args=["--enable-automation"],
            channel="chromium",
        )

        pages = context.pages
        page = pages[0] if pages else context.new_page()
        try:
            page.goto("https://web.whatsapp.com", wait_until="domcontentloaded", timeout=20000)
        except Exception:
            pass

        global _wa_playwright, _wa_context
        with _wa_lock:
            _wa_playwright = pw
            _wa_context = context

        return jsonify({"status": "ok", "message": "WhatsApp Web launched. Scan QR if needed."})
    except Exception as e:
        return jsonify({"error": f"Failed: {e}"}), 500

@amazon_bp.route("/amazon-download/wa-status", methods=["GET"])
def ad_wa_status():
    """Check if WhatsApp Web browser is running."""
    with _wa_lock:
        if _wa_context:
            try:
                _wa_context.pages
                return jsonify({"running": True})
            except Exception:
                pass
        return jsonify({"running": False})

@amazon_bp.route("/amazon-download/wa-group", methods=["POST"])
def ad_wa_set_group():
    """Save the WhatsApp group name."""
    data = request.get_json(silent=True) or {}
    group = (data.get("group_name") or "").strip()
    if not group:
        return jsonify({"error": "Group name required."}), 400
    
    set_setting(SETTING_WA_GROUP, group)
    return jsonify({"status": "ok", "group": group})

@amazon_bp.route("/amazon-download/wa-group", methods=["GET"])
def ad_wa_get_group():
    """Get saved WhatsApp group name."""
    group = get_setting(SETTING_WA_GROUP, "")
    return jsonify({"group_name": group})

@amazon_bp.route("/amazon-download/wa-sync", methods=["POST"])
def ad_wa_sync():
    """Start syncing WhatsApp media to order folders."""
    base_folder = get_setting(SETTING_BASE_FOLDER, "")
    if not base_folder or not Path(base_folder).is_dir():
        return jsonify({"error": "Base folder not set or not found."}), 400

    group_name = get_setting(SETTING_WA_GROUP, "")
    if not group_name:
        return jsonify({"error": "WhatsApp group name not set."}), 400

    task_id = f"wasync_{int(time.time() * 1000)}"
    _new_q(task_id)

    thread = threading.Thread(
        target=_run_wa_sync, args=(base_folder, group_name, task_id), daemon=True
    )
    thread.start()
    
    return jsonify({"task_id": task_id, "group": group_name})

@amazon_bp.route("/amazon-download/wa-progress/<task_id>")
def ad_wa_progress(task_id: str):
    """SSE stream for WhatsApp sync progress."""
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

def _run_wa_sync(base_folder_path: str, group_name: str, task_id: str):
    """Worker thread: scan WhatsApp group and download media to order folders."""
    q = _get_q(task_id)
    if q is None:
        return

    t0 = time.perf_counter()
    base_folder = Path(base_folder_path)

    try:
        _emit(q, stage="browser", pct=0, detail="Launching WhatsApp Web...", done=False)

        # Close any existing WA browser from another thread
        _close_wa_browser()

        from playwright.sync_api import sync_playwright

        _ensure_wa_dir()
        try:
            pw = sync_playwright().start()
            context = pw.chromium.launch_persistent_context(
                user_data_dir=str(WA_BROWSER_DATA_DIR),
                headless=False,
                viewport={"width": 1300, "height": 850},
                args=[
                    "--disable-blink-features=AutomationControlled",
                    "--no-first-run",
                    "--no-default-browser-check",
                ],
                ignore_default_args=["--enable-automation"],
                channel="chromium",
            )
        except Exception as e:
            _emit(q, stage="error", pct=0, detail="", done=True,
                  error=f"Failed to launch browser: {e}")
            return

        global _wa_playwright, _wa_context
        with _wa_lock:
            _wa_playwright = pw
            _wa_context = context
            
        pages = context.pages
        page = pages[0] if pages else context.new_page()

        # Navigate to WhatsApp Web
        page.goto("https://web.whatsapp.com", wait_until="domcontentloaded", timeout=20000)

        # Wait for WhatsApp to be ready (chat list loaded)
        _emit(q, stage="login", pct=2, detail="Waiting for WhatsApp Web to load...", done=False)
        try:
            page.wait_for_selector(
                "div[data-tab='3'], #pane-side, div[aria-label='Chat list']",
                timeout=60000
            )
        except Exception:
            # Maybe user needs to scan QR code
            _emit(q, stage="login", pct=2,
                  detail="Please scan the QR code in WhatsApp Web...", done=False)
            try:
                page.wait_for_selector(
                    "div[data-tab='3'], #pane-side, div[aria-label='Chat list']",
                    timeout=120000
                )
            except Exception:
                _emit(q, stage="error", pct=0, detail="", done=True,
                      error="WhatsApp Web did not load. Please scan QR and try again.")
                return

        time.sleep(2)
        _emit(q, stage="navigate", pct=5, detail=f"Opening group: {group_name}...", done=False)

        # Search for and open the group
        if not _wa_open_group(page, group_name):
            _emit(q, stage="error", pct=0, detail="", done=True,
                  error=f"Could not find group: {group_name}")
            return

        time.sleep(2)
        _emit(q, stage="scan", pct=10, detail="Scanning messages for order IDs...", done=False)

        # Scan messages and build order-media map
        order_media_map = _wa_scan_messages(page, q)

        if not order_media_map:
            _emit(q, stage="done", pct=100, done=True,
                  detail="No order IDs with media found in recent messages.",
                  result={"downloaded": 0, "matched": 0, "not_found": 0})
            return

        log.info("WA SYNC: Found %d order IDs with media", len(order_media_map))
        for oid4, media_list in order_media_map.items():
            log.info("  %s: %d media items", oid4, len(media_list))

        # Match order IDs to folders and download
        _emit(q, stage="download", pct=20, detail="Downloading media to folders...", done=False)

        total_items = sum(len(v) for v in order_media_map.values())
        downloaded = 0
        not_found = 0
        matched = 0
        item_idx = 0

        for oid4, media_elements in order_media_map.items():
            if _is_stopped(task_id):
                _emit(q, stage="stopped", pct=0, detail="Stopped by user.", done=True)
                return

            # Find matching folder in base location
            target_folder = _find_folder_for_order(base_folder, oid4)
            if not target_folder:
                log.warning("  No folder found for order ID -%s", oid4)
                not_found += len(media_elements)
                item_idx += len(media_elements)
                continue

            matched += 1
            log.info("  Downloading %d files -> %s", len(media_elements), target_folder.name)

            for media_info in media_elements:
                if _is_stopped(task_id):
                    _emit(q, stage="stopped", pct=0, detail="Stopped by user.", done=True)
                    return

                item_idx += 1
                pct = 20 + int((item_idx / total_items) * 70)
                _emit(q, stage="download", pct=pct,
                      detail=f"Downloading {item_idx}/{total_items} -> {target_folder.name}",
                      done=False)

                success = _wa_download_media(page, media_info, target_folder)
                if success:
                    downloaded += 1

        elapsed = round(time.perf_counter() - t0, 1)
        _emit(q, stage="done", pct=100, done=True,
              detail=f"Done! {downloaded} files downloaded in {elapsed}s",
              result={
                  "downloaded": downloaded,
                  "matched": matched,
                  "not_found": not_found,
                  "total_orders": len(order_media_map),
                  "elapsed": elapsed,
              })

        log.info("WA SYNC DONE: %d downloaded, %d orders matched, %d not found [%.1fs]",
                 downloaded, matched, not_found, elapsed)

    except Exception as exc:
        log.exception("WhatsApp sync failed")
        _emit(q, stage="error", pct=0, detail="", done=True, error=str(exc))


def _wa_open_group(page, group_name: str) -> bool:
    """Open the first chat in the chat list (assumed to be the target group)."""
    try:
        # Simply click the first chat in the list - user keeps the target group on top
        first_chat = page.locator(
            "#pane-side div[role='listitem'], "
            "#pane-side div[role='row'], "
            "#pane-side div[data-testid='cell-frame-container']"
        )
        if first_chat.count() > 0:
            first_chat.first.click()
            time.sleep(2)
            log.info("Opened first chat in list (assumed group: %s)", group_name)
            return True

        # Fallback: try clicking first span with a title in the chat list
        chat_title = page.locator("#pane-side span[title]")
        if chat_title.count() > 0:
            chat_title.first.click()
            time.sleep(2)
            log.info("Opened first chat title in list")
            return True
            
        return False
    except Exception as e:
        log.error("Failed to open first chat: %s", e)
        return False
def _wa_scan_messages(page, q: queue.Queue) -> dict:
    """Scan visible messages in the chat. Returns {order_id_4: [media_info_list]}."""
    
    # Reads messages from bottom to top. When a 4-digit text message is found,
    # all media messages ABOVE it (until the next 4-digit marker) belong to that order.
    order_media_map = {}

    try:
        # Get all message elements in the chat
        # WhatsApp messages are in a scrollable container
        msg_container = page.locator(
            "div[data-tab='8'], "
            "div[role='application'], "
            "div.copyable-area"
        )

        # Get all individual message rows
        messages = page.locator(
            "div.message-in, div.message-out, "
            "div[data-pre-plain-text], "
            "div[class*='message']"
        )

        # Fallback: get messages by role
        if messages.count() == 0:
            messages = page.locator("div[role='row']")

        msg_count = messages.count()
        log.info("WA SCAN: Found %d message elements", msg_count)

        if msg_count == 0:
            return {}

        # Collect all messages with their type (text/media) and content
        # Process from bottom (newest) to top (oldest)
        collected = []
        for i in range(msg_count - 1, -1, -1):
            try:
                msg_el = messages.nth(i)
                msg_text = ""
                is_media = False

                # Check if it's a media message (image/video/document)
                media_el = msg_el.locator(
                    "img[src*='blob:'], img[src*='media'], "
                    "div[data-testid='media-url-provider'], "
                    "a[href*='blob:'], "
                    "div[data-testid='document-thumb'], "
                    "span[data-icon='audio-download'], "
                    "img[data-testid='image-thumb']"
                )

                if media_el.count() > 0:
                    is_media = True

                # Also check for downloadable content
                dl_el = msg_el.locator(
                    "button[aria-label='Download'], "
                    "span[data-icon='download'], "
                    "span[data-icon='audio-download']"
                )
                if dl_el.count() > 0:
                    is_media = True
                    
                # Get text content
                text_el = msg_el.locator(
                    "span.selectable-text, "
                    "span[dir='ltr'], "
                    "span.copyable-text"
                )
                if text_el.count() > 0:
                    msg_text = text_el.first.inner_text().strip()

                collected.append({
                    "index": i,
                    "text": msg_text,
                    "is_media": is_media,
                    "element_index": i,
                })
            except Exception:
                continue

        # Now process: find 4-digit order markers and assign media above them
        current_order_id = None
        current_media = []

        for msg in collected:  # bottom to top
            text = msg["text"].strip()

            # Check if this is a 4-digit order ID marker
            if re.match(r"^\d{4}$", text):
                # Save previous order's media
                if current_order_id and current_media:
                    order_media_map[current_order_id] = current_media
                current_order_id = text
                current_media = []
                log.info("  Found order marker: %s", text)
            elif msg["is_media"] and current_order_id:
                # This media belongs to the current order (above the marker)
                current_media.append(msg)

        # Save last order's media
        if current_order_id and current_media:
            order_media_map[current_order_id] = current_media

    except Exception as e:
        log.error("WA scan failed: %s", e)

    return order_media_map


def _wa_download_media(page, media_info: dict, target_folder: Path) -> bool:
    """Download a single media item from WhatsApp to the target folder."""
    try:
        target_folder.mkdir(parents=True, exist_ok=True)
        msg_el = page.locator(
            "div.message-in, div.message-out, "
            "div[data-pre-plain-text], "
            "div[class*='message'], "
            "div[role='row']"
        ).nth(media_info["element_index"])

        # Try to click on the media to open it
        media_clickable = msg_el.locator(
            "img[src*='blob:'], img[data-testid='image-thumb'], "
            "div[data-testid='media-url-provider'], "
            "div[role='button']"
        )

        if media_clickable.count() > 0:
            media_clickable.first.click()
            time.sleep(1)

            # Look for download button in the lightbox/overlay
            dl_btn = page.locator(
                "span[data-icon='download'], "
                "button[aria-label='Download'], "
                "div[aria-label='Download']"
            )

            if dl_btn.count() > 0:
                with page.expect_download(timeout=30000) as download_info:
                    dl_btn.first.click()

                download = download_info.value
                filename = download.suggested_filename or f"media_{int(time.time()*1000)}"
                download.save_as(str(target_folder / filename))
                log.info("      Downloaded: %s", filename)

                # Close the overlay
                close_btn = page.locator(
                    "span[data-icon='x'], button[aria-label='Close']"
                )
                if close_btn.count() > 0:
                    close_btn.first.click()
                    time.sleep(0.5)

                return True

        # Fallback: try direct download button on message
        dl_btn = msg_el.locator(
            "button[aria-label='Download'], "
            "span[data-icon='download']"
        )
        if dl_btn.count() > 0:
            with page.expect_download(timeout=30000) as download_info:
                dl_btn.first.click()

            download = download_info.value
            filename = download.suggested_filename or f"media_{int(time.time()*1000)}"
            download.save_as(str(target_folder / filename))
            log.info("      Downloaded: %s", filename)
            return True

        return False
    except Exception as e:
        log.warning("      Download failed: %s", e)
        return False


def _find_folder_for_order(base_folder: Path, order_id_4: str) -> Path | None:
    """Find the folder in base_folder that matches the given 4-digit order ID."""
    tag = f"-{order_id_4}"

    # Check base folder
    if base_folder.exists():
        for d in base_folder.iterdir():
            if d.is_dir() and d.name != "WhatsApp" and tag in d.name.lower():
                return d

    # Check WhatsApp subfolder
    wa_dir = base_folder / "WhatsApp"
    if wa_dir.exists():
        for d in wa_dir.iterdir():
            if d.is_dir() and tag in d.name.lower():
                return d

    return None
            