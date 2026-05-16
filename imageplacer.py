"""Basa - Web-based Pre-Printing Workflow"""
import io
import os
import uuid
import shutil
import logging
import tempfile
from pathlib import Path

from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.utils import secure_filename

from PIL import Image, ExifTags, ImageCms
from reportlab.lib.units import inch
from reportlab.lib.colors import CMYKColor
from reportlab.pdfgen import canvas
from PIL import ImageEnhance
from PIL import ImageFilter
from PIL import ImageStat
from natsort import natsorted
try:
    import pillow_heif
    pillow_heif.register_heif_opener()
    HEIC_SUPPORTED = True
except ImportError:
    HEIC_SUPPORTED = False

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)-8s %(message)s")
log = logging.getLogger("basa-web")

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 500 * 1024 * 1024

SUPPORTED_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif"}
if HEIC_SUPPORTED:
    SUPPORTED_EXTENSIONS |= {".heic", ".heif"}

PAGE_W = 13.0
PAGE_H = 19.0

CUT_MARK_LEN = 0.25
INNER_CUT_MARK_LEN = 0.10
CUT_MARK_THICKNESS = 0.75
CUT_MARK_COLOR = CMYKColor(0, 0, 0, 1.0)
INNER_CUT_MARK_THICKNESS = 0.4
INNER_CUT_MARK_COLOR = CMYKColor(0,0,0,0.60)
LABEL_COLOR = CMYKColor(0, 0, 0, 1.0)
LABEL_FONT = "Helvetica"
LABEL_SIZE = 14

JPEG_QUALITY = 100

LAYOUTS = {
    "4x3_polaroid_18": {
        "label": "4x3 Polaroid (18 per sheet)",
        "cols": 3, "rows": 6, "max_images": 18,
        "cell_w": 4.0, "cell_h": 3.0,
        "grid_left": 0.5, "grid_bottom": 0.65,
        "photo_w": 3.0, "photo_h": 2.5,
        "offset_x": 0.75, "offset_y": 0.25,
        "frame_w_px": 750, "frame_h_px": 900,
        "rotate": True, "label_y": 0.2,
    },
    "3x2_polaroid_36": {
        "label": "3x2 Polaroid (36 per sheet)",
        "cols": 4, "rows": 9, "max_images": 36,
        "cell_w": 3.0, "cell_h": 2.0,
        "grid_left": 0.5, "grid_bottom": 0.5,
        "photo_w": 2.25, "photo_h": 1.7,
        "offset_x": 0.57, "offset_y": 0.15,
        "frame_w_px": 510, "frame_h_px": 675,
        "rotate": True, "label_y": 0.08,
    },
    "3x3_square_24": {
        "label": "3x3 Square (24 per sheet)",
        "cols": 4, "rows": 6, "max_images": 24,
        "cell_w": 3.0, "cell_h": 3.0,
        "grid_left": 0.5, "grid_bottom": 0.5,
        "photo_w": 2.6, "photo_h": 2.6,
        "offset_x": 0.2, "offset_y": 0.2,
        "frame_w_px": 780, "frame_h_px": 780,
        "rotate": False, "label_y": 0.08,
    },
}

def _apply_exif_orientation(img: Image.Image) -> Image.Image:
    try:
        exif = img.getexif()
        orientation_tag = None
        for tag, name in ExifTags.TAGS.items():
            if name == "Orientation":
                orientation_tag = tag
                break
        
        if orientation_tag and orientation_tag in exif:
            orientation = exif[orientation_tag]
            rotations = {3: 180, 6: 270, 8: 90}
            if orientation in rotations:
                img = img.rotate(rotations[orientation], expand=True)
            elif orientation in (2, 4, 5, 7):
                img = img.transpose(Image.Transpose.FLIP_LEFT_RIGHT)
                if orientation == 4:
                    img = img.rotate(180, expand=True)
                elif orientation == 5:
                    img = img.rotate(270, expand=True)
                elif orientation == 7:
                    img = img.rotate(90, expand=True)
    except Exception:
        pass
    return img

def _flatten_alpha(img: Image.Image) -> Image.Image:
    if img.mode in ("RGBA", "LA", "PA"):
        background = Image.new("RGB", img.size, (255, 255, 255))
        if img.mode == "RGBA":
            background.paste(img, mask=img.split()[3])
        else:
            background.paste(img.convert("RGBA"), mask=img.convert("RGBA").split()[3])
        return background
    return img
def _fill_frame(img: Image.Image, frame_w: int, frame_h: int,
                offset_x: float = 0.5, offset_y: float = 0.5) -> Image.Image:
    """Scale image to *cover* the frame (no white bars), then crop.
    
    offset_x / offset_y are 0.0..1.0 controlling where the crop window
    sits on the oversized axis (0.5 = centre, 0.0 = left/top, 1.0 = right/bottom).
    """
    img_w, img_h = img.size
    scale = max(frame_w / img_w, frame_h / img_h)
    new_w = round(img_w * scale)
    new_h = round(img_h * scale)
    img = img.resize((new_w, new_h), Image.Resampling.LANCZOS)
    # Crop from the oversized image using the offset
    crop_x = round((new_w - frame_w) * max(0.0, min(1.0, offset_x)))
    crop_y = round((new_h - frame_h) * max(0.0, min(1.0, offset_y)))
    img = img.crop((crop_x, crop_y, crop_x + frame_w, crop_y + frame_h))
    return img
def _fit_to_frame(img: Image.Image, frame_w: int, frame_h: int) -> Image.Image:
    img_w, img_h = img.size
    scale = min(frame_w / img_w, frame_h / img_h)
    new_w = round(img_w * scale)
    new_h = round(img_h * scale)
    img = img.resize((new_w, new_h), Image.Resampling.LANCZOS)
    background = Image.new("RGB", (frame_w, frame_h), (255, 255, 255))
    off_x = (frame_w - new_w) // 2
    off_y = (frame_h - new_h) // 2
    background.paste(img, (off_x, off_y))
    return background

def convert_to_cmyk_properly(img: Image.Image) -> Image.Image:
    cmyk_profile_path = "ISOcoated_v2_eci.icc" 

    if os.path.exists(cmyk_profile_path):
        try:
            import io
            
            # 1. Get the Input Profile (Source)
            embedded_profile = img.info.get('icc_profile')
            if embedded_profile:
                source_profile = ImageCms.ImageCmsProfile(io.BytesIO(embedded_profile))
            else:
                source_profile = ImageCms.createProfile("sRGB")
            
            # 2. Get the Output Profile (Destination)
            target_profile = ImageCms.getOpenProfile(cmyk_profile_path)
            
            
            return ImageCms.profileToProfile(
                img,
                source_profile,
                target_profile,
                renderingIntent=1,  # Relative Colorimetric
                outputMode="CMYK",
		flags=8192
                   
            )
            
        except Exception as e:
            log.error(f"ICC Profile conversion failed: {e}. Falling back to naive conversion.")
            return img.convert("CMYK")
    else:
        log.warning(f"ICC profile '{cmyk_profile_path}' not found in script directory. Using naive conversion.")
        return img.convert("CMYK")
def process_image(filepath: str, out_path: str, layout: dict,
offset_x:float = 0.5,offset_y: float = 0.5,mode: str = "fill") -> None:
    img = Image.open(filepath)
    img = _apply_exif_orientation(img)
    img = _flatten_alpha(img)
    if img.mode != "RGB":
        img = img.convert("RGB")
    if layout["rotate"]:
        w, h = img.size
        if w > h:
            img = img.rotate(-90, expand=True)
        if mode == "fit":
            img = _fit_to_frame(img, layout["frame_w_px"], layout["frame_h_px"])
        else:
            fill_ox = offset_y
            fill_oy = 1.0 - offset_x
            img = _fill_frame(img, layout["frame_w_px"], layout["frame_h_px"],
                              fill_ox, fill_oy)
        img = img.rotate(-90, expand=True)
    else:
        if mode == "fit":
            img = _fit_to_frame(img, layout["frame_w_px"], layout["frame_h_px"])
        else:
            img = _fill_frame(img, layout["frame_w_px"], layout["frame_h_px"],
                              offset_x, offset_y)
        
 
    # ------------------------------
    #img = enhancer.enhance(1.08)
    img = convert_to_cmyk_properly(img)
    # img = img.filter(ImageFilter.UnsharpMask(radius=2, percent=150, threshold=3))
    img.save(out_path, "JPEG", quality=JPEG_QUALITY)

def _create_placeholder(out_path: str, layout: dict) -> None:
    fw, fh = layout["frame_w_px"], layout["frame_h_px"]
    if layout["rotate"]:
        img = Image.new("RGB", (fh, fw), (230, 230, 230))
    else:
        img = Image.new("RGB", (fw, fh), (230, 230, 230))
    img = img.convert("CMYK")
    img.save(out_path, "JPEG", quality=JPEG_QUALITY)

def _draw_cut_marks(c: canvas.Canvas, layout: dict, layout_key:str = "") -> None:
    c.setStrokeColor(CUT_MARK_COLOR)
    c.setLineWidth(CUT_MARK_THICKNESS)
    cols, rows = layout["cols"], layout["rows"]
    cell_w, cell_h = layout["cell_w"], layout["cell_h"]
    grid_left, grid_bottom = layout["grid_left"], layout["grid_bottom"]
    mark = CUT_MARK_LEN * inch
    inner = INNER_CUT_MARK_LEN * inch
    grid_right = grid_left + cols * cell_w
    grid_top = grid_bottom + rows * cell_h
    for row in range(rows + 1):
        y = (grid_bottom + row * cell_h) * inch
        lx = grid_left * inch
        rx = grid_right * inch
        c.line(lx - mark, y, lx, y)
        c.line(rx, y, rx + mark, y)
    for col in range(cols + 1):
        x = (grid_left + col * cell_w) * inch
        by = grid_bottom * inch
        ty = grid_top * inch
        c.line(x, by - mark, x, by)
        c.line(x, ty, x, ty + mark)
    # Inner cut marks for 3x3 and 3x2 variants
    if layout_key.startswith("3x3") or layout_key.startswith("3x2"):
        c.setStrokeColor(INNER_CUT_MARK_COLOR)
        c.setLineWidth(INNER_CUT_MARK_THICKNESS)
    # Interior intersections: small cross marks
        for col in range(1, cols):
            for row in range(1, rows):
                x = (grid_left + col * cell_w) * inch
                y = (grid_bottom + row * cell_h) * inch
                c.line(x - inner, y, x + inner, y)   # horizontal
                c.line(x, y - inner, x, y + inner)   # vertical


def _draw_order_label(c: canvas.Canvas, order_name: str, label_y: float) -> None:
    c.setFillColor(LABEL_COLOR)
    c.setFont(LABEL_FONT, LABEL_SIZE)
    c.drawCentredString((PAGE_W / 2.0) * inch, label_y * inch, order_name)

def generate_pdf(image_paths: list[str], output_path: str, order_name: str, layout: dict, layout_key: str = "",
offsets: list[dict] | None = None) -> None:
    tmp_dir = tempfile.mkdtemp(prefix="basa_web_")
    tmp_files: list[str] = []

    cols, rows = layout["cols"], layout["rows"]
    grid_left, grid_bottom = layout["grid_left"], layout["grid_bottom"]
    cell_w, cell_h = layout["cell_w"], layout["cell_h"]
    photo_w, photo_h = layout["photo_w"], layout["photo_h"]
    ox, oy = layout["offset_x"], layout["offset_y"]
    max_slots = cols * rows

    try:
        processed: list[str] = []
        for idx, src in enumerate(image_paths):
            try:
                tmp_path = os.path.join(tmp_dir, f"img_{idx:02d}.jpg")
                off = (offsets[idx] if offsets and idx < len(offsets)
                   else {"x": 0.5, "y": 0.5, "mode": "fill"})
                process_image(src, tmp_path, layout,
                          off.get("x", 0.5), off.get("y", 0.5),
                          mode=off.get("mode", "fill"))
                tmp_files.append(tmp_path)
                processed.append(tmp_path)
            except Exception:
                log.exception("Skipping corrupted image: %s", src)

        while len(processed) < max_slots:
            ph_path = os.path.join(tmp_dir, f"placeholder_{len(processed):02d}.jpg")
            _create_placeholder(ph_path, layout)
            tmp_files.append(ph_path)
            processed.append(ph_path)

        c = canvas.Canvas(output_path, pagesize=(PAGE_W * inch, PAGE_H * inch))

        for seq, tmp_path in enumerate(processed):
            col = seq % cols
            row_from_top = seq // cols
            row = (rows - 1) - row_from_top

            cell_x = (grid_left + col * cell_w) * inch
            cell_y = (grid_bottom + row * cell_h) * inch
            photo_x = cell_x + ox * inch
            photo_y = cell_y + oy * inch

            c.drawImage(tmp_path, photo_x, photo_y,
                        width=photo_w * inch, height=photo_h * inch,
                        preserveAspectRatio=True)

        _draw_cut_marks(c, layout,layout_key)
        _draw_order_label(c, order_name, layout["label_y"])
        c.save()
    
    finally:
        for f in tmp_files:
            try:
                os.remove(f)
            except OSError:
                pass
        try:
            os.rmdir(tmp_dir)
        except OSError:
            pass
def prepare_preview(image_paths: list[str], layout: dict,
                    max_cell_px: int = 200) -> list[dict]:

    fw, fh = layout["frame_w_px"], layout["frame_h_px"]
    results = []
    for idx, src in enumerate(image_paths):
        try:
            img = Image.open(src)
            img = _apply_exif_orientation(img)
            img = _flatten_alpha(img)
            if img.mode != "RGB":
                img = img.convert("RGB")
            if layout["rotate"]:
                w, h = img.size
                if w > h:
                    img = img.rotate(-90, expand=True)

            # --- Compute overflow at full resolution (for ratio) ---
            img_w, img_h = img.size
            scale = max(fw / img_w, fh / img_h)
            full_w = round(img_w * scale)
            full_h = round(img_h * scale)
            ov_x_full = full_w - fw  # overflow in full-res pixels
            ov_y_full = full_h - fh

            # --- Downsample to thumbnail size ---
            # Determine the thumbnail frame size (the visible area)
            if layout["rotate"]:
                # After -90° rotation: displayed frame is fh x fw
                disp_w, disp_h = fh, fw
            else:
                disp_w, disp_h = fw, fh
            thumb_scale = min(max_cell_px / disp_w, max_cell_px / disp_h)
            if thumb_scale >= 1.0:
                thumb_scale = 1.0  # don't upscale

            # Scale the oversized image down to thumbnail proportions
            thumb_w = round(full_w * thumb_scale)
            thumb_h = round(full_h * thumb_scale)
            img = img.resize((thumb_w, thumb_h), Image.Resampling.LANCZOS)
            thumb_frame_w = round(fw * thumb_scale)
            thumb_frame_h = round(fh * thumb_scale)
            overflow_x = thumb_w - thumb_frame_w
            overflow_y = thumb_h - thumb_frame_h

            if layout["rotate"]:
                img = img.rotate(-90, expand=True)
                overflow_x, overflow_y = overflow_y, overflow_x

            results.append({
                "index": idx,
                "img": img,
                "overflow_x": overflow_x,
                "overflow_y": overflow_y,
            })
        except Exception:
            log.exception("Preview: skipping corrupted image %s", src)
    return results
@app.route("/")
def index():
    return render_template("index.html", layouts=LAYOUTS)

@app.route("/generate", methods=["POST"])
def generate():
    layout_key = request.form.get("layout", "4x3_polaroid_18")
    if layout_key not in LAYOUTS:
        return jsonify({"error": "Invalid layout"}), 400
    layout = LAYOUTS[layout_key]

    order_name = request.form.get("order_name", "").strip()
    if not order_name:
        order_name = "Order"
    order_name = secure_filename(order_name) or "Order"

    files = request.files.getlist("images")
    if not files or len(files) == 0:
        return jsonify({"error": "No images uploaded"}), 400

    valid_files = [f for f in files if f.filename and
                   Path(f.filename).suffix.lower() in SUPPORTED_EXTENSIONS]

    if not valid_files:
        return jsonify({"error": "No supported image files found"}), 400

    if len(valid_files) > layout["max_images"]:
        return jsonify({
            "error": f"Too many images. This layout supports max {layout['max_images']} photos, "
                     f"but you uploaded {len(valid_files)}."
        }), 400

    work_dir = tempfile.mkdtemp(prefix="basa_upload_")
    try:
        saved_paths: list[str] = []
        for f in natsorted(valid_files, key=lambda x: x.filename.lower()):
            safe_name = secure_filename(f.filename)
            if not safe_name:
                continue
            dest = os.path.join(work_dir, safe_name)
            f.save(dest)
            saved_paths.append(dest)

        if not saved_paths:
            return jsonify({"error": "Failed to save uploaded files"}), 500

        pdf_path = os.path.join(work_dir, f"{order_name}.pdf")
        generate_pdf(saved_paths, pdf_path, order_name, layout,layout_key)
        buf = io.BytesIO()
        with open(pdf_path,"rb") as f:
            buf.write(f.read())
        buf.seek(0)
        shutil.rmtree(work_dir, ignore_errors=True)

        return send_file(
            buf,
            mimetype="application/pdf",
            as_attachment=True,
            download_name=f"{order_name}.pdf",
        )
    except Exception as e:
        log.exception("PDF generation failed")
        shutil.rmtree(work_dir,ignore_errors=True)
        return jsonify({"error": str(e)}), 500
    
if __name__ == "__main__":
    import socket
    hostname = socket.gethostname()
    local_ip = socket.gethostbyname(hostname)
    print(f"\n  Basa Web App running at:")
    print(f"  Local:   http://127.0.0.1:5000")
    print(f"  Mobile:  http://{local_ip}:5000\n")
    app.run(host="0.0.0.0", port=5000, debug=False)