"""Basa - Web-based Pre-Printing Workflow"""

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

CUT_MARK_LEN = 0.08
CUT_MARK_THICKNESS = 0.25
CUT_MARK_COLOR = CMYKColor(0, 0, 0, 0.15)

LABEL_COLOR = CMYKColor(0, 0, 0, 1.0)
LABEL_FONT = "Helvetica"
LABEL_SIZE = 14

JPEG_QUALITY = 100

LAYOUTS = {
    "4x3_polaroid_18": {
        "label": "4x3 Polaroid (18 per sheet)",
        "cols": 3, "rows": 6, "max_images": 18,
        "cell_w": 4.0, "cell_h": 3.0,
        "grid_left": 0.5, "grid_bottom": 1.0,
        "photo_w": 3.0, "photo_h": 2.5,
        "offset_x": 0.75, "offset_y": 0.25,
        "frame_w_px": 750, "frame_h_px": 900,
        "rotate": True, "label_y": 0.5,
    },
    "3x2_polaroid_36": {
        "label": "3x2 Polaroid (36 per sheet)",
        "cols": 4, "rows": 9, "max_images": 36,
        "cell_w": 3.0, "cell_h": 2.0,
        "grid_left": 0.5, "grid_bottom": 0.5,
        "photo_w": 2.25, "photo_h": 1.7,
        "offset_x": 0.57, "offset_y": 0.15,
        "frame_w_px": 510, "frame_h_px": 675,
        "rotate": True, "label_y": 0.2,
    },
    "3x3_square_24": {
        "label": "3x3 Square (24 per sheet)",
        "cols": 4, "rows": 6, "max_images": 24,
        "cell_w": 3.0, "cell_h": 3.0,
        "grid_left": 0.5, "grid_bottom": 0.5,
        "photo_w": 2.6, "photo_h": 2.6,
        "offset_x": 0.2, "offset_y": 0.2,
        "frame_w_px": 780, "frame_h_px": 780,
        "rotate": False, "label_y": 0.2,
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

def _fit_to_frame(img: Image.Image, frame_w: int, frame_h: int) -> Image.Image:
    img_w, img_h = img.size
    scale = min(frame_w / img_w, frame_h / img_h)
    new_w = round(img_w * scale)
    new_h = round(img_h * scale)
    img = img.resize((new_w, new_h), Image.Resampling.LANCZOS)
    background = Image.new("RGB", (frame_w, frame_h), (255, 255, 255))
    offset_x = (frame_w - new_w) // 2
    offset_y = (frame_h - new_h) // 2
    background.paste(img, (offset_x, offset_y))
    return background

def convert_to_cmyk_properly(img: Image.Image) -> Image.Image:
    cmyk_profile_path = "CoatedGRACoL2006.icc" 

    if os.path.exists(cmyk_profile_path):
        try:
            import io
            
            # 1. Get the Input Profile (Source)
            # Try to extract the embedded profile from the original image first
            embedded_profile = img.info.get('icc_profile')
            if embedded_profile:
                source_profile = ImageCms.ImageCmsProfile(io.BytesIO(embedded_profile))
            else:
                # Fallback to standard sRGB if no profile is embedded
                source_profile = ImageCms.createProfile("sRGB")
            
            # 2. Get the Output Profile (Destination)
            # Use getOpenProfile() to load the physical file from your folder
            target_profile = ImageCms.getOpenProfile(cmyk_profile_path)
            
            # 3. Apply the Transformation
# 3. Apply the Transformation
            return ImageCms.profileToProfile(
                img,
                source_profile,
                target_profile,
                renderingIntent=0,  # 0 is the universal code for Perceptual Intent
                outputMode="CMYK"
            )
            
        except Exception as e:
            log.error(f"ICC Profile conversion failed: {e}. Falling back to naive conversion.")
            return img.convert("CMYK")
    else:
        log.warning(f"ICC profile '{cmyk_profile_path}' not found in script directory. Using naive conversion.")
        return img.convert("CMYK")
def process_image(filepath: str, out_path: str, layout: dict) -> None:
    img = Image.open(filepath)
    img = _apply_exif_orientation(img)
    img = _flatten_alpha(img)
    if img.mode != "RGB":
        img = img.convert("RGB")
    if layout["rotate"]:
        w, h = img.size
        if w > h:
            img = img.rotate(-90, expand=True)
        img = _fit_to_frame(img, layout["frame_w_px"], layout["frame_h_px"])
        img = img.rotate(-90, expand=True)
    else:
        img = _fit_to_frame(img, layout["frame_w_px"], layout["frame_h_px"])
        
    # Apply the proper, profile-based conversion instead of naive conversion
    img = convert_to_cmyk_properly(img)
    
    img.save(out_path, "JPEG", quality=JPEG_QUALITY)

def _create_placeholder(out_path: str, layout: dict) -> None:
    fw, fh = layout["frame_w_px"], layout["frame_h_px"]
    if layout["rotate"]:
        img = Image.new("RGB", (fh, fw), (230, 230, 230))
    else:
        img = Image.new("RGB", (fw, fh), (230, 230, 230))
    img = img.convert("CMYK")
    img.save(out_path, "JPEG", quality=JPEG_QUALITY)

def _draw_cut_marks(c: canvas.Canvas, layout: dict) -> None:
    c.setStrokeColor(CUT_MARK_COLOR)
    c.setLineWidth(CUT_MARK_THICKNESS)
    cols, rows = layout["cols"], layout["rows"]
    cell_w, cell_h = layout["cell_w"], layout["cell_h"]
    grid_left, grid_bottom = layout["grid_left"], layout["grid_bottom"]
    mark = CUT_MARK_LEN * inch
    for col in range(cols + 1):
        for row in range(rows + 1):
            x = (grid_left + col * cell_w) * inch
            y = (grid_bottom + row * cell_h) * inch
            if col > 0:
                c.line(x - mark, y, x, y)
            if col < cols:
                c.line(x, y, x + mark, y)
            if row > 0:
                c.line(x, y - mark, x, y)
            if row < rows:
                c.line(x, y, x, y + mark)

def _draw_order_label(c: canvas.Canvas, order_name: str, label_y: float) -> None:
    c.setFillColor(LABEL_COLOR)
    c.setFont(LABEL_FONT, LABEL_SIZE)
    c.drawCentredString((PAGE_W / 2.0) * inch, label_y * inch, order_name)

def generate_pdf(image_paths: list[str], output_path: str, order_name: str, layout: dict) -> None:
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
                process_image(src, tmp_path, layout)
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

        _draw_cut_marks(c, layout)
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
        for f in sorted(valid_files, key=lambda x: x.filename.lower()):
            safe_name = secure_filename(f.filename)
            if not safe_name:
                continue
            dest = os.path.join(work_dir, safe_name)
            f.save(dest)
            saved_paths.append(dest)

        if not saved_paths:
            return jsonify({"error": "Failed to save uploaded files"}), 500

        pdf_path = os.path.join(work_dir, f"{order_name}.pdf")
        generate_pdf(saved_paths, pdf_path, order_name, layout)

        return send_file(
            pdf_path,
            mimetype="application/pdf",
            as_attachment=True,
            download_name=f"{order_name}.pdf",
        )
    except Exception as e:
        log.exception("PDF generation failed")
        return jsonify({"error": str(e)}), 500
    finally:
        try:
            shutil.rmtree(work_dir, ignore_errors=True)
        except Exception:
            pass

if __name__ == "__main__":
    import socket
    hostname = socket.gethostname()
    local_ip = socket.gethostbyname(hostname)
    print(f"\n  Basa Web App running at:")
    print(f"  Local:   http://127.0.0.1:5000")
    print(f"  Mobile:  http://{local_ip}:5000\n")
    app.run(host="0.0.0.0", port=5000, debug=False)