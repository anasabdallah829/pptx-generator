# streamlit_app.py
# Arabic technical UI + PPTX processing
# Requirements: streamlit, python-pptx, Pillow

import streamlit as st
import zipfile
import os
import io
import tempfile
import shutil
import random
from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER, MSO_SHAPE_TYPE
from pptx.util import Inches
from io import BytesIO

# ---------------------------
# Configuration / Styling
# ---------------------------
st.set_page_config(page_title="Slide Sync - PPTX Image Replacer", layout="wide")
# Minimal CSS to give nicer layout similar to the referenced repo
st.markdown(
    """
    <style>
    .app-header {display:flex; align-items:center; gap:12px;}
    .logo {width:48px;height:48px;border-radius:8px;background:#0ea5a4;display:inline-block;}
    .title {font-size:24px; font-weight:700; margin:0;}
    .subtitle {color:#6b7280; margin:0; font-size:13px;}
    .panel {background: #ffffff; padding:18px; border-radius:10px; box-shadow: 0 1px 3px rgba(0,0,0,0.06);}
    .small {font-size:13px; color:#6b7280;}
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="app-header">
      <div class="logo"></div>
      <div>
        <div class="title">Slide Sync — PPTX Image Replacer</div>
        <div class="subtitle">Upload a template (.pptx) and a ZIP of folders — one slide per folder, preserving layout & formats.</div>
      </div>
    </div>
    <hr/>
    """,
    unsafe_allow_html=True,
)

# ---------------------------
# Helper: session-state lists for logging details
# ---------------------------
if 'processing_details' not in st.session_state:
    st.session_state.processing_details = []

def add_detail(message, level="info"):
    st.session_state.processing_details.append({"msg": message, "level": level})

def clear_details():
    st.session_state.processing_details = []

def show_details(expanded=False):
    if st.session_state.processing_details:
        with st.expander("📋 Processing details", expanded=expanded):
            for d in st.session_state.processing_details:
                if d["level"] == "error":
                    st.error(d["msg"])
                elif d["level"] == "warning":
                    st.warning(d["msg"])
                elif d["level"] == "success":
                    st.success(d["msg"])
                else:
                    st.info(d["msg"])

# ---------------------------
# UI: Inputs & options
# ---------------------------
with st.container():
    with st.form("main_form"):
        col1, col2 = st.columns([1, 1])
        with col1:
            uploaded_pptx = st.file_uploader("📂 PowerPoint template (.pptx)", type=["pptx"])
        with col2:
            uploaded_zip = st.file_uploader("🗜 ZIP of folders (each folder = 1 slide)", type=["zip"])
        st.markdown("---")
        st.markdown("### ⚙️ Processing options")
        col3, col4 = st.columns(2)
        with col3:
            order_opt = st.radio("Image ordering", ("alphabetical", "random"), index=0)
        with col4:
            mismatch_opt = st.selectbox("If folder images ≠ template slots", ("truncate", "repeat", "skip_folder"))
        submit = st.form_submit_button("🚀 Start processing")

# ---------------------------
# Utility functions for formatting extraction & application
# ---------------------------
def get_shape_formatting(shape):
    """Collect left, top, width, height, rotation and some line/ shadow info if available."""
    fmt = {}
    fmt['left'] = shape.left
    fmt['top'] = shape.top
    fmt['width'] = shape.width
    fmt['height'] = shape.height
    fmt['rotation'] = getattr(shape, 'rotation', 0)
    # optional attributes - best-effort
    try:
        if hasattr(shape, 'line') and shape.line:
            fmt['line'] = {'width': getattr(shape.line, 'width', None),
                           'color': getattr(shape.line.color, 'rgb', None)}
    except Exception:
        pass
    try:
        if hasattr(shape, 'shadow') and shape.shadow:
            fmt['shadow'] = {'visible': getattr(shape.shadow, 'visible', None)}
    except Exception:
        pass
    return fmt

def apply_shape_formatting(new_shape, fmt):
    """Apply a limited set of formatting properties back to a newly added picture shape."""
    try:
        new_shape.left = fmt.get('left', new_shape.left)
        new_shape.top = fmt.get('top', new_shape.top)
        new_shape.width = fmt.get('width', new_shape.width)
        new_shape.height = fmt.get('height', new_shape.height)
        if 'rotation' in fmt and fmt['rotation']:
            try:
                new_shape.rotation = fmt['rotation']
            except Exception:
                pass
        # line & shadow are best-effort
        try:
            if 'line' in fmt and fmt['line'].get('width') is not None:
                new_shape.line.width = fmt['line']['width']
        except Exception:
            pass
    except Exception as e:
        add_detail(f"warn: apply_shape_formatting failed: {e}", "warning")

def extract_image_shapes_info(slide):
    """Return ordered list of image placeholders and regular pictures with formatting info."""
    infos = []
    # placeholders first (preserve order found)
    for shape in slide.shapes:
        if shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.PICTURE:
            infos.append({"shape": shape, "type": "placeholder", "fmt": get_shape_formatting(shape)})
    # then regular pictures
    for shape in slide.shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            # avoid duplicates if same object appeared
            if not any(s['shape'] == shape for s in infos):
                infos.append({"shape": shape, "type": "picture", "fmt": get_shape_formatting(shape)})
    # sort by top,left for deterministic order
    infos.sort(key=lambda x: (x['fmt']['top'], x['fmt']['left']))
    return infos

def get_template_slots(first_slide):
    """Return both shape infos and fallback positions extracted from shapes (if none)."""
    infos = extract_image_shapes_info(first_slide)
    positions = [info['fmt'] for info in infos]
    return infos, positions

# ---------------------------
# Core replacement functions
# ---------------------------
def replace_image_in_shape(slide, shape_info, image_path):
    """Replace the given shape (placeholder or picture) with image_path, preserving formatting."""
    try:
        # attempt placeholder insert if it is placeholder
        if shape_info['type'] == 'placeholder':
            shape = shape_info['shape']
            try:
                # Try insert_picture (works for placeholders)
                with open(image_path, 'rb') as fb:
                    shape.insert_picture(fb)
                add_detail(f"Replaced placeholder with {os.path.basename(image_path)}", "success")
                return True
            except Exception as e:
                add_detail(f"placeholder insert failed -> fallback: {e}", "warning")
                # fallback: remove and add picture at same coordinates
                fmt = shape_info['fmt']
                try:
                    el = shape._element
                    el.getparent().remove(el)
                except Exception:
                    pass
                try:
                    new_shape = slide.shapes.add_picture(image_path, fmt['left'], fmt['top'], fmt['width'], fmt['height'])
                    apply_shape_formatting(new_shape, fmt)
                    add_detail(f"Fallback replaced placeholder with {os.path.basename(image_path)}", "success")
                    return True
                except Exception as e2:
                    add_detail(f"fallback failed: {e2}", "error")
                    return False

        elif shape_info['type'] == 'picture':
            shape = shape_info['shape']
            fmt = shape_info['fmt']
            # remove old
            try:
                el = shape._element
                el.getparent().remove(el)
            except Exception:
                pass
            # add new
            try:
                new_shape = slide.shapes.add_picture(image_path, fmt['left'], fmt['top'], fmt['width'], fmt['height'])
                apply_shape_formatting(new_shape, fmt)
                add_detail(f"Replaced picture with {os.path.basename(image_path)}", "success")
                return True
            except Exception as e:
                add_detail(f"replace picture failed: {e}", "error")
                return False

    except Exception as e:
        add_detail(f"general replace error: {e}", "error")
        return False

def add_images_by_positions(slide, image_paths, positions):
    """Use positions list (formatting dicts) to add images preserving sizes and positions."""
    added = 0
    for i, fmt in enumerate(positions):
        if i >= len(image_paths):
            break
        path = image_paths[i]
        try:
            new_shape = slide.shapes.add_picture(path, fmt['left'], fmt['top'], fmt['width'], fmt['height'])
            apply_shape_formatting(new_shape, fmt)
            added += 1
            add_detail(f"Added image {os.path.basename(path)} by template position", "success")
        except Exception as e:
            add_detail(f"add by pos failed for {path}: {e}", "warning")
    return added

# ---------------------------
# Processing pipeline
# ---------------------------
def process_uploads(pptx_bytes, zip_bytes, image_order="alphabetical", mismatch_action="truncate"):
    """
    Main processing:
    - extract zip to temp dir
    - gather subfolders that contain images
    - for each folder: create a new slide (based on first slide's layout), set title, replace images/placeholders
    - return bytes of final pptx and stats
    """
    # setup temp
    tmpdir = tempfile.mkdtemp(prefix="slide_sync_")
    try:
        # save uploaded pptx to temp path (so Presentation can open by path or bytes)
        template_path = os.path.join(tmpdir, "template.pptx")
        with open(template_path, "wb") as f:
            f.write(pptx_bytes.getvalue())

        # extract zip
        zip_path = os.path.join(tmpdir, "upload.zip")
        with open(zip_path, "wb") as f:
            f.write(zip_bytes.getvalue())

        extract_dir = os.path.join(tmpdir, "extracted")
        os.makedirs(extract_dir, exist_ok=True)
        with zipfile.ZipFile(zip_path, "r") as z:
            z.extractall(extract_dir)

        # find subfolders that contain images
        all_subs = [entry.path for entry in os.scandir(extract_dir) if entry.is_dir()]
        valid_folders = []
        for fld in sorted(all_subs):
            imgs = [f for f in os.listdir(fld) if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.webp'))]
            if imgs:
                valid_folders.append((fld, sorted(imgs)))
            else:
                add_detail(f"Folder {os.path.basename(fld)} has no images and will be skipped", "warning")

        if not valid_folders:
            raise ValueError("No valid folders with images found in ZIP.")

        # Open presentation
        prs = Presentation(template_path)
        if len(prs.slides) == 0:
            raise ValueError("Template PPTX has no slides.")

        # analyze first slide
        first_slide = prs.slides[0]
        template_infos, template_positions = get_template_slots(first_slide)
        expected_slots = max(len(template_infos), len(template_positions))
        add_detail(f"Template has {len(template_infos)} image shapes and {len(template_positions)} template positions (expected slots {expected_slots})", "info")

        created_slides = 0
        replaced_total = 0

        # For each valid folder create a new slide and populate
        for idx, (fld, imgs) in enumerate(valid_folders):
            add_detail(f"Processing folder {os.path.basename(fld)} ({len(imgs)} images)...", "info")
            # image ordering
            if image_order == "random":
                random.shuffle(imgs)
            else:
                imgs.sort()

            # Create new slide using the same layout as first slide
            try:
                new_slide = prs.slides.add_slide(first_slide.slide_layout)
            except Exception:
                # fallback to blank layout if needed
                new_slide = prs.slides.add_slide(prs.slide_layouts[6])

            # set slide title from folder name if possible
            try:
                title_shapes = [s for s in new_slide.shapes if s.is_placeholder and s.placeholder_format.type == PP_PLACEHOLDER.TITLE]
                if title_shapes:
                    title_shapes[0].text = os.path.basename(fld)
                else:
                    # optional: add small textbox title if no placeholder
                    tx = new_slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(0.6))
                    tf = tx.text_frame
                    tf.text = os.path.basename(fld)
            except Exception as e:
                add_detail(f"Could not set title for slide {os.path.basename(fld)}: {e}", "warning")

            # collect shape infos of the newly added slide (placeholders + pics)
            new_infos = extract_image_shapes_info(new_slide)

            # prepare full image paths
            full_image_paths = [os.path.join(fld, im) for im in imgs]

            # mismatch handling options
            if new_infos:
                # if user wants to skip folder when mismatch and counts differ
                if mismatch_action == "skip_folder" and len(full_image_paths) != len(new_infos):
                    add_detail(f"Skipping folder {os.path.basename(fld)} due to mismatch (images {len(full_image_paths)} vs slots {len(new_infos)})", "info")
                    # remove the added slide because skipped
                    try:
                        # remove last slide
                        prs.slides._sldIdLst.remove(prs.slides._sldIdLst[-1])
                    except Exception:
                        pass
                    continue

                # For 'truncate' or 'repeat' handle mapping
                for i, shape_info in enumerate(new_infos):
                    if mismatch_action == "truncate" and i >= len(full_image_paths):
                        break
                    # select image path with repeat if needed
                    img_path = full_image_paths[i % len(full_image_paths)]
                    success = replace_image_in_shape(new_slide, shape_info, img_path)
                    if success:
                        replaced_total += 1

            elif template_positions:
                # fallback: use template positions
                added = add_images_by_positions(new_slide, full_image_paths, template_positions)
                replaced_total += added
            else:
                # ultimate fallback: insert first image in default area
                try:
                    fp = full_image_paths[0]
                    new_shape = new_slide.shapes.add_picture(fp, Inches(1), Inches(1.5), Inches(8), Inches(4.5))
                    replaced_total += 1
                except Exception as e:
                    add_detail(f"Failed to add fallback image for {os.path.basename(fld)}: {e}", "error")

            created_slides += 1
            add_detail(f"Finished folder {os.path.basename(fld)}: created slide, replaced images where possible.", "success")

        # Save result to bytes buffer
        out_buf = BytesIO()
        prs.save(out_buf)
        out_buf.seek(0)

        # return bytes + stats
        stats = {"created_slides": created_slides, "replaced_total": replaced_total, "folders_processed": len(valid_folders)}
        return out_buf, stats

    finally:
        # cleanup
        try:
            shutil.rmtree(tmpdir)
        except Exception:
            pass

# ---------------------------
# Main UI flow
# ---------------------------
if submit:
    clear_details()
    if not uploaded_pptx or not uploaded_zip:
        st.error("Please upload both a .pptx and a .zip file.")
    else:
        try:
            st.info("Preparing files...")
            # Read uploaded files as BytesIO (so we can reuse)
            pptx_bytes = BytesIO(uploaded_pptx.read())
            zip_bytes = BytesIO(uploaded_zip.read())

            add_detail("Starting processing pipeline...", "info")
            # process
            result_buf, stats = process_uploads(pptx_bytes, zip_bytes, image_order=order_opt, mismatch_action=mismatch_opt)
            add_detail(f"Created slides: {stats['created_slides']}, Replaced images: {stats['replaced_total']}", "success")

            if stats['created_slides'] == 0:
                st.error("No slides were created. Check logs/details for issues.")
                show_details(expanded=True)
            else:
                # Build output filename
                original_name = os.path.splitext(uploaded_pptx.name)[0]
                out_name = f"{original_name}_Modified.pptx"
                st.success(f"Processing complete — {stats['created_slides']} slides created, {stats['replaced_total']} images replaced.")
                st.download_button("⬇️ Download modified PPTX", data=result_buf.getvalue(), file_name=out_name, mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
                show_details(expanded=False)

        except Exception as e:
            st.error(f"Processing failed: {e}")
            add_detail(f"Processing exception: {e}", "error")
            show_details(expanded=True)
else:
    st.info("Upload a PowerPoint template and a ZIP (folders with images). See instructions below.")
    with st.expander("Instructions & notes"):
        st.markdown("""
        - The app uses the first slide as template (layout, placeholders & picture positions).
        - Each subfolder inside the ZIP will produce one slide; the folder name will be used as slide title.
        - Supported images: PNG/JPG/GIF/BMP/TIFF/WEBP.
        - If template has placeholders it prefers to fill them; otherwise it uses picture shapes or template positions.
        - Options: image ordering (alphabetical/random) and mismatch handling (truncate/repeat/skip_folder).
        - Output file name = `<original>_Modified.pptx`.
        """)

# show details panel at bottom
show_details(expanded=False)
