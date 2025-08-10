import streamlit as st
import zipfile
import os
import io
from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
import shutil
from pptx.util import Inches
import random

# Set Streamlit page configuration
st.set_page_config(page_title="Slide-Sync-Images", layout="wide", initial_sidebar_state="expanded")

# --- App Header and Description ---
st.title("🖼️ Slide-Sync-Images: PowerPoint Image Replacer")
st.markdown("""
_An easy-to-use tool for quickly generating new PowerPoint slides from templates and image folders._
""")
st.markdown("---")

# --- Sidebar for Instructions ---
with st.sidebar:
    st.header("📖 Instructions")
    st.markdown("""
    1.  **Upload a PowerPoint Template (.pptx)**: This file's first slide will be used as a template. It should contain placeholders or regular images where you want new images to appear.
    2.  **Upload a ZIP file**: This file must contain one or more folders, with each folder containing the images for a single new slide.
    3.  **Choose your settings**: Decide whether you want to place images sequentially or randomly.
    4.  **Click "Start Processing"**: The app will generate a new PowerPoint file with a slide for each folder in your ZIP file.
    """)

# --- Main Interface ---

st.header("📂 File Uploads")
uploaded_pptx = st.file_uploader("Upload PowerPoint Template (.pptx)", type=["pptx"], key="pptx_uploader")
uploaded_zip = st.file_uploader("Upload ZIP file with image folders", type=["zip"], key="zip_uploader")

st.markdown("---")

st.header("⚙️ Processing Settings")
image_order_option = st.radio(
    "How should images be placed in the slides?",
    ("In order (Default)", "Randomly"),
    index=0
)

st.markdown("---")

show_details = st.checkbox("Show detailed processing log", value=False)


def analyze_first_slide(prs):
    """
    Analyzes the first slide of the presentation to find all image shapes.
    """
    if len(prs.slides) == 0:
        return False, "No slides found in the template."

    first_slide = prs.slides[0]
    
    PICTURE_SHAPE_TYPES = (13, 21)
    
    picture_placeholders = [
        shape for shape in first_slide.shapes
        if shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.PICTURE
    ]
    regular_pictures = [
        shape for shape in first_slide.shapes
        if hasattr(shape, 'shape_type') and shape.shape_type in PICTURE_SHAPE_TYPES
    ]
    
    total_image_slots = len(picture_placeholders) + len(regular_pictures)

    return True, {
        'placeholders': len(picture_placeholders),
        'regular_pictures': len(regular_pictures),
        'total_slots': total_image_slots,
        'slide_layout': first_slide.slide_layout
    }


def get_image_shapes(slide):
    """
    Extracts and sorts all image shapes (placeholders and regular pictures) from a slide.
    """
    PICTURE_SHAPE_TYPES = (13, 21)
    
    image_shapes = []
    for shape in slide.shapes:
        if shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.PICTURE:
            image_shapes.append(shape)
        elif hasattr(shape, 'shape_type') and shape.shape_type in PICTURE_SHAPE_TYPES:
            image_shapes.append(shape)
            
    image_shapes.sort(key=lambda s: (s.top, s.left))
    return image_shapes


def main():
    if uploaded_pptx and uploaded_zip:
        if "process_started" not in st.session_state:
            st.session_state.process_started = False

        if st.button("🚀 Start Processing") or st.session_state.process_started:
            st.session_state.process_started = True
            
            temp_dir = None
            try:
                with st.spinner("Checking and extracting files..."):
                    zip_bytes = io.BytesIO(uploaded_zip.read())
                    with zipfile.ZipFile(zip_bytes, "r") as zip_ref:
                        temp_dir = "temp_images"
                        if os.path.exists(temp_dir):
                            shutil.rmtree(temp_dir)
                        os.makedirs(temp_dir)
                        zip_ref.extractall(temp_dir)
                    
                    all_items = os.listdir(temp_dir)
                    folder_paths = [os.path.join(temp_dir, item) for item in all_items if os.path.isdir(os.path.join(temp_dir, item))]
                
                if not folder_paths:
                    st.error("❌ The ZIP file doesn't contain any image folders.")
                    st.stop()
                
                folder_paths.sort()
                st.success(f"✅ Found **{len(folder_paths)}** image folders to process.")

                prs = Presentation(io.BytesIO(uploaded_pptx.read()))
                
                st.info("🔍 Analyzing the template slide...")
                ok, analysis_result = analyze_first_slide(prs)
                if not ok:
                    st.error(f"❌ Error: {analysis_result}")
                    st.stop()
                
                st.success("✅ Template analysis complete.")
                col1, col2, col3 = st.columns(3)
                with col1: st.metric("Image Placeholders", analysis_result['placeholders'])
                with col2: st.metric("Regular Images", analysis_result['regular_pictures'])
                with col3: st.metric("Total Image Slots", analysis_result['total_slots'])
                
                first_slide = prs.slides[0]
                template_image_shapes = get_image_shapes(first_slide)
                
                if not template_image_shapes:
                    st.warning("⚠ The template slide has no image slots. We'll add one image per slide.")
                    slide_layout = prs.slide_layouts[6]
                else:
                    slide_layout = analysis_result['slide_layout']

                mismatch_folders = []
                for fp in folder_paths:
                    imgs = [f for f in os.listdir(fp) if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.webp'))]
                    if len(imgs) != len(template_image_shapes):
                        mismatch_folders.append((os.path.basename(fp), len(imgs), len(template_image_shapes)))
                
                if mismatch_folders and 'mismatch_action' not in st.session_state:
                    with st.form("mismatch_form"):
                        st.warning("⚠️ Mismatch detected! Some folders have a different number of images than the template slots.")
                        for name, img_count, _ in mismatch_folders:
                            st.write(f"- Folder `{name}` has {img_count} images.")
                        st.markdown(f"**Number of image slots in template: {len(template_image_shapes)}**")

                        choice_text = st.radio(
                            "How should we handle folders with a different number of images?",
                            ("Truncate (use only up to the number of slots)", "Repeat (cycle through images to fill all slots)", "Skip (ignore folders with a mismatch)", "Stop (abort the entire process)"),
                            index=0
                        )
                        submit_choice = st.form_submit_button("✅ Confirm and Continue")

                    if submit_choice:
                        st.session_state['mismatch_action'] = {
                            "Truncate (use only up to the number of slots)": 'truncate',
                            "Repeat (cycle through images to fill all slots)": 'repeat',
                            "Skip (ignore folders with a mismatch)": 'skip_folder',
                            "Stop (abort the entire process)": 'stop'
                        }.get(choice_text)
                    else:
                        st.stop()
                
                if 'mismatch_action' in st.session_state:
                    mismatch_action = st.session_state['mismatch_action']
                else:
                    mismatch_action = 'truncate'

                if mismatch_action == 'stop':
                    st.error("❌ Process aborted by user choice.")
                    st.stop()

                st.info("🔄 Generating new slides...")
                total_replaced = 0
                created_slides = 0

                progress_bar = st.progress(0)
                status_text = st.empty()

                for folder_idx, folder_path in enumerate(folder_paths):
                    folder_name = os.path.basename(folder_path)
                    status_text.text(f"Processing folder {folder_idx + 1}/{len(folder_paths)}: **{folder_name}**")

                    imgs = [f for f in os.listdir(folder_path) if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.webp'))]
                    
                    if not imgs:
                        if show_details:
                            st.warning(f"⚠ Folder '{folder_name}' is empty. Skipping.")
                        continue
                    
                    if image_order_option == "Randomly":
                        random.shuffle(imgs)
                    else:
                        imgs.sort()

                    if mismatch_action == 'skip_folder' and len(imgs) != len(template_image_shapes):
                        if show_details:
                            st.info(f"ℹ Skipping folder '{folder_name}' due to image count mismatch.")
                        continue

                    new_slide = prs.slides.add_slide(slide_layout)
                    created_slides += 1
                    
                    new_image_shapes = get_image_shapes(new_slide)
                    
                    replaced_count = 0
                    for i, new_shape in enumerate(new_image_shapes):
                        if mismatch_action == 'truncate' and i >= len(imgs):
                            break
                        
                        image_filename = imgs[i % len(imgs)]
                        image_path = os.path.join(folder_path, image_filename)
                        
                        try:
                            # Use insert_picture if available (for placeholders)
                            new_shape.insert_picture(image_path)
                            replaced_count += 1
                        except AttributeError:
                            # Fallback for regular pictures to maintain position/size
                            left, top, width, height = new_shape.left, new_shape.top, new_shape.width, new_shape.height
                            new_shape.element.getparent().remove(new_shape.element)
                            new_slide.shapes.add_picture(
                                image_path, left, top, width, height
                            )
                            replaced_count += 1
                            
                    total_replaced += replaced_count
                    
                    try:
                        title_shapes = [shape for shape in new_slide.shapes
                                        if shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.TITLE]
                        if title_shapes:
                            title_shapes[0].text = folder_name
                        else:
                            textbox = new_slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
                            text_frame = textbox.text_frame
                            text_frame.text = folder_name
                    except Exception:
                        pass
                    
                    if show_details:
                        st.success(f"✅ Created slide for '{folder_name}' and replaced {replaced_count} images.")

                    progress_bar.progress((folder_idx + 1) / len(folder_paths))

                progress_bar.empty()
                status_text.empty()

                st.success("🎉 **Processing complete!**")
                
                col1, col2, col3 = st.columns(3)
                with col1: st.metric("Slides Added", created_slides)
                with col2: st.metric("Images Replaced", total_replaced)
                with col3: st.metric("Folders Processed", len(folder_paths))

                if created_slides == 0:
                    st.error("❌ No slides were added to the presentation.")
                    st.stop()

                original_name = os.path.splitext(uploaded_pptx.name)[0]
                output_filename = f"{original_name}_Updated.pptx"
                output_buffer = io.BytesIO()
                prs.save(output_buffer)
                output_buffer.seek(0)

                st.download_button(
                    label="⬇️ Download Updated Presentation",
                    data=output_buffer.getvalue(),
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    key="download_button"
                )

            except Exception as e:
                st.error(f"❌ An error occurred during processing: {e}")
                if show_details:
                    import traceback
                    st.error(f"Error details: {traceback.format_exc()}")
            finally:
                if temp_dir and os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)
    else:
        st.info("Awaiting file uploads... Please provide both a PowerPoint template and a ZIP file.")

if __name__ == '__main__':
    main()
