import streamlit as st
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import os
import zipfile
import tempfile
import shutil

st.set_page_config(page_title="PowerPoint Generator", layout="centered")

st.title("📊 PowerPoint Generator")
st.write("Upload a PowerPoint template and a ZIP file with folders of images. The app will generate a slide for each folder, with up to 6 images placed as in the template.")

# --- File Upload ---
pptx_file = st.file_uploader("Upload PowerPoint Template (.pptx)", type=["pptx"])
zip_file = st.file_uploader("Upload ZIP file containing folders of images", type=["zip"])

if st.button("Generate Presentation") and pptx_file and zip_file:
    with st.spinner("Processing..."):

        # Temporary working directory
        with tempfile.TemporaryDirectory() as tmpdir:
            pptx_path = os.path.join(tmpdir, "template.pptx")
            zip_path = os.path.join(tmpdir, "images.zip")
            output_path = os.path.join(tmpdir, "generated.pptx")

            # Save uploaded files
            with open(pptx_path, "wb") as f:
                f.write(pptx_file.read())
            with open(zip_path, "wb") as f:
                f.write(zip_file.read())

            # Unzip images
            extract_dir = os.path.join(tmpdir, "unzipped")
            with zipfile.ZipFile(zip_path, "r") as zip_ref:
                zip_ref.extractall(extract_dir)

            # --- PowerPoint generation function ---
            def generate_pptx(template_path, folders_path, output_path):
                prs_template = Presentation(template_path)
                base_slide = prs_template.slides[0]

                # Get image placeholder positions
                image_shapes = [shape for shape in base_slide.shapes if shape.shape_type == MSO_SHAPE_TYPE.PICTURE]
                image_positions = [(shape.left, shape.top, shape.height) for shape in image_shapes]

                # Create a new presentation with same dimensions
                prs = Presentation()
                prs.slide_width = prs_template.slide_width
                prs.slide_height = prs_template.slide_height

                folders = sorted([f for f in os.listdir(folders_path) if os.path.isdir(os.path.join(folders_path, f))])

                for folder in folders:
                    images = sorted([
                        os.path.join(folders_path, folder, f)
                        for f in os.listdir(os.path.join(folders_path, folder))
                        if f.lower().endswith((".jpg", ".jpeg", ".png"))
                    ])

                    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank layout

                    # Add folder name as title
                    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(1))
                    text_frame = title_box.text_frame
                    p = text_frame.paragraphs[0]
                    run = p.add_run()
                    run.text = folder
                    font = run.font
                    font.size = Pt(28)
                    font.bold = True
                    font.color.rgb = RGBColor(0, 0, 128)

                    # Add images to slide
                    for idx, (left, top, height) in enumerate(image_positions):
                        if idx < len(images):
                            slide.shapes.add_picture(images[idx], left, top, height=height)

                prs.save(output_path)

            # Call generator
            generate_pptx(pptx_path, extract_dir, output_path)

            # Read final file
            with open(output_path, "rb") as f:
                final_pptx = f.read()

            st.success("✅ Presentation generated successfully!")
            st.download_button("📥 Download PowerPoint", data=final_pptx, file_name="generated_presentation.pptx")

else:
    st.info("Please upload both a .pptx file and a .zip file to start.")
