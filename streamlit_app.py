import streamlit as st
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import zipfile
import os
import tempfile
import shutil
from io import BytesIO

st.set_page_config(page_title="PPTX Image Replacer", layout="centered")
st.title("📸 PowerPoint Image Replacer Tool")

pptx_file = st.file_uploader("Upload your PowerPoint (.pptx) template", type=["pptx"])
zip_file = st.file_uploader("Upload a ZIP file with folders of images", type=["zip"])

if pptx_file and zip_file:
    if st.button("Start Processing"):
        with tempfile.TemporaryDirectory() as tmpdir:
            # Save and extract uploaded files
            pptx_path = os.path.join(tmpdir, "template.pptx")
            with open(pptx_path, "wb") as f:
                f.write(pptx_file.read())

            zip_path = os.path.join(tmpdir, "images.zip")
            with open(zip_path, "wb") as f:
                f.write(zip_file.read())

            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(os.path.join(tmpdir, "images"))

            # Load template
            prs = Presentation(pptx_path)
            base_slide = prs.slides[0]

            # Process each folder
            image_base_path = os.path.join(tmpdir, "images")
            folders = sorted([f for f in os.listdir(image_base_path) if os.path.isdir(os.path.join(image_base_path, f))])

            # New presentation
            final_prs = Presentation()
            final_prs.slide_width = prs.slide_width
            final_prs.slide_height = prs.slide_height

            for folder in folders:
                image_folder = os.path.join(image_base_path, folder)
                images = sorted([
                    os.path.join(image_folder, f)
                    for f in os.listdir(image_folder)
                    if f.lower().endswith((".png", ".jpg", ".jpeg"))
                ])

                # Duplicate base slide
                slide_layout = prs.slide_layouts[0]
                new_slide = final_prs.slides.add_slide(slide_layout)

                for shape in base_slide.shapes:
                    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                        new_slide.shapes._spTree.remove(shape._element)

                idx = 0
                for shape in base_slide.shapes:
                    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE and idx < len(images):
                        left = shape.left
                        top = shape.top
                        height = shape.height
                        final_shape = new_slide.shapes.add_picture(images[idx], left, top, height=height)
                        idx += 1
                    elif shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
                        el = shape.element
                        new_slide.shapes._spTree.insert_element_before(el, 'p:extLst')

            # Save final file
            output = BytesIO()
            final_prs.save(output)
            output.seek(0)

            st.success("✅ Done! Download your new PPTX:")
            st.download_button("📥 Download Final PPTX", output, file_name="final_presentation.pptx")
