import streamlit as st
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Inches
import zipfile
import os
import tempfile
import shutil
from io import BytesIO

st.set_page_config(page_title="📸 PPTX Image Slide Generator", layout="centered")
st.title("📸 PowerPoint Generator from Images")

pptx_file = st.file_uploader("Upload PowerPoint Template (.pptx)", type=["pptx"])
zip_file = st.file_uploader("Upload ZIP of Folders with Images", type=["zip"])

if pptx_file and zip_file and st.button("Generate PowerPoint"):
    with tempfile.TemporaryDirectory() as tmpdir:
        # Save uploaded files
        pptx_path = os.path.join(tmpdir, "template.pptx")
        with open(pptx_path, "wb") as f:
            f.write(pptx_file.read())

        zip_path = os.path.join(tmpdir, "images.zip")
        with open(zip_path, "wb") as f:
            f.write(zip_file.read())

        # Unzip images
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(os.path.join(tmpdir, "images"))

        # Load template
        prs = Presentation(pptx_path)
        base_slide = prs.slides[0]
        base_shapes = list(base_slide.shapes)

        # Create new presentation
        final_prs = Presentation()
        final_prs.slide_width = prs.slide_width
        final_prs.slide_height = prs.slide_height

        # Process folders
        image_base_path = os.path.join(tmpdir, "images")
        folders = sorted([f for f in os.listdir(image_base_path) if os.path.isdir(os.path.join(image_base_path, f))])

        for folder in folders:
            folder_path = os.path.join(image_base_path, folder)
            images = sorted([
                os.path.join(folder_path, f)
                for f in os.listdir(folder_path)
                if f.lower().endswith((".png", ".jpg", ".jpeg"))
            ])

            # Create slide with same layout
            layout = prs.slide_layouts[0]
            slide = final_prs.slides.add_slide(layout)

            img_idx = 0
            for shape in base_shapes:
                # Replace text placeholder (title)
                if shape.has_text_frame and "title" in shape.name.lower():
                    new_shape = slide.shapes.title
                    if new_shape:
                        new_shape.text = folder
                # Copy textboxes or images
                elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE and img_idx < len(images):
                    slide.shapes.add_picture(images[img_idx], shape.left, shape.top, height=shape.height)
                    img_idx += 1
                elif shape.shape_type != MSO_SHAPE_TYPE.PICTURE and not shape.has_text_frame:
                    new_shape = slide.shapes._spTree.insert_element_before(shape.element, 'p:extLst')

        # Output
        output = BytesIO()
        final_prs.save(output)
        output.seek(0)

        st.success("✅ Done! Download your customized PowerPoint:")
        st.download_button("📥 Download Final PPTX", output, file_name="generated_presentation.pptx")
