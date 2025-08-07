import streamlit as st
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import os
import zipfile
import tempfile
from io import BytesIO

st.set_page_config(page_title="PPTX Generator", layout="centered")
st.title("📸 PowerPoint Slide Generator")

pptx_file = st.file_uploader("Upload PowerPoint Template (.pptx)", type=["pptx"])
zip_file = st.file_uploader("Upload ZIP of Folders with Images", type=["zip"])

def clone_slide(pres, slide):
    """Clone a slide and return the copy"""
    slide_id = slide.slide_id
    slide_layout = slide.slide_layout
    new_slide = pres.slides.add_slide(slide_layout)
    for shape in slide.shapes:
        el = shape.element
        new_slide.shapes._spTree.insert_element_before(el, 'p:extLst')
    return new_slide

def replace_images_on_slide(slide, new_images):
    pic_shapes = [shape for shape in slide.shapes if shape.shape_type == MSO_SHAPE_TYPE.PICTURE]
    for idx, shape in enumerate(pic_shapes):
        if idx < len(new_images):
            left, top, height = shape.left, shape.top, shape.height
            slide.shapes._spTree.remove(shape._element)
            slide.shapes.add_picture(new_images[idx], left, top, height=height)

def replace_title(slide, new_title):
    for shape in slide.shapes:
        if shape.has_text_frame:
            shape.text = new_title
            break

if pptx_file and zip_file and st.button("Generate PPTX"):
    with tempfile.TemporaryDirectory() as tmpdir:
        pptx_path = os.path.join(tmpdir, "template.pptx")
        zip_path = os.path.join(tmpdir, "images.zip")

        # Save uploaded files
        with open(pptx_path, "wb") as f:
            f.write(pptx_file.read())
        with open(zip_path, "wb") as f:
            f.write(zip_file.read())

        # Extract images
        extract_dir = os.path.join(tmpdir, "extracted")
        with zipfile.ZipFile(zip_path, "r") as zip_ref:
            zip_ref.extractall(extract_dir)

        prs = Presentation(pptx_path)
        template_slide = prs.slides[0]

        folders = sorted([f for f in os.listdir(extract_dir) if os.path.isdir(os.path.join(extract_dir, f))])

        # Create new presentation using same template
        final_prs = Presentation(pptx_path)
        while len(final_prs.slides) > 0:
            r_id = final_prs.slides._sldIdLst[0].rId
            final_prs.part.drop_rel(r_id)
            del final_prs.slides._sldIdLst[0]

        for folder in folders:
            folder_path = os.path.join(extract_dir, folder)
            image_files = sorted([
                os.path.join(folder_path, f)
                for f in os.listdir(folder_path)
                if f.lower().endswith((".png", ".jpg", ".jpeg"))
            ])

            # Clone and update the slide
            new_slide = clone_slide(final_prs, template_slide)
            replace_images_on_slide(new_slide, image_files)
            replace_title(new_slide, folder)

        # Export final presentation
        output = BytesIO()
        final_prs.save(output)
        output.seek(0)

        st.success("✅ Done! Download your presentation:")
        st.download_button("📥 Download PPTX", output, file_name="final_presentation.pptx")
