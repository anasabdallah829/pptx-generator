import streamlit as st
from pptx import Presentation
from pptx.util import Inches
from pptx.enum.shapes import MSO_SHAPE_TYPE
import os
import zipfile
import tempfile
from io import BytesIO

st.set_page_config(page_title="PPTX Generator", layout="centered")
st.title("📸 PowerPoint Slide Generator")

pptx_file = st.file_uploader("Upload PowerPoint Template (.pptx)", type=["pptx"])
zip_file = st.file_uploader("Upload ZIP of Folders with Images", type=["zip"])

def get_picture_placeholders(slide):
    return [shape for shape in slide.shapes if shape.shape_type == MSO_SHAPE_TYPE.PICTURE]

def add_images_to_slide(slide, image_paths, positions):
    for img_path, ref_shape in zip(image_paths, positions):
        slide.shapes.add_picture(
            img_path,
            left=ref_shape.left,
            top=ref_shape.top,
            width=ref_shape.width,
            height=ref_shape.height
        )

def add_title(slide, text):
    for shape in slide.shapes:
        if shape.has_text_frame:
            shape.text = text
            break

if pptx_file and zip_file and st.button("Generate PPTX"):
    with tempfile.TemporaryDirectory() as tmpdir:
        pptx_path = os.path.join(tmpdir, "template.pptx")
        zip_path = os.path.join(tmpdir, "images.zip")

        with open(pptx_path, "wb") as f:
            f.write(pptx_file.read())

        with open(zip_path, "wb") as f:
            f.write(zip_file.read())

        extract_dir = os.path.join(tmpdir, "unzipped")
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(extract_dir)

        prs = Presentation(pptx_path)
        template_slide = prs.slides[0]
        layout = template_slide.slide_layout
        ref_images = get_picture_placeholders(template_slide)

        folders = sorted([
            f for f in os.listdir(extract_dir)
            if os.path.isdir(os.path.join(extract_dir, f))
        ])

        # الشريحة الأولى تبقى كما هي
        for idx, folder in enumerate(folders):
            folder_path = os.path.join(extract_dir, folder)
            image_files = sorted([
                os.path.join(folder_path, f)
                for f in os.listdir(folder_path)
                if f.lower().endswith((".png", ".jpg", ".jpeg"))
            ])

            if idx == 0:
                # الشريحة الأولى يتم تعديل عنوانها وصورها
                slide = template_slide
            else:
                slide = prs.slides.add_slide(layout)

            add_title(slide, folder)
            add_images_to_slide(slide, image_files, ref_images)

        output = BytesIO()
        prs.save(output)
        output.seek(0)

        st.success("✅ تم إنشاء العرض بنجاح!")
        st.download_button("📥 تحميل العرض النهائي", output, file_name="final_presentation.pptx")
