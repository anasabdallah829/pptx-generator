import streamlit as st
import zipfile
import os
import io
from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.util import Inches

st.set_page_config(page_title="PowerPoint Image Replacer", layout="centered")
st.title("🔄 PowerPoint Image & Placeholder Replacer")

uploaded_pptx = st.file_uploader("📂 اختر ملف PowerPoint (.pptx)", type=["pptx"])
uploaded_zip = st.file_uploader("🗜️ اختر ملف ZIP يحتوي على مجلدات صور", type=["zip"])

if uploaded_pptx and uploaded_zip:
    if st.button("🚀 بدء المعالجة"):
        try:
            st.info("📦 جاري استخراج الصور من ملف ZIP...")
            zip_bytes = io.BytesIO(uploaded_zip.read())
            with zipfile.ZipFile(zip_bytes, "r") as zip_ref:
                temp_dir = "temp_images"
                if os.path.exists(temp_dir):
                    import shutil
                    shutil.rmtree(temp_dir)
                os.makedirs(temp_dir)
                zip_ref.extractall(temp_dir)

            # قراءة البوربوينت
            st.info("📄 جاري قراءة ملف PowerPoint...")
            prs = Presentation(io.BytesIO(uploaded_pptx.read()))

            # جمع مجلدات الصور
            folder_paths = [os.path.join(temp_dir, d) for d in os.listdir(temp_dir) if os.path.isdir(os.path.join(temp_dir, d))]
            if not folder_paths:
                st.error("❌ ملف ZIP لا يحتوي على مجلدات صور.")
                st.stop()

            slide_count = len(prs.slides)
            st.info(f"📊 الملف يحتوي على {slide_count} شريحة.")

            slide_index = 0
            replaced_count = 0
            for folder in folder_paths:
                images = [f for f in os.listdir(folder) if f.lower().endswith(('.png', '.jpg', '.jpeg'))]
                if not images:
                    st.warning(f"⚠ المجلد {os.path.basename(folder)} لا يحتوي على صور، تم تجاوزه.")
                    continue

                if slide_index >= slide_count:
                    break

                slide = prs.slides[slide_index]

                # وضع عنوان الشريحة من اسم المجلد
                title_shapes = [shape for shape in slide.shapes if shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.TITLE]
                if title_shapes:
                    title_shapes[0].text = os.path.basename(folder)

                img_idx = 0
                for shape in slide.shapes:
                    # استبدال في placeholder
                    if shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.PICTURE:
                        with open(os.path.join(folder, images[img_idx % len(images)]), "rb") as img_file:
                            shape.insert_picture(img_file)
                        replaced_count += 1
                        img_idx += 1

                    # استبدال الصور العادية
                    elif shape.shape_type == 13:  # 13 = Picture
                        left, top, width, height = shape.left, shape.top, shape.width, shape.height
                        slide.shapes._spTree.remove(shape._element)
                        with open(os.path.join(folder, images[img_idx % len(images)]), "rb") as img_file:
                            pic = slide.shapes.add_picture(img_file, left, top, width, height)
                        replaced_count += 1
                        img_idx += 1

                slide_index += 1

            if replaced_count == 0:
                st.error("❌ لم يتم العثور على أي صور أو Placeholders في العرض التقديمي.")
                st.stop()

            # حفظ الملف الجديد بنفس الاسم + _Modified
            original_name = os.path.splitext(uploaded_pptx.name)[0]
            output_filename = f"{original_name}_Modified.pptx"
            output_path = os.path.join(".", output_filename)
            prs.save(output_path)

            with open(output_path, "rb") as f:
                st.success(f"✅ تم استبدال {replaced_count} صورة بنجاح!")
                st.download_button("⬇ تحميل الملف المعدل", f, file_name=output_filename)

        except Exception as e:
            st.error(f"❌ خطأ أثناء المعالجة: {e}")
