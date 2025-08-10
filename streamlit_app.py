import streamlit as st
import zipfile
import os
import tempfile
from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.util import Inches, Pt
from io import BytesIO

st.set_page_config(page_title="PowerPoint Image Replacer", page_icon="📊")

st.title("📊 PowerPoint Image Replacer with Placeholders")

uploaded_pptx = st.file_uploader("📂 ارفع ملف PowerPoint (.pptx)", type=["pptx"])
uploaded_zip = st.file_uploader("🖼️ ارفع ملف الصور (.zip)", type=["zip"])

if uploaded_pptx and uploaded_zip:
    with st.status("⏳ جاري معالجة الملفات...", expanded=True) as status:
        # إنشاء مجلد مؤقت
        with tempfile.TemporaryDirectory() as tmpdir:
            pptx_path = os.path.join(tmpdir, uploaded_pptx.name)
            zip_path = os.path.join(tmpdir, uploaded_zip.name)

            # حفظ الملفات
            with open(pptx_path, "wb") as f:
                f.write(uploaded_pptx.getbuffer())
            with open(zip_path, "wb") as f:
                f.write(uploaded_zip.getbuffer())

            # فك ضغط الصور
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(tmpdir)

            # قراءة مجلدات الصور
            folders = [os.path.join(tmpdir, d) for d in os.listdir(tmpdir)
                       if os.path.isdir(os.path.join(tmpdir, d))]
            if not folders:
                st.error("❌ ملف ZIP لا يحتوي على مجلدات صور!")
                st.stop()

            # فتح العرض التقديمي
            prs = Presentation(pptx_path)

            # إحصاء الـ placeholders
            placeholder_count = 0
            for slide in prs.slides:
                for shape in slide.shapes:
                    if shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.PICTURE:
                        placeholder_count += 1

            st.write(f"📊 عدد الـ placeholders: {placeholder_count}")

            # معالجة الشرائح
            slide_index = 0
            for folder in folders:
                images = [os.path.join(folder, img) for img in os.listdir(folder)
                          if img.lower().endswith((".png", ".jpg", ".jpeg"))]

                if not images:
                    st.warning(f"⚠️ المجلد {os.path.basename(folder)} لا يحتوي على صور.")
                    continue

                # نسخ الشريحة الأولى
                template_slide = prs.slides[0]
                slide = prs.slides.add_slide(template_slide.slide_layout)

                # تعيين عنوان الشريحة
                for shape in slide.shapes:
                    if shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.TITLE:
                        shape.text = os.path.basename(folder)

                # استبدال الصور في الـ placeholders
                img_idx = 0
                for shape in slide.shapes:
                    if shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.PICTURE:
                        if img_idx < len(images):
                            pic = images[img_idx]
                            left, top, width, height = shape.left, shape.top, shape.width, shape.height
                            sp = slide.shapes.add_picture(pic, left, top, width, height)
                            slide.shapes._spTree.remove(shape._element)  # إزالة القديم
                            img_idx += 1

                slide_index += 1
                st.write(f"✅ تم إنشاء الشريحة {slide_index} بعنوان {os.path.basename(folder)}")

            if slide_index == 0:
                st.error("❌ لم يتم إنشاء أي شريحة جديدة. تحقق من أن الملف يحتوي على placeholders للصور.")
                st.stop()

            # حفظ الملف المعدل
            output_filename = uploaded_pptx.name.replace(".pptx", "_Modified.pptx")
            output_path = os.path.join(tmpdir, output_filename)
            prs.save(output_path)

            # تنزيل الملف
            with open(output_path, "rb") as f:
                st.download_button("📥 تحميل العرض المعدل", f, file_name=output_filename)

            status.update(label="✅ تم الانتهاء من المعالجة", state="complete")
