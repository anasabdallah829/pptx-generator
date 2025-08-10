import streamlit as st
import os
import zipfile
import io
import tempfile
from pptx import Presentation
from pptx.util import Inches
from pptx.enum.shapes import MSO_SHAPE

IMAGE_EXTENSIONS = ('.png', '.jpg', '.jpeg', '.gif', '.bmp')

def find_images_in_folder(folder_path):
    """إرجاع قائمة الصور في المجلد مرتبة أبجدياً"""
    return sorted([
        os.path.join(folder_path, f)
        for f in os.listdir(folder_path)
        if f.lower().endswith(IMAGE_EXTENSIONS)
    ])

def process_files_with_template(zip_file, pptx_file, continue_on_mismatch=True):
    try:
        prs = Presentation(pptx_file)

        if not prs.slides:
            raise ValueError("ملف PowerPoint لا يحتوي على شرائح.")

        # حفظ خصائص الصور من الشريحة الأولى
        template_slide = prs.slides[0]
        image_props = []
        for shape in template_slide.shapes:
            if shape.shape_type == MSO_SHAPE.PICTURE:
                image_props.append({
                    "left": shape.left,
                    "top": shape.top,
                    "width": shape.width,
                    "height": shape.height
                })

        st.info(f"تم العثور على {len(image_props)} صورة في القالب.")

        # استخراج ملفات ZIP
        with tempfile.TemporaryDirectory() as temp_dir:
            zip_content = io.BytesIO(zip_file.read())
            with zipfile.ZipFile(zip_content, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)

            folders = sorted([
                d for d in os.listdir(temp_dir)
                if os.path.isdir(os.path.join(temp_dir, d)) and not d.startswith('.')
            ])

            # حذف الشريحة الأولى (القالب)
            prs.slides._sldIdLst.remove(prs.slides._sldIdLst[0])

            for folder_name in folders:
                folder_path = os.path.join(temp_dir, folder_name)
                images = find_images_in_folder(folder_path)

                if len(images) != len(image_props):
                    msg = f"مجلد {folder_name}: عدد الصور ({len(images)}) يختلف عن القالب ({len(image_props)})."
                    if not continue_on_mismatch:
                        st.warning(msg + " تم إيقاف المعالجة.")
                        return None
                    else:
                        st.warning(msg + " سيتم الاستمرار.")

                # إنشاء شريحة جديدة
                slide_layout = prs.slide_layouts[6]  # شريحة فارغة
                slide = prs.slides.add_slide(slide_layout)

                # وضع الصور
                for idx, prop in enumerate(image_props):
                    if idx < len(images):
                        slide.shapes.add_picture(images[idx], prop["left"], prop["top"], prop["width"], prop["height"])
                    else:
                        # ترك مكان فارغ إذا لم تتوفر صورة
                        pass

        # حفظ النتيجة
        output = io.BytesIO()
        prs.save(output)
        output.seek(0)
        return output

    except Exception as e:
        st.error(f"حدث خطأ: {e}")
        return None

# واجهة Streamlit
st.title("استبدال الصور في PowerPoint بالقالب")
zip_file = st.file_uploader("رفع ملف ZIP يحتوي على المجلدات والصور", type="zip")
pptx_file = st.file_uploader("رفع ملف PPTX القالب", type="pptx")
continue_on_mismatch = st.checkbox("الاستمرار عند اختلاف عدد الصور", value=True)

if st.button("بدء المعالجة"):
    if zip_file and pptx_file:
        result = process_files_with_template(zip_file, pptx_file, continue_on_mismatch)
        if result:
            st.success("تمت المعالجة بنجاح!")
            st.download_button("تحميل الملف المعدل", data=result, file_name="output.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
    else:
        st.warning("الرجاء رفع الملفات المطلوبة.")
