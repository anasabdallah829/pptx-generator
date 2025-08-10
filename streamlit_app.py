import streamlit as st
from pptx import Presentation
from pptx.util import Inches
import zipfile
import os
import tempfile
import shutil

# ====== دالة لاستبدال الصور في الشريحة ======
def replace_images_in_slide(slide, image_paths):
    """
    استبدال الصور في شريحة PowerPoint بقائمة من الصور الجديدة.
    يتم التعامل مع أي صور موجودة في الشريحة واستبدالها حسب الترتيب.
    """
    img_index = 0
    for shape in slide.shapes:
        if shape.shape_type == 13:  # رقم 13 هو نوع الصورة في PPTX
            if img_index < len(image_paths):
                # حذف الصورة القديمة
                x, y, cx, cy = shape.left, shape.top, shape.width, shape.height
                slide.shapes._spTree.remove(shape._element)
                # إدراج الصورة الجديدة
                slide.shapes.add_picture(image_paths[img_index], x, y, cx, cy)
                img_index += 1

# ====== الدالة الرئيسية للتطبيق ======
def process_pptx(template_pptx, images_zip):
    """
    قراءة ملف PowerPoint وقائمة صور من ملف مضغوط، واستبدال الصور في كل شريحة.
    """
    # إنشاء مجلد مؤقت
    temp_dir = tempfile.mkdtemp()

    # حفظ الملفات المرفوعة في مجلد مؤقت
    template_path = os.path.join(temp_dir, "template.pptx")
    with open(template_path, "wb") as f:
        f.write(template_pptx.getbuffer())

    zip_path = os.path.join(temp_dir, "images.zip")
    with open(zip_path, "wb") as f:
        f.write(images_zip.getbuffer())

    # فك ضغط الصور
    extract_path = os.path.join(temp_dir, "images")
    os.makedirs(extract_path, exist_ok=True)
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(extract_path)

    # تحميل البوربوينت
    prs = Presentation(template_path)

    # جلب المجلدات بالترتيب
    folders = sorted(os.listdir(extract_path))
    for i, folder in enumerate(folders):
        folder_path = os.path.join(extract_path, folder)
        if os.path.isdir(folder_path):
            image_files = sorted([
                os.path.join(folder_path, img)
                for img in os.listdir(folder_path)
                if img.lower().endswith((".png", ".jpg", ".jpeg"))
            ])
            if i < len(prs.slides):
                replace_images_in_slide(prs.slides[i], image_files)

    # حفظ النتيجة
    output_path = os.path.join(temp_dir, "output.pptx")
    prs.save(output_path)

    # قراءة الملف الناتج وإرجاعه للتحميل
    with open(output_path, "rb") as f:
        pptx_bytes = f.read()

    shutil.rmtree(temp_dir)  # تنظيف الملفات المؤقتة
    return pptx_bytes


# ====== واجهة Streamlit ======
st.title("📊 أداة استبدال الصور في PowerPoint")

template_pptx = st.file_uploader("ارفع قالب PowerPoint (.pptx)", type=["pptx"])
images_zip = st.file_uploader("ارفع ملف الصور المضغوط (.zip)", type=["zip"])

if st.button("بدء المعالجة"):
    if not template_pptx or not images_zip:
        st.error("الرجاء رفع كل من ملف PowerPoint وملف الصور المضغوط.")
    else:
        try:
            output_file = process_pptx(template_pptx, images_zip)
            st.success("✅ تم استبدال الصور بنجاح!")
            st.download_button(
                label="📥 تحميل الملف الناتج",
                data=output_file,
                file_name="output.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
        except Exception as e:
            st.error(f"حدث خطأ أثناء المعالجة: {e}")
