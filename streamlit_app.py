import streamlit as st
from pptx import Presentation
from pptx.util import Inches
from pptx.enum.shapes import MSO_SHAPE, PP_PLACEHOLDER
import zipfile
import os
import tempfile
import shutil

# امتدادات ملفات الصور المقبولة
IMAGE_EXTENSIONS = (".png", ".jpg", ".jpeg", ".gif", ".bmp")

# ====== دالة لاستبدال الصور في الشريحة ======
def replace_images_in_slide(slide, image_paths):
    """
    استبدال الصور والعناصر النائبة للصور في شريحة PowerPoint بقائمة من الصور الجديدة.
    """
    img_index = 0
    
    # قائمة بخصائص الصور (موضع وحجم) ليتم إضافتها لاحقًا
    image_replacements = []

    for shape in slide.shapes:
        # التحقق إذا كان الشكل عبارة عن صورة عادية
        # تم تصحيح الخطأ هنا: استخدام MSO_SHAPE.PICTURE
        if shape.shape_type == MSO_SHAPE.PICTURE:
            if img_index < len(image_paths):
                x, y, cx, cy = shape.left, shape.top, shape.width, shape.height
                image_replacements.append({
                    "path": image_paths[img_index],
                    "pos": (x, y),
                    "size": (cx, cy)
                })
                img_index += 1
                slide.shapes._spTree.remove(shape._element)
        
        # التحقق إذا كان الشكل عبارة عن عنصر نائب للصورة
        elif shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.PICTURE:
            if img_index < len(image_paths):
                shape.insert_picture(image_paths[img_index])
                img_index += 1

    # إضافة الصور المتبقية (التي كانت عادية)
    for replacement in image_replacements:
        slide.shapes.add_picture(
            replacement["path"],
            replacement["pos"][0],
            replacement["pos"][1],
            width=replacement["size"][0],
            height=replacement["size"][1]
        )

# ====== الدالة الرئيسية للتطبيق ======
def process_pptx(template_pptx, images_zip):
    """
    قراءة ملف PowerPoint وقائمة صور من ملف مضغوط، واستبدال الصور في كل شريحة.
    """
    # إنشاء مجلد مؤقت
    temp_dir = tempfile.mkdtemp()

    try:
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
                    if img.lower().endswith(IMAGE_EXTENSIONS)
                ])
                if i < len(prs.slides):
                    replace_images_in_slide(prs.slides[i], image_files)

        # حفظ النتيجة
        output_path = os.path.join(temp_dir, "output.pptx")
        prs.save(output_path)

        # قراءة الملف الناتج وإرجاعه للتحميل
        with open(output_path, "rb") as f:
            pptx_bytes = f.read()

        return pptx_bytes

    finally:
        # تنظيف الملفات المؤقتة
        shutil.rmtree(temp_dir, ignore_errors=True)

# ====== واجهة Streamlit ======
st.title("📊 أداة استبدال الصور في PowerPoint")
st.markdown("---")
st.info("💡 **ملاحظة:** ستقوم الأداة الآن باستبدال **الصور العادية** و **العناصر النائبة للصور** في قالب PowerPoint الخاص بك.")

template_pptx = st.file_uploader("ارفع قالب PowerPoint (.pptx)", type=["pptx"])
images_zip = st.file_uploader("ارفع ملف الصور المضغوط (.zip)", type=["zip"])

if st.button("بدء المعالجة"):
    if not template_pptx or not images_zip:
        st.error("الرجاء رفع كل من ملف PowerPoint وملف الصور المضغوط.")
    else:
        try:
            with st.spinner("جاري المعالجة..."):
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
