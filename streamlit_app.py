import streamlit as st
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Inches
from io import BytesIO
import zipfile
import os

# ---------------------------------------------
# دالة للتحقق من كون الشكل في الشريحة عبارة عن صورة
# ---------------------------------------------
def is_picture(shape):
    return shape.shape_type == MSO_SHAPE_TYPE.PICTURE

# ---------------------------------------------
# دالة لاستبدال الصور في الشريحة مع الحفاظ على الأبعاد والمكان
# ---------------------------------------------
def replace_images_in_slide(slide, images):
    img_index = 0
    for shape in slide.shapes:
        if is_picture(shape) and img_index < len(images):
            try:
                # الاحتفاظ بالموقع والأبعاد الأصلية
                left, top, width, height = shape.left, shape.top, shape.width, shape.height
                
                # حذف الصورة القديمة
                sp = shape._element
                sp.getparent().remove(sp)
                
                # إدراج الصورة الجديدة بنفس الأبعاد والمكان
                slide.shapes.add_picture(images[img_index], left, top, width, height)
                img_index += 1
            except Exception as e:
                st.warning(f"تعذر استبدال الصورة: {e}")

# ---------------------------------------------
# الدالة الرئيسية لمعالجة ملف PowerPoint
# ---------------------------------------------
def process_pptx(pptx_template, zip_images):
    # تحميل القالب
    try:
        prs = Presentation(pptx_template)
    except Exception as e:
        st.error(f"تعذر قراءة ملف PowerPoint: {e}")
        return None

    # استخراج الصور من ملف zip
    temp_dir = "temp_images"
    os.makedirs(temp_dir, exist_ok=True)
    try:
        with zipfile.ZipFile(zip_images, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
    except Exception as e:
        st.error(f"تعذر استخراج الصور من ملف ZIP: {e}")
        return None

    # فرز المجلدات لضمان ترتيب الصور
    folders = sorted(
        [os.path.join(temp_dir, d) for d in os.listdir(temp_dir) if os.path.isdir(os.path.join(temp_dir, d))]
    )

    # التأكد من أن هناك مجلدات للصور
    if not folders:
        st.error("ملف ZIP لا يحتوي على مجلدات صور صالحة.")
        return None

    # نسخة الشريحة الأولى كمصدر للقالب
    first_slide_layout = prs.slides[0]
    
    for folder in folders:
        images = sorted(
            [os.path.join(folder, img) for img in os.listdir(folder) if img.lower().endswith(('.png', '.jpg', '.jpeg'))]
        )

        if not images:
            st.warning(f"لا توجد صور في المجلد: {folder}")
            continue

        # إذا كانت هذه أول مجلد، استبدل الصور في الشريحة الأولى
        if folder == folders[0]:
            replace_images_in_slide(prs.slides[0], images)
        else:
            # نسخ الشريحة الأولى
            slide_clone = prs.slides.add_slide(first_slide_layout.slide_layout)
            # نسخ محتوى الشريحة الأصلية إلى الجديدة
            for shape in first_slide_layout.shapes:
                slide_clone.shapes._spTree.insert_element_before(shape.element.clone(), 'p:extLst')
            replace_images_in_slide(slide_clone, images)

    # حفظ النتيجة في ملف مؤقت
    output_pptx = BytesIO()
    prs.save(output_pptx)
    output_pptx.seek(0)

    # تنظيف الملفات المؤقتة
    for folder in folders:
        for f in os.listdir(folder):
            os.remove(os.path.join(folder, f))
        os.rmdir(folder)
    os.rmdir(temp_dir)

    return output_pptx

# ---------------------------------------------
# واجهة المستخدم بـ Streamlit
# ---------------------------------------------
st.set_page_config(page_title="استبدال الصور في PowerPoint", page_icon="📊", layout="centered")

st.title("📊 أداة استبدال الصور في PowerPoint")
st.write("قم بتحميل قالب PowerPoint وملف ZIP يحتوي على مجلدات صور، وسيتم إنشاء ملف PowerPoint جديد بنفس القالب.")

pptx_file = st.file_uploader("📂 اختر ملف PowerPoint (.pptx)", type=["pptx"])
zip_file = st.file_uploader("📂 اختر ملف ZIP للصور", type=["zip"])

if st.button("🔄 تنفيذ العملية"):
    if pptx_file and zip_file:
        result = process_pptx(pptx_file, zip_file)
        if result:
            st.success("✅ تم إنشاء الملف بنجاح!")
            st.download_button("⬇️ تحميل الملف الناتج", result, file_name="output.pptx")
    else:
        st.error("الرجاء تحميل الملفات المطلوبة.")
