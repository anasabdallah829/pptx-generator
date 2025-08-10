import streamlit as st
import os
import zipfile
import io
import tempfile
from pptx import Presentation
from pptx.util import Inches
from pptx.enum.shapes import MSO_SHAPE

# امتدادات ملفات الصور المقبولة
IMAGE_EXTENSIONS = ('.png', '.jpg', '.jpeg', '.gif', '.bmp')

def find_first_image_in_folder(folder_path):
    """
    يبحث عن أول ملف صورة بامتداد مسموح به في المسار المحدد.

    :param folder_path: المسار إلى المجلد.
    :return: المسار الكامل لأول صورة تم العثور عليها، أو None إذا لم يتم العثور على أي صورة.
    """
    for item in os.listdir(folder_path):
        if item.lower().endswith(IMAGE_EXTENSIONS):
            return os.path.join(folder_path, item)
    return None

def process_files_with_images(zip_file, pptx_file):
    """
    يعالج ملفات ZIP و PPTX لإضافة شريحة لكل مجلد، مع تضمين أول صورة من كل مجلد.

    :param zip_file: كائن ملف Streamlit المرفوع لملف ZIP.
    :param pptx_file: كائن ملف Streamlit المرفوع لملف PPTX.
    :return: كائن BytesIO لملف PPTX المعدل، أو None في حالة الفشل.
    """
    try:
        # تحميل العرض التقديمي من كائن الملف المرفوع
        prs = Presentation(pptx_file)
        
        # إنشاء دليل مؤقت لاستخراج ملف ZIP
        with tempfile.TemporaryDirectory() as temp_dir:
            # كتابة محتوى ملف ZIP المرفوع إلى ملف مؤقت
            with open(os.path.join(temp_dir, 'uploaded.zip'), 'wb') as f:
                f.write(zip_file.getbuffer())
            
            # استخراج محتويات ZIP إلى المجلد المؤقت
            with zipfile.ZipFile(os.path.join(temp_dir, 'uploaded.zip'), 'r') as zip_ref:
                zip_ref.extractall(temp_dir)
            
            # الحصول على قائمة المجلدات من المجلد المؤقت
            folders = [d for d in os.listdir(temp_dir) if os.path.isdir(os.path.join(temp_dir, d)) and not d.startswith('.')]
            
            # المرور على كل مجلد وإضافة شريحة جديدة
            for folder_name in sorted(folders):
                folder_path = os.path.join(temp_dir, folder_name)
                st.info(f"جاري معالجة المجلد: **{folder_name}**")
                
                # إضافة شريحة جديدة بعنوان
                slide_layout = prs.slide_layouts[5]  # تخطيط "Title Only"
                slide = prs.slides.add_slide(slide_layout)
                title = slide.shapes.title
                title.text = f"مجلد: {folder_name}"
                
                # البحث عن أول صورة في المجلد
                image_path = find_first_image_in_folder(folder_path)
                
                if image_path:
                    # إضافة الصورة إلى الشريحة
                    try:
                        # تحديد موضع الصورة وحجمها (يمكن تعديلها حسب الحاجة)
                        left = top = Inches(1.5)
                        width = Inches(7)
                        slide.shapes.add_picture(image_path, left, top, width=width)
                        st.success(f"تمت إضافة صورة إلى شريحة المجلد **{folder_name}**.")
                    except Exception as img_e:
                        st.warning(f"تعذر إضافة الصورة من المجلد **{folder_name}**: {img_e}")
                else:
                    st.warning(f"لم يتم العثور على أي صور في المجلد **{folder_name}**.")

            # حفظ العرض التقديمي المعدل في الذاكرة
            output_stream = io.BytesIO()
            prs.save(output_stream)
            output_stream.seek(0)
            
            return output_stream
        
    except Exception as e:
        st.error(f"حدث خطأ غير متوقع: {e}")
        return None

# --- واجهة المستخدم لتطبيق Streamlit ---
st.set_page_config(page_title="أداة معالجة مجلدات PowerPoint", layout="centered")
st.title("أداة معالجة مجلدات PowerPoint 📁🖼️")
st.markdown("---")

st.write(
    "يرجى رفع **ملف مضغوط (.zip)** يحتوي على مجلدات وصور، بالإضافة إلى **ملف PowerPoint (.pptx)**. "
    "ستقوم الأداة بإنشاء شريحة جديدة لكل مجلد في الملف المضغوط، وتضيف أول صورة تجدها داخل كل مجلد."
)

# عناصر رفع الملفات
zip_file_upload = st.file_uploader("1. قم برفع ملف ZIP:", type=["zip"])
pptx_file_upload = st.file_uploader("2. قم برفع ملف PowerPoint (.pptx):", type=["pptx"])

# زر المعالجة
if st.button("معالجة وإنشاء العرض التقديمي"):
    if zip_file_upload is not None and pptx_file_upload is not None:
        with st.spinner("جاري معالجة الملفات وإنشاء العرض التقديمي الجديد... 🔄"):
            modified_pptx_stream = process_files_with_images(zip_file_upload, pptx_file_upload)
            
            if modified_pptx_stream:
                st.success("اكتملت المعالجة بنجاح! ملفك جاهز للتنزيل. 🎉")
                # زر تنزيل الملف المعدل
                st.download_button(
                    label="تنزيل ملف PPTX المعدل",
                    data=modified_pptx_stream,
                    file_name="modified_presentation_with_images.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
            else:
                st.error("فشلت المعالجة. يرجى التأكد من تنسيقات الملفات ومحتوياتها والمحاولة مرة أخرى.")
    else:
        st.warning("يجب رفع كل من ملف ZIP وملف PPTX للمتابعة.")
