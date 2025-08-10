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

def find_images_in_folder(folder_path):
    """
    يبحث عن جميع ملفات الصور بامتدادات مسموح بها في المسار المحدد.

    :param folder_path: المسار إلى المجلد.
    :return: قائمة بجميع المسارات الكاملة لملفات الصور التي تم العثور عليها، مرتبة أبجدياً.
    """
    images = []
    for item in os.listdir(folder_path):
        if item.lower().endswith(IMAGE_EXTENSIONS):
            images.append(os.path.join(folder_path, item))
    return sorted(images)

def process_files_with_template(zip_file, pptx_file):
    """
    يعالج ملفات ZIP و PPTX باستخدام الشريحة الأولى كقالب.

    - يستخرج خصائص الصور والعنوان من الشريحة الأولى.
    - ينشئ شريحة جديدة لكل مجلد، ويستخدم الخصائص المخزنة لوضع الصور من المجلد.
    
    :param zip_file: كائن ملف Streamlit المرفوع لملف ZIP.
    :param pptx_file: كائن ملف Streamlit المرفوع لملف PPTX.
    :return: كائن BytesIO لملف PPTX المعدل، أو None في حالة الفشل.
    """
    try:
        # تحميل العرض التقديمي من كائن الملف المرفوع
        prs = Presentation(pptx_file)
        
        # التأكد من وجود شرائح في العرض التقديمي
        if not prs.slides:
            raise ValueError("ملف PowerPoint لا يحتوي على أي شرائح.")

        # --- الخطوة 1: استخلاص خصائص الصور والعنوان من الشريحة الأولى ---
        template_slide = prs.slides[0]
        image_properties = []
        title_shape_properties = None

        for shape in template_slide.shapes:
            if shape.has_text_frame and shape.is_placeholder and shape.placeholder_format.type == 1: # MSO_SHAPE.PLACEHOLDER
                 title_shape_properties = {
                    "left": shape.left,
                    "top": shape.top,
                    "width": shape.width,
                    "height": shape.height,
                    "text_frame": shape.text_frame,
                    "font_size": shape.text_frame.paragraphs[0].font.size
                 }
            elif shape.shape_type == MSO_SHAPE.PICTURE:
                image_properties.append({
                    "left": shape.left,
                    "top": shape.top,
                    "width": shape.width,
                    "height": shape.height
                })

        st.info(f"تم العثور على {len(image_properties)} صورة في الشريحة الأولى كقالب.")
        
        # --- الخطوة 2: معالجة الملف المضغوط وإنشاء الشرائح ---
        with tempfile.TemporaryDirectory() as temp_dir:
            zip_content = io.BytesIO(zip_file.read())
            with zipfile.ZipFile(zip_content, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)
            
            folders = [d for d in os.listdir(temp_dir) if os.path.isdir(os.path.join(temp_dir, d)) and not d.startswith('.')]
            
            # حذف الشريحة الأولى (القالب) بعد استخلاص خصائصها
            prs.slides._sldIdLst.remove(prs.slides._sldIdLst[0])

            # المرور على كل مجلد وإنشاء شريحة جديدة
            for folder_name in sorted(folders):
                folder_path = os.path.join(temp_dir, folder_name)
                st.info(f"جاري معالجة المجلد: **{folder_name}**")
                
                # إنشاء شريحة جديدة
                slide = prs.slides.add_slide(prs.slide_layouts[6]) # استخدام تخطيط فارغ تمامًا
                
                # إضافة العنوان بناءً على خصائص الشريحة الأولى
                if title_shape_properties:
                    title_shape = slide.shapes.add_textbox(
                        title_shape_properties["left"],
                        title_shape_properties["top"],
                        title_shape_properties["width"],
                        title_shape_properties["height"]
                    )
                    title_shape.text = f"مجلد: {folder_name}"
                    
                # البحث عن الصور في المجلد
                folder_images = find_images_in_folder(folder_path)
                
                # إضافة الصور مع الحفاظ على أماكن وأحجام القالب
                num_images_to_add = min(len(image_properties), len(folder_images))
                if num_images_to_add > 0:
                    for i in range(num_images_to_add):
                        img_prop = image_properties[i]
                        img_path = folder_images[i]
                        slide.shapes.add_picture(
                            img_path,
                            img_prop["left"],
                            img_prop["top"],
                            width=img_prop["width"],
                            height=img_prop["height"]
                        )
                    st.success(f"تمت إضافة {num_images_to_add} صورة إلى شريحة المجلد **{folder_name}**.")
                else:
                    st.warning(f"لم يتم العثور على صور في المجلد **{folder_name}**، أو أن القالب لا يحتوي على صور.")

            # حفظ العرض التقديمي المعدل في الذاكرة
            output_stream = io.BytesIO()
            prs.save(output_stream)
            output_stream.seek(0)
            
            return output_stream
        
    except ValueError as ve:
        st.error(f"خطأ في الملف: {ve}")
        return None
    except Exception as e:
        st.error(f"حدث خطأ غير متوقع: {e}")
        return None

# --- واجهة المستخدم لتطبيق Streamlit ---
st.set_page_config(page_title="أداة معالجة مجلدات PowerPoint", layout="centered")
st.title("أداة معالجة مجلدات PowerPoint 📁🖼️")
st.markdown("---")

st.write(
    "هذه الأداة تستخدم **الشريحة الأولى** من ملف PowerPoint كقالب. سيتم إنشاء شريحة جديدة "
    "لكل مجلد في الملف المضغوط، مع استبدال العنوان والصور بناءً على محتوى المجلد، مع الحفاظ على أماكنها وأحجامها الأصلية."
)

# عناصر رفع الملفات
zip_file_upload = st.file_uploader("1. قم برفع ملف ZIP:", type=["zip"])
pptx_file_upload = st.file_uploader("2. قم برفع ملف PowerPoint (.pptx):", type=["pptx"])

# زر المعالجة
if st.button("معالجة وإنشاء العرض التقديمي"):
    if zip_file_upload is not None and pptx_file_upload is not None:
        with st.spinner("جاري معالجة الملفات وإنشاء العرض التقديمي الجديد... 🔄"):
            modified_pptx_stream = process_files_with_template(zip_file_upload, pptx_file_upload)
            
            if modified_pptx_stream:
                st.success("اكتملت المعالجة بنجاح! ملفك جاهز للتنزيل. 🎉")
                # زر تنزيل الملف المعدل
                st.download_button(
                    label="تنزيل ملف PPTX المعدل",
                    data=modified_pptx_stream,
                    file_name="modified_presentation_with_template.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
            else:
                st.error("فشلت المعالجة. يرجى التأكد من تنسيقات الملفات ومحتوياتها والمحاولة مرة أخرى.")
    else:
        st.warning("يجب رفع كل من ملف ZIP وملف PPTX للمتابعة.")
