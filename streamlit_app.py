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
            # تم تصحيح الخطأ هنا: تم استخدام MSO_SHAPE.PICTURE
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
                folder_path
