import streamlit as st
from pptx import Presentation
from pptx.util import Inches
from pptx.enum.shapes import MSO_SHAPE, PP_PLACEHOLDER
import zipfile
import os
import tempfile
import shutil

# ====== دالة لاستبدال الصور في الشريحة ======
def replace_images_in_slide(slide, image_paths):
    """
    استبدال الصور والعناصر النائبة للصور في شريحة PowerPoint بقائمة من الصور الجديدة.
    يتم التعامل مع أي صور موجودة في الشريحة واستبدالها حسب الترتيب.
    """
    img_index = 0
    # قائمة بخصائص الصور (موضع وحجم) ليتم إضافتها لاحقًا
    image_replacements = []

    for shape in slide.shapes:
        # التحقق إذا كان الشكل عبارة عن صورة عادية أو عنصر نائب للصورة
        if shape.shape_type == MSO_SHAPE.PICTURE:
            if img_index < len(image_paths):
                # تخزين خصائص الصورة القديمة
                x, y, cx, cy = shape.left, shape.top, shape.width, shape.height
                image_replacements.append({
                    "path": image_paths[img_index],
                    "pos": (x, y),
                    "size": (cx, cy)
                })
                img_index += 1
                # حذف الصورة القديمة
                slide.shapes._spTree.remove(shape._element)
        elif shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.PICTURE:
            if img_index < len(image_paths):
                # استخدام طريقة replace_picture للتعامل مع العناصر النائبة
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
