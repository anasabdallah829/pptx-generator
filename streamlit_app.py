import streamlit as st 
import zipfile 
import os 
import io 
from pptx import Presentation 
from pptx.enum.shapes import PP_PLACEHOLDER 
from pptx.util import Inches 
import shutil 
 
st.set_page_config(page_title="PowerPoint Image Replacer", layout="centered") 
st.title("🔄 PowerPoint Image & Placeholder Replacer") 
st.markdown("---") 
 
# رفع الملفات 
uploaded_pptx = st.file_uploader("📂 اختر ملف PowerPoint (.pptx)", type=["pptx"]) 
uploaded_zip = st.file_uploader("🗜️ اختر ملف ZIP يحتوي على مجلدات صور", type=["zip"]) 
 
# خيار عرض التفاصيل 
show_details = st.checkbox("عرض التفاصيل المفصلة", value=False) 
 
def analyze_first_slide(prs): 
    """تحليل الشريحة الأولى لتحديد وجود صور أو placeholders""" 
    if len(prs.slides) == 0: 
        return False, "لا توجد شرائح في الملف" 
     
    first_slide = prs.slides[0] 
     
    # البحث عن placeholders للصور 
    picture_placeholders = [ 
        shape for shape in first_slide.shapes  
        if shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.PICTURE 
    ] 
     
    # البحث عن الصور العادية 
    regular_pictures = [ 
        shape for shape in first_slide.shapes  
        if hasattr(shape, 'shape_type') and shape.shape_type == 13  # 13 = Picture 
    ] 
     
    total_image_slots = len(picture_placeholders) + len(regular_pictures) 
     
    if total_image_slots > 0: 
        return True, { 
            'placeholders': len(picture_placeholders), 
            'regular_pictures': len(regular_pictures), 
            'total_slots': total_image_slots, 
            'slide_layout': first_slide.slide_layout 
        } 
    else: 
        return False, "لا توجد صور أو placeholders للصور في الشريحة الأولى" 
 
def get_image_positions(slide): 
    """استخراج مواقع وأحجام الصور من الشريحة""" 
    positions = [] 
     
    for shape in slide.shapes: 
        if shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.PICTURE: 
            positions.append({ 
                'shape': shape, 
                'type': 'placeholder', 
                'left': shape.left, 
                'top': shape.top, 
                'width': shape.width, 
                'height': shape.height, 
                'placeholder_type': shape.placeholder_format.type 
            }) 
        elif hasattr(shape, 'shape_type') and shape.shape_type == 13:  # Picture 
            positions.append({ 
                'shape': shape, 
                'type': 'picture', 
                'left': shape.left, 
                'top': shape.top, 
                'width': shape.width, 
                'height': shape.height 
            }) 
     
    # ترتيب حسب الموقع (من الأعلى للأسفل، من اليسار لليمين) 
    positions.sort(key=lambda x: (x['top'], x['left'])) 
    return positions 
 
def replace_images_in_slide(slide, images_folder, folder_name, image_positions, show_details=False): 
    """استبدال الصور في الشريحة مع الحفاظ على المواقع والأحجام""" 
     
    # الحصول على قائمة الصور 
    if not os.path.exists(images_folder): 
        return 0, f"المجلد {images_folder} غير موجود" 
     
    images = [f for f in os.listdir(images_folder)  
              if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.webp'))] 
     
    if not images: 
        return 0, f"لا توجد صور في المجلد {folder_name}" 
     
    # ترتيب الصور أبجدياً 
    images.sort() 
     
    replaced_count = 0 
     
    # استبدال عنوان الشريحة 
    try: 
        title_shapes = [shape for shape in slide.shapes  
                       if shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.TITLE] 
        if title_shapes: 
            title_shapes[0].text = folder_name 
            if show_details: 
                st.success(f"✅ تم تعيين العنوان: {folder_name}") 
        else: 
            # إضافة عنوان جديد إذا لم يوجد 
            textbox = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1)) 
            text_frame = textbox.text_frame 
            text_frame.text = folder_name 
            paragraph = text_frame.paragraphs[0] 
            paragraph.font.size = Inches(0.4) 
            paragraph.font.bold = True 
            if show_details: 
                st.success(f"✅ تم إضافة العنوان: {folder_name}") 
    except Exception as e: 
        if show_details: 
            st.warning(f"⚠ خطأ في تعيين العنوان: {e}") 
     
    # استبدال الصور 
    for i, pos_info in enumerate(image_positions): 
        if i >= len(images): 
            break 
             
        try: 
            image_path = os.path.join(images_folder, images[i]) 
             
            if pos_info['type'] == 'placeholder': 
                # استبدال placeholder 
                with open(image_path, "rb") as img_file: 
                    pos_info['shape'].insert_picture(img_file) 
                replaced_count += 1 
                if show_details: 
                    st.success(f"✅ تم استبدال placeholder بالصورة: {images[i]}") 
                     
            elif pos_info['type'] == 'picture': 
                # استبدال الصورة العادية 
                shape = pos_info['shape'] 
                left, top, width, height = pos_info['left'], pos_info['top'], pos_info['width'], pos_info['height'] 
                 
                # حذف الصورة القديمة 
                slide.shapes._spTree.remove(shape._element) 
                 
                # إضافة الصورة الجديدة بنفس المواقع والأحجام 
                with open(image_path, "rb") as img_file: 
                    slide.shapes.add_picture(img_file, left, top, width, height) 
                 
                replaced_count += 1 
                if show_details: 
                    st.success(f"✅ تم استبدال الصورة العادية: {images[i]}") 
