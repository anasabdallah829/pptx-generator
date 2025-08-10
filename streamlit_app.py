import streamlit as st 
import zipfile 
import os 
import io 
from pptx import Presentation 
from pptx.enum.shapes import PP_PLACEHOLDER 
import shutil 
from copy import deepcopy

st.set_page_config(page_title="PowerPoint Image Replacer", layout="centered") 
st.title("🔄 PowerPoint Image & Placeholder Replacer") 
st.markdown("---") 

# رفع الملفات
uploaded_pptx = st.file_uploader("📂 اختر ملف PowerPoint (.pptx)", type=["pptx"]) 
uploaded_zip = st.file_uploader("🗜️ اختر ملف ZIP يحتوي على مجلدات صور", type=["zip"]) 
show_details = st.checkbox("عرض التفاصيل المفصلة", value=False) 

def analyze_first_slide(prs):
    if len(prs.slides) == 0:
        return False, "لا توجد شرائح"
    first_slide = prs.slides[0]
    picture_placeholders = [
        s for s in first_slide.shapes 
        if s.is_placeholder and s.placeholder_format.type == PP_PLACEHOLDER.PICTURE
    ]
    regular_pictures = [
        s for s in first_slide.shapes
        if hasattr(s, 'shape_type') and s.shape_type == 13
    ]
    total = len(picture_placeholders) + len(regular_pictures)
    return (total > 0, {
        'placeholders': len(picture_placeholders),
        'regular_pictures': len(regular_pictures),
        'total_slots': total
    }) if total > 0 else (False, "لا توجد صور أو أماكن صور")

def get_template_info(slide):
    info = {'image_positions': []}
    for shape in slide.shapes:
        if shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.PICTURE:
            info['image_positions'].append({
                'type': 'placeholder',
                'left': shape.left, 'top': shape.top,
                'width': shape.width, 'height': shape.height
            })
        elif hasattr(shape, 'shape_type') and shape.shape_type == 13:
            info['image_positions'].append({
                'type': 'picture',
                'left': shape.left, 'top': shape.top,
                'width': shape.width, 'height': shape.height
            })
    info['image_positions'].sort(key=lambda x: (x['top'], x['left']))
    return info

def duplicate_slide(prs, index=0):
    slide_id = prs.slides._sldIdLst[index]
    new_slide_id = deepcopy(slide_id)
    prs.slides._sldIdLst.insert(len(prs.slides._sldIdLst), new_slide_id)
    return prs.slides[-1]

def replace_images_in_slide(slide, template_info, images_folder, show_details=False):
    images = [f for f in os.listdir(images_folder) if f.lower().endswith(('png','jpg','jpeg','gif','bmp','tiff','webp'))]
    images.sort()
    shapes_list = []
    for shape in slide.shapes:
        if shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.PICTURE:
            shapes_list.append({'shape': shape, 'type': 'placeholder'})
        elif hasattr(shape, 'shape_type') and shape.shape_type == 13:
            shapes_list.append({'shape': shape, 'type': 'picture'})
    shapes_list.sort(key=lambda x: (x['shape'].top, x['shape'].left))
    replaced = 0
    for i, s_info in enumerate(shapes_list):
        if i >= len(images): break
        img_path = os.path.join(images_folder, images[i])
        try:
            if s_info['type'] == 'placeholder':
                with open(img_path, "rb") as img_file:
                    s_info['shape'].insert_picture(img_file)
            else:
                left, top, width, height = s_info['shape'].left, s_info['shape'].top, s_info['shape'].width, s_info['shape'].height
                slide.shapes._spTree.remove(s_info['shape']._element)
                with open(img_path, "rb") as img_file:
                    slide.shapes.add_picture(img_file, left, top, width, height)
            replaced += 1
            if show_details:
                st.success(f"✅ استبدال صورة: {images[i]}")
        except Exception as e:
            if show_details:
                st.warning(f"⚠ خطأ: {e}")
    return replaced

if uploaded_pptx and uploaded_zip:
    if st.button("🚀 بدء المعالجة"):
        temp_dir = None
        try:
            # استخراج الصور من ZIP
            zip_bytes = io.BytesIO(uploaded_zip.read())
            with zipfile.ZipFile(zip_bytes, "r") as zip_ref:
                temp_dir = "temp_images"
                if os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)
                os.makedirs(temp_dir)
                zip_ref.extractall(temp_dir)

            folder_paths = [os.path.join(temp_dir, f) for f in os.listdir(temp_dir) if os.path.isdir(os.path.join(temp_dir, f))]
            folder_paths.sort()

            if not folder_paths:
                st.error("❌ لا توجد مجلدات صور")
                st.stop()

            prs = Presentation(io.BytesIO(uploaded_pptx.read()))
            has_images, analysis = analyze_first_slide(prs)
            if not has_images:
                st.error(f"❌ {analysis}")
                st.stop()

            template_info = get_template_info(prs.slides[0])
            total_replaced = 0

            progress_bar = st.progress(0)
            for idx, folder_path in enumerate(folder_paths):
                new_slide = duplicate_slide(prs, 0)  # نسخ الشريحة الأولى بالكامل
                replaced = replace_images_in_slide(new_slide, template_info, folder_path, show_details)
                total_replaced += replaced
                progress_bar.progress((idx + 1) / len(folder_paths))

            progress_bar.empty()
            st.success(f"✅ تم استبدال {total_replaced} صورة")

            output_buffer = io.BytesIO()
            prs.save(output_buffer)
            output_buffer.seek(0)

            st.download_button(
                label="⬇️ تحميل الملف المُحدث",
                data=output_buffer.getvalue(),
                file_name="Updated.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )

        finally:
            if temp_dir and os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
else:
    st.info("📋 ارفع ملف PowerPoint وملف ZIP للبدء")
