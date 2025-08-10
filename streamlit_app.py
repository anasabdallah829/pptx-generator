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
    if len(prs.slides) == 0: 
        return False, "لا توجد شرائح في الملف"
    first_slide = prs.slides[0]
    picture_placeholders = [
        shape for shape in first_slide.shapes  
        if shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.PICTURE
    ]
    regular_pictures = [
        shape for shape in first_slide.shapes  
        if hasattr(shape, 'shape_type') and shape.shape_type == 13
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
 
def get_template_info(slide): 
    template_info = { 
        'title_info': None, 
        'image_positions': [] 
    }
    for shape in slide.shapes:
        if shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.TITLE:
            template_info['title_info'] = {
                'left': shape.left,
                'top': shape.top,
                'width': shape.width,
                'height': shape.height,
                'text': shape.text
            }
        elif shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.PICTURE:
            template_info['image_positions'].append({
                'type': 'placeholder',
                'left': shape.left,
                'top': shape.top,
                'width': shape.width,
                'height': shape.height
            })
        elif hasattr(shape, 'shape_type') and shape.shape_type == 13:
            template_info['image_positions'].append({
                'type': 'picture',
                'left': shape.left,
                'top': shape.top,
                'width': shape.width,
                'height': shape.height
            })
    template_info['image_positions'].sort(key=lambda x: (x['top'], x['left']))
    return template_info
 
def create_slide_with_images(prs, slide_layout, template_info, images_folder, folder_name, show_details=False): 
    if not os.path.exists(images_folder):
        return 0, f"المجلد {images_folder} غير موجود"
    images = [f for f in os.listdir(images_folder)  
              if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.webp'))]
    if not images:
        return 0, f"لا توجد صور في المجلد {folder_name}"
    images.sort()
    try:
        # إنشاء شريحة جديدة بنفس الـ layout للحفاظ على التنسيق
        new_slide = prs.slides.add_slide(slide_layout)
        replaced_count = 0

        # إضافة أو تعديل العنوان
        try:
            title_shapes = [shape for shape in new_slide.shapes  
                           if shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.TITLE]
            if title_shapes:
                title_shapes[0].text = folder_name
                if show_details:
                    st.success(f"✅ تم تعيين العنوان: {folder_name}")
            elif template_info['title_info']:
                title_info = template_info['title_info']
                textbox = new_slide.shapes.add_textbox(
                    title_info['left'], title_info['top'],  
                    title_info['width'], title_info['height']
                )
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

        current_slide_shapes = list(new_slide.shapes)
        picture_placeholders = [
            shape for shape in current_slide_shapes
            if shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.PICTURE
        ]
        regular_pictures = [
            shape for shape in current_slide_shapes
            if hasattr(shape, 'shape_type') and shape.shape_type == 13
        ]
        all_image_shapes = []
        for shape in picture_placeholders:
            all_image_shapes.append({
                'shape': shape,
                'type': 'placeholder',
                'left': shape.left,
                'top': shape.top
            })
        for shape in regular_pictures:
            all_image_shapes.append({
                'shape': shape,
                'type': 'picture',
                'left': shape.left,
                'top': shape.top
            })
        all_image_shapes.sort(key=lambda x: (x['top'], x['left']))

        for i, shape_info in enumerate(all_image_shapes):
            if i >= len(images):
                break
            try:
                image_path = os.path.join(images_folder, images[i])
                shape = shape_info['shape']
                if shape_info['type'] == 'placeholder':
                    with open(image_path, "rb") as img_file:
                        shape.insert_picture(img_file)
                    replaced_count += 1
                elif shape_info['type'] == 'picture':
                    left, top, width, height = shape.left, shape.top, shape.width, shape.height
                    new_slide.shapes._spTree.remove(shape._element)
                    with open(image_path, "rb") as img_file:
                        new_slide.shapes.add_picture(img_file, left, top, width, height)
                    replaced_count += 1
            except Exception as e:
                if show_details:
                    st.warning(f"⚠ خطأ في استبدال الصورة {images[i]}: {e}")

        if len(all_image_shapes) == 0 and len(template_info['image_positions']) > 0:
            for i, pos_info in enumerate(template_info['image_positions']):
                if i >= len(images):
                    break
                try:
                    image_path = os.path.join(images_folder, images[i])
                    with open(image_path, "rb") as img_file:
                        new_slide.shapes.add_picture(
                            img_file,  
                            pos_info['left'], pos_info['top'],  
                            pos_info['width'], pos_info['height']
                        )
                    replaced_count += 1
                except Exception as e:
                    if show_details:
                        st.warning(f"⚠ خطأ في إضافة الصورة {images[i]}: {e}")
        return replaced_count, "تم بنجاح"
    except Exception as e:
        return 0, f"خطأ في إنشاء الشريحة: {e}"
 
if uploaded_pptx and uploaded_zip: 
    if st.button("🚀 بدء المعالجة"): 
        temp_dir = None 
        try: 
            st.info("📦 جاري فحص الملفات...")
            zip_bytes = io.BytesIO(uploaded_zip.read())
            with zipfile.ZipFile(zip_bytes, "r") as zip_ref:
                temp_dir = "temp_images"
                if os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)
                os.makedirs(temp_dir)
                zip_ref.extractall(temp_dir)
            all_items = os.listdir(temp_dir)
            folder_paths = []
            for item in all_items:
                item_path = os.path.join(temp_dir, item)
                if os.path.isdir(item_path):
                    images_in_folder = [f for f in os.listdir(item_path)  
                                      if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.webp'))]
                    if images_in_folder:
                        folder_paths.append(item_path)
            if not folder_paths:
                st.error("❌ لا توجد مجلدات تحتوي على صور في الملف المضغوط.")
                st.stop()
            folder_paths.sort()
            st.success(f"✅ تم العثور على {len(folder_paths)} مجلد يحتوي على صور")
            prs = Presentation(io.BytesIO(uploaded_pptx.read()))
            st.info("🔍 جاري تحليل الشريحة الأولى...")
            has_images, analysis_result = analyze_first_slide(prs)
            if not has_images:
                st.error("❌ الشريحة الأولى لا تحتوي على صور أو placeholders!")
                st.stop()
            st.success("✅ تم العثور على صور أو placeholders في الشريحة الأولى!")
            first_slide = prs.slides[0]
            template_info = get_template_info(first_slide)
            slide_layout = analysis_result['slide_layout']

            total_replaced = 0
            created_slides = 0
            progress_bar = st.progress(0)
            status_text = st.empty()

            for folder_idx, folder_path in enumerate(folder_paths):
                folder_name = os.path.basename(folder_path)
                status_text.text(f"🔄 معالجة المجلد {folder_idx + 1}/{len(folder_paths)}: {folder_name}")
                replaced_count, message = create_slide_with_images(
                    prs, slide_layout, template_info, folder_path, folder_name, show_details
                )
                if "تم بنجاح" in message:
                    created_slides += 1
                    total_replaced += replaced_count
                progress = (folder_idx + 1) / len(folder_paths)
                progress_bar.progress(progress)

            progress_bar.empty()
            status_text.empty()

            st.success("🎉 تم الانتهاء من المعالجة!")
            output_filename = f"{os.path.splitext(uploaded_pptx.name)[0]}_Updated.pptx"
            output_buffer = io.BytesIO()
            prs.save(output_buffer)
            output_buffer.seek(0)
            st.download_button(
                label="⬇️ تحميل الملف المُحدث",
                data=output_buffer.getvalue(),
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                key="download_button"
            )
        finally:
            if temp_dir and os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
else:
    st.info("📋 يُرجى رفع ملف PowerPoint وملف ZIP للبدء")
