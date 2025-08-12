import streamlit as st
import zipfile
import os
import io
from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER, MSO_SHAPE_TYPE
import shutil
from pptx.util import Inches
import random

# إعداد صفحة Streamlit
st.set_page_config(page_title="PowerPoint Image Replacer", layout="centered")
st.title("🔄 PowerPoint Image & Placeholder Replacer")
st.markdown("---")

# واجهة المستخدم لرفع الملفات
uploaded_pptx = st.file_uploader("📂 اختر ملف PowerPoint (.pptx)", type=["pptx"], key="pptx_uploader")
uploaded_zip = st.file_uploader("🗜️ اختر ملف ZIP يحتوي على مجلدات صور", type=["zip"], key="zip_uploader")

# خيارات جديدة
st.markdown("### ⚙️ إعدادات المعالجة")
image_order_option = st.radio(
    "كيف تريد ترتيب الصور في الشرائح؟",
    ("بالترتيب (افتراضي)", "عشوائي"),
    index=0
)

# خيار عرض التفاصيل
show_details = st.checkbox("عرض التفاصيل المفصلة", value=False)


def analyze_first_slide(prs):
    """
    تحليل الشريحة الأولى: إرجاع نتائج حتى لو لم توجد مواضع للصور.
    """
    if len(prs.slides) == 0:
        return False, "لا توجد شرائح في الملف"

    first_slide = prs.slides[0]
    
    picture_placeholders = [
        shape for shape in first_slide.shapes
        if shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.PICTURE
    ]
    
    # استخدام نفس طريقة الكود المرجعي للصور العادية
    regular_pictures = [
        shape for shape in first_slide.shapes 
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE
    ]
    
    total_image_slots = len(picture_placeholders) + len(regular_pictures)

    return True, {
        'placeholders': len(picture_placeholders),
        'regular_pictures': len(regular_pictures),
        'total_slots': total_image_slots,
        'slide_layout': first_slide.slide_layout
    }


def get_image_shapes_info(slide):
    """
    استخراج معلومات مفصلة عن أشكال الصور من الشريحة
    مع استخدام نفس طريقة الكود المرجعي
    """
    image_shapes_info = []
    
    # البحث عن placeholders للصور
    for shape in slide.shapes:
        if shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.PICTURE:
            image_shapes_info.append({
                'shape': shape,
                'type': 'placeholder',
                'left': shape.left,
                'top': shape.top,
                'width': shape.width,
                'height': shape.height,
                'position': (shape.top, shape.left)
            })
    
    # البحث عن الصور العادية باستخدام نفس طريقة الكود المرجعي
    regular_pictures = [
        shape for shape in slide.shapes 
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE
    ]
    
    # إضافة معلومات الصور العادية مع استخدام نفس البيانات من الكود المرجعي
    for shape in regular_pictures:
        image_shapes_info.append({
            'shape': shape,
            'type': 'picture',
            'left': shape.left,
            'top': shape.top,
            'width': shape.width,
            'height': shape.height,
            'position': (shape.top, shape.left)
        })
    
    # ترتيب حسب الموقع (من الأعلى للأسفل، من اليسار لليمين)
    image_shapes_info.sort(key=lambda x: x['position'])
    return image_shapes_info


def get_template_image_positions(slide):
    """
    استخراج مواقع الصور من القالب بنفس طريقة الكود المرجعي
    """
    # استخدام نفس الطريقة من الكود المرجعي
    image_shapes = [shape for shape in slide.shapes if shape.shape_type == MSO_SHAPE_TYPE.PICTURE]
    image_positions = [(shape.left, shape.top, shape.height) for shape in image_shapes]
    
    # إضافة placeholders أيضاً
    for shape in slide.shapes:
        if shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.PICTURE:
            image_positions.append((shape.left, shape.top, shape.height))
    
    return image_positions


def replace_image_in_shape(slide, shape_info, image_path, show_details=False):
    """
    استبدال صورة في شكل محدد مع استخدام طريقة الكود المرجعي للصور العادية
    """
    try:
        shape = shape_info['shape']
        shape_type = shape_info['type']
        
        if shape_type == 'placeholder':
            # معالجة placeholders بالطريقة العادية
            try:
                with open(image_path, 'rb') as img_file:
                    shape.insert_picture(img_file)
                if show_details:
                    st.success(f"✅ تم استبدال placeholder بنجاح: {os.path.basename(image_path)}")
                return True
            except Exception as e:
                if show_details:
                    st.warning(f"⚠ فشل في استبدال placeholder، محاولة طريقة بديلة: {e}")
                
                # طريقة بديلة للـ placeholders
                try:
                    left, top, width, height = shape_info['left'], shape_info['top'], shape_info['width'], shape_info['height']
                    
                    # حذف الشكل القديم
                    shape_element = shape._element
                    shape_element.getparent().remove(shape_element)
                    
                    # إضافة صورة جديدة
                    slide.shapes.add_picture(image_path, left, top, width, height)
                    
                    if show_details:
                        st.success(f"✅ تم استبدال placeholder بالطريقة البديلة: {os.path.basename(image_path)}")
                    return True
                except Exception as e2:
                    if show_details:
                        st.error(f"❌ فشل في استبدال placeholder: {e2}")
                    return False
        
        elif shape_type == 'picture':
            # استخدام نفس طريقة الكود المرجعي للصور العادية
            try:
                left, top, height = shape_info['left'], shape_info['top'], shape_info['height']
                
                # حذف الصورة القديمة (نفس طريقة الكود المرجعي)
                shape_element = shape._element
                shape_element.getparent().remove(shape_element)
                
                # إضافة الصورة الجديدة بنفس طريقة الكود المرجعي
                # استخدام height فقط كما في الكود المرجعي
                slide.shapes.add_picture(image_path, left, top, height=height)
                
                if show_details:
                    st.success(f"✅ تم استبدال الصورة العادية بطريقة الكود المرجعي: {os.path.basename(image_path)}")
                return True
            except Exception as e:
                if show_details:
                    st.error(f"❌ فشل في استبدال الصورة العادية: {e}")
                return False
        
        return False
        
    except Exception as e:
        if show_details:
            st.error(f"❌ خطأ عام في استبدال الصورة: {e}")
        return False


def add_images_using_template_positions(slide, images, image_positions, show_details=False):
    """
    إضافة الصور باستخدام مواقع القالب (نفس طريقة الكود المرجعي)
    """
    added_count = 0
    
    for idx, (left, top, height) in enumerate(image_positions):
        if idx < len(images):
            try:
                slide.shapes.add_picture(images[idx], left, top, height=height)
                added_count += 1
                if show_details:
                    st.success(f"✅ تم إضافة صورة بطريقة القالب: {os.path.basename(images[idx])}")
            except Exception as e:
                if show_details:
                    st.error(f"❌ فشل في إضافة صورة: {e}")
    
    return added_count


def add_title_to_slide(slide, folder_name, show_details=False):
    """
    إضافة أو تحديث عنوان الشريحة
    """
    try:
        # البحث عن placeholder للعنوان
        title_shapes = [
            shape for shape in slide.shapes
            if shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.TITLE
        ]
        
        if title_shapes:
            # تحديث العنوان الموجود
            title_shapes[0].text = folder_name
            if show_details:
                st.success(f"✅ تم تحديث العنوان: {folder_name}")
        else:
            # إضافة عنوان جديد
            try:
                textbox = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
                text_frame = textbox.text_frame
                text_frame.text = folder_name
                
                # تنسيق النص
                paragraph = text_frame.paragraphs[0]
                paragraph.font.size = Inches(0.4)
                paragraph.font.bold = True
                
                if show_details:
                    st.success(f"✅ تم إضافة عنوان جديد: {folder_name}")
            except Exception as e:
                if show_details:
                    st.warning(f"⚠ فشل في إضافة العنوان: {e}")
    except Exception as e:
        if show_details:
            st.warning(f"⚠ خطأ في معالجة العنوان: {e}")


def process_folder_images(slide, folder_path, folder_name, template_shapes_info, template_positions, mismatch_action, show_details=False):
    """
    معالجة صور مجلد واحد وإضافتها للشريحة
    """
    # الحصول على قائمة الصور
    imgs = [f for f in os.listdir(folder_path) 
            if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.webp'))]
    
    if not imgs:
        if show_details:
            st.warning(f"⚠ المجلد {folder_name} فارغ من الصور")
        return 0
    
    # ترتيب الصور بناءً على اختيار المستخدم
    if image_order_option == "عشوائي":
        random.shuffle(imgs)
    else:
        imgs.sort()
    
    # تحويل أسماء الصور إلى مسارات كاملة
    image_paths = [os.path.join(folder_path, img) for img in imgs]
    
    # إضافة العنوان
    add_title_to_slide(slide, folder_name, show_details)
    
    # الحصول على معلومات أشكال الصور في الشريحة الجديدة
    new_shapes_info = get_image_shapes_info(slide)
    
    replaced_count = 0
    
    if new_shapes_info:
        # إذا وجدت أشكال صور في الشريحة الجديدة، استبدلها
        if show_details:
            st.info(f"📸 وجدت {len(new_shapes_info)} شكل صورة في الشريحة الجديدة")
        
        # معالجة اختلاف عدد الصور
        if mismatch_action == 'skip_folder' and len(imgs) != len(new_shapes_info):
            if show_details:
                st.info(f"ℹ تم تخطي المجلد {folder_name} لوجود اختلاف في عدد الصور")
            return 0
        
        # استبدال الصور
        for i, shape_info in enumerate(new_shapes_info):
            if mismatch_action == 'truncate' and i >= len(imgs):
                break
            
            # اختيار الصورة (مع التكرار إذا لزم الأمر)
            image_path = image_paths[i % len(image_paths)]
            
            # التحقق من وجود الملف
            if not os.path.exists(image_path):
                if show_details:
                    st.warning(f"⚠ الملف غير موجود: {image_path}")
                continue
            
            # استبدال الصورة
            success = replace_image_in_shape(slide, shape_info, image_path, show_details)
            if success:
                replaced_count += 1
    
    elif template_positions:
        # إذا لم توجد أشكال في الشريحة الجديدة، استخدم مواقع القالب
        if show_details:
            st.info(f"📍 استخدام مواقع القالب ({len(template_positions)} موقع)")
        
        replaced_count = add_images_using_template_positions(
            slide, image_paths, template_positions, show_details
        )
    
    else:
        # إضافة الصورة الأولى في موقع افتراضي
        if show_details:
            st.warning(f"⚠ لا توجد مواضع للصور، إضافة الصورة الأولى في موقع افتراضي")
        
        if image_paths:
            try:
                slide.shapes.add_picture(image_paths[0], Inches(1), Inches(2), Inches(8), Inches(5))
                if show_details:
                    st.success(f"✅ تم إضافة الصورة الأولى في موقع افتراضي: {imgs[0]}")
                replaced_count = 1
            except Exception as e:
                if show_details:
                    st.error(f"❌ فشل في إضافة الصورة الافتراضية: {e}")
    
    return replaced_count


def main():
    if uploaded_pptx and uploaded_zip:
        if "process_started" not in st.session_state:
            st.session_state.process_started = False

        if st.button("🚀 بدء المعالجة") or st.session_state.process_started:
            st.session_state.process_started = True
            
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
                        # التحقق من وجود صور في المجلد
                        imgs_in_folder = [f for f in os.listdir(item_path) 
                                        if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.webp'))]
                        if imgs_in_folder:
                            folder_paths.append(item_path)
                            if show_details:
                                st.info(f"📁 المجلد '{item}' يحتوي على {len(imgs_in_folder)} صورة")
                
                if not folder_paths:
                    st.error("❌ لا توجد مجلدات تحتوي على صور في الملف المضغوط.")
                    st.stop()
                
                folder_paths.sort()
                st.success(f"✅ تم العثور على {len(folder_paths)} مجلد يحتوي على صور")

                prs = Presentation(io.BytesIO(uploaded_pptx.read()))
                
                st.info("🔍 جاري تحليل الشريحة الأولى...")
                ok, analysis_result = analyze_first_slide(prs)
                if not ok:
                    st.error(f"❌ {analysis_result}")
                    st.stop()
                
                st.success("✅ تحليل الشريحة الأولى جاهز")
                col1, col2, col3 = st.columns(3)
                with col1: st.metric("Placeholders للصور", analysis_result['placeholders'])
                with col2: st.metric("الصور العادية", analysis_result['regular_pictures'])
                with col3: st.metric("إجمالي أماكن الصور", analysis_result['total_slots'])
                
                first_slide = prs.slides[0]
                template_shapes_info = get_image_shapes_info(first_slide)
                template_positions = get_template_image_positions(first_slide)
                
                if not template_shapes_info and not template_positions:
                    st.warning("⚠ الشريحة الأولى لا تحتوي على مواضع صور. سيتم إضافة الصورة الأولى من كل مجلد فقط.")
                    slide_layout = prs.slide_layouts[6]  # Blank layout
                else:
                    slide_layout = analysis_result['slide_layout']

                # فحص التطابق في عدد الصور
                expected_count = max(len(template_shapes_info), len(template_positions))
                mismatch_folders = []
                for fp in folder_paths:
                    imgs = [f for f in os.listdir(fp) 
                           if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.webp'))]
                    if len(imgs) != expected_count:
                        mismatch_folders.append((os.path.basename(fp), len(imgs), expected_count))
                
                if mismatch_folders and 'mismatch_action' not in st.session_state:
                    with st.form("mismatch_form"):
                        st.warning("⚠ تم اكتشاف اختلاف في عدد الصور لبعض المجلدات مقارنة بعدد مواضع الصور في الشريحة الأولى.")
                        for name, img_count, _ in mismatch_folders:
                            st.write(f"- المجلد `{name}` يحتوي على {img_count} صورة.")
                        st.markdown(f"**عدد مواضع الصور في القالب: {expected_count}**")

                        choice_text = st.radio(
                            "اختر كيف تريد التعامل مع المجلدات التي يختلف عدد صورها:",
                            ("استبدال فقط حتى أقل عدد (truncate)", "تكرار الصور لملء جميع المواضع (repeat)", "تخطي المجلدات ذات الاختلاف (skip_folder)", "إيقاف العملية (stop)"),
                            index=0
                        )
                        submit_choice = st.form_submit_button("✅ تأكيد الاختيار والمتابعة")

                    if submit_choice:
                        if choice_text.startswith("استبدال فقط"): st.session_state['mismatch_action'] = 'truncate'
                        elif choice_text.startswith("تكرار"): st.session_state['mismatch_action'] = 'repeat'
                        elif choice_text.startswith("تخطي"): st.session_state['mismatch_action'] = 'skip_folder'
                        else: st.session_state['mismatch_action'] = 'stop'
                    else:
                        st.stop()
                
                if 'mismatch_action' in st.session_state:
                    mismatch_action = st.session_state['mismatch_action']
                else:
                    mismatch_action = 'truncate'

                if mismatch_action == 'stop':
                    st.error("❌ تم إيقاف العملية بناءً على اختيار المستخدم.")
                    st.stop()

                st.info("🔄 جاري إضافة الشرائح الجديدة...")
                total_replaced = 0
                created_slides = 0

                progress_bar = st.progress(0)
                status_text = st.empty()

                for folder_idx, folder_path in enumerate(folder_paths):
                    folder_name = os.path.basename(folder_path)
                    status_text.text(f"🔄 معالجة المجلد {folder_idx + 1}/{len(folder_paths)}: {folder_name}")

                    try:
                        # إنشاء شريحة جديدة
                        new_slide = prs.slides.add_slide(slide_layout)
                        created_slides += 1
                        
                        # معالجة صور المجلد
                        replaced_count = process_folder_images(
                            new_slide, folder_path, folder_name, 
                            template_shapes_info, template_positions, mismatch_action, show_details
                        )
                        
                        total_replaced += replaced_count
                        
                        if show_details:
                            st.success(f"✅ تم إنشاء شريحة للمجلد '{folder_name}' واستبدال {replaced_count} صورة")
                    
                    except Exception as e:
                        st.error(f"❌ خطأ في معالجة المجلد {folder_name}: {e}")
                        if show_details:
                            import traceback
                            st.error(f"تفاصيل الخطأ: {traceback.format_exc()}")

                    progress_bar.progress((folder_idx + 1) / len(folder_paths))

                progress_bar.empty()
                status_text.empty()

                st.success("🎉 تم الانتهاء من المعالجة!")
                
                # تنظيف session state
                if 'mismatch_action' in st.session_state: 
                    del st.session_state['mismatch_action']
                if 'process_started' in st.session_state: 
                    del st.session_state['process_started']

                col1, col2, col3 = st.columns(3)
                with col1: st.metric("الشرائح المُضافة", created_slides)
                with col2: st.metric("الصور المُستبدلة", total_replaced)
                with col3: st.metric("المجلدات المُعالجة", len(folder_paths))

                if created_slides == 0:
                    st.error("❌ لم يتم إضافة أي شرائح.")
                    st.stop()

                # حفظ الملف
                original_name = os.path.splitext(uploaded_pptx.name)[0]
                output_filename = f"{original_name}_Updated.pptx"
                output_buffer = io.BytesIO()
                prs.save(output_buffer)
                output_buffer.seek(0)

                st.success(f"✅ تم إنشاء ملف PowerPoint جديد بـ {created_slides} شريحة!")

                st.download_button(
                    label="⬇️ تحميل الملف المُحدث",
                    data=output_buffer.getvalue(),
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    key="download_button"
                )

            except Exception as e:
                st.error(f"❌ خطأ أثناء المعالجة: {e}")
                if show_details:
                    import traceback
                    st.error(f"تفاصيل الخطأ: {traceback.format_exc()}")
            finally:
                if temp_dir and os.path.exists(temp_dir):
                    try:
                        shutil.rmtree(temp_dir)
                    except Exception as cleanup_error:
                        if show_details:
                            st.warning(f"⚠ خطأ في تنظيف الملفات المؤقتة: {cleanup_error}")
    else:
        st.info("📋 يُرجى رفع ملف PowerPoint وملف ZIP للبدء")

        with st.expander("📖 تعليمات الاستخدام"):
            st.markdown("""
            ### كيفية استخدام التطبيق:

            1.  **ملف PowerPoint (.pptx):**
                - يجب أن يحتوي على شريحة واحدة على الأقل.
                - يتم استخدام تنسيق الشريحة الأولى كقالب.

            2.  **ملف ZIP:**
                - يجب أن يحتوي على مجلدات، وكل مجلد يحتوي على صور.
                - أسماء المجلدات ستصبح عناوين الشرائح.

            3.  **النتيجة:**
                - شريحة منفصلة لكل مجلد.
                - يتم استبدال الصور و placeholders في القالب بصور من المجلدات.
                - في حال عدم وجود مواضع للصور في القالب، تُضاف الصورة الأولى من كل مجلد.

            ### أنواع الصور المدعومة:
            - PNG, JPG, JPEG, GIF, BMP, TIFF, WEBP
            """)
            
if __name__ == '__main__':
    main()
