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

def get_template_info(slide):
    """استخراج معلومات القالب من الشريحة"""
    template_info = {
        'title_info': None,
        'image_positions': []
    }
    
    # البحث عن العنوان
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
        elif hasattr(shape, 'shape_type') and shape.shape_type == 13:  # Picture
            template_info['image_positions'].append({
                'type': 'picture',
                'left': shape.left,
                'top': shape.top,
                'width': shape.width,
                'height': shape.height
            })
    
    # ترتيب مواقع الصور
    template_info['image_positions'].sort(key=lambda x: (x['top'], x['left']))
    return template_info

def create_slide_with_images(prs, slide_layout, template_info, images_folder, folder_name, show_details=False):
    """إنشاء شريحة جديدة مع الصور"""
    
    # الحصول على قائمة الصور
    if not os.path.exists(images_folder):
        return 0, f"المجلد {images_folder} غير موجود"
    
    images = [f for f in os.listdir(images_folder) 
              if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.webp'))]
    
    if not images:
        return 0, f"لا توجد صور في المجلد {folder_name}"
    
    # ترتيب الصور أبجدياً
    images.sort()
    
    try:
        # إنشاء شريحة جديدة
        new_slide = prs.slides.add_slide(slide_layout)
        replaced_count = 0
        
        # إضافة العنوان
        try:
            title_shapes = [shape for shape in new_slide.shapes 
                           if shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.TITLE]
            if title_shapes:
                title_shapes[0].text = folder_name
                if show_details:
                    st.success(f"✅ تم تعيين العنوان: {folder_name}")
            elif template_info['title_info']:
                # إضافة عنوان جديد بنفس موقع القالب
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
        
        # إضافة الصور
        current_slide_shapes = list(new_slide.shapes)
        
        # البحث عن placeholders للصور في الشريحة الجديدة
        picture_placeholders = [
            shape for shape in current_slide_shapes
            if shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.PICTURE
        ]
        
        # البحث عن الصور العادية في الشريحة الجديدة
        regular_pictures = [
            shape for shape in current_slide_shapes
            if hasattr(shape, 'shape_type') and shape.shape_type == 13
        ]
        
        # دمج جميع أماكن الصور وترتيبها
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
        
        # ترتيب الأشكال حسب الموقع
        all_image_shapes.sort(key=lambda x: (x['top'], x['left']))
        
        # استبدال الصور
        for i, shape_info in enumerate(all_image_shapes):
            if i >= len(images):
                break
                
            try:
                image_path = os.path.join(images_folder, images[i])
                shape = shape_info['shape']
                
                if shape_info['type'] == 'placeholder':
                    # استبدال placeholder
                    with open(image_path, "rb") as img_file:
                        shape.insert_picture(img_file)
                    replaced_count += 1
                    if show_details:
                        st.success(f"✅ تم استبدال placeholder: {images[i]}")
                        
                elif shape_info['type'] == 'picture':
                    # استبدال الصورة العادية
                    left, top, width, height = shape.left, shape.top, shape.width, shape.height
                    
                    # حذف الصورة القديمة
                    new_slide.shapes._spTree.remove(shape._element)
                    
                    # إضافة الصورة الجديدة
                    with open(image_path, "rb") as img_file:
                        new_slide.shapes.add_picture(img_file, left, top, width, height)
                    
                    replaced_count += 1
                    if show_details:
                        st.success(f"✅ تم استبدال صورة عادية: {images[i]}")
                        
            except Exception as e:
                if show_details:
                    st.warning(f"⚠ خطأ في استبدال الصورة {images[i]}: {e}")
        
        # إذا لم توجد أماكن للصور، أضف الصور يدوياً
        if len(all_image_shapes) == 0 and len(template_info['image_positions']) > 0:
            if show_details:
                st.info("📸 إضافة الصور باستخدام مواقع القالب")
            
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
                    if show_details:
                        st.success(f"✅ تم إضافة صورة: {images[i]}")
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
            # الخطوة 1: فحص الملف المضغوط
            st.info("📦 جاري فحص الملفات...")
            
            # استخراج الملف المضغوط
            zip_bytes = io.BytesIO(uploaded_zip.read())
            with zipfile.ZipFile(zip_bytes, "r") as zip_ref:
                temp_dir = "temp_images"
                if os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)
                os.makedirs(temp_dir)
                zip_ref.extractall(temp_dir)
            
            # جمع المجلدات
            all_items = os.listdir(temp_dir)
            folder_paths = []
            for item in all_items:
                item_path = os.path.join(temp_dir, item)
                if os.path.isdir(item_path):
                    # التحقق من وجود صور في المجلد
                    images_in_folder = [f for f in os.listdir(item_path) 
                                      if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.webp'))]
                    if images_in_folder:
                        folder_paths.append(item_path)
                        if show_details:
                            st.info(f"📁 المجلد '{item}' يحتوي على {len(images_in_folder)} صورة")
            
            if not folder_paths:
                st.error("❌ لا توجد مجلدات تحتوي على صور في الملف المضغوط.")
                st.stop()
            
            folder_paths.sort()
            st.success(f"✅ تم العثور على {len(folder_paths)} مجلد يحتوي على صور")
            
            # قراءة ملف PowerPoint
            prs = Presentation(io.BytesIO(uploaded_pptx.read()))
            
            # الخطوة 2: تحليل الشريحة الأولى
            st.info("🔍 جاري تحليل الشريحة الأولى...")
            has_images, analysis_result = analyze_first_slide(prs)
            
            if not has_images:
                # الخطوة 3: إرسال تنبيه إذا لم توجد صور
                st.error("❌ تنبيه: الشريحة الأولى لا تحتوي على صور أو placeholders للصور!")
                st.error(f"📋 السبب: {analysis_result}")
                st.info("💡 يُرجى رفع ملف PowerPoint يحتوي على:")
                st.info("   • صور في الشريحة الأولى")
                st.info("   • أو placeholders للصور")
                st.stop()
            
            # عرض نتائج التحليل
            st.success("✅ تم العثور على صور أو placeholders في الشريحة الأولى!")
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Placeholders للصور", analysis_result['placeholders'])
            with col2:
                st.metric("الصور العادية", analysis_result['regular_pictures'])
            with col3:
                st.metric("إجمالي أماكن الصور", analysis_result['total_slots'])
            
            # الحصول على معلومات القالب من الشريحة الأولى
            first_slide = prs.slides[0]
            template_info = get_template_info(first_slide)
            slide_layout = analysis_result['slide_layout']
            
            if show_details:
                st.info(f"📍 تم تحديد {len(template_info['image_positions'])} موقع للصور في القالب")
            
            # إنشاء عرض تقديمي جديد بدلاً من حذف الشرائح
            st.info("🔄 جاري إنشاء عرض تقديمي جديد...")
            new_prs = Presentation()
            
            # حذف الشريحة الافتراضية إذا وجدت
            if len(new_prs.slides) > 0:
                slide_id = new_prs.slides._sldIdLst[0]
                new_prs.slides._sldIdLst.remove(slide_id)
            
            # نسخ slide_layout إلى العرض الجديد
            # استخدام layout افتراضي إذا لم نتمكن من نسخ الأصلي
            try:
                target_layout = new_prs.slide_layouts[1]  # استخدام layout "Title and Content"
            except:
                target_layout = new_prs.slide_layouts[0]  # استخدام layout الافتراضي
            
            total_replaced = 0
            created_slides = 0
            
            # شريط التقدم
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            # إنشاء شريحة لكل مجلد
            for folder_idx, folder_path in enumerate(folder_paths):
                folder_name = os.path.basename(folder_path)
                status_text.text(f"🔄 معالجة المجلد {folder_idx + 1}/{len(folder_paths)}: {folder_name}")
                
                try:
                    # إنشاء شريحة جديدة وإضافة الصور
                    replaced_count, message = create_slide_with_images(
                        new_prs, target_layout, template_info, folder_path, folder_name, show_details
                    )
                    
                    if "تم بنجاح" in message:
                        created_slides += 1
                        total_replaced += replaced_count
                        
                        if show_details:
                            st.success(f"✅ تم إنشاء شريحة للمجلد '{folder_name}' مع {replaced_count} صورة")
                    else:
                        st.warning(f"⚠ مشكلة في المجلد {folder_name}: {message}")
                    
                except Exception as e:
                    st.error(f"❌ خطأ في معالجة المجلد {folder_name}: {e}")
                
                # تحديث شريط التقدم
                progress = (folder_idx + 1) / len(folder_paths)
                progress_bar.progress(progress)
            
            # مسح شريط التقدم
            progress_bar.empty()
            status_text.empty()
            
            # النتائج النهائية
            st.success("🎉 تم الانتهاء من المعالجة!")
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("الشرائح المُنشأة", created_slides)
            with col2:
                st.metric("الصور المُستبدلة", total_replaced)
            with col3:
                st.metric("المجلدات المُعالجة", len(folder_paths))
            
            if created_slides == 0:
                st.error("❌ لم يتم إنشاء أي شرائح.")
                st.stop()
            
            # التحقق من عدد الشرائح النهائي
            final_slide_count = len(new_prs.slides)
            st.info(f"📋 العدد النهائي للشرائح: {final_slide_count}")
            
            if final_slide_count != len(folder_paths):
                st.warning(f"⚠ تحذير: عدد الشرائح ({final_slide_count}) لا يطابق عدد المجلدات ({len(folder_paths)})")
            
            # حفظ الملف الجديد
            original_name = os.path.splitext(uploaded_pptx.name)[0]
            output_filename = f"{original_name}_Updated.pptx"
            
            output_buffer = io.BytesIO()
            new_prs.save(output_buffer)
            output_buffer.seek(0)
            
            st.success(f"✅ تم إنشاء ملف PowerPoint جديد بـ {created_slides} شريحة!")
            
            # زر التحميل
            st.download_button(
                label="⬇️ تحميل الملف المُحدث",
                data=output_buffer.getvalue(),
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                key="download_button"
            )
            
        except Exception as e:
            st.error(f"❌ خطأ أثناء المعالجة: {e}")
            import traceback
            if show_details:
                st.error(f"تفاصيل الخطأ: {traceback.format_exc()}")
        
        finally:
            # تنظيف الملفات المؤقتة
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
        
        1. **ملف PowerPoint (.pptx):**
           - يجب أن يحتوي على شريحة واحدة على الأقل
           - الشريحة الأولى يجب أن تحتوي على صور أو placeholders للصور
           - سيتم استخدام تنسيق الشريحة الأولى كقالب
        
        2. **ملف ZIP:**
           - يجب أن يحتوي على مجلدات
           - كل مجلد يجب أن يحتوي على صور
           - أسماء المجلدات ستصبح عناوين الشرائح
        
        3. **النتيجة:**
           - شريحة منفصلة لكل مجلد
           - الصور ستحل محل الصور الأصلية أو placeholders
           - الحفاظ على نفس التنسيق والأحجام
        
        ### أنواع الصور المدعومة:
        - PNG, JPG, JPEG, GIF, BMP, TIFF, WEBP
        """)
