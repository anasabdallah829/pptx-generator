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
                    
        except Exception as e:
            if show_details:
                st.warning(f"⚠ خطأ في استبدال الصورة {images[i] if i < len(images) else 'غير محدد'}: {e}")
    
    return replaced_count, "تم بنجاح"

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
            
            # الحصول على مواقع الصور من الشريحة الأولى
            first_slide = prs.slides[0]
            image_positions = get_image_positions(first_slide)
            
            if show_details:
                st.info(f"📍 تم تحديد {len(image_positions)} موقع للصور في الشريحة الأولى")
            
            # حذف جميع الشرائح الموجودة
            st.info("🗑️ جاري حذف الشرائح الموجودة...")
            slides_to_remove = list(prs.slides)
            for slide in slides_to_remove:
                rId = prs.slides._sldIdLst[prs.slides.index(slide)].rId
                prs.part.drop_rel(rId)
                del prs.slides._sldIdLst[prs.slides.index(slide)]
            
            # الخطوة 4: إنشاء شريحة لكل مجلد
            st.info("🔄 جاري إنشاء الشرائح الجديدة...")
            
            total_replaced = 0
            created_slides = 0
            slide_layout = analysis_result['slide_layout']
            
            # شريط التقدم
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            for folder_idx, folder_path in enumerate(folder_paths):
                folder_name = os.path.basename(folder_path)
                status_text.text(f"🔄 معالجة المجلد {folder_idx + 1}/{len(folder_paths)}: {folder_name}")
                
                try:
                    # إنشاء شريحة جديدة
                    new_slide = prs.slides.add_slide(slide_layout)
                    created_slides += 1
                    
                    # الحصول على مواقع الصور في الشريحة الجديدة
                    new_image_positions = get_image_positions(new_slide)
                    
                    # استبدال الصور
                    replaced_count, message = replace_images_in_slide(
                        new_slide, folder_path, folder_name, new_image_positions, show_details
                    )
                    
                    total_replaced += replaced_count
                    
                    if show_details:
                        st.success(f"✅ تم إنشاء شريحة للمجلد '{folder_name}' واستبدال {replaced_count} صورة")
                    
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
            
            # حفظ الملف الجديد
            original_name = os.path.splitext(uploaded_pptx.name)[0]
            output_filename = f"{original_name}_Updated.pptx"
            
            output_buffer = io.BytesIO()
            prs.save(output_buffer)
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
