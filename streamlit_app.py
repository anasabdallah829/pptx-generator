import streamlit as st
import zipfile
import os
import io
from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.util import Inches

st.set_page_config(page_title="PowerPoint Image Replacer", layout="centered")
st.title("🔄 PowerPoint Image & Placeholder Replacer")

uploaded_pptx = st.file_uploader("📂 اختر ملف PowerPoint (.pptx)", type=["pptx"])
uploaded_zip = st.file_uploader("🗜️ اختر ملف ZIP يحتوي على مجلدات صور", type=["zip"])

if uploaded_pptx and uploaded_zip:
    if st.button("🚀 بدء المعالجة"):
        temp_dir = None
        try:
            st.info("📦 جاري استخراج الصور من ملف ZIP...")
            zip_bytes = io.BytesIO(uploaded_zip.read())
            with zipfile.ZipFile(zip_bytes, "r") as zip_ref:
                temp_dir = "temp_images"
                if os.path.exists(temp_dir):
                    import shutil
                    shutil.rmtree(temp_dir)
                os.makedirs(temp_dir)
                zip_ref.extractall(temp_dir)

            # قراءة البوربوينت
            st.info("📄 جاري قراءة ملف PowerPoint...")
            pptx_bytes = io.BytesIO(uploaded_pptx.read())
            prs = Presentation(pptx_bytes)

            # جمع مجلدات الصور مع التحقق من وجودها
            all_items = []
            if os.path.exists(temp_dir):
                all_items = os.listdir(temp_dir)
            
            folder_paths = []
            for item in all_items:
                item_path = os.path.join(temp_dir, item)
                if os.path.exists(item_path) and os.path.isdir(item_path):
                    folder_paths.append(item_path)
            
            if not folder_paths:
                st.error("❌ ملف ZIP لا يحتوي على مجلدات صور.")
                st.stop()

            # ترتيب المجلدات أبجدياً
            folder_paths.sort()

            # التحقق من وجود شرائح في البوربوينت
            if len(prs.slides) == 0:
                st.error("❌ ملف PowerPoint لا يحتوي على أي شرائح.")
                st.stop()
            
            # الحصول على layout الشريحة الأولى كقالب
            template_slide_layout = prs.slides[0].slide_layout
            
            # إنشاء عرض تقديمي جديد بدلاً من حذف الشرائح
            new_prs = Presentation()
            
            # نسخ الـ slide master من العرض الأصلي إذا أمكن
            try:
                # استخدام نفس القالب من العرض الأصلي
                slide_layouts = prs.slide_layouts
                if len(slide_layouts) > 0:
                    template_layout = slide_layouts[0]
                else:
                    template_layout = template_slide_layout
            except:
                # في حالة فشل الوصول للـ layouts، استخدم layout افتراضي
                template_layout = new_prs.slide_layouts[0]

            st.info(f"📁 تم العثور على {len(folder_paths)} مجلد. جاري إنشاء الشرائح...")
            
            total_replaced = 0
            created_slides = 0
            
            # إنشاء شريحة جديدة لكل مجلد
            for folder_idx, folder in enumerate(folder_paths):
                if not os.path.exists(folder):
                    st.warning(f"⚠ المجلد غير موجود: {folder}")
                    continue
                    
                folder_name = os.path.basename(folder)
                
                try:
                    images = [f for f in os.listdir(folder) if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp'))]
                except OSError as e:
                    st.warning(f"⚠ خطأ في قراءة المجلد {folder_name}: {e}")
                    continue
                
                if not images:
                    st.warning(f"⚠ المجلد {folder_name} لا يحتوي على صور، تم تجاوزه.")
                    continue

                # إنشاء شريحة جديدة
                try:
                    slide = new_prs.slides.add_slide(template_layout)
                    created_slides += 1
                except Exception as e:
                    st.warning(f"⚠ خطأ في إنشاء شريحة للمجلد {folder_name}: {e}")
                    continue
                
                # تحديث شريط التقدم
                progress = (folder_idx + 1) / len(folder_paths)
                st.progress(progress, text=f"معالجة المجلد: {folder_name}")

                # وضع عنوان الشريحة من اسم المجلد
                try:
                    title_shapes = [shape for shape in slide.shapes if shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.TITLE]
                    if title_shapes:
                        title_shapes[0].text = folder_name
                except Exception as e:
                    st.warning(f"⚠ خطأ في تعيين العنوان للمجلد {folder_name}: {e}")

                img_idx = 0
                folder_replaced_count = 0
                
                # نسخ الشرائح من العرض الأصلي إذا كان هناك محتوى
                if len(prs.slides) > 0:
                    original_slide = prs.slides[0]
                    
                    for shape in slide.shapes:
                        if img_idx >= len(images):
                            break
                            
                        # استبدال في placeholder للصور
                        if shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.PICTURE:
                            try:
                                image_path = os.path.join(folder, images[img_idx])
                                if os.path.exists(image_path):
                                    with open(image_path, "rb") as img_file:
                                        shape.insert_picture(img_file)
                                    folder_replaced_count += 1
                                    img_idx += 1
                            except Exception as e:
                                st.warning(f"⚠ خطأ في استبدال الصورة {images[img_idx] if img_idx < len(images) else 'غير معروف'}: {e}")
                                img_idx += 1

                        # استبدال الصور العادية
                        elif hasattr(shape, 'shape_type') and shape.shape_type == 13:  # 13 = Picture
                            try:
                                left, top, width, height = shape.left, shape.top, shape.width, shape.height
                                slide.shapes._spTree.remove(shape._element)
                                
                                image_path = os.path.join(folder, images[img_idx])
                                if os.path.exists(image_path):
                                    with open(image_path, "rb") as img_file:
                                        pic = slide.shapes.add_picture(img_file, left, top, width, height)
                                    folder_replaced_count += 1
                                    img_idx += 1
                            except Exception as e:
                                st.warning(f"⚠ خطأ في استبدال الصورة {images[img_idx] if img_idx < len(images) else 'غير معروف'}: {e}")
                                img_idx += 1

                total_replaced += folder_replaced_count
                st.info(f"📊 المجلد {folder_name}: تم استبدال {folder_replaced_count} صورة")

            if created_slides == 0:
                st.error("❌ لم يتم إنشاء أي شرائح.")
                st.stop()

            if total_replaced == 0:
                st.warning("⚠ لم يتم استبدال أي صور. تأكد من وجود placeholders للصور في القالب.")

            # حفظ الملف الجديد
            original_name = os.path.splitext(uploaded_pptx.name)[0]
            output_filename = f"{original_name}_Modified.pptx"
            
            # حفظ في الذاكرة بدلاً من القرص
            output_buffer = io.BytesIO()
            new_prs.save(output_buffer)
            output_buffer.seek(0)

            st.success(f"✅ تم إنشاء {created_slides} شريحة جديدة!")
            st.success(f"✅ تم استبدال {total_replaced} صورة إجمالياً!")
            st.download_button(
                "⬇ تحميل الملف المعدل", 
                output_buffer.getvalue(), 
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )

        except Exception as e:
            st.error(f"❌ خطأ أثناء المعالجة: {e}")
            import traceback
            st.error(f"تفاصيل الخطأ: {traceback.format_exc()}")
        
        finally:
            # تنظيف نهائي للملفات المؤقتة
            if temp_dir and os.path.exists(temp_dir):
                try:
                    import shutil
                    shutil.rmtree(temp_dir)
                except Exception as cleanup_error:
                    st.warning(f"⚠ خطأ في تنظيف الملفات المؤقتة: {cleanup_error}")
