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
            prs = Presentation(io.BytesIO(uploaded_pptx.read()))

            # جمع مجلدات الصور
            folder_paths = [os.path.join(temp_dir, d) for d in os.listdir(temp_dir) if os.path.isdir(os.path.join(temp_dir, d))]
            if not folder_paths:
                st.error("❌ ملف ZIP لا يحتوي على مجلدات صور.")
                st.stop()

            # ترتيب المجلدات أبجدياً
            folder_paths.sort()

            # الحصول على layout الشريحة الأولى كقالب
            if len(prs.slides) == 0:
                st.error("❌ ملف PowerPoint لا يحتوي على أي شرائح.")
                st.stop()
            
            template_slide_layout = prs.slides[0].slide_layout
            
            # حذف جميع الشرائح الموجودة
            slide_ids = [slide.slide_id for slide in prs.slides]
            for slide_id in slide_ids:
                slide_to_remove = None
                for slide in prs.slides:
                    if slide.slide_id == slide_id:
                        slide_to_remove = slide
                        break
                if slide_to_remove:
                    xml_slides = prs.slides._sldIdLst
                    slides = list(xml_slides)
                    xml_slides.remove(slides[prs.slides.index(slide_to_remove)])

            st.info(f"📁 تم العثور على {len(folder_paths)} مجلد. جاري إنشاء الشرائح...")
            
            total_replaced = 0
            
            # إنشاء شريحة جديدة لكل مجلد
            for folder_idx, folder in enumerate(folder_paths):
                folder_name = os.path.basename(folder)
                images = [f for f in os.listdir(folder) if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp'))]
                
                if not images:
                    st.warning(f"⚠ المجلد {folder_name} لا يحتوي على صور، تم تجاوزه.")
                    continue

                # إنشاء شريحة جديدة
                slide = prs.slides.add_slide(template_slide_layout)
                
                # تحديث شريط التقدم
                progress = (folder_idx + 1) / len(folder_paths)
                st.progress(progress, text=f"معالجة المجلد: {folder_name}")

                # وضع عنوان الشريحة من اسم المجلد
                title_shapes = [shape for shape in slide.shapes if shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.TITLE]
                if title_shapes:
                    title_shapes[0].text = folder_name

                img_idx = 0
                folder_replaced_count = 0
                
                for shape in slide.shapes:
                    if img_idx >= len(images):
                        break
                        
                    # استبدال في placeholder للصور
                    if shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.PICTURE:
                        try:
                            image_path = os.path.join(folder, images[img_idx])
                            with open(image_path, "rb") as img_file:
                                shape.insert_picture(img_file)
                            folder_replaced_count += 1
                            img_idx += 1
                        except Exception as e:
                            st.warning(f"⚠ خطأ في استبدال الصورة {images[img_idx]}: {e}")
                            img_idx += 1

                    # استبدال الصور العادية
                    elif hasattr(shape, 'shape_type') and shape.shape_type == 13:  # 13 = Picture
                        try:
                            left, top, width, height = shape.left, shape.top, shape.width, shape.height
                            slide.shapes._spTree.remove(shape._element)
                            
                            image_path = os.path.join(folder, images[img_idx])
                            with open(image_path, "rb") as img_file:
                                pic = slide.shapes.add_picture(img_file, left, top, width, height)
                            folder_replaced_count += 1
                            img_idx += 1
                        except Exception as e:
                            st.warning(f"⚠ خطأ في استبدال الصورة {images[img_idx]}: {e}")
                            img_idx += 1

                total_replaced += folder_replaced_count
                st.info(f"📊 المجلد {folder_name}: تم استبدال {folder_replaced_count} صورة")

            if total_replaced == 0:
                st.error("❌ لم يتم العثور على أي صور أو Placeholders قابلة للاستبدال في القالب.")
                st.stop()

            # حفظ الملف الجديد
            original_name = os.path.splitext(uploaded_pptx.name)[0]
            output_filename = f"{original_name}_Modified.pptx"
            output_path = os.path.join(".", output_filename)
            prs.save(output_path)

            # تنظيف الملفات المؤقتة
            if os.path.exists(temp_dir):
                import shutil
                shutil.rmtree(temp_dir)

            with open(output_path, "rb") as f:
                file_data = f.read()
                
            # حذف الملف المؤقت
            if os.path.exists(output_path):
                os.remove(output_path)
                
            st.success(f"✅ تم إنشاء {len([f for f in folder_paths if os.listdir(f)])} شريحة جديدة!")
            st.success(f"✅ تم استبدال {total_replaced} صورة إجمالياً!")
            st.download_button(
                "⬇ تحميل الملف المعدل", 
                file_data, 
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )

        except Exception as e:
            st.error(f"❌ خطأ أثناء المعالجة: {e}")
            import traceback
            st.error(f"تفاصيل الخطأ: {traceback.format_exc()}")
        
        finally:
            # تنظيف نهائي للملفات المؤقتة
            if 'temp_dir' in locals() and os.path.exists(temp_dir):
                import shutil
                shutil.rmtree(temp_dir)
