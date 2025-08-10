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

            # عرض محتويات ملف ZIP للتشخيص
            st.info("🔍 فحص محتويات ملف ZIP...")
            all_items = os.listdir(temp_dir)
            st.write(f"العناصر الموجودة في ZIP: {all_items}")

            # جمع مجلدات الصور مع تشخيص مفصل
            folder_paths = []
            for item in all_items:
                item_path = os.path.join(temp_dir, item)
                if os.path.isdir(item_path):
                    folder_paths.append(item_path)
                    # عرض محتويات كل مجلد
                    folder_contents = os.listdir(item_path)
                    images_in_folder = [f for f in folder_contents if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp'))]
                    st.write(f"📁 المجلد '{item}' يحتوي على {len(images_in_folder)} صورة: {images_in_folder[:3]}{'...' if len(images_in_folder) > 3 else ''}")

            if not folder_paths:
                st.error("❌ ملف ZIP لا يحتوي على مجلدات صور.")
                st.stop()

            # ترتيب المجلدات أبجدياً
            folder_paths.sort()
            st.info(f"📊 تم العثور على {len(folder_paths)} مجلد للمعالجة")

            # قراءة البوربوينت
            st.info("📄 جاري قراءة ملف PowerPoint...")
            prs = Presentation(io.BytesIO(uploaded_pptx.read()))

            # التحقق من وجود شرائح
            if len(prs.slides) == 0:
                st.error("❌ ملف PowerPoint لا يحتوي على أي شرائح.")
                st.stop()
            
            st.info(f"📋 العرض التقديمي الأصلي يحتوي على {len(prs.slides)} شريحة")
            
            # الحصول على layout الشريحة الأولى كقالب
            template_slide_layout = prs.slides[0].slide_layout
            
            # طريقة محسنة لحذف جميع الشرائح
            st.info("🗑️ جاري حذف الشرائح الموجودة...")
            slides_to_remove = list(prs.slides)
            for slide in slides_to_remove:
                rId = prs.slides._sldIdLst[prs.slides.index(slide)].rId
                prs.part.drop_rel(rId)
                del prs.slides._sldIdLst[prs.slides.index(slide)]
            
            st.info(f"✅ تم حذف جميع الشرائح. العدد الحالي: {len(prs.slides)}")
            
            total_replaced = 0
            created_slides_count = 0
            
            # إنشاء شريحة جديدة لكل مجلد
            for folder_idx, folder in enumerate(folder_paths):
                folder_name = os.path.basename(folder)
                st.info(f"🔄 معالجة المجلد {folder_idx + 1}/{len(folder_paths)}: {folder_name}")
                
                images = [f for f in os.listdir(folder) if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp'))]
                
                if not images:
                    st.warning(f"⚠ المجلد {folder_name} لا يحتوي على صور، تم تجاوزه.")
                    continue

                # إنشاء شريحة جديدة
                try:
                    slide = prs.slides.add_slide(template_slide_layout)
                    created_slides_count += 1
                    st.success(f"✅ تم إنشاء الشريحة رقم {created_slides_count} للمجلد: {folder_name}")
                except Exception as e:
                    st.error(f"❌ خطأ في إنشاء شريحة للمجلد {folder_name}: {e}")
                    continue
                
                # تحديث شريط التقدم
                progress = (folder_idx + 1) / len(folder_paths)
                st.progress(progress, text=f"معالجة المجلد: {folder_name}")

                # وضع عنوان الشريحة من اسم المجلد
                try:
                    title_shapes = [shape for shape in slide.shapes if shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.TITLE]
                    if title_shapes:
                        title_shapes[0].text = folder_name
                        st.info(f"📝 تم تعيين العنوان: {folder_name}")
                    else:
                        st.warning(f"⚠ لم يتم العثور على placeholder للعنوان في الشريحة")
                except Exception as e:
                    st.warning(f"⚠ خطأ في تعيين العنوان: {e}")

                img_idx = 0
                folder_replaced_count = 0
                
                # عد الـ placeholders والصور الموجودة
                picture_placeholders = [shape for shape in slide.shapes if shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.PICTURE]
                regular_pictures = [shape for shape in slide.shapes if hasattr(shape, 'shape_type') and shape.shape_type == 13]
                
                st.info(f"🖼️ الشريحة تحتوي على {len(picture_placeholders)} placeholder للصور و {len(regular_pictures)} صورة عادية")
                
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
                            st.success(f"✅ تم استبدال صورة في placeholder: {images[img_idx-1]}")
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
                            st.success(f"✅ تم استبدال صورة عادية: {images[img_idx-1]}")
                        except Exception as e:
                            st.warning(f"⚠ خطأ في استبدال الصورة {images[img_idx]}: {e}")
                            img_idx += 1

                total_replaced += folder_replaced_count
                st.info(f"📊 المجلد {folder_name}: تم استبدال {folder_replaced_count} صورة")

            # التحقق النهائي
            st.info(f"📋 العدد النهائي للشرائح في العرض: {len(prs.slides)}")
            st.info(f"🎯 تم إنشاء {created_slides_count} شريحة فعلياً")

            if created_slides_count == 0:
                st.error("❌ لم يتم إنشاء أي شرائح.")
                st.stop()

            if total_replaced == 0:
                st.warning("⚠ لم يتم استبدال أي صور. تأكد من وجود placeholders للصور في القالب.")

            # حفظ الملف الجديد
            original_name = os.path.splitext(uploaded_pptx.name)[0]
            output_filename = f"{original_name}_Modified.pptx"
            
            # حفظ في الذاكرة
            output_buffer = io.BytesIO()
            prs.save(output_buffer)
            output_buffer.seek(0)

            st.success(f"✅ تم إنشاء {created_slides_count} شريحة جديدة!")
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
