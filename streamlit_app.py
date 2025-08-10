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

# إضافة خيار للمستخدم
operation_mode = st.radio(
    "اختر طريقة المعالجة:",
    ["إضافة شرائح جديدة (الحفاظ على الشرائح الأصلية)", "استبدال جميع الشرائح"],
    index=0
)

# خيار إظهار التفاصيل
show_details = st.checkbox("عرض التفاصيل المفصلة", value=False)

if uploaded_pptx and uploaded_zip:
    if st.button("🚀 بدء المعالجة"):
        temp_dir = None
        try:
            if show_details:
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
            if show_details:
                st.info("🔍 فحص محتويات ملف ZIP...")
                all_items = os.listdir(temp_dir)
                st.write(f"العناصر الموجودة في ZIP: {all_items}")
            else:
                all_items = os.listdir(temp_dir)

            # جمع مجلدات الصور مع تشخيص مفصل
            folder_paths = []
            for item in all_items:
                item_path = os.path.join(temp_dir, item)
                if os.path.isdir(item_path):
                    folder_paths.append(item_path)
                    if show_details:
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
            if show_details:
                st.info("📄 جاري قراءة ملف PowerPoint...")
            
            prs = Presentation(io.BytesIO(uploaded_pptx.read()))

            # التحقق من وجود شرائح
            if len(prs.slides) == 0:
                st.error("❌ ملف PowerPoint لا يحتوي على أي شرائح.")
                st.stop()

            original_slides_count = len(prs.slides)
            st.info(f"📋 العرض التقديمي الأصلي يحتوي على {original_slides_count} شريحة")

            # الحصول على layout الشريحة الأولى كقالب
            template_slide_layout = prs.slides[0].slide_layout

            # معالجة الشرائح حسب الخيار المحدد
            if operation_mode == "استبدال جميع الشرائح":
                if show_details:
                    st.info("🗑️ جاري حذف الشرائح الموجودة...")
                
                slides_to_remove = list(prs.slides)
                for slide in slides_to_remove:
                    rId = prs.slides._sldIdLst[prs.slides.index(slide)].rId
                    prs.part.drop_rel(rId)
                    del prs.slides._sldIdLst[prs.slides.index(slide)]
                
                if show_details:
                    st.info(f"✅ تم حذف جميع الشرائح. العدد الحالي: {len(prs.slides)}")
            else:
                if show_details:
                    st.info("📝 سيتم الحفاظ على الشرائح الأصلية وإضافة شرائح جديدة")

            total_replaced = 0
            created_slides_count = 0

            # شريط التقدم الرئيسي
            progress_bar = st.progress(0)
            status_text = st.empty()

            # إنشاء شريحة جديدة لكل مجلد
            for folder_idx, folder in enumerate(folder_paths):
                folder_name = os.path.basename(folder)
                status_text.text(f"🔄 معالجة المجلد {folder_idx + 1}/{len(folder_paths)}: {folder_name}")

                images = [f for f in os.listdir(folder) if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp'))]

                if not images:
                    if show_details:
                        st.warning(f"⚠ المجلد {folder_name} لا يحتوي على صور، تم تجاوزه.")
                    continue

                # إنشاء شريحة جديدة
                try:
                    slide = prs.slides.add_slide(template_slide_layout)
                    created_slides_count += 1
                    current_slide_number = len(prs.slides)
                    if show_details:
                        st.success(f"✅ تم إنشاء الشريحة رقم {current_slide_number} للمجلد: {folder_name}")
                except Exception as e:
                    st.error(f"❌ خطأ في إنشاء شريحة للمجلد {folder_name}: {e}")
                    continue

                # تحديث شريط التقدم
                progress = (folder_idx + 1) / len(folder_paths)
                progress_bar.progress(progress)

                # وضع عنوان الشريحة من اسم المجلد
                try:
                    title_shapes = [shape for shape in slide.shapes if shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.TITLE]
                    if title_shapes:
                        title_shapes[0].text = folder_name
                        if show_details:
                            st.info(f"📝 تم تعيين العنوان: {folder_name}")
                    else:
                        # إذا لم يوجد placeholder للعنوان، أضف نص في أعلى الشريحة
                        try:
                            textbox = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
                            text_frame = textbox.text_frame
                            text_frame.text = folder_name
                            # تنسيق النص
                            paragraph = text_frame.paragraphs[0]
                            paragraph.font.size = Inches(0.3)
                            paragraph.font.bold = True
                            if show_details:
                                st.info(f"📝 تم إضافة العنوان كنص: {folder_name}")
                        except Exception as title_error:
                            if show_details:
                                st.warning(f"⚠ لم يتم العثور على placeholder للعنوان: {title_error}")
                except Exception as e:
                    if show_details:
                        st.warning(f"⚠ خطأ في تعيين العنوان: {e}")

                img_idx = 0
                folder_replaced_count = 0

                # جمع معلومات الصور والـ placeholders
                picture_placeholders = [shape for shape in slide.shapes if shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.PICTURE]
                regular_pictures = [shape for shape in slide.shapes if hasattr(shape, 'shape_type') and shape.shape_type == 13]

                if show_details:
                    st.info(f"🖼️ الشريحة تحتوي على {len(picture_placeholders)} placeholder للصور و {len(regular_pictures)} صورة عادية")

                # إذا لم توجد placeholders للصور، أضف الصور يدوياً
                if len(picture_placeholders) == 0 and len(regular_pictures) == 0:
                    if show_details:
                        st.info("📸 لا توجد placeholders للصور، سيتم إضافة الصور يدوياً")
                    
                    # إضافة الصور في شبكة
                    images_per_row = 3
                    image_width = Inches(2.5)
                    image_height = Inches(2)
                    start_left = Inches(1)
                    start_top = Inches(2)
                    
                    for i, image_name in enumerate(images[:9]):  # حد أقصى 9 صور
                        try:
                            row = i // images_per_row
                            col = i % images_per_row
                            left = start_left + col * (image_width + Inches(0.5))
                            top = start_top + row * (image_height + Inches(0.5))
                            
                            image_path = os.path.join(folder, image_name)
                            with open(image_path, "rb") as img_file:
                                slide.shapes.add_picture(img_file, left, top, image_width, image_height)
                                folder_replaced_count += 1
                                if show_details:
                                    st.success(f"✅ تم إضافة الصورة: {image_name}")
                        except Exception as e:
                            if show_details:
                                st.warning(f"⚠ خطأ في إضافة الصورة {image_name}: {e}")
                else:
                    # جمع معلومات الصور الموجودة مع مواقعها وأحجامها
                    shapes_info = []
                    
                    for shape in slide.shapes:
                        if shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.PICTURE:
                            shapes_info.append({
                                'shape': shape,
                                'type': 'placeholder',
                                'left': shape.left,
                                'top': shape.top,
                                'width': shape.width,
                                'height': shape.height
                            })
                        elif hasattr(shape, 'shape_type') and shape.shape_type == 13:  # Picture
                            shapes_info.append({
                                'shape': shape,
                                'type': 'picture',
                                'left': shape.left,
                                'top': shape.top,
                                'width': shape.width,
                                'height': shape.height
                            })
                    
                    # ترتيب الصور حسب الموقع (من اليسار لليمين، من الأعلى للأسفل)
                    shapes_info.sort(key=lambda x: (x['top'], x['left']))
                    
                    # استبدال الصور مع الحفاظ على المواقع والأحجام
                    for shape_info in shapes_info:
                        if img_idx >= len(images):
                            break
                            
                        try:
                            shape = shape_info['shape']
                            left = shape_info['left']
                            top = shape_info['top']
                            width = shape_info['width']
                            height = shape_info['height']
                            
                            # حذف الصورة الأصلية
                            if shape_info['type'] == 'placeholder':
                                # للـ placeholder، نحتاج لمعالجة خاصة
                                try:
                                    image_path = os.path.join(folder, images[img_idx])
                                    with open(image_path, "rb") as img_file:
                                        shape.insert_picture(img_file)
                                    folder_replaced_count += 1
                                    if show_details:
                                        st.success(f"✅ تم استبدال صورة في placeholder: {images[img_idx]}")
                                except Exception as e:
                                    if show_details:
                                        st.warning(f"⚠ خطأ في استبدال placeholder: {e}")
                            else:
                                # للصور العادية، احذف وأضف جديدة بنفس المواقع والأحجام
                                slide.shapes._spTree.remove(shape._element)
                                
                                image_path = os.path.join(folder, images[img_idx])
                                with open(image_path, "rb") as img_file:
                                    new_pic = slide.shapes.add_picture(img_file, left, top, width, height)
                                
                                folder_replaced_count += 1
                                if show_details:
                                    st.success(f"✅ تم استبدال صورة عادية: {images[img_idx]}")
                            
                            img_idx += 1
                            
                        except Exception as e:
                            if show_details:
                                st.warning(f"⚠ خطأ في استبدال الصورة {images[img_idx] if img_idx < len(images) else 'غير محدد'}: {e}")
                            img_idx += 1

                total_replaced += folder_replaced_count
                if show_details:
                    st.info(f"📊 المجلد {folder_name}: تم معالجة {folder_replaced_count} صورة")

            # مسح شريط التقدم والحالة
            progress_bar.empty()
            status_text.empty()

            # التحقق النهائي
            final_slides_count = len(prs.slides)
            st.success(f"📋 العدد النهائي للشرائح في العرض: {final_slides_count}")
            
            if operation_mode == "إضافة شرائح جديدة (الحفاظ على الشرائح الأصلية)":
                st.info(f"📊 الشرائح الأصلية: {original_slides_count}")
                st.info(f"🆕 الشرائح الجديدة المضافة: {created_slides_count}")
            else:
                st.info(f"🎯 تم إنشاء {created_slides_count} شريحة جديدة (استبدال كامل)")

            if created_slides_count == 0:
                st.error("❌ لم يتم إنشاء أي شرائح جديدة.")
                st.stop()

            # حفظ الملف الجديد
            original_name = os.path.splitext(uploaded_pptx.name)[0]
            if operation_mode == "إضافة شرائح جديدة (الحفاظ على الشرائح الأصلية)":
                output_filename = f"{original_name}_Enhanced.pptx"
            else:
                output_filename = f"{original_name}_Replaced.pptx"

            # حفظ في الذاكرة
            output_buffer = io.BytesIO()
            prs.save(output_buffer)
            output_buffer.seek(0)

            st.success(f"✅ تم إنشاء {created_slides_count} شريحة جديدة!")
            st.success(f"✅ تم معالجة {total_replaced} صورة إجمالياً!")
            st.success(f"📋 العرض النهائي يحتوي على {final_slides_count} شريحة")
            
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
                    if show_details:
                        st.warning(f"⚠ خطأ في تنظيف الملفات المؤقتة: {cleanup_error}")
