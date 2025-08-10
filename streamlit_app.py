import streamlit as st
import zipfile
import os
import io
from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER, MSO_SHAPE
from pptx.util import Inches
import shutil
import pptx
from pptx.oxml.ns import qn

# إعداد صفحة Streamlit
st.set_page_config(page_title="PowerPoint Image Replacer", layout="centered")
st.title("🔄 PowerPoint Image & Placeholder Replacer")
st.markdown("---")

# واجهة المستخدم لرفع الملفات
uploaded_pptx = st.file_uploader("📂 اختر ملف PowerPoint (.pptx)", type=["pptx"])
uploaded_zip = st.file_uploader("🗜️ اختر ملف ZIP يحتوي على مجلدات صور", type=["zip"])

# خيار عرض التفاصيل
show_details = st.checkbox("عرض التفاصيل المفصلة", value=False)


def analyze_first_slide(prs):
    """
    تحليل الشريحة الأولى: إرجاع نتائج حتى لو لم توجد مواضع للصور.
    """
    if len(prs.slides) == 0:
        return False, "لا توجد شرائح في الملف"

    first_slide = prs.slides[0]

    # البحث عن placeholders للصور
    picture_placeholders = [
        shape for shape in first_slide.shapes
        if shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.PICTURE
    ]

    # البحث عن الصور العادية باستخدام MSO_SHAPE.PICTURE
    regular_pictures = [
        shape for shape in first_slide.shapes
        if hasattr(shape, 'shape_type') and shape.shape_type == MSO_SHAPE.PICTURE.value
    ]

    total_image_slots = len(picture_placeholders) + len(regular_pictures)

    return True, {
        'placeholders': len(picture_placeholders),
        'regular_pictures': len(regular_pictures),
        'total_slots': total_image_slots,
        'slide_layout': first_slide.slide_layout
    }


def get_image_positions(slide):
    """
    استخراج مواقع وأحجام الصور من الشريحة، سواء كانت placeholders أو صور عادية.
    """
    positions = []

    for shape in slide.shapes:
        if shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.PICTURE:
            positions.append({
                'shape': shape,
                'type': 'placeholder',
                'left': shape.left,
                'top': shape.top,
                'width': shape.width,
                'height': shape.height
            })
        elif hasattr(shape, 'shape_type') and shape.shape_type == MSO_SHAPE.PICTURE.value:
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


def replace_images_in_slide(prs, slide, images_folder, folder_name, image_positions,
                            show_details=False, mismatch_action='truncate'):
    """
    استبدال الصور في الشريحة مع الحفاظ على المواقع والأحجام.
    """
    if not os.path.exists(images_folder):
        return 0, f"المجلد {images_folder} غير موجود"

    images = sorted([f for f in os.listdir(images_folder)
                     if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.webp'))])

    if not images:
        return 0, f"لا توجد صور في المجلد {folder_name}"

    replaced_count = 0

    # استبدال عنوان الشريحة
    try:
        title_shapes = [shape for shape in slide.shapes
                        if shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.TITLE]
        if title_shapes:
            title_shapes[0].text = folder_name
        else:
            # إضافة عنوان جديد إذا لم يوجد
            textbox = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
            text_frame = textbox.text_frame
            text_frame.text = folder_name
    except Exception:
        if show_details:
            st.warning(f"⚠ خطأ في تعيين العنوان، تم تجاهله.")

    # حالة: لا توجد مواضع صور في القالب => نضيف صورة أولى مملوءة بعرض الشريحة
    if not image_positions:
        if images:
            image_filename = images[0]
            image_path = os.path.join(images_folder, image_filename)
            try:
                slide.shapes.add_picture(image_path, 0, 0, prs.slide_width)
                replaced_count += 1
                if show_details:
                    st.success(f"✅ تم إضافة صورة مملوءة للشريحة (لا توجد مواضع): {image_filename}")
            except Exception as e:
                if show_details:
                    st.warning(f"⚠ فشل إضافة الصورة المملوءة: {e}")
        return replaced_count, "تم بنجاح (بدون مواضع)"

    # معالجة كل موضع صورة
    for i, pos_info in enumerate(image_positions):
        # اختيار الصورة وفق سياسة الاختلاف
        if mismatch_action == 'truncate':
            if i >= len(images):
                break
            image_filename = images[i]
        elif mismatch_action == 'repeat':
            image_filename = images[i % len(images)]
        else:
            # حالات 'skip_folder' أو 'stop' تُعالج في الدالة الرئيسية
            if mismatch_action == 'skip_folder':
                return 0, f"تم تخطي المجلد {folder_name} بطلب المستخدم"
            elif mismatch_action == 'stop':
                raise RuntimeError("تم إيقاف العملية بطلب المستخدم")
            else:
                if i >= len(images):
                    break
                image_filename = images[i]

        image_path = os.path.join(images_folder, image_filename)

        try:
            shape = pos_info['shape']
            left, top, width, height = pos_info['left'], pos_info['top'], pos_info['width'], pos_info['height']
            
            # حذف الشكل القديم
            sp_tree = slide.shapes._spTree
            sp_tree.remove(shape._element)

            # إضافة صورة جديدة بنفس الموضع والأبعاد
            new_pic = slide.shapes.add_picture(image_path, left, top, width, height)
            
            replaced_count += 1
            if show_details:
                st.success(f"✅ تم استبدال الصورة (حذف وإضافة): {image_filename}")
        except Exception as e:
            if show_details:
                st.warning(f"⚠ فشل استبدال الصورة {image_filename}: {e}")
            # نستمر دون مقاطعة العملية

    return replaced_count, "تم بنجاح"


if uploaded_pptx and uploaded_zip:
    if st.button("🚀 بدء المعالجة"):
        temp_dir = None
        try:
            # الخطوة 1: فحص الملف المضغوط
            st.info("📦 جاري فحص الملفات...")
            zip_bytes = io.BytesIO(uploaded_zip.read())
            with zipfile.ZipFile(zip_bytes, "r") as zip_ref:
                temp_dir = "temp_images"
                if os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)
                os.makedirs(temp_dir)
                zip_ref.extractall(temp_dir)

            all_items = os.listdir(temp_dir)
            folder_paths = [os.path.join(temp_dir, item) for item in all_items if os.path.isdir(os.path.join(temp_dir, item))]
            
            if not folder_paths:
                st.error("❌ لا توجد مجلدات في الملف المضغوط.")
                st.stop()
            
            folder_paths.sort()
            st.success(f"✅ تم العثور على {len(folder_paths)} مجلد يحتوي على صور")

            # قراءة ملف PowerPoint
            prs = Presentation(io.BytesIO(uploaded_pptx.read()))

            # الخطوة 2: تحليل الشريحة الأولى
            st.info("🔍 جاري تحليل الشريحة الأولى...")
            ok, analysis_result = analyze_first_slide(prs)
            if not ok:
                st.error(f"❌ {analysis_result}")
                st.stop()

            st.success("✅ تحليل الشريحة الأولى جاهز")
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Placeholders للصور", analysis_result['placeholders'])
            with col2:
                st.metric("الصور العادية", analysis_result['regular_pictures'])
            with col3:
                st.metric("إجمالي أماكن الصور", analysis_result['total_slots'])

            # الحصول على مواقع الصور من الشريحة الأولى (قد يكون فارغاً)
            first_slide = prs.slides[0]
            image_positions = get_image_positions(first_slide)

            # معالجة حالة عدم وجود مواضع صور
            if analysis_result['total_slots'] == 0:
                st.warning("⚠ الشريحة الأولى لا تحتوي على مواضع صور. سيتم إنشاء شرائح جديدة وإضافة الصورة الأولى من كل مجلد.")
                mismatch_action = 'truncate' # لا يوجد خيار آخر منطقي هنا
            else:
                mismatch_action = 'truncate' # القيمة الافتراضية

            # التحقق من اختلافات عدد الصور في المجلدات
            mismatch_folders = []
            folder_info_list = []
            for fp in folder_paths:
                imgs = [f for f in os.listdir(fp) if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.webp'))]
                folder_info_list.append((os.path.basename(fp), len(imgs)))
                if len(imgs) != len(image_positions):
                    mismatch_folders.append((os.path.basename(fp), len(imgs), len(image_positions)))

            # إذا كانت هناك اختلافات ولم يتم اتخاذ إجراء بعد
            if mismatch_folders and 'mismatch_action' not in st.session_state:
                with st.form("mismatch_form"):
                    st.warning("⚠ تم اكتشاف اختلاف في عدد الصور لبعض المجلدات مقارنة بعدد مواضع الصور في الشريحة الأولى.")
                    for name, img_count, _ in mismatch_folders:
                        st.write(f"- المجلد `{name}` يحتوي على {img_count} صورة.")
                    st.markdown(f"**عدد مواضع الصور في القالب: {len(image_positions)}**")

                    choice_text = st.radio(
                        "اختر كيف تريد التعامل مع المجلدات التي يختلف عدد صورها:",
                        ("استبدال فقط حتى أقل عدد (truncate)", "تكرار الصور لملء جميع المواضع (repeat)", "تخطي المجلدات ذات الاختلاف (skip_folder)", "إيقاف العملية (stop)"),
                        index=0
                    )
                    submit_choice = st.form_submit_button("✅ تأكيد الاختيار والمتابعة")

                if submit_choice:
                    if choice_text.startswith("استبدال فقط"):
                        st.session_state['mismatch_action'] = 'truncate'
                    elif choice_text.startswith("تكرار"):
                        st.session_state['mismatch_action'] = 'repeat'
                    elif choice_text.startswith("تخطي"):
                        st.session_state['mismatch_action'] = 'skip_folder'
                    else:
                        st.session_state['mismatch_action'] = 'stop'
                    # إعادة تشغيل التطبيق بعد حفظ الاختيار
                    st.experimental_rerun()
                else:
                    st.stop()
            
            if 'mismatch_action' in st.session_state:
                mismatch_action = st.session_state['mismatch_action']
                if mismatch_action == 'stop':
                    st.error("❌ تم إيقاف العملية بناءً على اختيار المستخدم.")
                    st.stop()

            # حذف جميع الشرائح الموجودة (طريقة آمنة)
            st.info("🗑️ جاري حذف الشرائح الموجودة...")
            while prs.slides:
                slide_id = prs.slides[0].slide_id
                prs.part.delete_slide(slide_id)
            st.success("✅ تم حذف جميع الشرائح القديمة.")
            
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

                imgs = [f for f in os.listdir(folder_path) if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.webp'))]
                if not imgs:
                    if show_details:
                        st.warning(f"⚠ المجلد {folder_name} فارغ من الصور، سيتم تخطيه.")
                    continue

                if mismatch_action == 'skip_folder' and len(imgs) != len(image_positions):
                    if show_details:
                        st.info(f"ℹ تم تخطي المجلد {folder_name} لوجود اختلاف في عدد الصور.")
                    continue

                # إنشاء شريحة جديدة
                new_slide = prs.slides.add_slide(slide_layout)
                created_slides += 1
                
                # الحصول على مواقع الصور من الشريحة الجديدة
                new_image_positions = get_image_positions(new_slide)

                replaced_count, message = replace_images_in_slide(
                    prs, new_slide, folder_path, folder_name, new_image_positions, show_details, mismatch_action
                )

                total_replaced += replaced_count
                if show_details:
                    st.success(f"✅ تم إنشاء شريحة للمجلد '{folder_name}' واستبدال {replaced_count} صورة")

                # تحديث شريط التقدم
                progress = (folder_idx + 1) / len(folder_paths)
                progress_bar.progress(progress)

            # مسح شريط التقدم
            progress_bar.empty()
            status_text.empty()

            # النتائج النهائية
            st.success("🎉 تم الانتهاء من المعالجة!")
            if 'mismatch_action' in st.session_state:
                del st.session_state['mismatch_action']

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
                shutil.rmtree(temp_dir)
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
        """)
