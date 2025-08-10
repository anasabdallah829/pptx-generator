import streamlit as st
import zipfile
import os
import io
from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.util import Inches
import shutil
import pptx
from pptx.oxml.ns import qn

st.set_page_config(page_title="PowerPoint Image Replacer", layout="centered")
st.title("🔄 PowerPoint Image & Placeholder Replacer")
st.markdown("---")

# رفع الملفات
uploaded_pptx = st.file_uploader("📂 اختر ملف PowerPoint (.pptx)", type=["pptx"])
uploaded_zip = st.file_uploader("🗜️ اختر ملف ZIP يحتوي على مجلدات صور", type=["zip"])

# خيار عرض التفاصيل
show_details = st.checkbox("عرض التفاصيل المفصلة", value=False)


def analyze_first_slide(prs):
    """تحليل الشريحة الأولى: إرجاع نتائج حتى لو لم توجد مواضع للصور (لا نوقف التشغيل هنا)."""
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

    # نعيد دائماً dict تحليلي (باستثناء حالة عدم وجود شرائح)
    return True, {
        'placeholders': len(picture_placeholders),
        'regular_pictures': len(regular_pictures),
        'total_slots': total_image_slots,
        'slide_layout': first_slide.slide_layout
    }


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


def replace_images_in_slide(prs, slide, images_folder, folder_name, image_positions,
                            show_details=False, mismatch_action='truncate'):
    """
    استبدال الصور في الشريحة مع الحفاظ على المواقع والأحجام.
    prs: Presentation object (مطلوب لعرض الشريحة إذا لم توجد مواضع).
    mismatch_action: 'truncate' | 'repeat' | 'skip_folder' | 'stop'
    """
    if not os.path.exists(images_folder):
        return 0, f"المجلد {images_folder} غير موجود"

    images = [f for f in os.listdir(images_folder)
              if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.webp'))]

    if not images:
        return 0, f"لا توجد صور في المجلد {folder_name}"

    images.sort()
    replaced_count = 0

    # استبدال عنوان الشريحة (نحافظ على منطقك الأصلي)
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

    # حالة: لا توجد مواضع صور في القالب => نضيف صورة أولى مملوءة بعرض الشريحة
    if not image_positions:
        # نختار الصورة الأولى أو بحسب سياسة التكرار (repeat لا معنى هنا لأن سنعرض صورة واحدة)
        image_filename = images[0]
        image_path = os.path.join(images_folder, image_filename)
        try:
            # نفذ إضافة الصورة بمقياس يتناسب مع عرض الشريحة (يحافظ على النسبة)
            slide.shapes.add_picture(image_path, 0, 0, prs.slide_width)
            replaced_count += 1
            if show_details:
                st.success(f"✅ تم إضافة صورة مملوءة للشريحة (لا مواضع في القالب): {image_filename}")
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
            # 'skip_folder' أو 'stop' يتم التعامل معها قبل الاستدعاء عادة
            if mismatch_action == 'skip_folder':
                return 0, f"تم تخطي المجلد {folder_name} بطلب المستخدم"
            elif mismatch_action == 'stop':
                raise RuntimeError("تم إيقاف العملية بطلب المستخدم")
            else:
                # افتراضي: truncate
                if i >= len(images):
                    break
                image_filename = images[i]

        image_path = os.path.join(images_folder, image_filename)

        try:
            if pos_info['type'] == 'placeholder':
                # أسلوب آمن لاستبدال الplaceholder (يحافظ على التنسيق قدر الإمكان)
                try:
                    # insert_picture يقبل مسار الملف أو ملف باينري
                    pos_info['shape'].insert_picture(image_path)
                    replaced_count += 1
                    if show_details:
                        st.success(f"✅ تم استبدال placeholder بالصورة: {image_filename}")
                except Exception as e:
                    # محاولة احتياطية: حذف وإضافة صورة جديدة بنفس الموضع
                    try:
                        left, top, width, height = pos_info['left'], pos_info['top'], pos_info['width'], pos_info['height']
                        # حذف عنصر الplaceholder
                        slide.shapes._spTree.remove(pos_info['shape']._element)
                        new_pic = slide.shapes.add_picture(image_path, left, top, width, height)
                        replaced_count += 1
                        if show_details:
                            st.success(f"✅ (احتياطي) تم استبدال placeholder بالصورة: {image_filename}")
                    except Exception as e2:
                        if show_details:
                            st.warning(f"⚠ فشل استبدال placeholder (احتياطي): {e2}")

            elif pos_info['type'] == 'picture':
                shape = pos_info['shape']
                # الطريقة المفضلة: إضافة image part جديدة لشريحة وتغيير r:embed في blip (يحافظ على التنسيقات)
                try:
                    # حاول الحصول أو إضافة image part جديد (قد يعتمد على نسخة python-pptx)
                    image_part, new_rId = shape.part.get_or_add_image_part(image_path)
                    # إيجاد عنصر blip وتعيين embed إلى rId الجديد
                    blip = None
                    if shape._element is not None:
                        try:
                            blip_list = shape._element.xpath('.//a:blip', namespaces={
                                'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
                            })
                            if blip_list:
                                blip = blip_list[0]
                        except Exception:
                            blip = None

                    if blip is not None:
                        blip.set(qn('r:embed'), new_rId)
                        replaced_count += 1
                        if show_details:
                            st.success(f"✅ تم استبدال الصورة (محفوظ التنسيقات): {image_filename}")
                    else:
                        # إذا لم نتمكن من الوصول إلى blip، نستخدم الحل الاحتياطي (حذف واضافة)
                        raise RuntimeError("عنصر blip غير موجود لتحديث embed")

                except Exception as e:
                    # حل احتياطي: حذف الشكل وإضافة صورة جديدة بنفس الخصائص الممكنة
                    if show_details:
                        st.warning(f"⚠ تعذر الاستبدال الآمن للصورة '{image_filename}': {e}. سيتم المحاولة باحتياط.")
                    try:
                        left, top, width, height = pos_info['left'], pos_info['top'], pos_info['width'], pos_info['height']
                        # حفظ بعض الخصائص إذا كانت متاحة
                        rotation = None
                        crop_attrs = {}
                        try:
                            rotation = shape.rotation
                        except Exception:
                            rotation = None
                        try:
                            crop_attrs['left'] = getattr(shape, 'crop_left', None)
                            crop_attrs['top'] = getattr(shape, 'crop_top', None)
                            crop_attrs['right'] = getattr(shape, 'crop_right', None)
                            crop_attrs['bottom'] = getattr(shape, 'crop_bottom', None)
                        except Exception:
                            crop_attrs = {}

                        # حذف الشكل القديم
                        slide.shapes._spTree.remove(shape._element)

                        # إضافة الصورة الجديدة بنفس الموضع والأبعاد
                        new_pic = slide.shapes.add_picture(image_path, left, top, width, height)

                        # محاولة استعادة rotation و crop
                        try:
                            if rotation is not None:
                                new_pic.rotation = rotation
                        except Exception:
                            pass
                        try:
                            if crop_attrs.get('left') is not None:
                                new_pic.crop_left = crop_attrs['left']
                            if crop_attrs.get('top') is not None:
                                new_pic.crop_top = crop_attrs['top']
                            if crop_attrs.get('right') is not None:
                                new_pic.crop_right = crop_attrs['right']
                            if crop_attrs.get('bottom') is not None:
                                new_pic.crop_bottom = crop_attrs['bottom']
                        except Exception:
                            pass

                        replaced_count += 1
                        if show_details:
                            st.success(f"✅ (احتياطي) تم إضافة الصورة: {image_filename}")
                    except Exception as e2:
                        if show_details:
                            st.warning(f"⚠ فشل الاستبدال الاحتياطي للصورة {image_filename}: {e2}")
                        # نستمر دون مقاطعة العملية

        except Exception as e:
            if show_details:
                st.warning(f"⚠ خطأ في استبدال الصورة {image_filename if 'image_filename' in locals() else 'غير محدد'}: {e}")
            # لا نوقف التنفيذ عند خطأ في صورة واحدة

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
            ok, analysis_result = analyze_first_slide(prs)
            if not ok:
                st.error(f"❌ {analysis_result}")
                st.stop()

            # عرض نتائج التحليل
            st.success("✅ تحليل الشريحة الأولى جاهز")
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Placeholders للصور", analysis_result['placeholders'])
            with col2:
                st.metric("الصور العادية", analysis_result['regular_pictures'])
            with col3:
                st.metric("إجمالي أماكن الصور", analysis_result['total_slots'])

            # إذا كانت الشريحة الأولى لا تحتوي على مواضع صور - اسأل المستخدم إن أراد المتابعة
            if analysis_result['total_slots'] == 0:
                with st.form("no_slots_form"):
                    st.warning("⚠ الشريحة الأولى لا تحتوي على مواضع صور (total_slots = 0). سيتم إنشاء شرائح جديدة لكل مجلد، وستُضاف الصورة الأولى من كل مجلد مملوءة بعرض الشريحة.")
                    cont = st.form_submit_button("📌 المتابعة وإنشاء الشرائح بدون مواضع")
                    stop_btn = st.form_submit_button("⛔ إيقاف")
                if stop_btn:
                    st.info("❌ أوقف المستخدم العملية.")
                    st.stop()
                if not cont:
                    # المستخدم لم يضغط أي زر بعد
                    st.stop()
                # إذا تابع المستخدم، نكمل بدون تغيير (لا حاجة لتخزين flag إضافي)

            # الحصول على مواقع الصور من الشريحة الأولى (قد يكون فارغاً)
            first_slide = prs.slides[0]
            image_positions = get_image_positions(first_slide)

            if show_details:
                st.info(f"📍 تم تحديد {len(image_positions)} موقع للصور في الشريحة الأولى")

            # التحقق من اختلافات عدد الصور في المجلدات مقارنة بعدد مواضع الصور
            mismatch_folders = []
            folder_info_list = []
            for fp in folder_paths:
                imgs = [f for f in os.listdir(fp)
                        if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.webp'))]
                folder_info_list.append((os.path.basename(fp), len(imgs)))
                if len(imgs) != len(image_positions):
                    mismatch_folders.append((os.path.basename(fp), len(imgs), len(image_positions)))

            # تحقق من وجود اختيار سابق مخزن في session_state
            if 'mismatch_action' in st.session_state:
                mismatch_action = st.session_state['mismatch_action']
            else:
                mismatch_action = None

            if mismatch_folders and mismatch_action is None:
                # إظهار تفاصيل والطلب عبر form لتجنب rerun غير مرغوب
                with st.form("mismatch_form"):
                    st.warning("⚠ تم اكتشاف اختلاف في عدد الصور لبعض المجلدات مقارنة بعدد مواضع الصور في الشريحة الأولى.")
                    for name, img_count in folder_info_list:
                        st.write(f"- {name}: {img_count} صورة")
                    st.markdown(f"**عدد مواضع الصور (من الشريحة الأولى): {len(image_positions)}**")

                    choice = st.radio(
                        "اختر كيف تريد التعامل مع المجلدات التي يختلف عدد صورها عن عدد المواضع:",
                        (
                            "استبدال فقط حتى أقل عدد (truncate)",
                            "تكرار الصور لملء جميع المواضع (repeat)",
                            "تخطي المجلدات ذات الاختلاف (skip_folder)",
                            "إيقاف العملية (stop)"
                        ),
                        key='mismatch_choice'
                    )
                    submit_choice = st.form_submit_button("✅ تأكيد الاختيار والمتابعة")
                if not submit_choice:
                    st.stop()
                # ترجمة الاختيار إلى رمز داخلي
                if choice.startswith("استبدال فقط"):
                    mismatch_action = 'truncate'
                elif choice.startswith("تكرار"):
                    mismatch_action = 'repeat'
                elif choice.startswith("تخطي"):
                    mismatch_action = 'skip_folder'
                else:
                    mismatch_action = 'stop'
                # حفظ الاختيار في الجلسة لتجنب السلوك الذي يبعث التطبيق لإعادة البداية
                st.session_state['mismatch_action'] = mismatch_action

            # إذا لم يكن هناك اختلاف أو الاختيار مخزن سابقاً - نضبط الافتراضي
            if mismatch_action is None:
                mismatch_action = 'truncate'

            # حذف جميع الشرائح الموجودة (طريقة آمنة)
            st.info("🗑️ جاري حذف الشرائح الموجودة...")
            sldIdLst = prs.slides._sldIdLst
            for idx in range(len(sldIdLst) - 1, -1, -1):
                sldId = sldIdLst[idx]
                rId = getattr(sldId, 'rId', None)
                if rId:
                    try:
                        prs.part.drop_rel(rId)
                    except KeyError:
                        if show_details:
                            st.warning(f"⚠ العلاقة {rId} غير موجودة (تجاهل).")
                    except Exception as e:
                        if show_details:
                            st.warning(f"⚠ خطأ أثناء حذف العلاقة {rId}: {e}")
                try:
                    del sldIdLst[idx]
                except Exception as e:
                    if show_details:
                        st.warning(f"⚠ خطأ أثناء حذف شريحة عند الفهرس {idx}: {e}")

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
                    # قراءة عدد الصور في المجلد للتحكم في سياسة الاختلاف (skip_folder)
                    imgs = [f for f in os.listdir(folder_path)
                            if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.webp'))]
                    if len(imgs) == 0:
                        if show_details:
                            st.warning(f"⚠ المجلد {folder_name} لا يحتوي على صور، سيتم تخطيه.")
                        continue

                    if mismatch_action == 'skip_folder' and len(imgs) != len(image_positions):
                        if show_details:
                            st.info(f"ℹ تم تخطي المجلد {folder_name} لوجود اختلاف في عدد الصور.")
                        continue
                    if mismatch_action == 'stop' and len(imgs) != len(image_positions):
                        st.error("❌ تم إيقاف العملية بناءً على اختيار المستخدم.")
                        break

                    # إنشاء شريحة جديدة
                    new_slide = prs.slides.add_slide(slide_layout)
                    created_slides += 1

                    # الحصول على مواقع الصور في الشريحة الجديدة
                    new_image_positions = get_image_positions(new_slide)

                    # استبدال الصور (نمرر prs لاستخدام عرض الشريحة إذا لزم)
                    replaced_count, message = replace_images_in_slide(
                        prs, new_slide, folder_path, folder_name, new_image_positions, show_details, mismatch_action
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
           - إن لم تحتوي الشريحة الأولى على مواضع صور، سيُطلب منك التأكيد للمتابعة
           - سيتم استخدام تنسيق الشريحة الأولى كقالب

        2. **ملف ZIP:**
           - يجب أن يحتوي على مجلدات
           - كل مجلد يجب أن يحتوي على صور
           - أسماء المجلدات ستصبح عناوين الشرائح

        3. **النتيجة:**
           - شريحة منفصلة لكل مجلد
           - الصور ستحل محل الصور الأصلية أو placeholders
           - إذا لم توجد مواضع في القالب، تُضاف الصورة الأولى مملوءة بعرض الشريحة
           - الحفاظ على نفس التنسيق والأحجام قدر الإمكان

        ### أنواع الصور المدعومة:
        - PNG, JPG, JPEG, GIF, BMP, TIFF, WEBP
        """)

