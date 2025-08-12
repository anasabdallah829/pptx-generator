import streamlit as st
import zipfile
import os
import io
from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
import shutil
from pptx.util import Inches
import random
import tempfile
import copy
from io import BytesIO

# Set Streamlit page configuration
st.set_page_config(page_title="Slide-Sync-Images (Fixed)", layout="centered", initial_sidebar_state="expanded")

# --- (مقتصر) CSS لتنسيق واجهة بسيطة ---
st.markdown("""
<style>
    .stApp { background-color: #f7f9fc; }
    .main-header { text-align: center; font-size: 2em; color: #004d99; }
</style>
""", unsafe_allow_html=True)

st.markdown('<h1 class="main-header">🔄 Slide-Sync-Images — Fixed</h1>', unsafe_allow_html=True)

# --- واجهة المستخدم ---
uploaded_pptx = st.file_uploader("اختر ملف PowerPoint للقالب (.pptx)", type=["pptx"], key="pptx_uploader")
uploaded_zip = st.file_uploader("اختر ملف ZIP يحتوي على مجلدات الصور", type=["zip"], key="zip_uploader")

image_order_option = st.radio(
    "كيف تريد ترتيب الصور في الشرائح؟",
    ("بالترتيب (افتراضي)", "عشوائي"),
    index=0
)

show_details = st.checkbox("عرض التفاصيل المفصلة", value=False)

# --- مساعدة ---
st.sidebar.markdown("""
**تعليمات قصيرة:**
- ZIP يجب أن يحتوي على مجلدات (كل مجلد لشريحة واحدة).
- اسم المجلد سيصبح عنوان الشريحة.
""")

# --- وظائف مساعدة ---
PICTURE_SHAPE_TYPES = (13, 21)  # as in original: picture and picture frame


def get_image_shapes(slide):
    """
    إرجاع قائمة الأشكال التي تحتوي على صور في الشريحة.
    يشمل: picture placeholders و shapes لديها خاصية image.
    يتم ترتيبها بموقع (top, left) لتحديد ترتيب ثابت.
    """
    image_shapes = []
    for shape in slide.shapes:
        try:
            if shape.is_placeholder and hasattr(shape, 'placeholder_format') and shape.placeholder_format.type == PP_PLACEHOLDER.PICTURE:
                image_shapes.append(shape)
                continue
        except Exception:
            # shape قد لا يملك placeholder_format
            pass

        # شكل صور عادي (مثل صورة مضافة مباشرة)
        if hasattr(shape, 'image'):
            image_shapes.append(shape)
            continue

        # أحياناً يتم التعرف عليها عبر shape_type
        if hasattr(shape, 'shape_type') and shape.shape_type in PICTURE_SHAPE_TYPES:
            image_shapes.append(shape)

    # ترتيب ثابت (من أعلى لأسفل ثم من اليسار لليمين)
    image_shapes.sort(key=lambda s: (getattr(s, 'top', 0), getattr(s, 'left', 0)))
    return image_shapes


def duplicate_slide(presentation, source_slide):
    """
    استنساخ الشريحة source_slide إلى نهاية العرض (presentation).
    الطريقة: نسخ كل الأشكال (deepcopy) ثم إضافة الصور من blob للحفاظ على الربط الصحيح.
    هذه الطريقة أكثر ثباتاً من محاولة استخدام slide_layout مباشرة.
    """
    # محاولة اختيار قالب فارغ مناسب
    try:
        blank_layout = presentation.slide_layouts[-1]
    except Exception:
        blank_layout = presentation.slide_layouts[0]

    new_slide = presentation.slides.add_slide(blank_layout)

    # حذف أي أشكال ابتدائية جلبها layout
    for shp in list(new_slide.shapes):
        try:
            new_slide.shapes._spTree.remove(shp._element)
        except Exception:
            try:
                shp._element.getparent().remove(shp._element)
            except Exception:
                pass

    # نجمع معلومات الصور أولاً (so we can re-add them later to avoid relationship collisions)
    images_to_add = []  # list of (left, top, width, height, image_blob)

    for shp in source_slide.shapes:
        if hasattr(shp, 'image'):
            try:
                blob = shp.image.blob
                images_to_add.append((shp.left, shp.top, shp.width, shp.height, blob))
            except Exception:
                # في حال فشل الحصول على blob، نتجاهل
                pass
        else:
            # نسخ العنصر XML كاملاً
            try:
                el = shp._element
                new_el = copy.deepcopy(el)
                new_slide.shapes._spTree.insert_element_before(new_el, 'p:extLst')
            except Exception:
                # إذا فشل نسخ عنصر واحد، نستمر
                if show_details:
                    st.warning(f"تعذر نسخ شكل: {getattr(shp, 'name', 'unknown')}")

    # نضيف الصور في نهاية لظهورها في المقدمة
    for left, top, width, height, blob in images_to_add:
        try:
            img_stream = BytesIO(blob)
            new_slide.shapes.add_picture(img_stream, left, top, width, height)
        except Exception:
            if show_details:
                st.warning("فشل إضافة صورة عند استنساخ الشريحة.")

    return new_slide


# --- الوظيفة الرئيسية ---

def main():
    if not uploaded_pptx or not uploaded_zip:
        st.info("👋 قم بتحميل ملف PowerPoint وملف ZIP للبدء.")
        return

    if st.button("🚀 بدء المعالجة", use_container_width=True):
        temp_dir = None
        try:
            with st.spinner("📦 جاري فحص واستخراج الملفات..."):
                zip_bytes = io.BytesIO(uploaded_zip.read())
                tmp = tempfile.mkdtemp(prefix="slide_sync_")
                temp_dir = tmp
                with zipfile.ZipFile(zip_bytes, 'r') as z:
                    z.extractall(tmp)

            # جمع مسارات المجلدات (المجلدات المباشرة داخل temp_dir)
            all_items = sorted(os.listdir(temp_dir))
            folder_paths = [os.path.join(temp_dir, it) for it in all_items if os.path.isdir(os.path.join(temp_dir, it))]

            if not folder_paths:
                st.error("❌ ملف ZIP لا يحتوي على أي مجلدات صور في المستوى الأول.")
                return

            prs = Presentation(io.BytesIO(uploaded_pptx.read()))

            # تحليل الشريحة الأولى (ستكون القالب)
            if len(prs.slides) == 0:
                st.error("❌ الملف لا يحتوي على شرائح.")
                return

            template_slide = prs.slides[0]
            template_image_shapes = get_image_shapes(template_slide)

            st.success(f"✅ تم تحليل القالب: عدد أماكن الصور في الشريحة الأولى = {len(template_image_shapes)}")

            # التحقق من التوافق بين عدد الصور وعدد أماكن القالب
            mismatch_folders = []
            for fp in folder_paths:
                imgs = [f for f in os.listdir(fp) if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.webp'))]
                if len(imgs) != len(template_image_shapes):
                    mismatch_folders.append((os.path.basename(fp), len(imgs), len(template_image_shapes)))

            mismatch_action = 'truncate'
            if mismatch_folders:
                st.warning("⚠️ تم اكتشاف اختلافات في عدد الصور لبعض المجلدات مقارنة بالقالب.")
                for name, img_count, expected in mismatch_folders:
                    st.write(f"- المجلد `{name}` يحتوي على {img_count} صورة. (المتوقع: {expected})")

                choice_text = st.radio(
                    "اختر كيفية التعامل مع المجلدات التي يختلف عدد صورها:",
                    ("اقتصاص (استبدال حتى أقل عدد)", "تكرار (ملء كل الأماكن بتكرار الصور)", "تخطي (تجاهل المجلدات التي بها اختلاف)", "إيقاف العملية"),
                    index=0
                )
                mapping = {
                    "اقتصاص (استبدال حتى أقل عدد)": 'truncate',
                    "تكرار (ملء كل الأماكن بتكرار الصور)": 'repeat',
                    "تخطي (تجاهل المجلدات التي بها اختلاف)": 'skip_folder',
                    "إيقاف العملية": 'stop'
                }
                mismatch_action = mapping[choice_text]

                if mismatch_action == 'stop':
                    st.error("❌ ألغيت العملية بناءً على اختيارك.")
                    return

            # بدء إنشاء الشرائح والبدء بالاستنساخ ثم الاستبدال
            total_replaced = 0
            created_slides = 0

            progress = st.progress(0)
            status = st.empty()

            for idx, folder in enumerate(folder_paths):
                folder_name = os.path.basename(folder)
                status.text(f"جاري معالجة {idx+1}/{len(folder_paths)}: {folder_name}")

                imgs = [f for f in os.listdir(folder) if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.webp'))]
                if not imgs:
                    if show_details:
                        st.warning(f"المجلد '{folder_name}' لا يحتوي على صور. تم التخطي.")
                    progress.progress(int(((idx+1)/len(folder_paths))*100))
                    continue

                if image_order_option == "عشوائي":
                    random.shuffle(imgs)
                else:
                    imgs.sort()

                if mismatch_action == 'skip_folder' and len(imgs) != len(template_image_shapes):
                    if show_details:
                        st.info(f"تخطي المجلد '{folder_name}' بسبب اختلاف عدد الصور.")
                    progress.progress(int(((idx+1)/len(folder_paths))*100))
                    continue

                # استنساخ الشريحة القالب (يحفظ الهيكل تماماً)
                new_slide = duplicate_slide(prs, template_slide)
                created_slides += 1

                # الحصول على أشكال الصور في الشريحة الجديدة
                new_image_shapes = get_image_shapes(new_slide)

                # حسب اختيار التعامل مع mismatch
                replaced_count = 0
                for i, shape in enumerate(new_image_shapes):
                    if mismatch_action == 'truncate' and i >= len(imgs):
                        break

                    image_filename = imgs[i % len(imgs)] if mismatch_action == 'repeat' or i < len(imgs) else None
                    if not image_filename:
                        break

                    image_path = os.path.join(folder, image_filename)

                    # استبدال الصورة حسب نوع الشكل
                    try:
                        if hasattr(shape, 'is_placeholder') and shape.is_placeholder and hasattr(shape, 'placeholder_format') and shape.placeholder_format.type == PP_PLACEHOLDER.PICTURE:
                            # picture placeholder
                            shape.insert_picture(image_path)
                            replaced_count += 1
                        elif hasattr(shape, 'image'):
                            # شكل صورة عادي: نحذف الشكل ونضيف صورة جديدة بنفس الموضع
                            left, top, width, height = shape.left, shape.top, shape.width, shape.height
                            try:
                                shape._element.getparent().remove(shape._element)
                            except Exception:
                                # fallback: try to remove via shapes api
                                try:
                                    new_slide.shapes._spTree.remove(shape._element)
                                except Exception:
                                    pass
                            new_slide.shapes.add_picture(image_path, left, top, width, height)
                            replaced_count += 1
                        else:
                            # حالة احتياطية: نحاول إضافة الصورة بنفس موقع الشكل
                            left, top, width, height = getattr(shape, 'left', Inches(1)), getattr(shape, 'top', Inches(1)), getattr(shape, 'width', Inches(5)), getattr(shape, 'height', Inches(3))
                            try:
                                shape._element.getparent().remove(shape._element)
                            except Exception:
                                pass
                            new_slide.shapes.add_picture(image_path, left, top, width, height)
                            replaced_count += 1
                    except Exception as e:
                        if show_details:
                            st.warning(f"فشل استبدال صورة في شريحة '{folder_name}'. الخطأ: {e}")

                total_replaced += replaced_count

                # إضافة عنوان من اسم المجلد
                try:
                    title_placeholders = [s for s in new_slide.shapes if s.is_placeholder and hasattr(s, 'placeholder_format') and s.placeholder_format.type == PP_PLACEHOLDER.TITLE]
                    if title_placeholders:
                        title_placeholders[0].text = folder_name
                    else:
                        textbox = new_slide.shapes.add_textbox(Inches(1), Inches(0.3), Inches(8), Inches(0.6))
                        textbox.text_frame.text = folder_name
                except Exception:
                    pass

                if show_details:
                    st.success(f"تم إنشاء شريحة '{folder_name}' واستبدال {replaced_count} صورة.")

                progress.progress(int(((idx+1)/len(folder_paths))*100))

            progress.empty()
            status.empty()

            st.success("🎉 انتهت المعالجة.")
            st.markdown(f"- الشرائح المضافة: **{created_slides}**\n- الصور المستبدلة: **{total_replaced}**\n- المجلدات المعالجة: **{len(folder_paths)}**")

            # حفظ الملف للمستخدم
            output_buffer = io.BytesIO()
            original_name = os.path.splitext(uploaded_pptx.name)[0]
            output_filename = f"{original_name}_Updated.pptx"
            prs.save(output_buffer)
            output_buffer.seek(0)

            st.download_button("⬇️ تحميل العرض التقديمي المحدث", data=output_buffer.getvalue(), file_name=output_filename, mime="application/vnd.openxmlformats-officedocument.presentationml.presentation", use_container_width=True)

        except Exception as e:
            st.error(f"حدث خطأ أثناء المعالجة: {e}")
            if show_details:
                import traceback
                st.error(traceback.format_exc())
        finally:
            if temp_dir and os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)


if __name__ == '__main__':
    main()
