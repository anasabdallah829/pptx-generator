import streamlit as st
import zipfile
import os
import io
from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
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
    
    PICTURE_SHAPE_TYPES = (13, 21)
    
    picture_placeholders = [
        shape for shape in first_slide.shapes
        if shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.PICTURE
    ]
    regular_pictures = [
        shape for shape in first_slide.shapes
        if hasattr(shape, 'shape_type') and shape.shape_type in PICTURE_SHAPE_TYPES
    ]
    
    total_image_slots = len(picture_placeholders) + len(regular_pictures)

    return True, {
        'placeholders': len(picture_placeholders),
        'regular_pictures': len(regular_pictures),
        'total_slots': total_image_slots,
        'slide_layout': first_slide.slide_layout
    }


def get_image_shapes(slide):
    """
    استخراج جميع أشكال الصور من الشريحة، سواء كانت placeholders أو صور عادية.
    """
    PICTURE_SHAPE_TYPES = (13, 21)
    
    image_shapes = []
    for shape in slide.shapes:
        if shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.PICTURE:
            image_shapes.append(shape)
        elif hasattr(shape, 'shape_type') and shape.shape_type in PICTURE_SHAPE_TYPES:
            image_shapes.append(shape)
            
    image_shapes.sort(key=lambda s: (s.top, s.left))
    return image_shapes


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
                folder_paths = [os.path.join(temp_dir, item) for item in all_items if os.path.isdir(os.path.join(temp_dir, item))]
                
                if not folder_paths:
                    st.error("❌ لا توجد مجلدات في الملف المضغوط.")
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
                template_image_shapes = get_image_shapes(first_slide)
                
                if not template_image_shapes:
                    st.warning("⚠ الشريحة الأولى لا تحتوي على مواضع صور. سيتم إضافة الصورة الأولى من كل مجلد فقط.")
                    slide_layout = prs.slide_layouts[6]
                else:
                    slide_layout = analysis_result['slide_layout']

                mismatch_folders = []
                for fp in folder_paths:
                    imgs = [f for f in os.listdir(fp) if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.webp'))]
                    if len(imgs) != len(template_image_shapes):
                        mismatch_folders.append((os.path.basename(fp), len(imgs), len(template_image_shapes)))
                
                if mismatch_folders and 'mismatch_action' not in st.session_state:
                    with st.form("mismatch_form"):
                        st.warning("⚠ تم اكتشاف اختلاف في عدد الصور لبعض المجلدات مقارنة بعدد مواضع الصور في الشريحة الأولى.")
                        for name, img_count, _ in mismatch_folders:
                            st.write(f"- المجلد `{name}` يحتوي على {img_count} صورة.")
                        st.markdown(f"**عدد مواضع الصور في القالب: {len(template_image_shapes)}**")

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

                    imgs = [f for f in os.listdir(folder_path) if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.webp'))]
                    
                    if not imgs:
                        if show_details:
                            st.warning(f"⚠ المجلد {folder_name} فارغ من الصور، سيتم تخطيه.")
                        continue
                    
                    # ترتيب الصور بناءً على اختيار المستخدم
                    if image_order_option == "عشوائي":
                        random.shuffle(imgs)
                    else:
                        imgs.sort()

                    if mismatch_action == 'skip_folder' and len(imgs) != len(template_image_shapes):
                        if show_details:
                            st.info(f"ℹ تم تخطي المجلد {folder_name} لوجود اختلاف في عدد الصور.")
                        continue

                    new_slide = prs.slides.add_slide(slide_layout)
                    created_slides += 1
                    
                    new_image_shapes = get_image_shapes(new_slide)
                    
                    replaced_count = 0
                    for i, new_shape in enumerate(new_image_shapes):
                        if mismatch_action == 'truncate' and i >= len(imgs):
                            break
                        
                        image_filename = imgs[i % len(imgs)]
                        image_path = os.path.join(folder_path, image_filename)
                        
                        try:
                            new_shape.insert_picture(image_path)
                            replaced_count += 1
                        except AttributeError:
                            left, top, width, height = new_shape.left, new_shape.top, new_shape.width, new_shape.height
                            new_shape.element.getparent().remove(new_shape.element)
                            new_slide.shapes.add_picture(
                                image_path, left, top, width, height
                            )
                            replaced_count += 1
                            
                    total_replaced += replaced_count
                    
                    try:
                        title_shapes = [shape for shape in new_slide.shapes
                                        if shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.TITLE]
                        if title_shapes:
                            title_shapes[0].text = folder_name
                        else:
                            textbox = new_slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
                            text_frame = textbox.text_frame
                            text_frame.text = folder_name
                    except Exception:
                        pass
                    
                    if show_details:
                        st.success(f"✅ تم إنشاء شريحة للمجلد '{folder_name}' واستبدال {replaced_count} صورة")

                    progress_bar.progress((folder_idx + 1) / len(folder_paths))

                progress_bar.empty()
                status_text.empty()

                st.success("🎉 تم الانتهاء من المعالجة!")
                if 'mismatch_action' in st.session_state: del st.session_state['mismatch_action']
                if 'process_started' in st.session_state: del st.session_state['process_started']

                col1, col2, col3 = st.columns(3)
                with col1: st.metric("الشرائح المُضافة", created_slides)
                with col2: st.metric("الصور المُستبدلة", total_replaced)
                with col3: st.metric("المجلدات المُعالجة", len(folder_paths))

                if created_slides == 0:
                    st.error("❌ لم يتم إضافة أي شرائح.")
                    st.stop()

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
            
if __name__ == '__main__':
    main()
