import streamlit as st
import zipfile
import os
import io
from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
import shutil
from pptx.util import Inches
import random

# Set Streamlit page configuration
st.set_page_config(page_title="Slide-Sync-Images", layout="centered", initial_sidebar_state="expanded")

# Custom CSS for a modern, elegant design
st.markdown("""
<style>
    .stApp {
        background-color: #f0f2f6;
        color: #1a1a1a;
    }
    .main-header {
        text-align: center;
        font-size: 2.5em;
        font-weight: 700;
        color: #004d99;
        margin-bottom: 0.5em;
    }
    .sub-header {
        text-align: center;
        font-size: 1.2em;
        color: #666;
        margin-bottom: 2em;
    }
    .st-emotion-cache-1kyx11f {
        background-color: #f8f9fa;
        border-radius: 10px;
        padding: 20px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        margin-bottom: 20px;
    }
    .stButton>button {
        background-color: #004d99;
        color: white;
        border-radius: 8px;
        border: none;
        padding: 12px 24px;
        font-size: 1.1em;
        transition: background-color 0.3s;
    }
    .stButton>button:hover {
        background-color: #003366;
    }
    .st-emotion-cache-1g88h6 {
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05);
        border: 1px solid #e0e0e0;
    }
    .st-emotion-cache-163m3e9 {
        padding: 1rem 1rem 1rem 1rem;
    }
    .st-emotion-cache-1kyx11f > div:first-child > h3 {
        color: #004d99;
        font-weight: 600;
        border-bottom: 2px solid #e0e0e0;
        padding-bottom: 10px;
        margin-bottom: 20px;
    }
    .metric-container {
        padding: 15px;
        border-radius: 8px;
        background-color: #e6f7ff;
        border: 1px solid #b3e0ff;
        text-align: center;
    }
    .metric-label {
        font-size: 1em;
        color: #333;
        font-weight: 600;
    }
    .metric-value {
        font-size: 1.8em;
        font-weight: 700;
        color: #004d99;
    }
    .sidebar-header {
        color: #004d99;
        font-weight: 600;
        border-bottom: 2px solid #e0e0e0;
        padding-bottom: 10px;
        margin-bottom: 20px;
    }
    .st-emotion-cache-v063l {
      text-align: right;
    }
    .st-emotion-cache-h601 {
      direction: rtl;
    }
</style>
""", unsafe_allow_html=True)

# --- App Header and Description ---
st.markdown('<h1 class="main-header">🔄 Slide-Sync-Images</h1>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">أداة سهلة وسريعة لإنشاء شرائح PowerPoint جديدة من قالب ومجلدات صور.</p>', unsafe_allow_html=True)

# --- Main Interface ---

st.subheader("📂 تحميل الملفات")
uploaded_pptx = st.file_uploader("اختر ملف PowerPoint للقالب (.pptx)", type=["pptx"], key="pptx_uploader")
uploaded_zip = st.file_uploader("اختر ملف ZIP يحتوي على مجلدات الصور", type=["zip"], key="zip_uploader")

st.markdown("---")

st.subheader("⚙️ إعدادات المعالجة")
image_order_option = st.radio(
    "كيف تريد ترتيب الصور في الشرائح؟",
    ("بالترتيب (افتراضي)", "عشوائي"),
    index=0
)

show_details = st.checkbox("عرض التفاصيل المفصلة", value=False)

st.markdown("---")

def analyze_first_slide(prs):
    """
    تحليل الشريحة الأولى: إرجاع نتائج حتى لو لم توجد مواضع للصور.
    """
    if len(prs.slides) == 0:
        return False, "لا توجد شرائح في الملف."

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
    استخراج وترتيب جميع أشكال الصور من الشريحة، سواء كانت placeholders أو صور عادية.
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

        if st.button("🚀 بدء المعالجة", use_container_width=True) or st.session_state.process_started:
            st.session_state.process_started = True
            
            temp_dir = None
            try:
                with st.spinner("📦 جاري فحص واستخراج الملفات..."):
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
                    st.error("❌ ملف ZIP لا يحتوي على أي مجلدات صور.")
                    st.stop()
                
                folder_paths.sort()
                st.success(f"✅ تم العثور على **{len(folder_paths)}** مجلد صور للمعالجة.")

                prs = Presentation(io.BytesIO(uploaded_pptx.read()))
                
                st.info("🔍 جاري تحليل الشريحة الأولى...")
                ok, analysis_result = analyze_first_slide(prs)
                if not ok:
                    st.error(f"❌ خطأ: {analysis_result}")
                    st.stop()
                
                st.success("✅ تم الانتهاء من تحليل القالب.")
                col1, col2, col3 = st.columns(3)
                with col1: st.markdown(f'<div class="metric-container"><div class="metric-label">عدد placeholders</div><div class="metric-value">{analysis_result["placeholders"]}</div></div>', unsafe_allow_html=True)
                with col2: st.markdown(f'<div class="metric-container"><div class="metric-label">عدد الصور العادية</div><div class="metric-value">{analysis_result["regular_pictures"]}</div></div>', unsafe_allow_html=True)
                with col3: st.markdown(f'<div class="metric-container"><div class="metric-label">إجمالي أماكن الصور</div><div class="metric-value">{analysis_result["total_slots"]}</div></div>', unsafe_allow_html=True)
                
                st.markdown("---")
                
                first_slide = prs.slides[0]
                template_image_shapes = get_image_shapes(first_slide)
                
                if not template_image_shapes:
                    st.warning("⚠ لا توجد أماكن للصور في القالب. سيتم إضافة صورة واحدة لكل شريحة.")
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
                        st.warning("⚠️ تم اكتشاف اختلاف في عدد الصور لبعض المجلدات مقارنة بأماكن الصور في القالب.")
                        for name, img_count, _ in mismatch_folders:
                            st.write(f"- المجلد `{name}` يحتوي على {img_count} صورة.")
                        st.markdown(f"**عدد أماكن الصور في القالب: {len(template_image_shapes)}**")

                        choice_text = st.radio(
                            "اختر كيف تريد التعامل مع المجلدات التي يختلف عدد صورها:",
                            ("اقتصاص (استبدال فقط حتى أقل عدد)", "تكرار (ملء كل الأماكن بتكرار الصور)", "تخطي (تجاهل المجلدات التي بها اختلاف)", "إيقاف (إلغاء العملية بالكامل)"),
                            index=0
                        )
                        submit_choice = st.form_submit_button("✅ تأكيد المتابعة")

                    if submit_choice:
                        st.session_state['mismatch_action'] = {
                            "اقتصاص (استبدال فقط حتى أقل عدد)": 'truncate',
                            "تكرار (ملء كل الأماكن بتكرار الصور)": 'repeat',
                            "تخطي (تجاهل المجلدات التي بها اختلاف)": 'skip_folder',
                            "إيقاف (إلغاء العملية بالكامل)": 'stop'
                        }.get(choice_text)
                    else:
                        st.stop()
                
                if 'mismatch_action' in st.session_state:
                    mismatch_action = st.session_state['mismatch_action']
                else:
                    mismatch_action = 'truncate'

                if mismatch_action == 'stop':
                    st.error("❌ تم إلغاء العملية بناءً على اختيارك.")
                    st.stop()

                st.info("🔄 جاري إنشاء الشرائح الجديدة...")
                total_replaced = 0
                created_slides = 0

                progress_bar = st.progress(0)
                status_text = st.empty()

                for folder_idx, folder_path in enumerate(folder_paths):
                    folder_name = os.path.basename(folder_path)
                    status_text.text(f"جاري معالجة المجلد {folder_idx + 1}/{len(folder_paths)}: **{folder_name}**")

                    imgs = [f for f in os.listdir(folder_path) if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.webp'))]
                    
                    if not imgs:
                        if show_details:
                            st.warning(f"⚠ المجلد '{folder_name}' فارغ من الصور. تم التخطي.")
                        continue
                    
                    if image_order_option == "عشوائي":
                        random.shuffle(imgs)
                    else:
                        imgs.sort()

                    if mismatch_action == 'skip_folder' and len(imgs) != len(template_image_shapes):
                        if show_details:
                            st.info(f"ℹ تم تخطي المجلد '{folder_name}' بسبب اختلاف عدد الصور.")
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
                        st.success(f"✅ تم إنشاء شريحة للمجلد '{folder_name}' واستبدال {replaced_count} صورة.")

                    progress_bar.progress((folder_idx + 1) / len(folder_paths))

                progress_bar.empty()
                status_text.empty()
                
                st.markdown("---")
                st.success("🎉 **تم الانتهاء من المعالجة بنجاح!**")
                
                col1, col2, col3 = st.columns(3)
                with col1: st.markdown(f'<div class="metric-container"><div class="metric-label">الشرائح المضافة</div><div class="metric-value">{created_slides}</div></div>', unsafe_allow_html=True)
                with col2: st.markdown(f'<div class="metric-container"><div class="metric-label">الصور المستبدلة</div><div class="metric-value">{total_replaced}</div></div>', unsafe_allow_html=True)
                with col3: st.markdown(f'<div class="metric-container"><div class="metric-label">المجلدات التي تمت معالجتها</div><div class="metric-value">{len(folder_paths)}</div></div>', unsafe_allow_html=True)


                if created_slides == 0:
                    st.error("❌ لم يتم إضافة أي شرائح إلى العرض التقديمي.")
                    st.stop()

                original_name = os.path.splitext(uploaded_pptx.name)[0]
                output_filename = f"{original_name}_Updated.pptx"
                output_buffer = io.BytesIO()
                prs.save(output_buffer)
                output_buffer.seek(0)

                st.download_button(
                    label="⬇️ تحميل العرض التقديمي المحدث",
                    data=output_buffer.getvalue(),
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True
                )

            except Exception as e:
                st.error(f"❌ حدث خطأ أثناء المعالجة: {e}")
                if show_details:
                    import traceback
                    st.error(f"تفاصيل الخطأ: {traceback.format_exc()}")
            finally:
                if temp_dir and os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)
    else:
        st.info("👋 مرحباً! قم بتحميل ملف PowerPoint والقالب المضغوط للبدء.")
        
        st.sidebar.markdown('<h3 class="sidebar-header">📖 تعليمات</h3>', unsafe_allow_html=True)
        st.sidebar.markdown("""
        **1. ملف PowerPoint (.pptx):**
        - يجب أن يحتوي على شريحة واحدة على الأقل.
        - سيتم استخدام الشريحة الأولى كقالب.

        **2. ملف ZIP:**
        - يجب أن يحتوي على مجلدات، وكل مجلد يضم الصور المخصصة لشريحة واحدة.
        - سيتم استخدام أسماء المجلدات كعناوين للشرائح.
        """)


if __name__ == '__main__':
    main()
