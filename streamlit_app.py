import streamlit as st
import zipfile
import os
import io
from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER, MSO_SHAPE_TYPE

def get_image_shapes(slide):
    """إرجاع قائمة بكل أشكال الصور في الشريحة مرتبة بمواقعها"""
    image_shapes = []
    for shape in slide.shapes:
        if shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.PICTURE:
            image_shapes.append(shape)
        elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            image_shapes.append(shape)
    image_shapes.sort(key=lambda s: (s.top, s.left))
    return image_shapes


st.title("📑 مزامنة الصور مع الشرائح")

uploaded_pptx = st.file_uploader("📂 ارفع ملف PowerPoint", type=["pptx"])
uploaded_zip = st.file_uploader("🖼️ ارفع ملف الصور (ZIP)", type=["zip"])
mismatch_action = st.selectbox("📏 عند اختلاف عدد الصور عن الشرائح:", ["truncate", "repeat"])
show_details = st.checkbox("إظهار تفاصيل المعالجة", value=True)

if uploaded_pptx and uploaded_zip:
    pptx_bytes = uploaded_pptx.read()
    prs = Presentation(io.BytesIO(pptx_bytes))
    
    zip_bytes = uploaded_zip.read()
    with zipfile.ZipFile(io.BytesIO(zip_bytes), "r") as zip_ref:
        temp_dir = "temp_images"
        os.makedirs(temp_dir, exist_ok=True)
        zip_ref.extractall(temp_dir)

    replaced_count = 0

    for idx, slide in enumerate(prs.slides):
        folder_name = f"slide{idx + 1}"
        folder_path = os.path.join(temp_dir, folder_name)

        if not os.path.exists(folder_path):
            continue

        imgs = [f for f in os.listdir(folder_path) if f.lower().endswith((".png", ".jpg", ".jpeg"))]
        imgs.sort()

        new_image_shapes = get_image_shapes(slide)

        for i, new_shape in enumerate(new_image_shapes):
            if mismatch_action == 'truncate' and i >= len(imgs):
                break

            image_filename = imgs[i % len(imgs)]
            image_path = os.path.join(folder_path, image_filename)

            try:
                if new_shape.is_placeholder and new_shape.placeholder_format.type == PP_PLACEHOLDER.PICTURE:
                    new_shape.insert_picture(image_path)
                else:
                    left, top, width, height = new_shape.left, new_shape.top, new_shape.width, new_shape.height
                    new_shape.element.getparent().remove(new_shape.element)
                    slide.shapes.add_picture(image_path, left, top, width, height)
                replaced_count += 1
            except Exception as e:
                if show_details:
                    st.error(f"❌ خطأ أثناء استبدال الصورة {image_filename} في الشريحة {idx+1}: {e}")

    output_stream = io.BytesIO()
    prs.save(output_stream)
    output_stream.seek(0)

    st.success(f"✅ تم استبدال {replaced_count} صورة بنجاح!")
    st.download_button(
        label="📥 تحميل العرض المعدل",
        data=output_stream,
        file_name="presentation_updated.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
