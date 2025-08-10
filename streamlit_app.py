from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import os

def generate_pptx_from_template(template_path, folders_path, output_path):
    prs_template = Presentation(template_path)
    base_slide = prs_template.slides[0]

    # Get positions of images in the template
    image_shapes = [shape for shape in base_slide.shapes if shape.shape_type == MSO_SHAPE_TYPE.PICTURE]
    image_positions = [(shape.left, shape.top, shape.height) for shape in image_shapes]

    # Create new presentation
    prs = Presentation()
    prs.slide_width = prs_template.slide_width
    prs.slide_height = prs_template.slide_height

    folders = sorted([f for f in os.listdir(folders_path) if os.path.isdir(os.path.join(folders_path, f))])

    for folder in folders:
        images = sorted([
            os.path.join(folders_path, folder, f)
            for f in os.listdir(os.path.join(folders_path, folder))
            if f.lower().endswith((".jpg", ".jpeg", ".png"))
        ])

        slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank layout

        # Add title (folder name)
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(1))
        text_frame = title_box.text_frame
        p = text_frame.paragraphs[0]
        run = p.add_run()
        run.text = folder
        font = run.font
        font.size = Pt(28)
        font.bold = True
        font.color.rgb = RGBColor(0, 0, 128)

        # Insert images
        for idx, (left, top, height) in enumerate(image_positions):
            if idx < len(images):
                slide.shapes.add_picture(images[idx], left, top, height=height)

    prs.save(output_path)
