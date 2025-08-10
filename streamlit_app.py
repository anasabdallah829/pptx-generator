import streamlit as st
import os
import zipfile
import io
from pptx import Presentation
from pptx.util import Inches

def process_files(zip_file, pptx_file):
    """
    Processes the uploaded ZIP and PPTX files to create a new presentation.

    - Extracts folder names from the ZIP file.
    - Creates a new title slide for each folder in the existing PPTX file.
    
    :param zip_file: A Streamlit UploadedFile object for the ZIP file.
    :param pptx_file: A Streamlit UploadedFile object for the PPTX file.
    :return: A BytesIO object of the modified PPTX file, or None if an error occurs.
    """
    try:
        # Load the PowerPoint presentation from the uploaded file object
        prs = Presentation(pptx_file)
        
        # Read the contents of the ZIP file into a BytesIO object
        zip_content = io.BytesIO(zip_file.read())
        
        # Extract unique top-level folder names from the ZIP file
        folder_names = set()
        with zipfile.ZipFile(zip_content, 'r') as zip_ref:
            for member in zip_ref.infolist():
                # Split the path to get the top-level folder name
                path_parts = member.filename.split(os.sep)
                # Ensure the path is not empty and is a directory
                if path_parts and path_parts[0] and member.is_dir():
                    folder_names.add(path_parts[0])

        # Add a new title slide for each folder found in the ZIP
        # We sort the folder names to ensure a consistent slide order
        for folder_name in sorted(list(folder_names)):
            slide_layout = prs.slide_layouts[5] # Using the "Title Only" layout
            slide = prs.slides.add_slide(slide_layout)
            title = slide.shapes.title
            title.text = f"Folder: {folder_name}"

        # Save the modified presentation to an in-memory buffer
        output_stream = io.BytesIO()
        prs.save(output_stream)
        output_stream.seek(0)
        
        return output_stream
        
    except Exception as e:
        st.error(f"An unexpected error occurred during processing: {e}")
        return None

# --- Streamlit Application UI ---

# Set up the page configuration
st.set_page_config(page_title="PowerPoint Folder Processor", layout="centered")
st.title("PowerPoint Folder Processor 📁")
st.markdown("---")

st.write(
    "Upload a **ZIP file** containing folders and an existing **PowerPoint file (.pptx)**. "
    "This tool will create a new slide for each top-level folder in the ZIP file and append it to the presentation."
)

# File upload widgets
zip_file_upload = st.file_uploader("1. Upload your ZIP file:", type=["zip"])
pptx_file_upload = st.file_uploader("2. Upload your PowerPoint (.pptx) file:", type=["pptx"])

# Processing button
if st.button("Process and Generate Presentation"):
    if zip_file_upload is not None and pptx_file_upload is not None:
        with st.spinner("Processing files and generating new presentation... 🔄"):
            modified_pptx_stream = process_files(zip_file_upload, pptx_file_upload)
            
            if modified_pptx_stream:
                st.success("Processing complete! Your presentation is ready for download. 🎉")
                # Download button for the modified file
                st.download_button(
                    label="Download Modified PPTX",
                    data=modified_pptx_stream,
                    file_name="modified_presentation.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
            else:
                st.error("Failed to process files. Please check the file formats and contents and try again.")
    else:
        st.warning("Please upload both a ZIP and a PPTX file to proceed.")
