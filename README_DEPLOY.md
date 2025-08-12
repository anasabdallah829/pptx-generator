# Slide Sync - Streamlit PPTX Image Replacer

## Files
- streamlit_app.py
- requirements.txt

## Deploy locally
1. Create venv:
   python -m venv venv
   source venv/bin/activate  # or venv\Scripts\activate on Windows
2. Install deps:
   pip install -r requirements.txt
3. Run:
   streamlit run streamlit_app.py

## Deploy to Streamlit Cloud
1. Create a public GitHub repository and push these files.
2. Go to https://streamlit.io/cloud and connect your GitHub account.
3. Create a new app and select the repository & main file `streamlit_app.py`.
4. Deploy.

Notes:
- Keep uploaded PPTX and ZIP sizes reasonable (Streamlit Cloud upload limits).
- If you need custom fonts or large assets, consider using an external storage or hosting.
