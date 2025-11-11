import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from PIL import Image
import os

# NEW: minimal helpers for URL support
import requests, io, tempfile, csv
from pathlib import Path

# ------------------------------------------------------------------------------
# Load data (UNCHANGED path — keep your Excel beside app.py)
# ------------------------------------------------------------------------------
data = pd.read_excel("all companys database.xlsx")

# Local fallback bases (original behavior kept)
IMAGE_BASE = "images"
LOGO_BASE = "static/logo"
FIRST_PATH = Path("static/img/first.png")  # <-- add
LAST_PATH  = Path("static/img/last.png")
# ------------------------------------------------------------------------------
# NEW: Optional manifest support (Cloudflare R2)
#   - If st.secrets.IMAGE_MANIFEST_URL is set, load from URL
#   - Else if image_manifest.csv exists locally, load it
#   - Else fall back to local folders under IMAGE_BASE
# ------------------------------------------------------------------------------
BASE_DIR = Path(__file__).parent
LOCAL_MANIFEST = BASE_DIR / "image_manifest.csv"

@st.cache_data(show_spinner=False)
def load_manifest():
    """
    Manifest CSV columns: Company,Product,Type,ImageURLs
    ImageURLs is '|' separated list of absolute URLs (R2 public links)
    """
    url = (st.secrets.get("IMAGE_MANIFEST_URL", "") or "").strip()
    rows = []
    try:
        if url:
            txt = requests.get(url, timeout=30).text.splitlines()
            reader = csv.DictReader(txt)
            rows = list(reader)
        elif LOCAL_MANIFEST.exists():
            with open(LOCAL_MANIFEST, newline="", encoding="utf-8") as f:
                reader = csv.DictReader(f)
                rows = list(reader)
    except Exception:
        rows = []

    manifest = {}
    for r in rows:
        c = (r.get("Company") or "").strip()
        p = (r.get("Product") or "").strip()
        t = (r.get("Type") or "").strip()
        urls = [u.strip() for u in (r.get("ImageURLs") or "").split("|") if u.strip()]
        if c and p and t and urls:
            manifest[(_norm(c), _norm(p), _norm(t))] = urls
    return manifest

MANIFEST = load_manifest()
from pathlib import Path

def _norm(s):
    # case-insensitive & collapses multiple spaces
    return " ".join(str(s or "").strip().split()).lower()


def show_image_safe(src):
    try:
        s = str(src).strip()
        if not s:
            st.caption("⚠️ Empty image reference")
            return
        if "://" in s:
            r = requests.get(s, timeout=30)
            r.raise_for_status()
            img = Image.open(io.BytesIO(r.content)).convert("RGB")
        else:
            img = Image.open(s).convert("RGB")
        st.image(img, use_column_width=True)
    except Exception as e:
        st.caption(f"⚠️ Failed to preview image ({src}): {e}")
# ------------------------------------------------------------------------------
# Session state (UNCHANGED)
# ------------------------------------------------------------------------------
if 'ppt_items' not in st.session_state:
    st.session_state.ppt_items = {}
if 'temp_selection' not in st.session_state:
    st.session_state.temp_selection = {}
if 'last_temp_key' not in st.session_state:
    st.session_state.last_temp_key = None

# ------------------------------------------------------------------------------
# Utility functions
# ------------------------------------------------------------------------------

def get_image_list(company, product, ptype):
    key_norm = (_norm(company), _norm(product), _norm(ptype))

    # Prefer manifest (URLs)
    if MANIFEST and key_norm in MANIFEST:
        return MANIFEST[key_norm]

    # Fallback: try a softer match (same company+product, type startswith)
    if MANIFEST:
        for (c, p, t), urls in MANIFEST.items():
            if c == key_norm[0] and p == key_norm[1] and t.startswith(key_norm[2][:6]):  # small tolerance
                return urls

    # Final fallback: local filesystem (original behavior)
    folder = os.path.join(IMAGE_BASE, company, product, ptype)
    images = []
    if os.path.exists(folder):
        for file in os.listdir(folder):
            if file.lower().endswith(('.jpg', '.jpeg', '.png', '.webp')):
                images.append(os.path.join(folder, file))
    return images


def get_scaled_dimensions(img, max_width, max_height):
    img_width_px, img_height_px = img.size
    aspect_ratio = img_width_px / img_height_px
    box_aspect = max_width / max_height
    if aspect_ratio > box_aspect:
        width = max_width
        height = width / aspect_ratio
    else:
        height = max_height
        width = height * aspect_ratio
    return width, height

# NEW: open image from URL or local path for dimension calculation
def open_pil_image(path_or_url):
    if "://" in str(path_or_url):
        resp = requests.get(path_or_url, timeout=60)
        resp.raise_for_status()
        return Image.open(io.BytesIO(resp.content))
    return Image.open(path_or_url)

# NEW: for python-pptx add_picture() which needs a path/stream; easiest is temp file
def fetch_to_tempfile(path_or_url):
    if "://" not in str(path_or_url):
        return path_or_url
    resp = requests.get(path_or_url, timeout=60)
    resp.raise_for_status()
    # infer ext from URL
    ext = ".png" if path_or_url.lower().endswith(".png") else ".jpg"
    fd, tpath = tempfile.mkstemp(suffix=ext)
    with os.fdopen(fd, "wb") as f:
        f.write(resp.content)
    return tpath

def create_beautiful_ppt(slide_data_list, include_intro_outro=True):
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]
    # keep your original first/last paths (you placed 'img/' in the repo)
    first_slide_path = str(FIRST_PATH)
    last_slide_path  = str(LAST_PATH)

    if include_intro_outro and os.path.exists(first_slide_path):
        slide = prs.slides.add_slide(blank)
        slide.shapes.add_picture(first_slide_path, Inches(0), Inches(0), width=prs.slide_width, height=prs.slide_height)

    for slide_data in slide_data_list:
        slide = prs.slides.add_slide(blank)
        company = slide_data.get('company', '')
        product = slide_data.get('product', '')
        link = slide_data.get('link', '')

        # Title (left aligned)
        title_shape = slide.shapes.add_textbox(Inches(0.4), Inches(0.3), Inches(12), Inches(0.7))
        frame = title_shape.text_frame
        frame.text = product
        p = frame.paragraphs[0]
        p.font.size = Pt(16)
        p.font.italic = True
        p.font.color.rgb = RGBColor(0, 0, 0)

        # Images
        y_img_top = 1.2
        y_img_bottom = 6.9
        max_img_height = y_img_bottom - y_img_top
        slide_width_in = prs.slide_width.inches
        imgs = slide_data['images']
        img_count = len(imgs)

        if img_count > 0:
            padding = 0.2
            columns = min(img_count, 3)
            rows = (img_count + columns - 1) // columns
            available_width = slide_width_in - (padding * (columns + 1))
            available_height = max_img_height - ((rows - 1) * padding)
            cell_width = available_width / columns
            cell_height = available_height / rows

            for i, img_src in enumerate(imgs):
                row = i // columns
                col = i % columns
                # open PIL image from URL or path (for sizing only)
                try:
                    with open_pil_image(img_src) as img:
                        img_width, img_height = get_scaled_dimensions(img, max_width=cell_width, max_height=cell_height)
                except Exception:
                    # If PIL fails, default fit box to avoid crash
                    img_width, img_height = cell_width, cell_height

                x = padding + col * (cell_width + padding) + (cell_width - img_width) / 2
                y = y_img_top + row * (cell_height + padding) + (cell_height - img_height) / 2
                # ensure python-pptx gets a local path
                add_path = fetch_to_tempfile(img_src)
                slide.shapes.add_picture(add_path, Inches(x), Inches(y), width=Inches(img_width), height=Inches(img_height))

        # Logo (original behavior — local 'logo/<Company>/logo.*')
        logo_dir = os.path.join(LOGO_BASE, company)
        logo_path = None
        for ext in ['png', 'jpg', 'jpeg']:
            candidate = os.path.join(logo_dir, f"logo.{ext}")
            if os.path.exists(candidate):
                logo_path = candidate
                break
        
        if logo_path:
            slide.shapes.add_picture(logo_path, prs.slide_width - Inches(1.2), Inches(0.1), width=Inches(1.1))

        cp_box = slide.shapes.add_textbox(prs.slide_width - Inches(3.6), prs.slide_height - Inches(0.3), Inches(3.6), Inches(0.4))
        cp_frame = cp_box.text_frame
        cp_frame.text = "Copyright © 2025 Altossa Projects LLp. All Rights Reserved."
        cp_para = cp_frame.paragraphs[0]
        cp_para.font.size = Pt(10)
        cp_para.font.color.rgb = RGBColor(128, 128, 128)

        if link:
            link_box = slide.shapes.add_textbox(Inches(0.1), prs.slide_height - Inches(0.3), Inches(7), Inches(0.4))
            link_frame = link_box.text_frame
            link_frame.text = str(link)
            p = link_frame.paragraphs[0]
            p.font.size = Pt(10)
            p.font.color.rgb = RGBColor(0, 102, 204)

    if include_intro_outro and os.path.exists(last_slide_path):
        slide = prs.slides.add_slide(blank)
        slide.shapes.add_picture(last_slide_path, Inches(0), Inches(0), width=prs.slide_width, height=prs.slide_height)

    return prs

# ------------------------------------------------------------------------------
# UI (UNCHANGED)
# ------------------------------------------------------------------------------
st.title("Product Selector with Search")
search_query = st.text_input("Search by Type", "")

if 'search_selection_keys' not in st.session_state:
    st.session_state.search_selection_keys = set()

if search_query:
    filtered_data = data[data['Type'].str.contains(search_query, case=False, na=False)]
    for idx, row in filtered_data.iterrows():
        company, product, ptype, link = row['Company'], row['Product'], row['Type'], row.get('Link', '')
        img_paths = get_image_list(company, product, ptype)

        st.markdown(f"### {product} - {ptype}")
        if len(img_paths) > 0:
            cols = st.columns(min(4, len(img_paths)))
            selected_imgs = []
            for i, path in enumerate(img_paths):
                with cols[i % len(cols)]:
                    # st.image supports both local paths and URLs
                    show_image_safe(path)
                    key = f"search_{company}_{product}_{ptype}_{i}".replace(" ", "_")
                    if st.checkbox("Include", key=key):
                        selected_imgs.append(path)
            if selected_imgs:
                st.session_state.ppt_items[f"{company}_{product}_{ptype}".replace(" ", "_")] = {
                    "company": company,
                    "product": product,
                    "link": link,
                    "images": selected_imgs
                }
else:
    company = st.selectbox("Select Company", sorted(data['Company'].dropna().unique()), key="company")
    products = sorted(data[data['Company'] == company]['Product'].dropna().unique())
    product = st.selectbox("Select Product", products, key="product")
    filtered_rows = data[(data['Company'] == company) & (data['Product'] == product)]

    current_base_key = f"{company}_{product}"
    if st.session_state.last_temp_key and st.session_state.last_temp_key != current_base_key:
        for k, v in st.session_state.temp_selection.items():
            if v['images']:
                st.session_state.ppt_items[f"{v['company']}_{v['product']}_{v['ptype']}"] = v
        st.session_state.temp_selection = {}
    st.session_state.last_temp_key = current_base_key

    for idx, row in filtered_rows.iterrows():
        ptype, link = row['Type'], row.get("Link", "")
        img_paths = get_image_list(company, product, ptype)

        st.markdown(f"### {ptype}")
        if len(img_paths) > 0:
            img_cols = st.columns(min(4, len(img_paths)))
            selected_imgs = st.session_state.temp_selection.get(f"{company}_{product}_{ptype}".replace(" ", "_"), {}).get("images", [])
            for i, path in enumerate(img_paths):
                with img_cols[i % len(img_cols)]:
                    show_image_safe(path)  # works for URLs too
                    key = f"manual_{company}_{product}_{ptype}_{i}".replace(" ", "_")
                    if st.checkbox("Include", key=key):
                        if path not in selected_imgs:
                            selected_imgs.append(path)
            st.session_state.temp_selection[f"{company}_{product}_{ptype}".replace(" ", "_")] = {
                "company": company,
                "product": product,
                "ptype": ptype,
                "link": link,
                "images": selected_imgs
            }
        else:
            st.info("No images found for this type.")

    for k, v in st.session_state.temp_selection.items():
        if v['images']:
            st.session_state.ppt_items[f"{v['company']}_{v['product']}_{v['ptype']}"] = v

# Sidebar generate PPTs (UNCHANGED)
with st.sidebar:
    st.markdown("## Ready to Download")

    if 'ppt_ready' not in st.session_state:
        st.session_state.ppt_ready = False
    if 'ppt_path' not in st.session_state:
        st.session_state.ppt_path = None

    if st.button("📦 Generate Combined PPT"):
        all_items = list(st.session_state.ppt_items.values())
        if all_items:
            prs = create_beautiful_ppt(all_items, include_intro_outro=True)
            ppt_path = "combined_presentation.pptx"
            prs.save(ppt_path)
            st.session_state.ppt_path = ppt_path
            st.session_state.ppt_ready = True
            st.success("PPT generated successfully!")
            # Clear selections after generation
            st.session_state.ppt_items = {}
            st.session_state.temp_selection = {}
            st.session_state.last_temp_key = None
        else:
            st.warning("No items selected for presentation!")

    if st.session_state.ppt_ready and st.session_state.ppt_path and os.path.exists(st.session_state.ppt_path):
        with open(st.session_state.ppt_path, "rb") as f:
            st.download_button(
                "Download Combined PPT",
                f,
                file_name="combined_presentation.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
