# app.py
# Altossa PPT Generator — Streamlit (Option A small assets in repo + product images via R2 URLs)
# - 16:9 (Widescreen) PPT
# - One slide per product (Title = product type; Link = product URL; One image per slide)
# - Uses tiny assets from repo: static/img/first.png, static/img/last.png, static/logo/<Company>/logo.png
# - Product images are read from a manifest CSV (URLs hosted on Cloudflare R2 public bucket)
# - Optional: Upload finished PPT to Cloudflare via Worker (set secrets in Streamlit Cloud)

import io
import os
import csv
import tempfile
from pathlib import Path
from typing import Dict, List, Tuple

import requests
import streamlit as st
import pandas as pd
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt

# --------------------------------------------------------------------------------------
# CONFIG (edit if needed)
# --------------------------------------------------------------------------------------

BASE_DIR = Path(__file__).parent
STATIC_DIR = BASE_DIR / "static"
FIRST_PATH = STATIC_DIR / "img" / "first.png"   # small file in repo
LAST_PATH  = STATIC_DIR / "img" / "last.png"    # small file in repo

# Manifest can be a local CSV in repo (default) OR a URL (hosted on R2)
# If you host the manifest on R2, set IMAGE_MANIFEST_URL in Streamlit Secrets
LOCAL_MANIFEST_PATH = BASE_DIR / "image_manifest.csv"
MANIFEST_URL = st.secrets.get("IMAGE_MANIFEST_URL", "").strip()

# Optional: Excel can be local or remote (if you need it). Leave as None if not used.
EXCEL_URL = st.secrets.get("EXCEL_URL", "").strip() or None
LOCAL_EXCEL_PATH = BASE_DIR / "data" / "all_companys_database.xlsx"

# Optional (only if you want to upload PPTs to Cloudflare R2 via Worker and get a shareable link)
WORKER_UPLOAD_URL = st.secrets.get("WORKER_UPLOAD_URL", "").strip()  # e.g., https://altossa-ppt-worker.<acct>.workers.dev/upload
UPLOAD_TOKEN = st.secrets.get("UPLOAD_TOKEN", "").strip()            # same token stored as Worker secret

# Slide style
SLIDE_TITLE_FONT_PT = 28
SLIDE_LINK_FONT_PT = 14
IMG_W, IMG_H = 11.3, 5.1  # inches (contain-fit area)

# --------------------------------------------------------------------------------------
# UTILITIES
# --------------------------------------------------------------------------------------

@st.cache_data(show_spinner=False)
def load_manifest() -> Dict[Tuple[str, str, str], List[str]]:
    """
    Manifest format (CSV):
    Company,Product,Type,ImageURLs
    Ditre Italia,Alta Sofa,sofa,https://img.domain/images/Ditre%20Italia/Alta%20Sofa/sofa/1.jpg|https://.../2.png

    Returns: dict keyed by (Company, Product, Type) -> [image_url1, image_url2, ...]
    """
    rows = []
    if MANIFEST_URL:
        try:
            text = requests.get(MANIFEST_URL, timeout=30).text
            lines = text.splitlines()
            reader = csv.DictReader(lines)
            rows = list(reader)
        except Exception as e:
            st.warning(f"Failed to load remote manifest: {e}")
    elif LOCAL_MANIFEST_PATH.exists():
        with open(LOCAL_MANIFEST_PATH, newline="", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            rows = list(reader)
    else:
        return {}

    manifest = {}
    for row in rows:
        company = (row.get("Company") or "").strip()
        product = (row.get("Product") or "").strip()
        ptype   = (row.get("Type") or "").strip()
        urls = [u.strip() for u in (row.get("ImageURLs") or "").split("|") if u.strip()]
        if company and product and ptype and urls:
            manifest[(company, product, ptype)] = urls
    return manifest


def list_companies_products_types(manifest: Dict[Tuple[str, str, str], List[str]]):
    companies = sorted({k[0] for k in manifest.keys()})
    products_by_company = {}
    types_by_company_product = {}
    for (c, p, t) in manifest.keys():
        products_by_company.setdefault(c, set()).add(p)
        types_by_company_product.setdefault((c, p), set()).add(t)
    # Convert sets to sorted lists
    products_by_company = {c: sorted(list(v)) for c, v in products_by_company.items()}
    types_by_company_product = {(c, p): sorted(list(v)) for (c, p), v in types_by_company_product.items()}
    return companies, products_by_company, types_by_company_product


def fetch_to_tempfile(url_or_path: str) -> str:
    """
    If path is local, return as is. If URL, download to a temp file and return path.
    """
    if "://" not in url_or_path:
        return url_or_path
    r = requests.get(url_or_path, timeout=60)
    r.raise_for_status()
    content_type = r.headers.get("content-type", "").lower()
    ext = ".png" if "png" in content_type else ".jpg"
    fd, path = tempfile.mkstemp(suffix=ext)
    with os.fdopen(fd, "wb") as f:
        f.write(r.content)
    return path


def logo_path_for_company(company: str) -> str | None:
    """
    Look for a tiny logo in static/logo/<Company>/logo.(png|jpg|jpeg)
    """
    if not company:
        return None
    safe = str(company).strip()
    for ext in ("png", "jpg", "jpeg"):
        p = STATIC_DIR / "logo" / safe / f"logo.{ext}"
        if p.exists():
            return str(p)
    return None


def add_title_and_link(slide, title: str, link_url: str):
    left = Inches(0.5)
    title_top = Inches(0.3)
    link_top = Inches(1.2)
    width = Inches(12.3)
    height_title = Inches(0.8)
    height_link = Inches(0.5)

    title_shape = slide.shapes.add_textbox(left, title_top, width, height_title)
    p = title_shape.text_frame
    p.clear()
    run = p.paragraphs[0].add_run()
    run.text = title
    font = run.font
    font.size = Pt(SLIDE_TITLE_FONT_PT)
    font.bold = True

    if link_url:
        link_shape = slide.shapes.add_textbox(left, link_top, width, height_link)
        p2 = link_shape.text_frame
        p2.clear()
        run2 = p2.paragraphs[0].add_run()
        run2.text = link_url
        font2 = run2.font
        font2.size = Pt(SLIDE_LINK_FONT_PT)
        font2.color.rgb = None  # leave default; PPT will show theme color
        # Hyperlink
        try:
            run2.hyperlink.address = link_url
        except Exception:
            pass


def add_image_contain(slide, image_path: str):
    """
    Place image to fit inside a defined box (contain). Tuned for clean layout.
    """
    x = Inches(1.0)
    y = Inches(1.9)
    w = Inches(IMG_W)
    h = Inches(IMG_H)
    slide.shapes.add_picture(image_path, x, y, width=w, height=h)


def build_ppt(slides_data: List[dict], ppt_name: str, include_first_last: bool = True) -> str:
    """
    slides_data: list of dicts:
      {
        "product_type": str,
        "product_link": str,
        "image_url": str,
        "company": str | None
      }

    Returns: local path to saved PPTX
    """
    prs = Presentation()
    # Set 16:9
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]  # blank

    # First slide (from repo tiny image)
    if include_first_last and FIRST_PATH.exists():
        s = prs.slides.add_slide(blank)
        fp = fetch_to_tempfile(str(FIRST_PATH))
        try:
            add_image_contain(s, fp)
        except Exception:
            pass

    # Content slides
    for item in slides_data:
        s = prs.slides.add_slide(blank)
        product_type = item.get("product_type", "").strip() or "Product"
        product_link = item.get("product_link", "").strip()
        img_url      = item.get("image_url", "").strip()
        company      = item.get("company", "").strip()

        add_title_and_link(s, product_type, product_link)

        # Try to add logo (tiny, top-right)
        lp = logo_path_for_company(company)
        if lp:
            try:
                s.shapes.add_picture(lp, prs.slide_width - Inches(1.7), Inches(0.2), width=Inches(1.5))
            except Exception:
                pass

        # Add product image
        if img_url:
            try:
                local_img = fetch_to_tempfile(img_url)
                add_image_contain(s, local_img)
            except Exception as e:
                # Continue even if one image fails
                print("Image add error:", e)

    # Last slide (from repo tiny image)
    if include_first_last and LAST_PATH.exists():
        s = prs.slides.add_slide(blank)
        lp = fetch_to_tempfile(str(LAST_PATH))
        try:
            add_image_contain(s, lp)
        except Exception:
            pass

    out_name = ppt_name or "Altossa_Selection.pptx"
    out_path = str((BASE_DIR / out_name).resolve())
    prs.save(out_path)
    return out_path


def upload_ppt_to_r2_worker(local_ppt_path: str, final_filename: str) -> str:
    """
    Optional: Upload PPT bytes to Cloudflare Worker -> R2
    Returns: public URL (string)
    """
    if not WORKER_UPLOAD_URL or not UPLOAD_TOKEN:
        raise RuntimeError("Worker upload is not configured. Set WORKER_UPLOAD_URL and UPLOAD_TOKEN in secrets.")
    with open(local_ppt_path, "rb") as f:
        r = requests.post(
            f"{WORKER_UPLOAD_URL}?filename={final_filename}",
            headers={
                "x-upload-token": UPLOAD_TOKEN,
                "content-type": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            },
            data=f.read(),
            timeout=120
        )
    r.raise_for_status()
    data = r.json()
    if not data.get("ok"):
        raise RuntimeError(f"Upload failed: {data}")
    return data["url"]


# --------------------------------------------------------------------------------------
# UI
# --------------------------------------------------------------------------------------

st.set_page_config(page_title="Altossa PPT Generator", page_icon="📑", layout="centered")
st.title("Altossa — Fast PPT Generator (16:9)")

st.markdown(
    "One image per slide • Title shows **Product Type** • Link added as clickable text • "
    "Logos + First/Last slides from repo • Product images from **R2 URLs via manifest**"
)

# Load optional Excel (if you need to show product table or reference)
if EXCEL_URL:
    try:
        data_df = pd.read_excel(EXCEL_URL)
    except Exception as e:
        data_df = None
        st.info(f"Could not load Excel from URL. ({e})")
else:
    if LOCAL_EXCEL_PATH.exists():
        try:
            data_df = pd.read_excel(LOCAL_EXCEL_PATH)
        except Exception as e:
            data_df = None
            st.info(f"Could not load local Excel. ({e})")
    else:
        data_df = None

with st.expander("Optional: View data table (Excel)"):
    if data_df is not None:
        st.dataframe(data_df, use_container_width=True)
    else:
        st.caption("No Excel loaded.")

# Manifest / selection
manifest = load_manifest()
if not manifest:
    st.warning("No image manifest found. Add image_manifest.csv to repo or set IMAGE_MANIFEST_URL in secrets.")
else:
    companies, products_by_company, types_by_company_product = list_companies_products_types(manifest)

# Session storage for selected slides
if "slides" not in st.session_state:
    st.session_state.slides = []

st.subheader("Add Slides")

colA, colB = st.columns(2)
with colA:
    product_type = st.text_input("Product Type (slide title)", placeholder="e.g., sofa, armchair, table")
with colB:
    product_link = st.text_input("Product Link (clickable URL)", placeholder="https://brand.com/product")

if manifest:
    c1, c2 = st.columns(2)
    with c1:
        company = st.selectbox("Company", [""] + companies, index=0)
    product = ""
    ptype = ""
    img_choices = []
    if company:
        prods = products_by_company.get(company, [])
        with c2:
            product = st.selectbox("Product", [""] + prods, index=0)
        if product:
            types = types_by_company_product.get((company, product), [])
            ptype = st.selectbox("Type", [""] + types, index=0)
            if ptype:
                urls = manifest.get((company, product, ptype), [])
                st.caption(f"{len(urls)} image(s) available for this selection.")
                if urls:
                    img_choices = st.multiselect("Choose image(s) to add as slides", options=urls, default=urls[:1])
else:
    company = st.text_input("Company (for logo lookup)", placeholder="e.g., Ditre Italia")
    product = st.text_input("Product (optional)", placeholder="e.g., Alta Sofa")
    ptype   = st.text_input("Type (optional)", placeholder="e.g., sofa")
    img_url_manual = st.text_input("Image URL (R2 public URL)", placeholder="https://img.altossa.xyz/images/...")

add_cols = st.columns(3)
with add_cols[0]:
    if st.button("Add Slide"):
        if not product_type:
            st.error("Please fill Product Type.")
        elif not product_link:
            st.error("Please fill Product Link.")
        else:
            if manifest:
                if ptype and img_choices:
                    for u in img_choices:
                        st.session_state.slides.append({
                            "product_type": product_type.strip(),
                            "product_link": product_link.strip(),
                            "image_url": u.strip(),
                            "company": company.strip()
                        })
                    st.success(f"Added {len(img_choices)} slide(s).")
                else:
                    st.error("Select Company, Product, Type and choose at least one image.")
            else:
                if img_url_manual:
                    st.session_state.slides.append({
                        "product_type": product_type.strip(),
                        "product_link": product_link.strip(),
                        "image_url": img_url_manual.strip(),
                        "company": company.strip()
                    })
                    st.success("Added 1 slide.")
                else:
                    st.error("Enter an Image URL.")
with add_cols[1]:
    if st.button("Clear List"):
        st.session_state.slides = []
        st.info("Cleared.")
with add_cols[2]:
    include_first_last = st.checkbox("Include first & last slides", value=True)

# Show current slides
st.write("### Slides in queue")
if st.session_state.slides:
    for i, s in enumerate(st.session_state.slides, start=1):
        st.markdown(
            f"**{i}. {s['product_type']}**  \n"
            f"<span style='font-size:12px;color:#666'>{s['product_link']}</span>  \n"
            f"<span style='font-size:12px;color:#666'>{s['image_url']}</span>  \n"
            f"<span style='font-size:12px;color:#666'>Company: {s.get('company') or '-'}</span>",
            unsafe_allow_html=True
        )
else:
    st.caption("No slides yet.")

st.write("---")
ppt_name = st.text_input("Final PPT filename", value="Altossa_Selection.pptx")
gen_cols = st.columns(3)
with gen_cols[0]:
    if st.button("Generate PPT (16:9)"):
        if not st.session_state.slides:
            st.error("No slides to generate.")
        else:
            try:
                out_path = build_ppt(st.session_state.slides, ppt_name=ppt_name, include_first_last=include_first_last)
                st.session_state["ppt_path"] = out_path
                st.session_state["ppt_ready"] = True
                st.success("PPT generated.")
            except Exception as e:
                st.session_state["ppt_ready"] = False
                st.error(f"PPT generation failed: {e}")

if st.session_state.get("ppt_ready") and st.session_state.get("ppt_path") and os.path.exists(st.session_state["ppt_path"]):
    with open(st.session_state["ppt_path"], "rb") as f:
        st.download_button(
            "⬇ Download Combined PPT",
            f,
            file_name=ppt_name or "Altossa_Selection.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

    with gen_cols[1]:
        if WORKER_UPLOAD_URL and UPLOAD_TOKEN:
            if st.button("☁ Upload to Cloudflare (get shareable link)"):
                try:
                    url = upload_ppt_to_r2_worker(st.session_state["ppt_path"], ppt_name or "Altossa_Selection.pptx")
                    st.success("Uploaded successfully.")
                    st.markdown(f"**Your PPT link:** {url}")
                    st.markdown(f"[Open PPT]({url})")
                except Exception as e:
                    st.error(f"Upload failed: {e}")
        else:
            st.caption("Set WORKER_UPLOAD_URL and UPLOAD_TOKEN in secrets to enable Cloudflare upload.")
