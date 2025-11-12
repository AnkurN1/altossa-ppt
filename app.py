import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from PIL import Image
import os

# Load data
data = pd.read_excel("all companys database.xlsx")
IMAGE_BASE = "images"
LOGO_BASE = "logo"

if 'ppt_items' not in st.session_state:
    st.session_state.ppt_items = {}
if 'temp_selection' not in st.session_state:
    st.session_state.temp_selection = {}
if 'last_temp_key' not in st.session_state:
    st.session_state.last_temp_key = None

# Utility functions
def get_image_list(company, product, ptype):
    folder = os.path.join(IMAGE_BASE, str(company).strip(), str(product).strip(), str(ptype).strip())
    images = []
    if os.path.exists(folder):
        for file in os.listdir(folder):
            if file.lower().endswith(('.jpg', '.jpeg', '.png')):
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

def create_beautiful_ppt(slide_data_list, include_intro_outro=True):
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]
    first_slide_path = "img/first.png"
    last_slide_path = "img/last.png"

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

            for i, img_path in enumerate(imgs):
                row = i // columns
                col = i % columns
                with Image.open(img_path) as img:
                    img_width, img_height = get_scaled_dimensions(img, max_width=cell_width, max_height=cell_height)
                    x = padding + col * (cell_width + padding) + (cell_width - img_width) / 2
                    y = y_img_top + row * (cell_height + padding) + (cell_height - img_height) / 2
                    slide.shapes.add_picture(img_path, Inches(x), Inches(y), width=Inches(img_width), height=Inches(img_height))

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

# UI
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
                    st.image(path, use_container_width=True)
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
                    st.image(path, use_container_width=True)
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

# Sidebar generate PPTs
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
            st.download_button("Download Combined PPT", f, file_name="combined_presentation.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
