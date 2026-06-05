"""IMH Sale Sticker Generator — upload a sale-list Excel, download a labels PPT."""

import datetime as dt
import html

import pandas as pd
import streamlit as st

from label_generator import (
    parse_excel, generate_ppt, fit_label_sizes, fit_note_size, format_price,
)

DEFAULT_NOTE = "No Return, No Exchange on Sale Items — All Sales Final"

# ---------------------------------------------------------------------------
# HTML preview that mirrors the PPT layout (10" x 7.5" page, 4 labels)
# ---------------------------------------------------------------------------

PAGE_W_PX = 660           # rendered page width
PX_PER_PT = PAGE_W_PX / 10 / 72  # 10 inches wide, 72pt per inch

LABEL_STYLE = (
    "position:absolute;width:49%;height:44%;border:2px solid #000;"
    "display:flex;flex-direction:column;justify-content:center;align-items:center;"
    "text-align:center;font-family:'Times New Roman',serif;font-weight:bold;"
    "color:#000;overflow:hidden;padding:0 4px;box-sizing:border-box;"
)
LABEL_POS = [(1.0, 1.33), (50.0, 1.33), (1.0, 46.67), (50.0, 46.67)]


def _px(pt):
    return f"{pt * PX_PER_PT * 72 / 72:.0f}px"


def label_html(label, exp_text, note):
    small_lines = (1 if exp_text else 0) + (1 if note else 0)
    brand_pt, name_pt, price_pt = fit_label_sizes(label, small_lines)
    unit_pt = max(18, round(price_pt * 32 / 72))
    parts = []
    if label["brand"]:
        parts.append(
            f'<div style="font-size:{_px(brand_pt)};line-height:1.2;">{html.escape(label["brand"])}</div>'
        )
    if label["name"]:
        parts.append(
            f'<div style="font-size:{_px(name_pt)};line-height:1.2;">{html.escape(label["name"])}</div>'
        )
    qty = f'<span style="font-size:{_px(price_pt)};">/</span>' \
          f'<span style="font-size:{_px(unit_pt)};">{html.escape(label["qty"])}</span>' if label["qty"] else ""
    parts.append(
        f'<div style="line-height:1.2;"><span style="font-size:{_px(unit_pt)};">$ </span>'
        f'<span style="font-size:{_px(price_pt)};">{format_price(label["price"])}</span>{qty}</div>'
    )
    if exp_text:
        parts.append(
            f'<div style="font-size:{_px(12)};color:#f00;line-height:1.4;">{html.escape(exp_text)}</div>'
        )
    if note:
        parts.append(
            f'<div style="font-size:{_px(fit_note_size(note))};color:#f00;line-height:1.4;'
            f'white-space:nowrap;">{html.escape(note)}</div>'
        )
    return "".join(parts)


def preview_html(labels, exp_text, note):
    page_h = int(PAGE_W_PX * 7.5 / 10)
    pages = [labels[i:i + 4] for i in range(0, len(labels), 4)]
    out = [
        '<div style="display:inline-flex;max-height:640px;overflow-y:auto;'
        'flex-direction:column;gap:14px;max-width:100%;">'
    ]
    for pg_num, page in enumerate(pages, 1):
        out.append(
            f'<div style="position:relative;width:{PAGE_W_PX}px;height:{page_h}px;'
            'background:#fff;flex-shrink:0;border:1px solid #aaa;'
            'box-shadow:0 1px 4px rgba(0,0,0,.4);">'
        )
        for (left, top), label in zip(LABEL_POS, page):
            out.append(
                f'<div style="{LABEL_STYLE}left:{left}%;top:{top}%;">{label_html(label, exp_text, note)}</div>'
            )
        out.append(
            f'<div style="position:absolute;right:6px;bottom:2px;font-size:11px;'
            f'color:#999;font-family:sans-serif;">page {pg_num}/{len(pages)}</div></div>'
        )
    out.append("</div>")
    return "".join(out)

st.set_page_config(page_title="Sale Sticker Generator", page_icon="🏷️", layout="wide")

st.title("🏷️ Sale Sticker Generator")
st.caption(
    "Upload the sale-list Excel (columns: item, price). Blank rows and 'ADD' "
    "noise are ignored. Review/fix the parsed labels below, then download the PPT "
    "(4 labels per page)."
)

uploaded = st.file_uploader("Sale list (.xlsx)", type=["xlsx", "xlsm"])

if uploaded is None:
    st.info("Upload an Excel file to get started.")
    st.stop()

try:
    labels = parse_excel(uploaded)
except Exception as e:
    st.error(f"Could not read the Excel file: {e}")
    st.stop()

if not labels:
    st.warning("No valid item/price rows found in this file.")
    st.stop()

st.success(f"Parsed **{len(labels)}** items.")

col1, col2 = st.columns([1, 2])
with col1:
    exp_date = st.date_input(
        "Sale expiry date (shown on each label)",
        value=dt.date.today() + dt.timedelta(days=7),
        format="DD/MM/YYYY",
    )
    include_exp = st.checkbox("Include 'Sale Exp' line", value=True)
with col2:
    note = st.text_input(
        "Note (shown under the expiry date — keep it short, it shrinks to stay on one line)",
        value=DEFAULT_NOTE,
        max_chars=70,
    ).strip()

st.subheader("Review labels")
st.caption(
    "Edit any cell to fix parsing (e.g. spelling, brand split, quantity). "
    "Untick **Include** to skip an item."
)

df = pd.DataFrame(labels)
df.insert(0, "include", True)

edited = st.data_editor(
    df,
    width="stretch",
    hide_index=True,
    num_rows="dynamic",
    column_config={
        "include": st.column_config.CheckboxColumn("Include", width="small"),
        "brand": st.column_config.TextColumn("Brand (line 1)"),
        "name": st.column_config.TextColumn("Product name (line 2)"),
        "price": st.column_config.NumberColumn("Price ($)", min_value=0.0, step=0.01),
        "qty": st.column_config.TextColumn("Qty / unit (after price)"),
    },
)

final = edited[edited["include"] & edited["price"].notna()]
final_labels = final.drop(columns=["include"]).fillna("").to_dict("records")

st.write(f"**{len(final_labels)}** labels → **{(len(final_labels) + 3) // 4}** pages")

if final_labels:
    effective_exp = exp_date if include_exp else None
    exp_text = (
        f"Sale Exp: {effective_exp.day}-{effective_exp.strftime('%B-%Y')}"
        if effective_exp else ""
    )

    ppt_buf = generate_ppt(final_labels, effective_exp, note)
    st.download_button(
        "⬇️ Download labels PPT",
        data=ppt_buf,
        file_name=f"sale_labels_{dt.date.today():%Y%m%d}.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        type="primary",
    )

    st.subheader("Preview")
    st.caption(
        "Exact replica of the PPT pages. To open the real file in Google Slides: "
        "download it, then drag it into [Google Drive](https://drive.google.com) "
        "and double-click it."
    )
    st.html(preview_html(final_labels, exp_text, note))
