# IMH Sale Sticker Generator

Streamlit app: upload a sale-list Excel (two columns — item, price) and download a
PowerPoint of sale labels, 4 per page.

## How it works

- Blank rows and noise rows (e.g. `ADD`) are ignored automatically.
- Each item is split into **brand** (line 1), **product name** (line 2) and
  **quantity/unit** (taken from the text after the `-`, e.g. `-500 GM`), shown
  after the price as `$ 8.99/500 gm`.
- A red `Sale Exp: <date>` line is added (date picked in the app, can be turned off).
- Parsed labels appear in an editable table — fix spellings/quantities and untick
  items before downloading.

## Run locally

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Files

- `streamlit_app.py` — the Streamlit UI
- `label_generator.py` — Excel parsing + PPT generation
- `test_fit.py` — regression test: label content always fits inside the box
  (`python test_fit.py`)

