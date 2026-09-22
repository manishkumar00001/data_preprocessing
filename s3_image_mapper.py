# """
# S3 Image URL Mapper
# -------------------
# Upload a product CSV / Excel file, and this tool fills the `images` column with
# S3 URLs built from each product's SKU (e.g. <base-url>/UK40204.jpg).

# Run:
#     pip install -r requirements.txt
#     streamlit run s3_image_mapper.py
# """
# import io
# from concurrent.futures import ThreadPoolExecutor, as_completed
# from pathlib import Path
# from urllib.parse import quote

# import pandas as pd
# import requests
# import streamlit as st
# from openpyxl import load_workbook

# DEFAULT_BASE_URL = "https://design247-ecommerce.s3.ap-south-1.amazonaws.com/products/"
# EXT_CHOICES = ["jpg", "jpeg", "png", "webp"]


# # --------------------------------------------------------------------------- #
# # Core logic
# # --------------------------------------------------------------------------- #
# def build_url(base_url: str, key: str, ext: str) -> str:
#     return f"{base_url.rstrip('/')}/{quote(key)}.{ext}"


# def http_status(url: str, timeout: int = 8) -> int:
#     """HTTP status of a HEAD request (0 = network error / timeout)."""
#     try:
#         return requests.head(url, timeout=timeout, allow_redirects=True).status_code
#     except requests.RequestException:
#         return 0


# def resolve_row(sku, existing, base_url, exts, verify, overwrite, check_existing):
#     """Return (new_image_value, status) for one product row."""
#     sku, existing = (sku or "").strip(), (existing or "").strip()

#     if not sku:
#         return existing, "skipped: empty SKU"

#     # Keep the URL that's already there
#     if existing and not overwrite:
#         if check_existing:
#             code = http_status(existing)
#             return existing, "existing OK" if code == 200 else f"existing BROKEN (HTTP {code})"
#         return existing, "existing kept"

#     # Build a new URL from the SKU
#     if not verify:
#         return build_url(base_url, sku, exts[0]), "filled (not verified)"

#     last_code = 0
#     for ext in exts:
#         url = build_url(base_url, sku, ext)
#         last_code = http_status(url)
#         if last_code == 200:
#             return url, "filled"
#     return existing, f"NOT FOUND in S3 (HTTP {last_code})"


# def process(df, sku_col, img_col, base_url, exts, verify, overwrite, check_existing,
#             workers=16, progress_cb=None):
#     n = len(df)
#     skus = df[sku_col].tolist()
#     existing = df[img_col].tolist() if img_col in df.columns else [""] * n
#     results = [None] * n
#     done = 0
#     with ThreadPoolExecutor(max_workers=workers) as pool:
#         futures = {
#             pool.submit(resolve_row, skus[i], existing[i], base_url, exts,
#                         verify, overwrite, check_existing): i
#             for i in range(n)
#         }
#         for fut in as_completed(futures):
#             results[futures[fut]] = fut.result()
#             done += 1
#             if progress_cb:
#                 progress_cb(done / n)
#     values = [r[0] for r in results]
#     report = pd.DataFrame({
#         "row": [i + 2 for i in range(n)],  # spreadsheet row number
#         "sku": skus,
#         "image_url": values,
#         "status": [r[1] for r in results],
#     })
#     return values, report


# # --------------------------------------------------------------------------- #
# # File I/O
# # --------------------------------------------------------------------------- #
# def read_csv(raw: bytes) -> pd.DataFrame:
#     for enc in ("utf-8-sig", "latin-1"):
#         try:
#             return pd.read_csv(io.BytesIO(raw), dtype=str, keep_default_na=False, encoding=enc)
#         except UnicodeDecodeError:
#             continue
#     raise ValueError("Could not decode CSV file.")


# def read_excel(raw: bytes, sheet: str) -> pd.DataFrame:
#     return pd.read_excel(io.BytesIO(raw), sheet_name=sheet, dtype=str, keep_default_na=False)


# def write_excel(raw: bytes, sheet: str, img_col: str, values: list) -> bytes:
#     """Edit only the images column in the original workbook, so other sheets,
#     formatting and column widths are preserved."""
#     wb = load_workbook(io.BytesIO(raw))
#     ws = wb[sheet]
#     headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
#     if img_col in headers:
#         col = headers.index(img_col) + 1
#     else:
#         col = len(headers) + 1
#         ws.cell(row=1, column=col, value=img_col)
#     for i, v in enumerate(values, start=2):
#         ws.cell(row=i, column=col, value=v if v else None)
#     out = io.BytesIO()
#     wb.save(out)
#     return out.getvalue()


# def write_csv(df: pd.DataFrame, img_col: str, values: list) -> bytes:
#     out = df.copy()
#     out[img_col] = values
#     return out.to_csv(index=False).encode("utf-8-sig")  # BOM so Excel opens it cleanly


# # --------------------------------------------------------------------------- #
# # UI
# # --------------------------------------------------------------------------- #
# def main():
#     st.set_page_config(page_title="S3 Image URL Mapper", page_icon="🖼️", layout="wide")
#     st.title("🖼️ S3 Image URL Mapper")
#     st.caption("Fill the `images` column of your product sheet with S3 URLs built from the SKU.")

#     # ---- Sidebar settings
#     with st.sidebar:
#         st.header("Settings")
#         base_url = st.text_input("S3 / CloudFront base URL", DEFAULT_BASE_URL,
#                                  help="Folder URL where the images live. The SKU + extension is appended.")
#         exts = st.multiselect("Image extensions to try (in order)", EXT_CHOICES, default=["jpg"],
#                               help="With verification on, each extension is tried until one exists in S3.")
#         verify = st.checkbox("Verify that each image exists in S3", value=True,
#                              help="Sends a HEAD request per URL. Needs the images to be publicly readable.")
#         overwrite = st.checkbox("Overwrite URLs that already exist", value=False,
#                                 help="Off = only blank cells are filled.")
#         check_existing = st.checkbox("Also verify existing URLs", value=True,
#                                      disabled=overwrite)
#         workers = st.slider("Parallel requests", 1, 32, 16)

#     if not exts:
#         st.warning("Select at least one image extension in the sidebar.")
#         st.stop()

#     # ---- Input
#     uploaded = st.file_uploader("Upload your product file", type=["csv", "xlsx"])
#     if not uploaded:
#         st.info("Upload a .csv or .xlsx file to begin.")
#         return

#     raw = uploaded.getvalue()
#     is_excel = uploaded.name.lower().endswith(".xlsx")
#     sheet = None

#     try:
#         if is_excel:
#             sheets = load_workbook(io.BytesIO(raw), read_only=True).sheetnames
#             default_idx = next((i for i, s in enumerate(sheets) if "product" in s.lower()), 0)
#             sheet = st.selectbox("Sheet", sheets, index=default_idx)
#             df = read_excel(raw, sheet)
#         else:
#             df = read_csv(raw)
#     except Exception as e:
#         st.error(f"Could not read the file: {e}")
#         return

#     df.columns = [str(c).strip() for c in df.columns]
#     st.write(f"**{len(df)} rows**, {len(df.columns)} columns")
#     st.dataframe(df.head(10), use_container_width=True)

#     c1, c2 = st.columns(2)
#     sku_default = next((i for i, c in enumerate(df.columns) if c.lower() == "sku"), 0)
#     sku_col = c1.selectbox("SKU column (used as the image filename)", df.columns, index=sku_default)
#     img_default = next((c for c in df.columns if c.lower() in ("images", "image", "image_url")), "images")
#     img_col = c2.text_input("Images column (created if missing)", img_default)

#     dup = df[sku_col].str.strip().replace("", pd.NA).dropna().duplicated().sum()
#     if dup:
#         st.warning(f"{dup} duplicate SKU(s) found. Those rows will point to the same image.")

#     # ---- Run
#     if st.button("Generate image URLs", type="primary"):
#         bar = st.progress(0.0, text="Working...")
#         values, report = process(
#             df, sku_col, img_col.strip(), base_url, exts, verify, overwrite,
#             check_existing and not overwrite, workers,
#             progress_cb=lambda p: bar.progress(p, text=f"Working... {int(p * 100)}%"),
#         )
#         bar.empty()

#         stem, suffix = Path(uploaded.name).stem, ".xlsx" if is_excel else ".csv"
#         out_bytes = (write_excel(raw, sheet, img_col.strip(), values) if is_excel
#                      else write_csv(df, img_col.strip(), values))
#         st.session_state["result"] = {
#             "report": report,
#             "bytes": out_bytes,
#             "name": f"{stem}_with_images{suffix}",
#             "mime": ("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#                      if is_excel else "text/csv"),
#         }

#     # ---- Output (kept in session_state so it survives the download-button rerun)
#     res = st.session_state.get("result")
#     if res:
#         report = res["report"]
#         counts = report["status"].str.replace(r"\s*\(HTTP.*\)", "", regex=True).value_counts()
#         st.subheader("Result")
#         cols = st.columns(max(len(counts), 1))
#         for col, (label, n) in zip(cols, counts.items()):
#             col.metric(label, int(n))

#         problems = report[report["status"].str.contains("NOT FOUND|BROKEN|skipped", regex=True)]
#         if len(problems):
#             st.error(f"{len(problems)} row(s) need attention:")
#             st.dataframe(problems, use_container_width=True, hide_index=True)
#         else:
#             st.success("All rows have a working image URL.")

#         with st.expander("Full report"):
#             st.dataframe(report, use_container_width=True, hide_index=True)

#         d1, d2 = st.columns(2)
#         d1.download_button("⬇️ Download updated file", res["bytes"], res["name"], res["mime"],
#                            type="primary", use_container_width=True)
#         d2.download_button("⬇️ Download report (CSV)", report.to_csv(index=False).encode("utf-8-sig"),
#                            "image_mapping_report.csv", "text/csv", use_container_width=True)


# if __name__ == "__main__":
#     main()



# ---------------------------------------v2


"""
S3 Image URL Mapper
-------------------
Upload a product CSV / Excel file, and this tool fills the `images` column with
S3 URLs built from each product's SKU (e.g. <base-url>/UK40204.jpg).

Run:
    pip install -r requirements.txt
    streamlit run s3_image_mapper.py
"""
import io
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from urllib.parse import quote

import pandas as pd
import requests
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font

DEFAULT_BASE_URL = "https://design247-ecommerce.s3.ap-south-1.amazonaws.com/products/"
EXT_CHOICES = ["jpg", "jpeg", "png", "webp"]


# --------------------------------------------------------------------------- #
# Core logic
# --------------------------------------------------------------------------- #
def build_url(base_url: str, key: str, ext: str) -> str:
    return f"{base_url.rstrip('/')}/{quote(key)}.{ext}"


def http_status(url: str, timeout: int = 8) -> int:
    """HTTP status of a HEAD request (0 = network error / timeout)."""
    try:
        return requests.head(url, timeout=timeout, allow_redirects=True).status_code
    except requests.RequestException:
        return 0


def resolve_row(sku, existing, base_url, exts, verify, overwrite, check_existing):
    """Return (new_image_value, status) for one product row."""
    sku, existing = (sku or "").strip(), (existing or "").strip()

    if not sku:
        return existing, "skipped: empty SKU"

    # Keep the URL that's already there
    if existing and not overwrite:
        if check_existing:
            code = http_status(existing)
            return existing, "existing OK" if code == 200 else f"existing BROKEN (HTTP {code})"
        return existing, "existing kept"

    # Build a new URL from the SKU
    if not verify:
        return build_url(base_url, sku, exts[0]), "filled (not verified)"

    last_code = 0
    for ext in exts:
        url = build_url(base_url, sku, ext)
        last_code = http_status(url)
        if last_code == 200:
            return url, "filled"
    return existing, f"NOT FOUND in S3 (HTTP {last_code})"


def process(df, sku_col, img_col, base_url, exts, verify, overwrite, check_existing,
            workers=16, progress_cb=None):
    n = len(df)
    skus = df[sku_col].tolist()
    existing = df[img_col].tolist() if img_col in df.columns else [""] * n
    results = [None] * n
    done = 0
    with ThreadPoolExecutor(max_workers=workers) as pool:
        futures = {
            pool.submit(resolve_row, skus[i], existing[i], base_url, exts,
                        verify, overwrite, check_existing): i
            for i in range(n)
        }
        for fut in as_completed(futures):
            results[futures[fut]] = fut.result()
            done += 1
            if progress_cb:
                progress_cb(done / n)
    values = [r[0] for r in results]
    report = pd.DataFrame({
        "row": [i + 2 for i in range(n)],  # spreadsheet row number
        "sku": skus,
        "image_url": values,
        "status": [r[1] for r in results],
    })
    return values, report


# --------------------------------------------------------------------------- #
# File I/O
# --------------------------------------------------------------------------- #
def read_csv(raw: bytes) -> pd.DataFrame:
    for enc in ("utf-8-sig", "latin-1"):
        try:
            return pd.read_csv(io.BytesIO(raw), dtype=str, keep_default_na=False, encoding=enc)
        except UnicodeDecodeError:
            continue
    raise ValueError("Could not decode CSV file.")


def read_excel(raw: bytes, sheet: str) -> pd.DataFrame:
    return pd.read_excel(io.BytesIO(raw), sheet_name=sheet, dtype=str, keep_default_na=False)


def write_excel(raw: bytes, sheet: str, img_col: str, values: list) -> bytes:
    """Edit only the images column in the original workbook, so other sheets,
    formatting and column widths are preserved."""
    wb = load_workbook(io.BytesIO(raw))
    ws = wb[sheet]
    headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
    if img_col in headers:
        col = headers.index(img_col) + 1
    else:
        col = len(headers) + 1
        ws.cell(row=1, column=col, value=img_col)
    link_font = Font(color="0563C1", underline="single")
    for i, v in enumerate(values, start=2):
        cell = ws.cell(row=i, column=col)
        if v:
            cell.value = v
            cell.hyperlink = v
            cell.font = link_font
        else:
            cell.value = None
            cell.hyperlink = None
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def write_csv(df: pd.DataFrame, img_col: str, values: list) -> bytes:
    out = df.copy()
    out[img_col] = values
    return out.to_csv(index=False).encode("utf-8-sig")  # BOM so Excel opens it cleanly


# --------------------------------------------------------------------------- #
# UI
# --------------------------------------------------------------------------- #
def main():
    st.set_page_config(page_title="S3 Image URL Mapper", page_icon="🖼️", layout="wide")
    st.title("🖼️ S3 Image URL Mapper")
    st.caption("Fill the `images` column of your product sheet with S3 URLs built from the SKU.")

    # ---- Sidebar settings
    with st.sidebar:
        st.header("Settings")
        base_url = st.text_input("S3 / CloudFront base URL", DEFAULT_BASE_URL,
                                 help="Folder URL where the images live. The SKU + extension is appended.")
        exts = st.multiselect("Image extensions to try (in order)", EXT_CHOICES, default=["jpg"],
                              help="With verification on, each extension is tried until one exists in S3.")
        verify = st.checkbox("Verify that each image exists in S3", value=True,
                             help="Sends a HEAD request per URL. Needs the images to be publicly readable.")
        overwrite = st.checkbox("Overwrite URLs that already exist", value=False,
                                help="Off = only blank cells are filled.")
        check_existing = st.checkbox("Also verify existing URLs", value=True,
                                     disabled=overwrite)
        workers = st.slider("Parallel requests", 1, 32, 16)

    if not exts:
        st.warning("Select at least one image extension in the sidebar.")
        st.stop()

    # ---- Input
    uploaded = st.file_uploader("Upload your product file", type=["csv", "xlsx"])
    if not uploaded:
        st.info("Upload a .csv or .xlsx file to begin.")
        return

    raw = uploaded.getvalue()
    is_excel = uploaded.name.lower().endswith(".xlsx")
    sheet = None

    try:
        if is_excel:
            sheets = load_workbook(io.BytesIO(raw), read_only=True).sheetnames
            default_idx = next((i for i, s in enumerate(sheets) if "product" in s.lower()), 0)
            sheet = st.selectbox("Sheet", sheets, index=default_idx)
            df = read_excel(raw, sheet)
        else:
            df = read_csv(raw)
    except Exception as e:
        st.error(f"Could not read the file: {e}")
        return

    df.columns = [str(c).strip() for c in df.columns]
    st.write(f"**{len(df)} rows**, {len(df.columns)} columns")
    st.dataframe(df.head(10), use_container_width=True)

    c1, c2 = st.columns(2)
    sku_default = next((i for i, c in enumerate(df.columns) if c.lower() == "sku"), 0)
    sku_col = c1.selectbox("SKU column (used as the image filename)", df.columns, index=sku_default)
    img_default = next((c for c in df.columns if c.lower() in ("images", "image", "image_url")), "images")
    img_col = c2.text_input("Images column (created if missing)", img_default)

    dup = df[sku_col].str.strip().replace("", pd.NA).dropna().duplicated().sum()
    if dup:
        st.warning(f"{dup} duplicate SKU(s) found. Those rows will point to the same image.")

    # ---- Run
    if st.button("Generate image URLs", type="primary"):
        bar = st.progress(0.0, text="Working...")
        values, report = process(
            df, sku_col, img_col.strip(), base_url, exts, verify, overwrite,
            check_existing and not overwrite, workers,
            progress_cb=lambda p: bar.progress(p, text=f"Working... {int(p * 100)}%"),
        )
        bar.empty()

        stem, suffix = Path(uploaded.name).stem, ".xlsx" if is_excel else ".csv"
        out_bytes = (write_excel(raw, sheet, img_col.strip(), values) if is_excel
                     else write_csv(df, img_col.strip(), values))
        st.session_state["result"] = {
            "report": report,
            "bytes": out_bytes,
            "name": f"{stem}_with_images{suffix}",
            "mime": ("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                     if is_excel else "text/csv"),
        }

    # ---- Output (kept in session_state so it survives the download-button rerun)
    res = st.session_state.get("result")
    if res:
        report = res["report"]
        counts = report["status"].str.replace(r"\s*\(HTTP.*\)", "", regex=True).value_counts()
        st.subheader("Result")
        cols = st.columns(max(len(counts), 1))
        for col, (label, n) in zip(cols, counts.items()):
            col.metric(label, int(n))

        problems = report[report["status"].str.contains("NOT FOUND|BROKEN|skipped", regex=True)]
        if len(problems):
            st.error(f"{len(problems)} row(s) need attention:")
            st.dataframe(problems, use_container_width=True, hide_index=True)
        else:
            st.success("All rows have a working image URL.")

        with st.expander("Full report"):
            st.dataframe(report, use_container_width=True, hide_index=True)

        d1, d2 = st.columns(2)
        d1.download_button("⬇️ Download updated file", res["bytes"], res["name"], res["mime"],
                           type="primary", use_container_width=True)
        d2.download_button("⬇️ Download report (CSV)", report.to_csv(index=False).encode("utf-8-sig"),
                           "image_mapping_report.csv", "text/csv", use_container_width=True)


if __name__ == "__main__":
    main()