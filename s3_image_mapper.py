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
# from openpyxl.styles import Font

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
#     link_font = Font(color="0563C1", underline="single")
#     for i, v in enumerate(values, start=2):
#         cell = ws.cell(row=i, column=col)
#         if v:
#             cell.value = v
#             cell.hyperlink = v
#             cell.font = link_font
#         else:
#             cell.value = None
#             cell.hyperlink = None
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


# --------------------------------multi images


"""
S3 Image URL Mapper
-------------------
Upload a product CSV / Excel file, and this tool fills in S3 image URLs
built from each product's SKU.

Two modes:
  - Single image per product: one URL per row, e.g. <base>/SKU.jpg
  - Multiple images (count column): a column holds how many images a
    product has (e.g. subImage = 8). Builds SKU_1.jpg .. SKU_8.jpg and
    writes each into its own column (images_1, images_2, ...) so every
    one of them is a real clickable hyperlink in Excel. (A single Excel
    cell can only carry one hyperlink, so a comma-separated list in one
    cell can never be more than plain text — this is an Excel limit,
    not something the script can work around.)

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
LINK_FONT = Font(color="0563C1", underline="single")


# --------------------------------------------------------------------------- #
# Core logic — single image per product
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

    if existing and not overwrite:
        if check_existing:
            code = http_status(existing)
            return existing, "existing OK" if code == 200 else f"existing BROKEN (HTTP {code})"
        return existing, "existing kept"

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
        "row": [i + 2 for i in range(n)],
        "sku": skus,
        "image_url": values,
        "status": [r[1] for r in results],
    })
    return values, report


# --------------------------------------------------------------------------- #
# Core logic — multiple images per product (count column)
# --------------------------------------------------------------------------- #
def resolve_row_multi(sku, count_raw, base_url, ext, verify):
    """Build <sku>_1.<ext> .. <sku>_N.<ext> for a product with N images.

    Returns (list_of_urls, status). With verify on, only URLs that
    return HTTP 200 are kept (in order); the status reports how many
    of the N were found.
    """
    sku = (sku or "").strip()
    count_raw = (count_raw or "").strip()

    if not sku:
        return [], "skipped: empty SKU"
    if not count_raw:
        return [], "skipped: empty count"
    try:
        n = int(float(count_raw))
    except ValueError:
        return [], f"skipped: count '{count_raw}' is not a number"
    if n <= 0:
        return [], "skipped: count is 0"

    urls = [build_url(base_url, f"{sku}_{i}", ext) for i in range(1, n + 1)]

    if not verify:
        return urls, f"filled (not verified, {n} images)"

    found = [u for u in urls if http_status(u) == 200]
    if len(found) == n:
        return found, f"filled ({n}/{n} found)"
    if found:
        return found, f"PARTIAL ({len(found)}/{n} found)"
    return [], f"NOT FOUND in S3 (0/{n} found)"


def process_multi(df, sku_col, count_col, base_url, ext, verify, workers=16, progress_cb=None):
    """Returns (url_matrix, report) where url_matrix is a list of lists
    (one list of URLs per row, not yet padded to equal length)."""
    n = len(df)
    skus = df[sku_col].tolist()
    counts = df[count_col].tolist()
    results = [None] * n
    done = 0
    with ThreadPoolExecutor(max_workers=workers) as pool:
        futures = {
            pool.submit(resolve_row_multi, skus[i], counts[i], base_url, ext, verify): i
            for i in range(n)
        }
        for fut in as_completed(futures):
            results[futures[fut]] = fut.result()
            done += 1
            if progress_cb:
                progress_cb(done / n)
    url_lists = [r[0] for r in results]
    report = pd.DataFrame({
        "row": [i + 2 for i in range(n)],
        "sku": skus,
        "count": counts,
        "found": [len(u) for u in url_lists],
        "status": [r[1] for r in results],
    })
    return url_lists, report


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


def write_excel_single(raw: bytes, sheet: str, img_col: str, values: list) -> bytes:
    """Write one column. If a cell's value is a single URL it gets a real
    hyperlink; a comma-separated list of URLs is written as plain text
    (Excel only supports one hyperlink per cell, so a joined list can
    never be individually clickable)."""
    wb = load_workbook(io.BytesIO(raw))
    ws = wb[sheet]
    headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
    col = headers.index(img_col) + 1 if img_col in headers else len(headers) + 1
    if col > len(headers):
        ws.cell(row=1, column=col, value=img_col)

    for i, v in enumerate(values, start=2):
        cell = ws.cell(row=i, column=col)
        if v and "," not in v:
            cell.value = v
            cell.hyperlink = v
            cell.font = LINK_FONT
        elif v:
            cell.value = v
            cell.hyperlink = None
            cell.font = Font()
        else:
            cell.value = None
            cell.hyperlink = None

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def write_excel_multi(raw: bytes, sheet: str, img_col: str, url_lists: list) -> bytes:
    """Multi-image mode: one column per image slot (images_1, images_2, ...),
    each filled cell gets a real, individually clickable hyperlink."""
    max_n = max((len(u) for u in url_lists), default=0)
    wb = load_workbook(io.BytesIO(raw))
    ws = wb[sheet]
    headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]

    # find or create columns images_1 .. images_max_n
    col_indices = []
    for i in range(1, max_n + 1):
        name = f"{img_col}_{i}"
        if name in headers:
            col_indices.append(headers.index(name) + 1)
        else:
            col = len(headers) + 1
            ws.cell(row=1, column=col, value=name)
            headers.append(name)
            col_indices.append(col)

    for row_i, urls in enumerate(url_lists, start=2):
        for slot in range(max_n):
            cell = ws.cell(row=row_i, column=col_indices[slot])
            if slot < len(urls):
                cell.value = urls[slot]
                cell.hyperlink = urls[slot]
                cell.font = LINK_FONT
            else:
                cell.value = None
                cell.hyperlink = None

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def write_csv_single(df: pd.DataFrame, img_col: str, values: list) -> bytes:
    out = df.copy()
    out[img_col] = values
    return out.to_csv(index=False).encode("utf-8-sig")


def write_csv_multi(df: pd.DataFrame, img_col: str, url_lists: list) -> bytes:
    out = df.copy()
    max_n = max((len(u) for u in url_lists), default=0)
    for slot in range(max_n):
        out[f"{img_col}_{slot + 1}"] = [u[slot] if slot < len(u) else "" for u in url_lists]
    return out.to_csv(index=False).encode("utf-8-sig")


# --------------------------------------------------------------------------- #
# UI
# --------------------------------------------------------------------------- #
def main():
    st.set_page_config(page_title="S3 Image URL Mapper", page_icon="🖼️", layout="wide")
    st.title("🖼️ S3 Image URL Mapper")
    st.caption("Fill in S3 image URLs for your product sheet, built from the SKU.")

    with st.sidebar:
        st.header("Settings")
        mode = st.radio(
            "Mode",
            ["Single image per product", "Multiple images (count column)"],
            help=("Single: one URL built from the SKU, e.g. SKU.jpg.\n\n"
                  "Multiple: a column holds how many images a product has "
                  "(e.g. subImage = 8); this builds SKU_1.jpg .. SKU_8.jpg, "
                  "one per column (images_1, images_2, ...), each a real "
                  "clickable link."),
        )
        base_url = st.text_input("S3 / CloudFront base URL", DEFAULT_BASE_URL,
                                 help="Folder URL where the images live. The SKU + extension is appended.")

        if mode == "Single image per product":
            exts = st.multiselect("Image extensions to try (in order)", EXT_CHOICES, default=["jpg"],
                                  help="With verification on, each extension is tried until one exists in S3.")
            verify = st.checkbox("Verify that each image exists in S3", value=False,
                                 help="Sends a HEAD request per URL. Needs the images to be publicly readable. "
                                      "Off = build every URL without checking S3.")
            overwrite = st.checkbox("Overwrite URLs that already exist", value=False,
                                    help="Off = only blank cells are filled.")
            check_existing = st.checkbox("Also verify existing URLs", value=False,
                                         disabled=overwrite)
            workers = st.slider("Parallel requests", 1, 32, 16)
            if not exts:
                st.warning("Select at least one image extension in the sidebar.")
                st.stop()
        else:
            ext = st.selectbox("Image extension", EXT_CHOICES, index=0)
            layout = st.radio(
                "Output layout",
                ["One column, comma-separated", "Separate column per image (clickable links)"],
                help=("One column: all URLs joined by commas in a single 'images' cell "
                      "(plain text — Excel can't make a comma list individually clickable).\n\n"
                      "Separate columns: images_1, images_2, ... each a real clickable link."),
            )
            separator = ","
            if layout == "One column, comma-separated":
                separator = st.text_input("Separator between URLs", ",")
            verify = st.checkbox("Verify that each image exists in S3", value=False,
                                 help="Sends a HEAD request per generated URL. If some of the N "
                                      "images for a product are missing, only the ones found are "
                                      "kept (marked PARTIAL in the report). Off = build every URL "
                                      "without checking S3.")
            workers = st.slider("Parallel requests", 1, 32, 16)

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
    st.dataframe(df.head(10), width='stretch')

    sku_default = next((i for i, c in enumerate(df.columns) if c.lower() == "sku"), 0)

    if mode == "Single image per product":
        c1, c2 = st.columns(2)
        sku_col = c1.selectbox("SKU column (used as the image filename)", df.columns, index=sku_default)
        img_default = next((c for c in df.columns if c.lower() in ("images", "image", "image_url")), "images")
        img_col = c2.text_input("Images column (created if missing)", img_default)
    else:
        c1, c2, c3 = st.columns(3)
        sku_col = c1.selectbox("SKU column (used as the image filename prefix)", df.columns, index=sku_default)

        count_match = next((i for i, c in enumerate(df.columns) if c.lower() in ("subimage", "sub_image")), None)
        # never silently default to the SKU column itself — if nothing looks
        # like a count column, fall back to the first *other* column and
        # rely on the validation warning below to flag it
        if count_match is None:
            count_match = next((i for i in range(len(df.columns)) if i != sku_default), 0)
        count_col = c2.selectbox("Count column (how many images per product)", df.columns, index=count_match)

        img_default = next((c for c in df.columns if c.lower() in ("images", "image", "image_url")), "images")
        img_col = c3.text_input("Images column prefix (creates images_1, images_2, ...)", img_default)

        if count_col == sku_col:
            st.warning("Count column is the same as the SKU column — that's almost certainly wrong. "
                       "Pick the column that holds a number of images per product (e.g. subImage).")
        else:
            sample = df[count_col].astype(str).str.strip()
            sample = sample[sample != ""]
            non_numeric = (~sample.str.match(r"^\d+(\.0+)?$")).sum()
            if len(sample) and non_numeric / len(sample) > 0.3:
                st.warning(f"'{count_col}' has {non_numeric} of {len(sample)} values that aren't plain numbers. "
                           "Make sure this is really the image-count column, not something like the SKU or name.")

    dup = df[sku_col].str.strip().replace("", pd.NA).dropna().duplicated().sum()
    if dup:
        st.warning(f"{dup} duplicate SKU(s) found. Those rows will point to the same image(s).")

    # invalidate any previous result if the file, mode, or key columns changed,
    # so a stale report from a different upload is never shown by mistake
    run_key = (uploaded.name, uploaded.size, mode, sku_col, img_col,
               count_col if mode != "Single image per product" else None,
               layout if mode != "Single image per product" else None)
    if st.session_state.get("result", {}).get("key") != run_key:
        st.session_state.pop("result", None)

    # ---- Run
    if st.button("Generate image URLs", type="primary"):
        bar = st.progress(0.0, text="Working...")
        stem, suffix = Path(uploaded.name).stem, ".xlsx" if is_excel else ".csv"

        if mode == "Single image per product":
            values, report = process(
                df, sku_col, img_col.strip(), base_url, exts, verify, overwrite,
                check_existing and not overwrite, workers,
                progress_cb=lambda p: bar.progress(p, text=f"Working... {int(p * 100)}%"),
            )
            out_bytes = (write_excel_single(raw, sheet, img_col.strip(), values) if is_excel
                         else write_csv_single(df, img_col.strip(), values))
        else:
            url_lists, report = process_multi(
                df, sku_col, count_col, base_url, ext, verify, workers,
                progress_cb=lambda p: bar.progress(p, text=f"Working... {int(p * 100)}%"),
            )
            if layout == "One column, comma-separated":
                values = [separator.join(u) for u in url_lists]
                out_bytes = (write_excel_single(raw, sheet, img_col.strip(), values) if is_excel
                             else write_csv_single(df, img_col.strip(), values))
            else:
                out_bytes = (write_excel_multi(raw, sheet, img_col.strip(), url_lists) if is_excel
                             else write_csv_multi(df, img_col.strip(), url_lists))
        bar.empty()

        st.session_state["result"] = {
            "key": run_key,
            "report": report,
            "bytes": out_bytes,
            "name": f"{stem}_with_images{suffix}",
            "mime": ("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                     if is_excel else "text/csv"),
        }

    # ---- Output
    res = st.session_state.get("result")
    if res:
        report = res["report"]
        counts = report["status"].str.replace(r"\s*\(.*\)", "", regex=True).value_counts()
        st.subheader("Result")
        cols = st.columns(max(len(counts), 1))
        for col, (label, n) in zip(cols, counts.items()):
            col.metric(label, int(n))

        problems = report[report["status"].str.contains("NOT FOUND|BROKEN|skipped|PARTIAL", regex=True)]
        if len(problems):
            st.error(f"{len(problems)} row(s) need attention:")
            st.dataframe(problems, width='stretch', hide_index=True)
        else:
            st.success("All rows have working image URL(s).")

        with st.expander("Full report"):
            st.dataframe(report, width='stretch', hide_index=True)

        d1, d2 = st.columns(2)
        d1.download_button("⬇️ Download updated file", res["bytes"], res["name"], res["mime"],
                           type="primary", width='stretch')
        d2.download_button("⬇️ Download report (CSV)", report.to_csv(index=False).encode("utf-8-sig"),
                           "image_mapping_report.csv", "text/csv", width='stretch')


if __name__ == "__main__":
    main()