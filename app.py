import copy
import hashlib
import html
import json
import re
import secrets
import sqlite3
import tempfile
from dataclasses import dataclass
from html.parser import HTMLParser
from pathlib import Path
from typing import Any
from urllib.parse import quote_plus, urlparse

import openpyxl
import pandas as pd
import requests
import streamlit as st


APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "data"
EXPORT_DIR = APP_DIR / "outputs"
DB_PATH = DATA_DIR / "users.sqlite3"

SITE_NAME = "Lebal Info Finder"

DISTRIBUTOR = (
    "DISTRIBUTED BY / DISTRIBUÉ PAR: Nakama Trading Ltd, Scarborough, "
    "Ontario, M1X 2E5 service1@nakamatrading.com"
)
DEFAULT_DIRECTION_EN = (
    "DIRECTION FOR USE: Apply a proper amount onto lips and cheeks using a "
    "brush or fingertips. Blend evenly."
)
DEFAULT_DIRECTION_FR = (
    "MODE D’EMPLOI: Appliquer une quantité appropriée sur les lèvres et les "
    "joues à l’aide d’un pinceau ou du bout des doigts. Estomper uniformément."
)
DEFAULT_CAUTION_EN = (
    "CAUTIONS: For external use only. Discontinue use if irritation occurs. "
    "Keep out of reach of children."
)
DEFAULT_CAUTION_FR = (
    "MISES EN GARDE: Pour usage externe seulement. Cesser l'utilisation si une "
    "irritation se manifeste. Garder hors de la portée des enfants."
)
HOTLIST_URL = (
    "https://www.canada.ca/en/health-canada/services/consumer-product-safety/"
    "cosmetics/cosmetic-ingredient-hotlist-prohibited-restricted-ingredients/"
    "hotlist.html"
)


def ensure_dirs() -> None:
    DATA_DIR.mkdir(exist_ok=True)
    EXPORT_DIR.mkdir(exist_ok=True)


def password_hash(password: str, salt: str | None = None) -> str:
    salt = salt or secrets.token_hex(16)
    digest = hashlib.pbkdf2_hmac(
        "sha256", password.encode("utf-8"), salt.encode("utf-8"), 180_000
    ).hex()
    return f"{salt}${digest}"


def verify_password(password: str, stored: str) -> bool:
    if "$" not in stored:
        return False
    salt, expected = stored.split("$", 1)
    return secrets.compare_digest(password_hash(password, salt).split("$", 1)[1], expected)


def db() -> sqlite3.Connection:
    ensure_dirs()
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS users (
            username TEXT PRIMARY KEY,
            password_hash TEXT NOT NULL,
            role TEXT NOT NULL CHECK(role IN ('admin', 'viewer')),
            active INTEGER NOT NULL DEFAULT 1
        )
        """
    )
    if conn.execute("SELECT COUNT(*) FROM users").fetchone()[0] == 0:
        conn.execute(
            "INSERT INTO users(username, password_hash, role, active) VALUES(?,?,?,1)",
            ("admin", password_hash("change-me-now"), "admin"),
        )
        conn.commit()
    return conn


def authenticate(username: str, password: str) -> dict[str, Any] | None:
    with db() as conn:
        row = conn.execute(
            "SELECT username, password_hash, role, active FROM users WHERE username=?",
            (username.strip(),),
        ).fetchone()
    if not row or not row["active"] or not verify_password(password, row["password_hash"]):
        return None
    return {"username": row["username"], "role": row["role"]}


def manage_users() -> None:
    st.subheader("Users")
    with db() as conn:
        rows = conn.execute(
            "SELECT username, role, active FROM users ORDER BY username"
        ).fetchall()
    users = pd.DataFrame([dict(r) for r in rows])
    st.dataframe(users, use_container_width=True, hide_index=True)

    with st.form("add_user"):
        st.caption("Add approved user")
        col1, col2, col3 = st.columns([2, 2, 1])
        username = col1.text_input("Username")
        password = col2.text_input("Temporary password", type="password")
        role = col3.selectbox("Role", ["viewer", "admin"])
        if st.form_submit_button("Add user"):
            if not username or not password:
                st.error("Username and password are required.")
            else:
                try:
                    with db() as conn:
                        conn.execute(
                            "INSERT INTO users(username, password_hash, role, active) "
                            "VALUES(?,?,?,1)",
                            (username.strip(), password_hash(password), role),
                        )
                        conn.commit()
                    st.success(f"Added {username}.")
                except sqlite3.IntegrityError:
                    st.error("That username already exists.")

    with st.form("update_user"):
        st.caption("Update approved user")
        selected = st.selectbox("User", users["username"].tolist() if not users.empty else [])
        col1, col2, col3 = st.columns(3)
        new_role = col1.selectbox("New role", ["viewer", "admin"])
        active = col2.checkbox("Active", value=True)
        new_password = col3.text_input("New password", type="password")
        if st.form_submit_button("Save user"):
            with db() as conn:
                conn.execute(
                    "UPDATE users SET role=?, active=? WHERE username=?",
                    (new_role, int(active), selected),
                )
                if new_password:
                    conn.execute(
                        "UPDATE users SET password_hash=? WHERE username=?",
                        (password_hash(new_password), selected),
                    )
                conn.commit()
            st.success("User updated.")


class LinkParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.links: list[tuple[str, str]] = []
        self._href: str | None = None
        self._text: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag == "a":
            attrs_dict = dict(attrs)
            href = attrs_dict.get("href")
            if href and href.startswith("http"):
                self._href = href
                self._text = []

    def handle_data(self, data: str) -> None:
        if self._href:
            self._text.append(data)

    def handle_endtag(self, tag: str) -> None:
        if tag == "a" and self._href:
            text = html.unescape(" ".join(self._text)).strip()
            if text and len(text) > 8:
                self.links.append((text, self._href))
            self._href = None
            self._text = []


@st.cache_data(ttl=60 * 60)
def search_web(query: str) -> list[dict[str, str]]:
    headers = {"User-Agent": "Mozilla/5.0 label-research-tool/1.0"}
    urls = [
        f"https://duckduckgo.com/html/?q={quote_plus(query)}",
        f"https://www.bing.com/search?q={quote_plus(query)}",
    ]
    results: list[dict[str, str]] = []
    for url in urls:
        try:
            response = requests.get(url, timeout=12, headers=headers)
            response.raise_for_status()
        except requests.RequestException:
            continue
        parser = LinkParser()
        parser.feed(response.text)
        for title, href in parser.links:
            domain = urlparse(href).netloc.lower()
            if any(bad in domain for bad in ["facebook", "instagram", "tiktok", "pinterest"]):
                continue
            results.append({"title": title, "url": href})
        if results:
            break
    deduped: list[dict[str, str]] = []
    seen = set()
    for result in results:
        key = result["url"].split("&")[0]
        if key not in seen:
            seen.add(key)
            deduped.append(result)
    return deduped[:8]


@st.cache_data(ttl=60 * 60 * 24)
def fetch_text(url: str) -> str:
    try:
        response = requests.get(
            url,
            timeout=15,
            headers={"User-Agent": "Mozilla/5.0 label-research-tool/1.0"},
        )
        response.raise_for_status()
    except requests.RequestException:
        return ""
    text = re.sub(r"<script.*?</script>|<style.*?</style>", " ", response.text, flags=re.S | re.I)
    text = re.sub(r"<[^>]+>", " ", text)
    return html.unescape(re.sub(r"\s+", " ", text)).strip()


@st.cache_data(ttl=60 * 60 * 24)
def hotlist_text() -> str:
    return fetch_text(HOTLIST_URL)


def normalized_headers(ws: openpyxl.worksheet.worksheet.Worksheet) -> dict[str, int]:
    headers = {}
    for cell in ws[1]:
        if cell.value:
            key = str(cell.value).strip().lower()
            headers[key] = cell.column
    return headers


def find_column(headers: dict[str, int], *needles: str) -> int | None:
    for needle in needles:
        needle_l = needle.lower()
        for header, col in headers.items():
            if needle_l == header or needle_l in header:
                return col
    return None


def copy_cell_style(src: openpyxl.cell.cell.Cell, dst: openpyxl.cell.cell.Cell) -> None:
    if src.has_style:
        dst._style = copy.copy(src._style)
    if src.number_format:
        dst.number_format = src.number_format
    if src.alignment:
        dst.alignment = copy.copy(src.alignment)


def add_audit_columns(ws: openpyxl.worksheet.worksheet.Worksheet) -> dict[str, int]:
    headers = normalized_headers(ws)
    wanted = ["source url", "source notes", "row status"]
    for label in wanted:
        if label not in headers:
            col = ws.max_column + 1
            ws.cell(1, col).value = label.title()
            copy_cell_style(ws.cell(1, 1), ws.cell(1, col))
            headers[label] = col
    return headers


def net_weight_from_name(name: str) -> str | None:
    match = re.search(r"(\d+(?:\.\d+)?)\s*(ml|mL|ML|g|G)\s*(?:\+\s*(pendant keyring|1\s*pcs|pcs))?", name)
    if not match:
        return None
    amount, unit, extra = match.group(1), match.group(2), match.group(3)
    unit = "mL" if unit.lower() == "ml" else "g"
    result = f"Net.{amount} {unit}"
    if extra:
        result += " + 1 PCS"
    return result


def coo_from_text(text: str) -> str | None:
    country_map = {
        "korea": "Made In Korea / Fabriqué En Corée",
        "south korea": "Made In Korea / Fabriqué En Corée",
        "japan": "Made In Japan / Fabriqué au Japon",
        "china": "Made In China / Fabriqué En Chine",
    }
    low = text.lower()
    for key, value in country_map.items():
        if f"made in {key}" in low or f"product of {key}" in low or key in low:
            return value
    return None


def ingredients_from_text(text: str) -> str | None:
    match = re.search(
        r"(ingredients?|inci)\s*[:：]\s*(.{40,1200}?)(?:directions?|how to use|caution|warning|made in|$)",
        text,
        flags=re.I,
    )
    if not match:
        return None
    ingredients = match.group(2).strip(" .;")
    if len(ingredients) < 40 or "," not in ingredients:
        return None
    return f"INGREDIENTS/INGRÉDIENTS: {ingredients} / need to review"


def check_hotlist(ingredients: str) -> tuple[list[str], str]:
    if not ingredients or ingredients == "need to review":
        return [], ""
    text = hotlist_text().lower()
    if not text:
        return [], "Could not fetch Health Canada Hotlist; review required."
    found = []
    candidates = []
    after_prefix = ingredients.split(":", 1)[-1].split("/", 1)[0]
    for item in after_prefix.split(","):
        clean = re.sub(r"\([^)]*\)", "", item).strip().lower()
        if len(clean) > 4:
            candidates.append(clean)
    for ingredient in candidates:
        if ingredient in text:
            found.append(ingredient.title())
    note = "Restricted/prohibited candidate found on Health Canada Hotlist." if found else ""
    return found, note


@dataclass
class FillResult:
    values: dict[str, str]
    status: str
    source_url: str
    notes: str


def fill_from_reference(
    ref_ws: openpyxl.worksheet.worksheet.Worksheet,
    ref_headers: dict[str, int],
    barcode: str,
) -> dict[str, str] | None:
    barcode_col = find_column(ref_headers, "barcode")
    if not barcode_col:
        return None
    for row in range(2, ref_ws.max_row + 1):
        ref_barcode = str(ref_ws.cell(row, barcode_col).value or "").strip()
        if ref_barcode and ref_barcode == barcode:
            values = {}
            for header, col in ref_headers.items():
                value = ref_ws.cell(row, col).value
                if value not in (None, ""):
                    values[header] = str(value)
            return values
    return None


def process_row(
    row: dict[str, Any],
    reference_values: dict[str, str] | None,
    use_defaults: bool,
) -> FillResult:
    product = str(row.get("product name") or "")
    barcode = str(row.get("barcode") or "").replace(".0", "")
    values: dict[str, str] = {}
    source_url = ""
    notes: list[str] = []

    if reference_values:
        values.update(reference_values)
        notes.append("Matched completed reference sheet by barcode.")
        return FillResult(values, "Completed", "Workbook Sheet1", " ".join(notes))

    query = barcode if barcode else product
    results = search_web(f"{query} ingredients net weight origin")
    page_text = ""
    if results:
        source_url = results[0]["url"]
        page_text = fetch_text(source_url)
        notes.append(f"Search result: {results[0]['title']}")
    else:
        notes.append("No online source found.")

    values["net weight"] = net_weight_from_name(product) or "need to review"
    values["ingredients/ingrédients"] = ingredients_from_text(page_text) or "need to review"
    values["coo"] = coo_from_text(page_text + " " + product) or "need to review"
    values["distributed by / distribué par:"] = DISTRIBUTOR

    if use_defaults and "lip&cheek" in product.lower().replace(" ", ""):
        values["direction for use"] = DEFAULT_DIRECTION_EN
        values["mode d’emploi"] = DEFAULT_DIRECTION_FR
        values["cautions"] = DEFAULT_CAUTION_EN
        values["mises en garde:"] = DEFAULT_CAUTION_FR

    restricted, hotlist_note = check_hotlist(values.get("ingredients/ingrédients", ""))
    if restricted:
        values["restricted ingredients"] = ", ".join(restricted)
        notes.append(hotlist_note)
    elif hotlist_note:
        notes.append(hotlist_note)

    missing = [k for k, v in values.items() if v == "need to review"]
    status = "Need to review" if missing or restricted else "Completed"
    if not source_url:
        status = "Missing source"
    return FillResult(values, status, source_url, " ".join(notes))


def dataframe_from_sheet(ws: openpyxl.worksheet.worksheet.Worksheet, rows: int = 20) -> pd.DataFrame:
    data = list(ws.iter_rows(values_only=True))
    if not data:
        return pd.DataFrame()
    headers = [str(h) if h is not None else f"Unnamed {i+1}" for i, h in enumerate(data[0])]
    return pd.DataFrame(data[1 : rows + 1], columns=headers)


def full_dataframe_from_sheet(ws: openpyxl.worksheet.worksheet.Worksheet) -> pd.DataFrame:
    data = list(ws.iter_rows(values_only=True))
    if not data:
        return pd.DataFrame()
    headers = [str(h) if h is not None else f"Unnamed {i+1}" for i, h in enumerate(data[0])]
    return pd.DataFrame(data[1:], columns=headers).fillna("")


def apply_dataframe_to_sheet(
    wb_path: Path,
    sheet_name: str,
    df: pd.DataFrame,
    output_name: str,
) -> Path:
    wb = openpyxl.load_workbook(wb_path)
    ws = wb[sheet_name]
    for col_idx, header in enumerate(df.columns, start=1):
        ws.cell(1, col_idx).value = header if not str(header).startswith("Unnamed ") else None
    for row_idx, row in enumerate(df.itertuples(index=False), start=2):
        for col_idx, value in enumerate(row, start=1):
            ws.cell(row_idx, col_idx).value = None if value == "" else value
    return export_workbook(wb, output_name)


def workbook_bytes(uploaded_file: Any) -> bytes:
    return uploaded_file.getvalue()


def save_upload(uploaded_file: Any) -> Path:
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx", dir=DATA_DIR)
    tmp.write(workbook_bytes(uploaded_file))
    tmp.close()
    return Path(tmp.name)


def export_workbook(wb: openpyxl.Workbook, name: str) -> Path:
    EXPORT_DIR.mkdir(exist_ok=True)
    safe = re.sub(r"[^A-Za-z0-9_.-]+", "_", name).strip("_") or "filled_labels.xlsx"
    output = EXPORT_DIR / safe
    wb.save(output)
    return output


def process_workbook(
    path: Path,
    reference_sheet: str,
    fill_sheet: str,
    use_defaults: bool,
    limit: int | None = None,
) -> tuple[openpyxl.Workbook, pd.DataFrame]:
    wb = openpyxl.load_workbook(path)
    ref_ws = wb[reference_sheet]
    fill_ws = wb[fill_sheet]
    ref_headers = normalized_headers(ref_ws)
    fill_headers = add_audit_columns(fill_ws)

    barcode_col = find_column(fill_headers, "barcode")
    if not barcode_col:
        raise ValueError("No barcode column found.")

    product_col = find_column(fill_headers, "product name")
    records = []
    max_row = fill_ws.max_row if limit is None else min(fill_ws.max_row, limit + 1)
    for row_idx in range(2, max_row + 1):
        barcode = str(fill_ws.cell(row_idx, barcode_col).value or "").strip().replace(".0", "")
        product = str(fill_ws.cell(row_idx, product_col).value or "") if product_col else ""
        row = {"barcode": barcode, "product name": product}
        reference_values = fill_from_reference(ref_ws, ref_headers, barcode)
        result = process_row(row, reference_values, use_defaults)

        for header, value in result.values.items():
            col = find_column(fill_headers, header)
            if col and value:
                fill_ws.cell(row_idx, col).value = value
                if row_idx > 2:
                    copy_cell_style(fill_ws.cell(row_idx - 1, col), fill_ws.cell(row_idx, col))

        fill_ws.cell(row_idx, fill_headers["source url"]).value = result.source_url
        fill_ws.cell(row_idx, fill_headers["source notes"]).value = result.notes
        fill_ws.cell(row_idx, fill_headers["row status"]).value = result.status
        records.append(
            {
                "row": row_idx,
                "barcode": barcode,
                "product": product,
                "status": result.status,
                "source_url": result.source_url,
                "notes": result.notes,
            }
        )

    return wb, pd.DataFrame(records)


def login_screen() -> None:
    st.title(SITE_NAME)
    st.caption("Private workbook processing for bilingual Nakama labels.")
    with st.form("login"):
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Sign in")
    if submitted:
        user = authenticate(username, password)
        if user:
            st.session_state.user = user
            st.rerun()
        st.error("Invalid username or password.")
    st.info("First run admin account: username `admin`, password `change-me-now`.")


def main_app() -> None:
    user = st.session_state.user
    st.sidebar.write(f"Signed in as **{user['username']}**")
    if st.sidebar.button("Sign out"):
        st.session_state.pop("user", None)
        st.rerun()

    page = st.sidebar.radio(
        "Navigation",
        ["Fill labels", "Manage users"] if user["role"] == "admin" else ["Fill labels"],
    )
    if page == "Manage users":
        manage_users()
        return

    st.title(SITE_NAME)
    st.caption("Upload an Excel workbook, preview sheets, process rows, edit, and export.")

    uploaded = st.file_uploader("Excel workbook", type=["xlsx"])
    if not uploaded:
        return

    path = save_upload(uploaded)
    wb = openpyxl.load_workbook(path)
    names = wb.sheetnames
    col1, col2 = st.columns(2)
    reference_sheet = col1.selectbox("Completed reference sheet", names, index=names.index("Sheet1") if "Sheet1" in names else 0)
    fill_sheet = col2.selectbox("Sheet to fill", names, index=names.index("Sheet2") if "Sheet2" in names else min(1, len(names) - 1))

    with st.expander("Preview completed reference", expanded=False):
        st.dataframe(dataframe_from_sheet(wb[reference_sheet]), use_container_width=True)
    with st.expander("Preview sheet to fill", expanded=True):
        st.dataframe(dataframe_from_sheet(wb[fill_sheet]), use_container_width=True)

    use_defaults = st.checkbox(
        "Use approved default lip/cheek direction and general cautions when source data is missing",
        value=True,
    )
    limit = st.number_input("Rows to process now (0 = all)", min_value=0, value=0, step=1)

    if st.button("Process workbook", type="primary"):
        with st.status("Processing rows one by one...", expanded=True) as status:
            processed_wb, report = process_workbook(
                path,
                reference_sheet,
                fill_sheet,
                use_defaults,
                None if limit == 0 else int(limit),
            )
            output_path = export_workbook(processed_wb, f"filled_{uploaded.name}")
            status.update(label="Processing complete", state="complete")
        st.session_state.report = report
        st.session_state.output_path = str(output_path)
        st.session_state.fill_sheet = fill_sheet
        st.session_state.uploaded_name = uploaded.name

    if "report" in st.session_state:
        st.subheader("Row status")
        st.dataframe(st.session_state.report, use_container_width=True, hide_index=True)
        output_path = Path(st.session_state.output_path)

        st.subheader("Manual edit before export")
        processed_wb = openpyxl.load_workbook(output_path)
        edit_df = full_dataframe_from_sheet(processed_wb[st.session_state.fill_sheet])
        edited_df = st.data_editor(
            edit_df,
            use_container_width=True,
            hide_index=True,
            num_rows="fixed",
            key="processed_sheet_editor",
        )
        if st.button("Save edited workbook"):
            edited_path = apply_dataframe_to_sheet(
                output_path,
                st.session_state.fill_sheet,
                edited_df,
                f"edited_{st.session_state.uploaded_name}",
            )
            st.session_state.output_path = str(edited_path)
            output_path = edited_path
            st.success("Edited workbook saved.")

        st.download_button(
            "Download completed Excel",
            data=output_path.read_bytes(),
            file_name=output_path.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


def main() -> None:
    st.set_page_config(page_title=SITE_NAME, layout="wide")
    ensure_dirs()
    db().close()
    if "user" not in st.session_state:
        login_screen()
    else:
        main_app()


if __name__ == "__main__":
    main()
