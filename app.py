import copy
import difflib
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
from urllib.parse import parse_qs, quote_plus, urlencode, unquote, urlparse, urlunparse

import openpyxl
import pandas as pd
import requests
import streamlit as st


APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "data"
EXPORT_DIR = APP_DIR / "outputs"
DB_PATH = DATA_DIR / "users.sqlite3"
BUILTIN_REFERENCE_PATH = APP_DIR / "reference_data.json"

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

REQUIRED_LABEL_FIELDS = [
    "product name french",
    "net weight",
    "direction for use",
    "mode d’emploi",
    "cautions",
    "mises en garde:",
    "ingredients/ingrédients",
    "manufacturer",
    "distributed by / distribué par:",
    "coo",
]

TRUSTED_DOMAINS = [
    "judydoll.com",
    "millefee.com",
    "joybeautyhub.shop",
    "intoyoucosmetics.com",
    "yesstyle.com",
    "asianbeautywholesale.com",
    "uniquebunny.com",
    "kiseki.ca",
    "oliveyoung.com",
    "stylevana.com",
]

SEARCH_LANGUAGE_HINTS = [
    "",
    "English",
    "official",
    "ingredients",
    "net weight",
    "how to use",
    "成分",
    "容量",
    "使い方",
    "성분",
]

KNOWN_ONLINE_PRODUCTS = {
    "1129343972": {
        "source_url": "https://www.asianbeautywholesale.com/en/into-you-glowing-lipstick-8-colors-gl08-red-brown-3g/info.html/pid.1129343972\nhttps://www.intoyoucosmetics.com/en-ca/products/into-you-glow-lipstick?variant=57958599524655\nhttps://www.intoyoucosmetics.com/en-ca/products/into-you-glow-lipstick?variant=57958599590191\nhttps://www.yesstyle.com/en/into-you-glowing-lipstick-8-colors-gl08-red-brown-3g/info.html/pid.1129343972\nhttps://www.uniquebunny.com/products/into-you-glow-lipstick\nhttps://www.intoyoucosmetics.com/en-ca/pages/about-us",
        "net weight": "Net. 3 g",
        "source_direction": "Apply a small amount directly to lips. Store below 25°C and refrigerate if the product softens.",
        "ingredients": "Isostearyl Isostearate, Polyglyceryl-2 Triisostearate, Diisostearyl Malate, Sorbitan Isostearate, Paraffin, Trimethylpentaphenyl Trisiloxane, Microcrystalline Wax, Pentaerythrityl Isostearate, Euphorbia Cerifera (Candelilla) Wax, 1,2-Hexanediol, PEG/PPG-10/1 Dimethicone, CI 77891, CI 19140, CI 45410, CI 77491, Fragrance, CI 77499, Pentaerythrityl Tetraisostearate",
        "manufacturer": "HONGKONG LETS INTERNATIONAL TRADING LIMITED",
    },
    "1126245093": {
        "source_url": "https://www.intoyoucosmetics.com/en-gb/products/airy-lip-cheek-mud\nhttps://www.uniquebunny.com/fr/products/into-you-airy-lip-cheek-mud\nhttps://www.uniquebunny.com/products/into-you-airy-lip-cheek-mud\nhttps://www.yesstyle.com/en/into-you-airy-lip-cheek-mud-5-colors-c1-c5-c5-mauve-taupe-1-8g/info.html/pid.1126244966",
        "net weight": "Net. 2 g",
        "source_direction": "Apply a proper amount evenly on lips, or dab onto cheeks and blend with fingertips.",
        "coo": "Made In China / Fabriqué En Chine",
        "ingredients_fr": "Diméthicone, polymère croisé de diméthicone, triisostéarate de polyglycéryle-2, cyclopentasiloxane, cyclohexasiloxane, CI 19140, 1,2-hexanediol, silylate de silice, cétyl PEG/PPG-10/1 diméthicone, talc, CI 77491, CI 77891, extrait de Melaleuca Alternifolia, huile d'onagre, acétate d'alpha-tocophéryle, C30-45 Alkyl Diméthicone",
    },
    "1137202898": {
        "source_url": "https://www.yesstyle.com/en/into-you-six-color-blush-palette-six-color-blush-palette-15g/info.html/pid.1137202898",
        "net weight": "Net. 15 g",
        "source_direction": "Apply the blush to cheeks with a brush. Mix, match, and layer shades as desired.",
        "ingredients": "Talc, Mica, CI 77891, Synthetic Fluorphlogopite, Silica, Magnesium Myristate, Vinyl Dimethicone/Methicone Silsesquioxane Crosspolymer, Octyldodecanol, Isostearyl Neopentanoate, Dimethicone, Methyl Methacrylate Crosspolymer, Aluminum Myristate, Diisostearyl Malate, CI 77491, CI 75470, Ethylhexylglycerin, CI 77492, Glyceryl Caprylate, Lauroyl Lysine, CI 77499, Dimethicone Crosspolymer, Triethoxycaprylylsilane, Cocos Nucifera (Coconut) Oil, Alumina, Water, Aluminum Hydroxide, CI 73360, CI 19140, CI 45410, CI 77007, Boron Nitride",
        "manufacturer": "HONGKONG LETS INTERNATIONAL TRADING LIMITED",
    },
}

JUDYDOLL_LIQUID_BLUSH_SERUM = {
    "source_url": (
        "https://www.yesstyle.com/en/judydoll-liquid-blush-serum-4-colors-01-fig-5g/info.html/pid.1136648925\n"
        "https://judydoll.com/products/liquid-blush-serum"
    ),
    "net weight": "Net. 5 g",
    "coo": "Made In China / Fabriqué En Chine",
    "source_direction": "Lightly dab after foundation and before setting powder, then blend evenly across cheeks.",
    "shades": {
        "01-fig": "water, CI 77891, cyclopentasiloxane, isododecane, diisopropyl sebacate, butylene glycol, lauryl PEG-10 tri(trimethylsiloxy) silane, 1,2-pentanediol, diphenylmethoxy silane, trimethylsilyl silicate, CI 77019, polydimethylsiloxane, cyclohexasiloxane, squalane, cetearyl PEG/PPG-10/1 polydimethylsiloxane, polymethyl methacrylate, distearyldimonium lithium montmorillonite, magnesium sulfate, phenoxyethanol, polydimethylsiloxane/vinylpolydimethylsiloxane crosspolymer, polyhydroxystearic acid, HDI/trihydroxymethyl hexyl lactone crosspolymer, CI 77491, aluminum hydroxide, triethoxy octylsilane, CI 77492, polydimethylsiloxane crosspolymer, ethylhexylglycerin, vinylpolydimethylsiloxane/polymethylsilsesquioxane crosspolymer, disodium ethylenediaminetetraacetate, CI 73360, CI 77499, silica, glycerin, 1,3-propanediol, tocopherol (vitamin E), pentaerythrityl tetra(butylated hydroxytoluene) ester, collagen, PPG-13-decyloleate-24, 1,2-hexanediol, acetyl hexapeptide-8, palmitoyl pentapeptide-4",
        "02-pink-milk": "water, CI 77891, cyclopentasiloxane, isododecane, diisopropyl sebacate, butylene glycol, lauryl PEG-10 tri(trimethylsiloxy) silane, 1,2-pentanediol, diphenylmethoxy silane, trimethylsilyl silicate, polydimethylsiloxane, cyclohexasiloxane, squalane, cetearyl PEG/PPG-10/1 polydimethylsiloxane, polymethyl methacrylate, distearyldimonium lithium montmorillonite, magnesium sulfate, phenoxyethanol, CI 77019, polydimethylsiloxane/vinylpolydimethylsiloxane crosspolymer, polyhydroxystearic acid, HDI/trihydroxymethyl hexyl lactone crosspolymer, CI 77492, aluminum hydroxide, CI 77007, triethoxy octylsilane, polydimethylsiloxane crosspolymer, ethylhexylglycerin, vinylpolydimethylsiloxane/polymethylsilsesquioxane crosspolymer, CI 73360, CI 77499, disodium ethylenediaminetetraacetate, silica, glycerin, 1,3-propanediol, tocopherol (vitamin E), pentaerythrityl tetra(butylated hydroxytoluene) ester, collagen, PPG-13-decyloleate-24, 1,2-hexanediol, acetyl hexapeptide-8, palmitoyl pentapeptide-4",
        "03-salmon": "water, CI 77891, cyclopentasiloxane, isododecane, diisopropyl sebacate, butylene glycol, lauryl PEG-10 tri(trimethylsiloxy) silane, 1,2-pentanediol, diphenylmethoxy silane, trimethylsilyl silicate, polydimethylsiloxane, CI 77019, cyclohexasiloxane, squalane, cetearyl PEG/PPG-10/1 polydimethylsiloxane, polymethyl methacrylate, CI 77491, distearyldimonium lithium montmorillonite, magnesium sulfate, phenoxyethanol, polydimethylsiloxane/vinylpolydimethylsiloxane crosspolymer, polyhydroxystearic acid, HDI/trihydroxymethyl hexyl lactone crosspolymer, CI 77492, aluminum hydroxide, triethoxy octylsilane, polydimethylsiloxane crosspolymer, ethylhexylglycerin, vinylpolydimethylsiloxane/polymethylsilsesquioxane crosspolymer, CI 77499, CI 73360, disodium ethylenediaminetetraacetate, silica, glycerin, 1,3-propanediol, tocopherol (vitamin E), pentaerythrityl tetra(butylated hydroxytoluene) ester, collagen, PPG-13-decyloleate-24, 1,2-hexanediol, acetyl hexapeptide-8, palmitoyl pentapeptide-4",
        "04-rosewood": "water, cyclopentasiloxane, CI 77019, isododecane, diisopropyl sebacate, CI 77891, butylene glycol, lauryl PEG-10 tri(trimethylsiloxy) silane, 1,2-pentanediol, diphenylmethoxy silane, trimethylsilyl silicate, polydimethylsiloxane, cyclohexasiloxane, squalane, cetearyl PEG/PPG-10/1 polydimethylsiloxane, polymethyl methacrylate, distearyldimonium lithium montmorillonite, magnesium sulfate, phenoxyethanol, polydimethylsiloxane/vinylpolydimethylsiloxane crosspolymer, polyhydroxystearic acid, HDI/trihydroxymethyl hexyl lactone crosspolymer, CI 77491, triethoxy octylsilane, CI 77492, CI 73360, CI 77007, CI 77499, aluminum hydroxide, polydimethylsiloxane crosspolymer, ethylhexylglycerin, vinylpolydimethylsiloxane/polymethylsilsesquioxane crosspolymer, disodium ethylenediaminetetraacetate, silica, glycerin, 1,3-propanediol, tocopherol (vitamin E), pentaerythrityl tetra(butylated hydroxytoluene) ester, collagen, PPG-13-decyloleate-24, 1,2-hexanediol, acetyl hexapeptide-8, palmitoyl pentapeptide-4",
    },
}

MILLEFEE_IDOL_HIGHLIGHTER_PALETTE = {
    "net weight": "Net. 11 g",
    "coo": "Made In China / Fabriqué En Chine",
    "source_direction": "Apply to high points of the face with a brush. Layer as desired for added glow.",
    "shades": {
        "01-ice-blue": {
            "source_url": (
                "https://www.yesstyle.com/en/millefee-idol-highlighter-palette-01-ice-blue/info.html/pid.1137196647\n"
                "https://millefee.com/products/idol-highlighter-palette"
            ),
            "ingredients": "Synthetic Fluorphlogopite, Titanium Dioxide, Dimethicone, Silica, Phenyl Trimethicone, Calcium Aluminum Borosilicate, Diisostearyl Malate, Petrolatum, Isononyl Isononanoate, Synthetic Wax, Polyisobutene, Dimethicone/Vinyl Dimethicone Crosspolymer, Trimethylsiloxysilicate, Methyl Methacrylate Crosspolymer, Tin Oxide, Microcrystalline Wax, Caprylyl Glycol, Hydroxystearic Acid, Polymethylsilsesquioxane, Iron Oxides, Dimethicone Crosspolymer, Ethylhexylglycerin, Triethoxycaprylylsilane, Stearic Acid, Red 226, Barium Sulfate, Yellow 4, Mica, Blue 1",
        },
        "02-rose-pink": {
            "source_url": (
                "https://www.yesstyle.com/en/millefee-idol-highlighter-palette-02-rose-pink/info.html/pid.1137196649\n"
                "https://millefee.com/products/idol-highlighter-palette"
            ),
            "ingredients": "Synthetic Fluorphlogopite, Titanium Dioxide, Isononyl Isononanoate, Silica, Dimethicone, Polyisobutene, Phenyl Trimethicone, Calcium Aluminum Borosilicate, Polymethylsilsesquioxane, Dimethicone Crosspolymer, Hydroxystearic Acid, Microcrystalline Wax, Caprylyl Glycol, Tin Oxide, Triethoxycaprylylsilane, Ethylhexylglycerin, Stearic Acid, Iron Oxides, Red 226, Barium Sulfate, Blue 1",
        },
    },
    "barcodes": {
        "1137196647": "01-ice-blue",
        "1137196649": "02-rose-pink",
    },
}


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
    conn.execute(
        """
        INSERT OR IGNORE INTO users(username, password_hash, role, active)
        VALUES(?,?,?,1)
        """,
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
            if href and (href.startswith("http") or href.startswith("/url?")):
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
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
        "Accept-Language": "en-US,en;q=0.9,fr;q=0.8",
    }
    urls = [
        f"https://www.google.com/search?q={quote_plus(query)}",
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
        result["url"] = clean_search_url(result["url"])
        if is_noise_url(result["url"]):
            continue
        key = result["url"].split("&")[0]
        if key not in seen:
            seen.add(key)
            deduped.append(result)
    return sorted(deduped, key=lambda item: source_rank(item["url"]))[:25]


def clean_search_url(url: str) -> str:
    parsed = urlparse(url)
    if parsed.path == "/url":
        target = parse_qs(parsed.query).get("q", [""])[0]
        if target.startswith("http"):
            return target
    if "bing.com" in parsed.netloc and parsed.path.startswith("/ck/"):
        target = parse_qs(parsed.query).get("u", [""])[0]
        if target.startswith("a1"):
            try:
                import base64

                decoded = base64.urlsafe_b64decode(target[2:] + "==").decode("utf-8", "ignore")
                if decoded.startswith("http"):
                    return decoded
            except Exception:
                pass
        if target.startswith("http"):
            return unquote(target)
    keep_params = {}
    for key, values in parse_qs(parsed.query).items():
        if key.lower() in {"variant", "pid"}:
            keep_params[key] = values[-1]
    clean_query = urlencode(keep_params)
    return urlunparse((parsed.scheme, parsed.netloc, parsed.path, "", clean_query, ""))


def is_noise_url(url: str) -> bool:
    parsed = urlparse(url)
    domain = parsed.netloc.lower()
    path = parsed.path.lower()
    if any(bad in domain for bad in ["facebook", "instagram", "tiktok", "pinterest"]):
        return True
    noisy_parts = [
        "/list.html",
        "/brand/",
        "/brands/",
        "/collections/",
        "/collection/",
        "/search",
        "/tag/",
        "/blog",
        "/blogs/",
        "/category/",
        "/categories/",
        "/cart",
        "/account",
    ]
    if any(part in path for part in noisy_parts):
        if "/products/" not in path and "/info.html" not in path:
            return True
    return False


def source_rank(url: str) -> int:
    domain = domain_from_url(url)
    for idx, trusted in enumerate(TRUSTED_DOMAINS):
        if trusted in domain:
            return idx
    if any(market in domain for market in ["amazon.", "ebay.", "temu.", "aliexpress."]):
        return 180
    if any(noisy in domain for noisy in ["baidu.", "wikipedia.", "reddit."]):
        return 220
    return 100


def product_detail_bonus(url: str) -> float:
    parsed = urlparse(url)
    domain = parsed.netloc.lower()
    path = parsed.path.lower()
    if "yesstyle.com" in domain and "/info.html" in path and "/pid." in path:
        return 7.0
    if "/products/" in path:
        return 6.5
    if path.endswith(".html") and not is_noise_url(url):
        return 4.0
    return 0.0


def exact_product_url(url: str, product: str) -> bool:
    if is_noise_url(url):
        return False
    tokens = search_tokens(product)
    if not tokens:
        return product_detail_bonus(url) > 0
    haystack = searchable_text(url)
    product_words = [token for token in tokens if token not in {"serum", "liquid", "color", "colors"}]
    hits = sum(1 for token in product_words if token in haystack)
    shade = shade_key(product)
    shade_ok = not shade or shade in haystack
    return product_detail_bonus(url) > 0 and hits >= max(2, min(4, len(product_words) - 1)) and shade_ok


@st.cache_data(ttl=60 * 60 * 24)
def fetch_text(url: str) -> str:
    text = fetch_direct_text(url)
    if len(text) > 200:
        return text
    text = fetch_shopify_product_text(url)
    if len(text) > 200:
        return text
    text = fetch_jina_text(url)
    if len(text) > 200:
        return text
    return ""


def fetch_direct_text(url: str) -> str:
    try:
        response = requests.get(
            url,
            timeout=8,
            headers={
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
                "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
                "Accept-Language": "en-US,en;q=0.9,fr;q=0.8,zh-CN;q=0.7,ja;q=0.6,ko;q=0.6",
            },
        )
        response.raise_for_status()
    except requests.RequestException:
        return ""
    text = re.sub(r"<script.*?</script>|<style.*?</style>", " ", response.text, flags=re.S | re.I)
    text = re.sub(r"<[^>]+>", " ", text)
    return html.unescape(re.sub(r"\s+", " ", text)).strip()


def fetch_jina_text(url: str) -> str:
    try:
        response = requests.get(
            f"https://r.jina.ai/{url}",
            timeout=12,
            headers={
                "User-Agent": "Mozilla/5.0 label-research-tool/1.0",
                "Accept": "text/plain, text/markdown, */*",
            },
        )
        response.raise_for_status()
    except requests.RequestException:
        return ""
    return html.unescape(re.sub(r"\s+", " ", response.text)).strip()


def fetch_shopify_product_text(url: str) -> str:
    parsed = urlparse(url)
    if "/products/" not in parsed.path:
        return ""
    handle = parsed.path.split("/products/", 1)[1].strip("/").split("/", 1)[0]
    if not handle:
        return ""
    product_json_url = f"{parsed.scheme}://{parsed.netloc}/products/{handle}.js"
    try:
        response = requests.get(
            product_json_url,
            timeout=8,
            headers={
                "User-Agent": "Mozilla/5.0 label-research-tool/1.0",
                "Accept": "application/json,text/plain,*/*",
            },
        )
        response.raise_for_status()
        data = response.json()
    except (requests.RequestException, ValueError):
        return ""
    pieces = [
        str(data.get("title", "")),
        str(data.get("vendor", "")),
        re.sub(r"<[^>]+>", " ", str(data.get("description", ""))),
    ]
    for variant in data.get("variants", []) or []:
        pieces.append(str(variant.get("title", "")))
        pieces.append(str(variant.get("sku", "")))
    return html.unescape(re.sub(r"\s+", " ", " ".join(pieces))).strip()


@st.cache_data(ttl=60 * 60 * 24)
def hotlist_text() -> str:
    return fetch_text(HOTLIST_URL)


@st.cache_data
def builtin_reference_data() -> dict[str, dict[str, str]]:
    if not BUILTIN_REFERENCE_PATH.exists():
        return {}
    return json.loads(BUILTIN_REFERENCE_PATH.read_text(encoding="utf-8"))


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
    for label in ["source websites", "source notes", "row status"]:
        col = headers.get(label)
        if col:
            ws.delete_cols(col)
            headers = normalized_headers(ws)

    if "source url" not in headers:
        col = ws.max_column + 1
        ws.cell(1, col).value = "Source Url"
        copy_cell_style(ws.cell(1, 1), ws.cell(1, col))
        headers = normalized_headers(ws)
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


def net_weight_from_text(text: str) -> str | None:
    patterns = [
        r"net\s*weight\s*[:：]?\s*(\d+(?:\.\d+)?)\s*(g|gram|grams|ml|mL|ML)",
        r"(\d+(?:\.\d+)?)\s*(g|gram|grams|ml|mL|ML)\s*/\s*0\.",
        r"(\d+(?:\.\d+)?)\s*(g|gram|grams|ml|mL|ML)\b",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.I)
        if match:
            amount, unit = match.group(1), match.group(2)
            unit = "mL" if unit.lower() == "ml" else "g"
            return f"Net. {amount} {unit}"
    return None


def coo_from_text(text: str) -> str | None:
    country_map = {
        "korea": "Made In Korea / Fabriqué En Corée",
        "south korea": "Made In Korea / Fabriqué En Corée",
        "japan": "Made In Japan / Fabriqué au Japon",
        "china": "Made In China / Fabriqué En Chine",
    }
    low = text.lower()
    for key, value in country_map.items():
        if (
            f"made in {key}" in low
            or f"product of {key}" in low
            or f"country of origin {key}" in low
            or f"country/region of origin {key}" in low
        ):
            return value
    if "chinese makeup brand" in low or "china makeup" in low or "chinese cosmetics" in low:
        return country_map["china"]
    return None


def domain_from_url(url: str) -> str:
    domain = urlparse(url).netloc.lower()
    return domain.removeprefix("www.")


def ingredients_from_text(text: str, product: str = "") -> str | None:
    if not text:
        return None
    shade_ingredients = ingredients_for_shade(text, product)
    if shade_ingredients:
        return ingredients_label(shade_ingredients)
    stop_words = (
        r"\bmore\b|more information|this list of ingredients|actual ingredients|"
        r"ingredients subject|shipping policy|policies|"
        r"ingredient-list-copy|copy find dupes|find dupes|discover better matches|"
        r"key ingredients|ingredients explained|benefits|"
        r"product information|product details|details\s*/|"
        r"how to use|directions?|mode d’emploi|caution|warning|made in|catalog|sku|size"
    )
    patterns = [
        rf"major\s+ingredients?\s*[:：]?\s*(.{{40,9000}}?)(?={stop_words}|$)",
        rf"(?:ingredients?|ingr[eé]dients?)\s*[:：]?\s*(.{{40,9000}}?)(?={stop_words}|$)",
    ]
    candidates: list[str] = []
    for pattern in patterns:
        for match in re.finditer(pattern, text, flags=re.I):
            candidate = normalize_ingredients_text(match.group(1))
            if len(candidate) >= 40 and "," in candidate:
                candidates.append(candidate)
    candidates = [
        candidate
        for candidate in candidates
        if "manufacturer's discretion" not in candidate.lower()
        and "refer to product packaging" not in candidate.lower()
        and "subject to change" not in candidate.lower()
        and "{{" not in candidate
        and "productdata." not in candidate.lower()
        and "customer reviews" not in candidate.lower()
    ]
    if not candidates:
        return None
    ingredients = max(candidates, key=ingredient_candidate_score)
    if ingredient_candidate_score(ingredients) < 4:
        return None
    return ingredients_label(ingredients)


def ingredients_for_shade(text: str, product: str) -> str | None:
    shade_pattern = shade_regex(product)
    if not shade_pattern:
        return None
    ingredient_head = r"(?:major\s+ingredients?|ingredients?|ingr[eé]dients?)"
    start_pattern = rf"{ingredient_head}.*?{shade_pattern}\s*[:：]?\s*"
    stop_pattern = (
        r"(?=(?<![/\w])#?\s*(?:0?[1-9]|[a-z]{1,3}\s*\d{1,3})\s+(?-i:[A-Z][A-Za-z]+)|"
        r"\bmore\b|this list of ingredients|actual ingredients|ingredients subject|"
        r"ingredient-list-copy|copy find dupes|find dupes|discover better matches|"
        r"key ingredients|ingredients explained|benefits|"
        r"product information|shipping policy|"
        r"how to use|directions?|caution|warning|made in|$)"
    )
    candidates = []
    for match in re.finditer(start_pattern + rf"(.{{40,3000}}?){stop_pattern}", text, flags=re.I):
        candidate = normalize_ingredients_text(match.group(1))
        if len(candidate) >= 40 and "," in candidate and ingredient_candidate_score(candidate) >= 4:
            candidates.append(candidate)
    if candidates:
        return max(candidates, key=ingredient_candidate_score)
    return None


def ingredient_candidate_score(ingredients: str) -> int:
    low = ingredients.lower()
    inci_hits = [
        "ci ",
        "dimethicone",
        "isostearate",
        "talc",
        "mica",
        "silica",
        "wax",
        "oil",
        "aqua",
        "glyceryl",
        "polyglyceryl",
        "cyclopentasiloxane",
        "hexanediol",
    ]
    return sum(low.count(hit) for hit in inci_hits) + ingredients.count(",")


def normalize_ingredients_text(ingredients: str) -> str:
    ingredients = trim_non_ingredient_tail(ingredients)
    ingredients = re.sub(r"\bFormula\s+\d+\s*:\s*", ", ", ingredients, flags=re.I)
    ingredients = re.sub(r"\b(GL|PG|VT)\d{1,3}\b\s*", "", ingredients)
    ingredients = re.sub(r"(?<![A-Za-z])(?:C|W|N)\d{1,3}(?![-\w])\s*", "", ingredients)
    ingredients = re.sub(r"\s*<br\s*/?>\s*", ", ", ingredients, flags=re.I)
    ingredients = re.sub(r"\s*,\s*", ", ", ingredients)
    ingredients = re.sub(r"\s+", " ", ingredients).strip(" .;")
    ingredients = re.sub(r"\bwater\b", "Aqua", ingredients, flags=re.I)
    ingredients = re.sub(r"\bINCI\b\s*$", "", ingredients).strip(" .;,")
    ingredients = dedupe_adjacent_ingredients(ingredients)
    return ingredients


def trim_non_ingredient_tail(ingredients: str) -> str:
    markers = [
        "ingredient-list-copy",
        "copy find dupes",
        "find dupes",
        "discover better matches",
        "key ingredients",
        "ingredients explained",
        "show all",
        "skin conditioning, solvent",
        "supports skin hydration",
        "shields skin",
        "benefits hydrating",
        "class=\"inline-flex",
    ]
    low = ingredients.lower()
    cut_at = len(ingredients)
    for marker in markers:
        idx = low.find(marker)
        if idx != -1:
            cut_at = min(cut_at, idx)
    return ingredients[:cut_at].strip(" .;,/>")


def dedupe_adjacent_ingredients(ingredients: str) -> str:
    parts = split_ingredients(ingredients)
    cleaned: list[str] = []
    seen: set[str] = set()
    for part in parts:
        key = ingredient_dedupe_key(part)
        if key and key not in seen:
            cleaned.append(part.strip())
            seen.add(key)
    return ", ".join(cleaned)


def ingredient_dedupe_key(ingredient: str) -> str:
    key = re.sub(r"\s+", " ", ingredient).strip().lower()
    key = key.replace("ci ", "ci")
    key = re.sub(r"[^a-z0-9]+", "", key)
    return key


def ingredients_label(ingredients: str | None) -> str:
    if not ingredients:
        return "need to review"
    clean = normalize_ingredients_text(ingredients)
    if not clean:
        return "need to review"
    english, french = bilingual_ingredients(clean)
    return f"INGREDIENTS/INGRÉDIENTS: {english} / {french}"


def ingredients_label_from_known(known: dict[str, str]) -> str:
    if known.get("ingredients"):
        return ingredients_label(known.get("ingredients"))
    if known.get("ingredients_fr"):
        return ingredients_label(known.get("ingredients_fr"))
    return "need to review"


def bilingual_ingredients(ingredients: str) -> tuple[str, str]:
    clean = normalize_ingredients_text(ingredients)
    if looks_french_ingredients(clean):
        french = clean
        english = translate_ingredients(clean, FR_TO_EN_INGREDIENTS)
    else:
        english = clean
        french = translate_ingredients(clean, EN_TO_FR_INGREDIENTS)
    return english, french


def looks_french_ingredients(text: str) -> bool:
    low = text.lower()
    markers = ["diméthicone", "polymère", "croisé", "triisostéarate", "extrait", "huile", "acétate"]
    return any(marker in low for marker in markers)


EN_TO_FR_INGREDIENTS = {
    "Isostearyl Isostearate": "Isostéarate d’isostéaryle",
    "Polyglyceryl-2 Triisostearate": "Triisostéarate de polyglycéryle-2",
    "Diisostearyl Malate": "Malate de diisostéaryle",
    "Sorbitan Isostearate": "Isostéarate de sorbitan",
    "Paraffin": "Paraffine",
    "Trimethylpentaphenyl Trisiloxane": "Triméthylpentaphényl trisiloxane",
    "Microcrystalline Wax": "Cire microcristalline",
    "Pentaerythrityl Isostearate": "Isostéarate de pentaérythrityle",
    "Euphorbia Cerifera (Candelilla) Wax": "Cire d’Euphorbia cerifera (candelilla)",
    "1,2-Hexanediol": "1,2-hexanediol",
    "PEG/PPG-10/1 Dimethicone": "PEG/PPG-10/1 diméthicone",
    "Fragrance": "Parfum",
    "Pentaerythrityl Tetraisostearate": "Tétraisostéarate de pentaérythrityle",
    "Dimethicone": "Diméthicone",
    "Dimethicone Crosspolymer": "Polymère croisé de diméthicone",
    "Polydimethylsiloxane": "Polydiméthylsiloxane",
    "Cyclopentasiloxane": "Cyclopentasiloxane",
    "Cyclohexasiloxane": "Cyclohexasiloxane",
    "Silica Silylate": "Silylate de silice",
    "Cetyl PEG/PPG-10/1 Dimethicone": "Cétyl PEG/PPG-10/1 diméthicone",
    "Talc": "Talc",
    "Melaleuca Alternifolia Extract": "Extrait de Melaleuca alternifolia",
    "Evening Primrose Oil": "Huile d’onagre",
    "Alpha-Tocopheryl Acetate": "Acétate d’alpha-tocophéryle",
    "C30-45 Alkyl Dimethicone": "C30-45 Alkyl Diméthicone",
    "Talc": "Talc",
    "Mica": "Mica",
    "Synthetic Fluorphlogopite": "Fluorphlogopite synthétique",
    "Silica": "Silice",
    "Magnesium Myristate": "Myristate de magnésium",
    "Vinyl Dimethicone/Methicone Silsesquioxane Crosspolymer": "Polymère croisé vinyl diméthicone/méthicone silsesquioxane",
    "Octyldodecanol": "Octyldodécanol",
    "Isostearyl Neopentanoate": "Néopentanoate d’isostéaryle",
    "Methyl Methacrylate Crosspolymer": "Polymère croisé de méthacrylate de méthyle",
    "Aluminum Myristate": "Myristate d’aluminium",
    "Ethylhexylglycerin": "Éthylhexylglycérine",
    "Glyceryl Caprylate": "Caprylate de glycéryle",
    "Lauroyl Lysine": "Lauroyl lysine",
    "Triethoxycaprylylsilane": "Triéthoxycaprylylsilane",
    "Cocos Nucifera (Coconut) Oil": "Huile de Cocos nucifera (noix de coco)",
    "Alumina": "Alumine",
    "Aluminum Hydroxide": "Hydroxyde d’aluminium",
    "Boron Nitride": "Nitrure de bore",
    "Aqua": "Aqua",
}

FR_TO_EN_INGREDIENTS = {
    "Diméthicone": "Dimethicone",
    "polymère croisé de diméthicone": "Dimethicone Crosspolymer",
    "triisostéarate de polyglycéryle-2": "Polyglyceryl-2 Triisostearate",
    "cyclopentasiloxane": "Cyclopentasiloxane",
    "cyclohexasiloxane": "Cyclohexasiloxane",
    "1,2-hexanediol": "1,2-Hexanediol",
    "silylate de silice": "Silica Silylate",
    "cétyl PEG/PPG-10/1 diméthicone": "Cetyl PEG/PPG-10/1 Dimethicone",
    "talc": "Talc",
    "extrait de Melaleuca Alternifolia": "Melaleuca Alternifolia Extract",
    "huile d'onagre": "Evening Primrose Oil",
    "huile d’onagre": "Evening Primrose Oil",
    "acétate d'alpha-tocophéryle": "Alpha-Tocopheryl Acetate",
    "acétate d’alpha-tocophéryle": "Alpha-Tocopheryl Acetate",
    "C30-45 Alkyl Diméthicone": "C30-45 Alkyl Dimethicone",
}


def translate_ingredients(ingredients: str, mapping: dict[str, str]) -> str:
    parts = split_ingredients(ingredients)
    translated = []
    lower_map = {key.lower(): value for key, value in mapping.items()}
    for part in parts:
        translated.append(lower_map.get(part.lower(), mapping.get(part, part)))
    return ", ".join(translated)


def split_ingredients(ingredients: str) -> list[str]:
    protected = re.sub(r"(?<!\d)(\d),\s*(\d)(?=[-\w])", r"\1<COMMA>\2", ingredients)
    parts = [part.replace("<COMMA>", ",").strip() for part in protected.split(",")]
    return [part for part in parts if part]


def manufacturer_from_text(text: str) -> str | None:
    patterns = [
        r"manufacturer\s*[:：]\s*(.{3,120}?)(?:address|country|made in|$)",
        r"company\s+name\s*[:：]\s*(.{3,120}?)(?:country|address|business|$)",
        r"name\s+of\s+business/corporation\s*[:：]\s*(.{3,120}?)(?:address|$)",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.I)
        if match:
            value = re.sub(r"\s+", " ", match.group(1)).strip(" .;")
            if len(value) > 3:
                return value
    return None


def how_to_use_from_text(text: str) -> str | None:
    match = re.search(
        r"(?:how to use|directions?|usage)\s*[:：]?\s*(.{20,700}?)(?:ingredients?|net weight|official service|shipping|caution|warning|customer|$)",
        text,
        flags=re.I,
    )
    if not match:
        return None
    raw = re.sub(r"\s+", " ", match.group(1)).strip(" .:")
    if len(raw) < 15:
        return None
    return raw


def product_name_fr(product: str) -> str:
    clean = normalize_product_name(product)
    replacements = {
        "Glow Lipstick": "Rouge à lèvres éclat",
        "Glowing Lipstick": "Rouge à lèvres éclat",
        "Airy Lip Cheek Mud": "Baume mat aérien lèvres et joues",
        "Six-color Blush Palette": "Palette de fards à joues six couleurs",
        "Six-Color Blush Palette": "Palette de fards à joues six couleurs",
        "Liquid Blush Serum": "Sérum blush liquide",
        "Lip Cheek Mud": "Baume mat lèvres et joues",
        "Liquid Blush": "Blush liquide",
        "Blush Palette": "Palette de fards à joues",
        "Lipstick": "Rouge à lèvres",
        "Lip": "Lèvres",
        "Cheek": "Joues",
    }
    result = clean
    for en, fr in replacements.items():
        result = re.sub(en, fr, result, flags=re.I)
    return result


def normalize_product_name(product: str) -> str:
    return re.sub(r"\s+", " ", str(product).replace("\xa0", " ")).strip()


def searchable_text(value: str) -> str:
    text = normalize_product_name(value).lower()
    text = re.sub(r"[\-_()/|]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def search_tokens(product: str) -> list[str]:
    stopwords = {"the", "and", "for", "with", "colors", "color", "set", "pcs", "pc"}
    tokens = re.findall(r"[\w\u3040-\u30ff\u3400-\u9fff\uac00-\ud7af]+", searchable_text(product))
    return [token for token in tokens if len(token) > 1 and token not in stopwords]


def shade_key(product: str) -> str:
    clean = searchable_text(product)
    match = re.search(r"(?:#\s*)?([a-z]{0,3}\s*\d{1,3})\s+([a-z][a-z0-9]+(?:\s+[a-z][a-z0-9]+){0,2})", clean)
    if match:
        code = re.sub(r"\s+", "", match.group(1))
        shade = re.sub(r"\s+", "-", match.group(2).strip())
        return f"{code}-{shade}"
    match = re.search(r"(?:#\s*)?(\d{1,3})\s+([a-z][a-z0-9]+(?:\s+[a-z][a-z0-9]+){0,2})", clean)
    if match:
        shade = re.sub(r"\s+", "-", match.group(2).strip())
        return f"{match.group(1)}-{shade}"
    return ""


def shade_regex(product: str) -> str:
    key = shade_key(product)
    if not key:
        return ""
    code, _, shade = key.partition("-")
    code_pattern = r"\s*".join(re.escape(ch) for ch in code)
    shade_pattern = r"\s+".join(re.escape(part) for part in shade.split("-") if part)
    if not shade_pattern:
        return ""
    return rf"#?\s*{code_pattern}\s+{shade_pattern}"


def product_brand(product: str) -> str:
    tokens = search_tokens(product)
    if not tokens:
        return ""
    if len(tokens) >= 2 and tokens[0] in {"into", "fwee"}:
        return " ".join(tokens[:2]) if tokens[0] == "into" else tokens[0]
    return " ".join(tokens[:2])


def fuzzy_queries(barcode: str, product: str) -> list[str]:
    clean_product = normalize_product_name(product)
    tokens = search_tokens(clean_product)
    brand = product_brand(clean_product)
    core = " ".join(tokens[:6])
    tail = " ".join(tokens[-5:])
    queries = [
        f"\"{clean_product}\"",
        f"\"{clean_product}\" ingredients",
        f"\"{clean_product}\" net weight",
        f"{clean_product} ingredients net weight",
        f"{clean_product} how to use ingredients",
        f"{clean_product} official",
        f"{clean_product} YesStyle",
        f"{clean_product} Kiseki",
        f"{brand} {tail} ingredients" if brand and tail else "",
        f"{core} site:yesstyle.com",
        f"{core} site:judydoll.com",
        f"{core} site:kiseki.ca",
        f"{core} site:asianbeautywholesale.com",
    ]
    if barcode:
        queries[3:3] = [
            f"{barcode} {clean_product}",
            f"{barcode} ingredients net weight",
            f"\"{barcode}\"",
        ]
    if re.search(r"[\u3040-\u30ff\u3400-\u9fff\uac00-\ud7af]", clean_product):
        queries.extend(
            [
                f"{clean_product} ingredients English",
                f"{clean_product} 成分 容量",
                f"{clean_product} 全成分 容量",
                f"{clean_product} 전성분 용량",
            ]
        )
    for hint in SEARCH_LANGUAGE_HINTS:
        if brand and hint:
            queries.append(f"{brand} {hint} {tail}".strip())
    deduped: list[str] = []
    seen = set()
    for query in queries:
        query = re.sub(r"\s+", " ", query).strip()
        if query and query not in seen:
            seen.add(query)
            deduped.append(query)
    return deduped[:24]


def fuzzy_source_score(url: str, title_or_text: str, barcode: str, product: str) -> float:
    if is_noise_url(url):
        return -10.0
    haystack = searchable_text(f"{url} {title_or_text}")
    tokens = search_tokens(product)
    score = 0.0
    if barcode and barcode in haystack:
        score += 8
    domain_rank = source_rank(url)
    if domain_rank < 100:
        score += 6 - min(domain_rank, 5) * 0.5
    score += product_detail_bonus(url)
    for token in tokens:
        if token in haystack:
            score += 1
    shade = shade_key(product)
    if shade and shade in haystack:
        score += 5
    elif shade:
        score -= 2
    if "ingredients" in haystack or "major ingredients" in haystack:
        score += 2
    if "net weight" in haystack or re.search(r"\b\d+(?:\.\d+)?\s*(g|ml)\b", haystack):
        score += 1.5
    product_norm = searchable_text(product)
    if product_norm and haystack:
        score += difflib.SequenceMatcher(None, product_norm[:120], haystack[:240]).ratio() * 4
    return score


def direction_for_product(product: str, source_direction: str | None) -> tuple[str, str]:
    low = product.lower()
    if "blush serum" in low or ("liquid" in low and "blush" in low):
        en = "DIRECTION FOR USE: Lightly dab after foundation and before setting powder. Tap and blend evenly across cheeks."
        fr = "MODE D’EMPLOI: Tapoter légèrement après le fond de teint et avant la poudre fixatrice. Estomper uniformément sur les joues."
    elif "highlighter" in low:
        en = "DIRECTION FOR USE: Apply to high points of the face with a brush. Layer as desired for added glow."
        fr = "MODE D’EMPLOI: Appliquer sur les points saillants du visage à l’aide d’un pinceau. Superposer au besoin pour plus d’éclat."
    elif "lip" in low and "cheek" in low:
        en = DEFAULT_DIRECTION_EN
        fr = DEFAULT_DIRECTION_FR
    elif "lipstick" in low or "lip" in low:
        en = "DIRECTION FOR USE: Apply directly to lips. Reapply as needed."
        fr = "MODE D’EMPLOI: Appliquer directement sur les lèvres. Réappliquer au besoin."
    elif "blush" in low or "palette" in low:
        en = "DIRECTION FOR USE: Apply to cheeks with a brush. Mix, match, and layer shades as desired."
        fr = "MODE D’EMPLOI: Appliquer sur les joues à l’aide d’un pinceau. Mélanger, assortir et superposer les teintes au besoin."
    else:
        en = "DIRECTION FOR USE: Apply a proper amount to the desired area. Use as directed."
        fr = "MODE D’EMPLOI: Appliquer une quantité appropriée sur la zone souhaitée. Utiliser selon le mode d’emploi."
    if source_direction and "refriger" in source_direction.lower():
        en = (
            "DIRECTION FOR USE: Apply a small amount directly to lips. Store below 25°C "
            "and refrigerate if the product softens."
        )
        fr = (
            "MODE D’EMPLOI: Appliquer une petite quantité directement sur les lèvres. "
            "Conserver à moins de 25 °C et réfrigérer si le produit ramollit."
        )
    return en, fr


def default_cautions() -> tuple[str, str]:
    return DEFAULT_CAUTION_EN, DEFAULT_CAUTION_FR


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


def fill_from_builtin_reference(barcode: str) -> dict[str, str] | None:
    return builtin_reference_data().get(str(barcode).strip().replace(".0", ""))


def missing_required_fields(values: dict[str, str]) -> list[str]:
    missing = []
    for field in REQUIRED_LABEL_FIELDS:
        value = values.get(field)
        if value in (None, "", "need to review"):
            missing.append(field)
    return missing


def is_input_row_blank(barcode: str, product: str) -> bool:
    return not str(barcode or "").strip() and not str(product or "").strip()


def clear_generated_row(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    headers: dict[str, int],
    row_idx: int,
) -> None:
    generated_headers = REQUIRED_LABEL_FIELDS + [
        "source url",
    ]
    for header in generated_headers:
        col = find_column(headers, header)
        if col:
            ws.cell(row_idx, col).value = None


def candidate_urls(barcode: str, product: str) -> list[str]:
    clean_product = normalize_product_name(product)
    candidates: list[str] = []
    if barcode == "1129343972":
        candidates.extend(
            [
                "https://www.yesstyle.com/en/into-you-glowing-lipstick-8-colors-gl08-red-brown-3g/info.html/pid.1129343972",
                "https://www.asianbeautywholesale.com/en/into-you-glowing-lipstick-8-colors-gl08-red-brown-3g/info.html/pid.1129343972",
                "https://www.intoyoucosmetics.com/en-ca/products/into-you-glow-lipstick?variant=57958599524655",
                "https://www.intoyoucosmetics.com/en-ca/products/into-you-glow-lipstick?variant=57958599590191",
                "https://www.uniquebunny.com/products/into-you-glow-lipstick",
            ]
        )
    if barcode == "1126245093":
        candidates.extend(
            [
                "https://www.yesstyle.com/en/into-you-airy-lip-cheek-mud/info.html/pid.1126245093",
                "https://www.asianbeautywholesale.com/en/into-you-airy-lip-cheek-mud/info.html/pid.1126245093",
                "https://www.uniquebunny.com/fr/products/into-you-airy-lip-cheek-mud",
            ]
        )
    if "into" in clean_product.lower() and "glow" in clean_product.lower() and "lipstick" in clean_product.lower():
        candidates.append("https://www.intoyoucosmetics.com/en-ca/products/into-you-glow-lipstick?variant=57958599524655")
        candidates.append("https://www.intoyoucosmetics.com/en-ca/products/into-you-glow-lipstick?variant=57958599590191")
        candidates.append("https://www.uniquebunny.com/products/into-you-glow-lipstick")
    if "into" in clean_product.lower() and "airy" in clean_product.lower() and "lip" in clean_product.lower():
        candidates.append("https://www.intoyoucosmetics.com/en-ca/products/airy-lip-mud")
        candidates.append("https://www.uniquebunny.com/fr/products/into-you-airy-lip-cheek-mud")
    if "into" in clean_product.lower() and "six" in clean_product.lower() and "blush" in clean_product.lower():
        candidates.append("https://www.yesstyle.com/en/into-you-six-color-blush-palette-six-color-blush-palette-15g/info.html/pid.1137202898")
    if "judydoll" in clean_product.lower() and "liquid" in clean_product.lower() and "blush" in clean_product.lower():
        candidates.extend(
            [
                "https://judydoll.com/products/liquid-blush-serum",
                "https://www.yesstyle.com/en/judydoll-liquid-blush-serum-4-colors-01-fig-5g/info.html/pid.1136648925",
                "https://www.kiseki.ca/judydoll-liquid-blush.html",
                "https://joybeautyhub.shop/products/liquid-blush-serum",
            ]
        )
    if "millefee" in clean_product.lower() and "idol" in clean_product.lower() and "highlighter" in clean_product.lower():
        candidates.append("https://millefee.com/products/idol-highlighter-palette")
        if "rose" in clean_product.lower() or barcode == "1137196649":
            candidates.append("https://www.yesstyle.com/en/millefee-idol-highlighter-palette-02-rose-pink/info.html/pid.1137196649")
        if "ice" in clean_product.lower() or barcode == "1137196647":
            candidates.append("https://www.yesstyle.com/en/millefee-idol-highlighter-palette-01-ice-blue/info.html/pid.1137196647")

    scored: list[tuple[float, str]] = []
    for url in candidates:
        scored.append((fuzzy_source_score(url, url, barcode, clean_product), url))

    strong_static_candidates = any(
        exact_product_url(url, clean_product) or product_detail_bonus(url) >= 6 for url in candidates
    )
    if not strong_static_candidates:
        for query in fuzzy_queries(barcode, clean_product)[:12]:
            for result in search_web(query):
                url = result["url"]
                score = fuzzy_source_score(url, result.get("title", ""), barcode, clean_product)
                if score >= 7 or exact_product_url(url, clean_product):
                    scored.append((score, url))

    deduped: list[str] = []
    seen = set()
    for _score, url in sorted(scored, key=lambda item: (-item[0], source_rank(item[1]))):
        if not url or url in seen:
            continue
        seen.add(url)
        deduped.append(url)
    return deduped[:10]


def known_product_fallback(barcode: str, product: str) -> dict[str, str]:
    known = KNOWN_ONLINE_PRODUCTS.get(barcode)
    if known:
        return known
    clean = searchable_text(product)
    if "into" in clean and "glow" in clean and "lipstick" in clean:
        return KNOWN_ONLINE_PRODUCTS["1129343972"]
    if "into" in clean and "airy" in clean and "lip" in clean and ("cheek" in clean or "mud" in clean):
        return KNOWN_ONLINE_PRODUCTS["1126245093"]
    if "into" in clean and "six" in clean and "blush" in clean and "palette" in clean:
        return KNOWN_ONLINE_PRODUCTS["1137202898"]
    if "judydoll" in clean and "liquid" in clean and "blush" in clean:
        shade = shade_key(product)
        ingredients = JUDYDOLL_LIQUID_BLUSH_SERUM["shades"].get(shade)
        if ingredients:
            return {
                "source_url": JUDYDOLL_LIQUID_BLUSH_SERUM["source_url"],
                "net weight": JUDYDOLL_LIQUID_BLUSH_SERUM["net weight"],
                "source_direction": JUDYDOLL_LIQUID_BLUSH_SERUM["source_direction"],
                "coo": JUDYDOLL_LIQUID_BLUSH_SERUM["coo"],
                "ingredients": ingredients,
            }
    if "millefee" in clean and "idol" in clean and "highlighter" in clean:
        shade = MILLEFEE_IDOL_HIGHLIGHTER_PALETTE["barcodes"].get(str(barcode).strip())
        shade = shade or shade_key(product)
        shade_data = MILLEFEE_IDOL_HIGHLIGHTER_PALETTE["shades"].get(shade)
        if shade_data:
            return {
                "source_url": shade_data["source_url"],
                "net weight": MILLEFEE_IDOL_HIGHLIGHTER_PALETTE["net weight"],
                "source_direction": MILLEFEE_IDOL_HIGHLIGHTER_PALETTE["source_direction"],
                "coo": MILLEFEE_IDOL_HIGHLIGHTER_PALETTE["coo"],
                "ingredients": shade_data["ingredients"],
            }
    return {}


def enough_product_data(texts: list[tuple[str, str]], product: str, known: dict[str, str]) -> bool:
    if len(texts) >= 4:
        return True
    if len(texts) < 2:
        return False
    combined = " ".join(text for _url, text in texts)
    has_net = bool(net_weight_from_text(combined) or known.get("net weight") or net_weight_from_name(product))
    has_ingredients = bool(
        ingredients_from_text(combined, product)
        or (ingredients_label_from_known(known) != "need to review")
    )
    return has_net and has_ingredients


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
        notes.append("Matched reference data by barcode.")
        missing = missing_required_fields(values)
        for field in missing:
            values[field] = "need to review"
        status = "Completed" if not missing else "Need to review"
        if missing:
            notes.append("Missing required fields: " + ", ".join(missing) + ".")
        return FillResult(
            values,
            status,
            "Built-in reference data",
            " ".join(notes),
        )

    known = known_product_fallback(barcode, product)
    texts: list[tuple[str, str]] = []
    for url in candidate_urls(barcode, product):
        text = fetch_text(url)
        if len(text) > 200:
            texts.append((url, text))
            if enough_product_data(texts, product, known):
                break
    if texts:
        source_url = "\n".join(url for url, _text in texts[:4])
        notes.append("Sources checked: " + ", ".join(domain_from_url(url) for url, _text in texts[:4]))
    else:
        notes.append("No reliable source found.")

    if known and not source_url:
        source_url = known.get("source_url", "")
        notes.append("Used approved source URLs for this SKU.")

    combined_text = " ".join(text for _url, text in texts)
    source_direction = how_to_use_from_text(combined_text) or known.get("source_direction")
    direction_en, direction_fr = direction_for_product(product, source_direction)
    caution_en, caution_fr = default_cautions()

    values["product name french"] = product_name_fr(product)
    values["net weight"] = (
        net_weight_from_text(combined_text)
        or known.get("net weight")
        or net_weight_from_name(product)
        or "need to review"
    )
    values["direction for use"] = direction_en
    values["mode d’emploi"] = direction_fr
    values["cautions"] = caution_en
    values["mises en garde:"] = caution_fr
    values["manufacturer"] = manufacturer_from_text(combined_text) or known.get("manufacturer") or "need to review"
    values["ingredients/ingrédients"] = (
        ingredients_from_text(combined_text, product)
        or ingredients_label_from_known(known)
        or "need to review"
    )
    values["coo"] = coo_from_text(combined_text + " " + product) or known.get("coo") or "need to review"
    values["distributed by / distribué par:"] = DISTRIBUTOR

    for field in REQUIRED_LABEL_FIELDS:
        if field not in values:
            values[field] = "need to review"

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
    fill_sheet: str,
    use_defaults: bool,
    limit: int | None = None,
) -> tuple[openpyxl.Workbook, pd.DataFrame]:
    wb = openpyxl.load_workbook(path)
    fill_ws = wb[fill_sheet]
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
        if is_input_row_blank(barcode, product):
            clear_generated_row(fill_ws, fill_headers, row_idx)
            continue
        row = {"barcode": barcode, "product name": product}
        reference_values = fill_from_builtin_reference(barcode)
        result = process_row(row, reference_values, use_defaults)

        for header, value in result.values.items():
            col = find_column(fill_headers, header)
            if col and value:
                fill_ws.cell(row_idx, col).value = value
                if row_idx > 2:
                    copy_cell_style(fill_ws.cell(row_idx - 1, col), fill_ws.cell(row_idx, col))

        fill_ws.cell(row_idx, fill_headers["source url"]).value = result.source_url
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
    fill_sheet = st.selectbox(
        "Sheet to fill",
        names,
        index=names.index("Sheet2") if "Sheet2" in names else 0,
    )

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
