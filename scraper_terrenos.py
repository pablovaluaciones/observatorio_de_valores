#!/usr/bin/env python3
"""Extrae todos los inmuebles en venta de Roca y Agostini y crea un Excel.

Instalación: py -m pip install requests beautifulsoup4 openpyxl geopandas
Ejecución:   py outputs/scraper_terrenos.py
"""

from __future__ import annotations

import re
import sys
import time
import unicodedata
from datetime import datetime
from pathlib import Path
from urllib.parse import unquote, urljoin, urlparse
from urllib.robotparser import RobotFileParser

import requests
from bs4 import BeautifulSoup
import geopandas as gpd
from openpyxl import Workbook
from shapely.geometry import Point


SOURCES = {
    "Roca Inmobiliaria": "https://www.rocavende.com/estado/en-venta/",
    "Agostini Inmobiliaria": "https://agostiniinmobiliaria.com/propiedades-venta/",
}
DELAY_SECONDS = 1.5
MAX_RECORDS_PER_SOURCE = 15  # Prueba inicial: hasta 15 avisos por inmobiliaria.
HEADERS = {
    "User-Agent": "TerrenosResearch/1.0 (+contacto: tu-email@example.com)",
    "Accept-Language": "es-AR,es;q=0.9",
}


def clean(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


def robots_allowed(url: str) -> bool:
    parsed = urlparse(url)
    robot_url = f"{parsed.scheme}://{parsed.netloc}/robots.txt"
    robots = RobotFileParser(robot_url)
    try:
        robots.read()
    except OSError as error:
        print(f"AVISO: no se pudo leer {robot_url}: {error}. Fuente omitida.")
        return False
    allowed = robots.can_fetch(HEADERS["User-Agent"], url)
    if not allowed:
        print(f"AVISO: robots.txt no permite consultar {url}. Fuente omitida.")
    return allowed


def request(session: requests.Session, url: str) -> requests.Response:
    time.sleep(DELAY_SECONDS)
    response = session.get(url, timeout=30)
    response.raise_for_status()
    return response


def collect_pages(soup: BeautifulSoup, current_url: str) -> set[str]:
    """Obtiene URLs de paginación sin asumir un tema de WordPress concreto."""
    domain = urlparse(current_url).netloc
    pages = {current_url}
    for link in soup.select("a[href]"):
        url = urljoin(current_url, link["href"])
        label = clean(link.get_text(" ", strip=True))
        if urlparse(url).netloc == domain and (re.search(r"(?:page|paged)[=/]\d+", url) or label.isdigit()):
            pages.add(url)
    return pages


def collect_details(soup: BeautifulSoup, listing_url: str) -> set[str]:
    """Roca etiqueta los enlaces como Detalle y Agostini usa /propiedad/."""
    domain = urlparse(listing_url).netloc
    details = set()
    for link in soup.select("a[href]"):
        url = urljoin(listing_url, link["href"]).split("#")[0]
        label = clean(link.get_text(" ", strip=True)).lower()
        if urlparse(url).netloc != domain:
            continue
        if "/propiedad/" in urlparse(url).path or label in {"detalle", "detalles", "ver detalle"}:
            details.add(url)
    return details


def first_match(pattern: str, text: str) -> str:
    match = re.search(pattern, text, flags=re.IGNORECASE)
    return clean(match.group(1)) if match else ""


def price(text: str) -> tuple[str, float | None]:
    match = re.search(r"(USD|US\$|\$)\s*(\d[\d.,]*)", text, flags=re.IGNORECASE)
    if not match:
        return "", None
    currency = "USD" if match.group(1).upper() in {"USD", "US$"} else "ARS"
    return currency, localized_number(match.group(2))


def area(text: str) -> float | None:
    candidates = re.findall(r"(\d{1,7}(?:[.,]\d+)?)\s*(?:m²|m2)\b", text, flags=re.IGNORECASE)
    values = []
    for candidate in candidates:
        try:
            number = localized_number(candidate)
            if number > 0:
                values.append(number)
        except ValueError:
            pass
    return max(values) if values else None


def localized_number(value: str) -> float:
    """Convierte 6172,00, 6172.00, 6.172,00 y 6,172.00 sin perder decimales."""
    value = value.strip().replace(" ", "")
    commas, dots = value.count(","), value.count(".")
    if commas and dots:
        decimal_separator = "," if value.rfind(",") > value.rfind(".") else "."
        thousands_separator = "." if decimal_separator == "," else ","
        return float(value.replace(thousands_separator, "").replace(decimal_separator, "."))
    separator = "," if commas else "." if dots else ""
    if separator:
        whole, fraction = value.rsplit(separator, 1)
        # Uno o dos dígitos posteriores al separador representan decimales.
        if len(fraction) <= 2:
            return float(f"{whole}.{fraction}")
        return float(value.replace(separator, ""))
    return float(value)


def section_text(soup: BeautifulSoup, keyword: str, limit: int = 1800) -> str:
    header = next((x for x in soup.find_all(["h2", "h3", "h4", "strong"])
                   if keyword in clean(x.get_text(" ", strip=True)).lower()), None)
    if not header:
        return ""
    parts = []
    for element in header.find_all_next(limit=14):
        if element is not header and element.name in {"h2", "h3", "h4"}:
            break
        value = clean(element.get_text(" ", strip=True))
        if value and value not in parts:
            parts.append(value)
    return clean(" ".join(parts))[:limit]


def publication_id(source: str, url: str, text: str) -> str:
    """Usa el ID publicado; si no existe, genera uno estable a partir de la URL."""
    listed_id = first_match(r"ID(?:\s+de\s+propiedad)?\s*:?\s*([A-Za-z0-9_-]+)", text)
    slug = urlparse(url).path.strip("/").split("/")[-1]
    source_code = "roca" if source.startswith("Roca") else "agostini"
    return f"{source_code}-{listed_id or slug or urlparse(url).netloc}"


def property_type(title: str, text: str) -> str:
    """Normaliza los tipos distintos de las dos inmobiliarias en categorías comparables."""
    type_from_page = first_match(r"([^\n]{1,80})\n\s*Tipo de propiedad", text)
    reference = f"{title} {type_from_page} {text[:2000]}".lower()
    if any(term in reference for term in ("terreno", "lote", "loteo", "finca")):
        return "Terreno"
    if any(term in reference for term in ("galpón", "galpon", "nave industrial", "depósito", "deposito")):
        return "Galpón"
    if any(term in reference for term in ("local", "oficina", "consultorio", "comercio", "negocio")):
        return "Negocio"
    if any(term in reference for term in ("casa", "departamento", "depto", "vivienda", "duplex", "dúplex", "ph ", "chalet")):
        return "Vivienda"
    return "Otro"


def roca_type(title: str) -> str:
    """Clasificación pedida para Roca: sólo a partir del título del aviso."""
    normalized = "".join(
        char for char in unicodedata.normalize("NFD", title.lower()) if unicodedata.category(char) != "Mn"
    )
    if "terreno" in normalized:
        return "terreno"
    if "lote" in normalized:
        return "lote"
    if "casa" in normalized or "vivienda" in normalized:
        return "casa"
    if "galpon" in normalized:
        return "galpon"
    if "finca" in normalized:
        return "finca"
    if "departamento" in normalized:
        return "departamento"
    return ""


def roca_publication_id(soup: BeautifulSoup, url: str, text: str) -> str:
    """Extrae el ID mostrado por Roca en div.fw-property-amenities-data."""
    for block in soup.select("div.fw-property-amenities-data"):
        block_text = clean(block.get_text(" ", strip=True))
        match = re.search(r"\bID\b\s*([0-9]+)\b", block_text, flags=re.IGNORECASE)
        if match:
            return f"roca-{match.group(1)}"
    return publication_id("Roca Inmobiliaria", url, text)


def roca_neighborhood(soup: BeautifulSoup) -> str:
    address = soup.select_one("address.item-address")
    if not address:
        return ""
    # La primera parte de "Huacalera, Huacalera, Tilcara" es la zona/barrio del aviso.
    return clean(address.get_text(" ", strip=True)).split(",")[0].strip()


def valid_coordinates(latitude: str, longitude: str) -> tuple[float, float] | None:
    try:
        lat, lon = float(latitude), float(longitude)
    except ValueError:
        return None
    return (lat, lon) if -90 <= lat <= 90 and -180 <= lon <= 180 else None


def extract_map_data(soup: BeautifulSoup, html: str) -> dict[str, object]:
    """Obtiene coordenadas publicadas en mapas Google/Leaflet de la ficha."""
    map_urls = []
    for element in soup.select("iframe[src], a[href]"):
        url = element.get("src") or element.get("href")
        if url and any(name in url.lower() for name in ("google.com/maps", "maps.google", "leaflet", "openstreetmap")):
            map_urls.append(url)
    map_url = map_urls[0] if map_urls else ""
    provider = ""
    searchable = unquote("\n".join(map_urls + [html]))
    lower = searchable.lower()
    if "google" in lower and "map" in lower:
        provider = "Google Maps"
    elif "leaflet" in lower:
        provider = "Leaflet"
    elif "openstreetmap" in lower:
        provider = "OpenStreetMap"

    patterns = (
        # URLs Google Maps: .../@-24.185,-65.299,17z or ?q=-24.185,-65.299
        r"@\s*(-?\d{1,2}(?:\.\d+)?)\s*,\s*(-?\d{1,3}(?:\.\d+)?)",
        r"(?:q|query|ll|center)=\s*(-?\d{1,2}(?:\.\d+)?)\s*,\s*(-?\d{1,3}(?:\.\d+)?)",
        # Leaflet / Google JavaScript markers.
        r"(?:L\.marker|setView|LatLng)\s*\(\s*\[?\s*(-?\d{1,2}(?:\.\d+)?)\s*,\s*(-?\d{1,3}(?:\.\d+)?)",
        # JSON de Houzez y otras variables JavaScript. Houzez publica: "lat":"-24...","lng":"-65...".
        r"(?:\"lat(?:itude)?\"|\blat(?:itude)?)\s*[:=]\s*[\"']?(-?\d{1,2}(?:\.\d+)?)[\"']?.{0,2000}?(?:\"(?:lng|lon|longitude)\"|\b(?:lng|lon|longitude))\s*[:=]\s*[\"']?(-?\d{1,3}(?:\.\d+)?)[\"']?",
    )
    for pattern in patterns:
        match = re.search(pattern, searchable, flags=re.IGNORECASE | re.DOTALL)
        if match:
            coordinates = valid_coordinates(match.group(1), match.group(2))
            if coordinates:
                return {
                    "Latitud": coordinates[0],
                    "Longitud": coordinates[1],
                    "Proveedor mapa": provider,
                    "URL mapa": map_url,
                    "Precisión ubicación": "Publicada en mapa; sin validar",
                }
    return {
        "Latitud": None,
        "Longitud": None,
        "Proveedor mapa": provider,
        "URL mapa": map_url,
        "Precisión ubicación": "Mapa sin coordenadas legibles" if map_url else "Sin mapa detectado",
    }


def parse_detail(source: str, url: str, html: str) -> dict[str, object]:
    soup = BeautifulSoup(html, "html.parser")
    text = soup.get_text("\n", strip=True)
    title_tag = soup.find("h1") or soup.find("title")
    title = clean(title_tag.get_text(" ", strip=True)) if title_tag else ""
    currency, numeric_price = price(text)
    is_roca = source.startswith("Roca")
    row = {
        "Fuente": source,
        "Título": title,
        "ID publicación": roca_publication_id(soup, url, text) if is_roca else publication_id(source, url, text),
        "Tipo": roca_type(title) if is_roca else property_type(title, text),
        "Moneda": currency,
        "Precio": numeric_price,
        "Superficie (m²)": area(text),
        "Barrio": roca_neighborhood(soup) if is_roca else first_match(r"Barrio\s*:?\s*([^\n]{1,120})", text),
        "Localidad": first_match(r"Localidad\s*:?\s*([^\n]{1,120})", text),
        "Servicios": section_text(soup, "caracter") or section_text(soup, "servicios"),
        "Descripción": section_text(soup, "descripci", 4000),
        "Actualizado en sitio": first_match(r"Actualizado(?:\s+el\s+d[ií]a)?\s*:?\s*([^\n]{1,80})", text),
        "URL": url,
        "Relevado el": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }
    row.update(extract_map_data(soup, html))
    return row


def scrape_source(session: requests.Session, source: str, start_url: str) -> list[dict[str, object]]:
    if not robots_allowed(start_url):
        return []
    first_page = request(session, start_url)
    pages = sorted(collect_pages(BeautifulSoup(first_page.text, "html.parser"), first_page.url))[:100]
    detail_urls = set()
    for page_url in pages:
        page = first_page if page_url == first_page.url else request(session, page_url)
        detail_urls.update(collect_details(BeautifulSoup(page.text, "html.parser"), page.url))
    print(f"{source}: {len(detail_urls)} fichas encontradas.")
    rows = []
    for position, detail_url in enumerate(sorted(detail_urls), start=1):
        try:
            detail = request(session, detail_url)
            row = parse_detail(source, detail.url, detail.text)
            rows.append(row)
            print(f"  {position}/{len(detail_urls)} {row['Título'][:70]}")
            if len(rows) >= MAX_RECORDS_PER_SOURCE:
                break
        except requests.RequestException as error:
            print(f"  ERROR: {detail_url} ({error})")
    return rows


def export_excel(rows: list[dict[str, object]], output: Path) -> None:
    columns = ["Fuente", "Título", "ID publicación", "Tipo", "Moneda", "Precio",
               "Superficie (m²)", "Barrio", "Localidad", "Servicios",
               "Descripción", "Latitud", "Longitud", "Proveedor mapa", "URL mapa", "Precisión ubicación",
               "Actualizado en sitio", "URL", "Relevado el"]
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Inmuebles en venta"
    sheet.append(columns)
    for item in rows:
        sheet.append([item[column] for column in columns])
    workbook.save(output)


def export_geopackage(rows: list[dict[str, object]], output: Path) -> int:
    """Exporta los avisos con coordenadas válidas como puntos POSGAR 2007."""
    georeferenced_rows = []
    for item in rows:
        coordinates = valid_coordinates(str(item.get("Latitud", "")), str(item.get("Longitud", "")))
        if coordinates:
            latitude, longitude = coordinates
            georeferenced_rows.append({
                "latitud": latitude,
                "longitud": longitude,
                "URL": item["URL"],
                "id": len(georeferenced_rows) + 1,
                "geometry": Point(longitude, latitude),
            })

    if not georeferenced_rows:
        print("AVISO: no hay inmuebles con coordenadas válidas; no se creó el GeoPackage.")
        return 0

    points = gpd.GeoDataFrame(georeferenced_rows, geometry="geometry", crs="EPSG:4326")
    points = points.to_crs("EPSG:5345")
    points.to_file(output, layer="inmuebles", driver="GPKG", index=False)
    return len(points)


def main() -> int:
    session = requests.Session()
    session.headers.update(HEADERS)
    rows = []
    for source, start_url in SOURCES.items():
        try:
            rows.extend(scrape_source(session, source, start_url))
        except requests.RequestException as error:
            print(f"ERROR en {source}: {error}")
    unique_rows = list({row["URL"]: row for row in rows}.values())
    output_dir = Path(__file__).parent / "output_files"
    output_dir.mkdir(exist_ok=True)
    output = output_dir / f"inmuebles_en_venta_{datetime.now():%Y-%m-%d_%H%M}.xlsx"
    export_excel(unique_rows, output)
    print(f"Excel creado: {output.resolve()} ({len(unique_rows)} inmuebles)")
    geopackage = output.with_suffix(".gpkg")
    georeferenced_count = export_geopackage(unique_rows, geopackage)
    if georeferenced_count:
        print(f"GeoPackage creado: {geopackage.resolve()} ({georeferenced_count} puntos)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
