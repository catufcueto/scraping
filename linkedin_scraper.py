"""
LinkedIn HR Decision-Maker Scraper
====================================
Busca decisores de RRHH en LinkedIn navegando a la página de empresa,
pestaña Personas, con filtro por Argentina.

Cargos buscados:
  Jefe de Recursos Humanos / Gerente de RRHH / Gerente de Capital Humano /
  People Manager / Director de Recursos Humanos / Director de Talento /
  CHRO / CPO / VP of People / People Operations / Talent Acquisition /
  Employee Experience / Culture Manager / Head of People

Uso:
  1. Primera vez:  python linkedin_scraper.py --login
  2. Test 2 emp:   python linkedin_scraper.py --input empresas.xlsx --test
  3. Completo:     python linkedin_scraper.py --input empresas.xlsx

Output: linkedin_rrhh_YYYYMMDD_HHMMSS.xlsx
"""

import asyncio
import json
import logging
import random
import re
import sys
import urllib.parse
from collections import Counter
from datetime import datetime
from pathlib import Path
from typing import Optional

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from playwright.async_api import BrowserContext, Page, async_playwright

# ─────────────────────────────────────────────
# CONFIGURACIÓN
# ─────────────────────────────────────────────

COOKIES_FILE = Path("linkedin_cookies.json")
OUTPUT_DIR   = Path(".")

TARGET_TITLES = [
    # ── Español — responsables de RRHH ───────────────────────────────────
    "Jefe de Recursos Humanos",
    "Jefa de Recursos Humanos",
    "Gerente de Recursos Humanos",
    "Gerente de RRHH",
    "Gerente de Capital Humano",
    "Gerente de Desarrollo Humano",
    "Jefe de Capital Humano",
    "Jefe de Desarrollo Humano",
    "Responsable de Recursos Humanos",
    "Responsable de RRHH",
    "Coordinador de Recursos Humanos",
    "Director de Recursos Humanos",
    "Director de Capital Humano",
    "Director de Talento",
    "Director de Desarrollo Humano",
    # ── Inglés / mixto ────────────────────────────────────────────────────
    "People Manager",
    "HR Manager",
    "HR Director",
    "HR Business Partner",
    "HRBP",
    "Chief Human Resources Officer",
    "Chief People Officer",
    "VP of People",
    "Head of People",
    "People Operations",
    "Talent Acquisition",
    "Talent Acquisition Manager",
    "Employee Experience",
    "Culture Manager",
]

DELAY_SHORT  = (1.5, 3.0)
DELAY_MEDIUM = (3.0, 6.0)
DELAY_LONG   = (6.0, 12.0)

MAX_RESULTS_PER_COMPANY = 20

# Se activa con --debug (guarda capturas cuando no encuentra empresa)
_DEBUG_SCREENSHOTS = False

# ─────────────────────────────────────────────
# LOGGING
# ─────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler("linkedin_scraper.log", encoding="utf-8"),
    ],
)
log = logging.getLogger(__name__)


# ─────────────────────────────────────────────
# UTILIDADES
# ─────────────────────────────────────────────

async def async_human_delay(range_: tuple) -> None:
    await asyncio.sleep(random.uniform(*range_))


def parse_full_name(full_name: str) -> tuple:
    parts = full_name.strip().split()
    if not parts:
        return ("", "")
    if len(parts) == 1:
        return (parts[0], "")
    return (" ".join(parts[:-1]), parts[-1])


# ─────────────────────────────────────────────
# SESIÓN / COOKIES
# ─────────────────────────────────────────────

async def save_cookies(context: BrowserContext) -> None:
    cookies = await context.cookies()
    COOKIES_FILE.write_text(
        json.dumps(cookies, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    log.info("Cookies guardadas en %s", COOKIES_FILE)


async def load_cookies(context: BrowserContext) -> bool:
    if not COOKIES_FILE.exists():
        return False
    try:
        cookies = json.loads(COOKIES_FILE.read_text(encoding="utf-8"))
        await context.add_cookies(cookies)
        log.info("Cookies cargadas desde %s", COOKIES_FILE)
        return True
    except Exception as e:
        log.warning("No se pudieron cargar cookies: %s", e)
        return False


async def is_logged_in(page: Page) -> bool:
    await page.goto("https://www.linkedin.com/feed/", wait_until="domcontentloaded")
    await async_human_delay(DELAY_MEDIUM)
    return "feed" in page.url or "mynetwork" in page.url


async def manual_login(page: Page) -> None:
    log.info("=" * 60)
    log.info("INICIO DE SESIÓN MANUAL")
    log.info("Iniciá sesión en LinkedIn en el navegador.")
    log.info("El script continúa automáticamente al detectar la sesión.")
    log.info("=" * 60)
    await page.goto("https://www.linkedin.com/login", wait_until="domcontentloaded")
    for _ in range(180):
        await asyncio.sleep(1)
        if "feed" in page.url or "mynetwork" in page.url:
            log.info("¡Sesión detectada! Continuando...")
            return
    raise TimeoutError("Timeout esperando inicio de sesión manual.")


# ─────────────────────────────────────────────
# BÚSQUEDA DE EMPRESA → URL
# ─────────────────────────────────────────────

# Palabras/frases que indican que una empresa está ubicada en Argentina.
# Se usan para priorizar resultados locales cuando hay varios con el mismo nombre.
_ARGENTINA_RE = re.compile(
    r"\b(argentina|buenos\s+aires|c\.?a\.?b\.?a\.?|ciudad\s+aut[oó]noma"
    r"|c[oó]rdoba|rosario|mendoza|tucum[aá]n|santa\s+fe|la\s+plata"
    r"|mar\s+del\s+plata|salta|neuqu[eé]n|bariloche|corrientes"
    r"|posadas|resistencia|san\s+luis|entre\s+r[ií]os"
    r"|catamarca|jujuy|r[ií]o\s+negro|formosa|chaco|tierra\s+del\s+fuego)\b",
    re.IGNORECASE,
)

# Palabras sin valor discriminativo para el matching de nombres de empresa
_COMPANY_STOP = {
    "de", "la", "el", "los", "las", "del", "al", "en", "y", "e", "o",
    "sa", "sl", "srl", "sac", "saci", "inc", "ltd", "llc",
    "the", "of", "and", "a", "an", "s", "cia", "co", "spa", "sas",
}
# Mínimo de palabras-clave compartidas (fracción) para aceptar un resultado
_MIN_NAME_SCORE = 0.30


def _keywords_from_name(text: str) -> set:
    """Extrae palabras-clave de 3+ chars de un nombre de empresa."""
    words = re.sub(r"[^\w\s]", " ", text.lower()).split()
    return {w for w in words if len(w) >= 3 and w not in _COMPANY_STOP}


def _name_score(query: str, candidate: str) -> float:
    """
    Fracción de palabras-clave de `query` que aparecen en `candidate`.
    Ej: query='BANCO PROVINCIA BUENOS AIRES', candidate='estudio bucle ia' → 0.0
        query='BANCO PROVINCIA BUENOS AIRES', candidate='banco provincia buenos aires' → 1.0
    """
    q_kw = _keywords_from_name(query)
    c_kw = _keywords_from_name(candidate)
    if not q_kw:
        return 0.0
    return len(q_kw & c_kw) / len(q_kw)


def _simplified_query(name: str, n: int = 3) -> str:
    """
    Versión simplificada de la query: primeras N palabras-clave en orden de aparición.
    'BANCO DE LA PROVINCIA DE BUENOS AIRES' → 'banco provincia buenos'
    """
    seen: set = set()
    kws: list = []
    for w in re.sub(r"[^\w\s]", " ", name.lower()).split():
        if len(w) >= 3 and w not in _COMPANY_STOP and w not in seen:
            seen.add(w)
            kws.append(w)
        if len(kws) >= n:
            break
    return " ".join(kws)


async def _is_argentina_card(container) -> bool:
    """
    Devuelve True si la tarjeta de resultado muestra una ubicación argentina.
    Revisa primero el subtítulo de ubicación; si no, todo el texto de la tarjeta.
    """
    try:
        # Subtítulos específicos de ubicación en LinkedIn company cards
        for sel in [
            ".entity-result__secondary-subtitle",
            ".entity-result__tertiary-subtitle",
            "[class*='secondary-subtitle']",
            "[class*='tertiary-subtitle']",
        ]:
            loc_el = await container.query_selector(sel)
            if loc_el:
                loc_text = (await loc_el.inner_text()).strip()
                if loc_text:
                    return bool(_ARGENTINA_RE.search(loc_text))
        # Fallback: texto completo de la tarjeta
        full_text = await container.inner_text()
        return bool(_ARGENTINA_RE.search(full_text))
    except Exception:
        return False


async def _best_company_url_from_page(
    page: Page,
    query: str,
    require_score: float = _MIN_NAME_SCORE,
) -> Optional[str]:
    """
    Lee los contenedores de resultado de la página actual y devuelve la URL
    de empresa cuyo nombre tenga mayor similitud con `query`.

    Criterios de ordenamiento (mejor primero):
      1. Empresa argentina  (True > False)
      2. Score de nombre    (mayor es mejor)

    Aplica filtro de posición (x < 860 px) para ignorar el sidebar.
    """
    candidates: list = []   # (score, is_argentina, url, name)

    # ── Iteración 1: contenedores bien estructurados ──────────────────────
    containers = await page.query_selector_all(
        "li.reusable-search__result-container, "
        ".entity-result, "
        "li[class*='entity-result']"
    )

    for container in containers:
        link = await container.query_selector("a[href*='/company/']")
        if not link:
            continue

        # Filtrar sidebar por posición horizontal
        try:
            box = await link.bounding_box()
            if box is None or box["x"] > 860:
                continue
        except Exception:
            pass

        href = await link.get_attribute("href")
        url  = _normalize_company_url(href)
        if not url:
            continue

        # Nombre del resultado para validar
        name = ""
        for ns in [
            ".entity-result__title-text",
            ".entity-result__title-lockup-text",
            "span[aria-hidden='true']",
            ".artdeco-entity-lockup__title",
        ]:
            ne = await container.query_selector(ns)
            if ne:
                raw = (await ne.inner_text()).strip()
                if raw:
                    name = raw
                    break
        if not name:
            name = url.split("/company/")[1].replace("-", " ")

        score       = _name_score(query, name)
        is_arg      = await _is_argentina_card(container) if score >= require_score else False
        log.debug("  Candidato '%-42s' score=%.2f arg=%s", name[:42], score, is_arg)

        if score >= require_score:
            candidates.append((score, is_arg, url, name))

    if candidates:
        # Prioridad: empresa argentina primero, luego mayor score de nombre
        candidates.sort(key=lambda c: (c[1], c[0]), reverse=True)
        best_score, best_arg, best_url, best_name = candidates[0]
        log.debug(
            "  Seleccionado: '%-42s' score=%.2f arg=%s",
            best_name[:42], best_score, best_arg,
        )
        return best_url

    # ── Iteración 2: links /company/ directos (cuando no hay contenedores) ─
    if not containers:
        fallback_best_url:   Optional[str] = None
        fallback_best_score: float         = 0.0
        links = await page.query_selector_all("a[href*='/company/']")
        for link in links:
            try:
                box = await link.bounding_box()
                if box is None or box["x"] > 860:
                    continue
            except Exception:
                pass
            href = await link.get_attribute("href")
            url  = _normalize_company_url(href)
            if not url:
                continue
            name  = url.split("/company/")[1].replace("-", " ")
            score = _name_score(query, name)
            log.debug("  Link candidato '%-42s' score=%.2f", name[:42], score)
            if score > fallback_best_score:
                fallback_best_score = score
                fallback_best_url   = url
            if score >= 0.85:
                break
        if fallback_best_url and fallback_best_score >= require_score:
            return fallback_best_url

    log.debug("  Sin candidato válido para '%s' (umbral=%.2f).", query[:50], require_score)
    return None


def _normalize_company_url(href: str) -> Optional[str]:
    """
    Extrae la URL canónica de empresa: https://www.linkedin.com/company/SLUG
    Descarta sub-paths (/posts, /people/, etc.) y parámetros de tracking.
    """
    if not href or "/company/" not in href:
        return None
    # Tomar todo hasta el primer segmento después de /company/
    after = href.split("/company/")[1]          # "slug/posts?trk=..." etc.
    slug  = after.split("/")[0].split("?")[0].strip()
    if not slug:
        return None
    return f"https://www.linkedin.com/company/{slug}"


async def _goto_companies_search(page: Page, keywords: str) -> None:
    url = (
        "https://www.linkedin.com/search/results/companies/?"
        + urllib.parse.urlencode({"keywords": keywords, "origin": "GLOBAL_SEARCH_HEADER"})
    )
    await page.goto(url, wait_until="domcontentloaded", timeout=30_000)
    try:
        await page.wait_for_selector(
            "li.reusable-search__result-container, .entity-result",
            timeout=10_000,
        )
    except Exception:
        await async_human_delay(DELAY_MEDIUM)
    await async_human_delay(DELAY_SHORT)


async def _do_company_search(page: Page, query: str) -> Optional[str]:
    """
    Busca la empresa en LinkedIn validando que el resultado coincida con `query`
    mediante similitud de palabras-clave. Si no se encuentra, retorna None
    inmediatamente (sin reintentos adicionales).

    Prioridad de resultados:
      - Empresas ubicadas en Argentina sobre empresas del mismo nombre en otros países
      - Mayor score de palabras-clave sobre menor score

    Flujo:
      1. search/results/companies/  →  candidatos validados con Argentina-priority
      2. search/results/all/        →  ídem (fallback)
      Si ninguno supera el umbral → None → la empresa se saltea.
    """
    # ── Paso 1: búsqueda específica de Empresas ───────────────────────────
    try:
        await _goto_companies_search(page, query)
    except Exception as e:
        log.debug("  goto companies error: %s", e)
        return None

    url = await _best_company_url_from_page(page, query)
    if url:
        log.debug("  URL via companies-search: %s", url)
        return url

    # ── Paso 2: fallback search/results/all/ ─────────────────────────────
    all_url = (
        "https://www.linkedin.com/search/results/all/?"
        + urllib.parse.urlencode({"keywords": query, "origin": "GLOBAL_SEARCH_HEADER"})
    )
    try:
        await page.goto(all_url, wait_until="domcontentloaded", timeout=30_000)
        try:
            await page.wait_for_selector("a[href*='/company/']", timeout=8_000)
        except Exception:
            await async_human_delay(DELAY_MEDIUM)
        await async_human_delay(DELAY_SHORT)
    except Exception:
        return None

    # Botón "Ver página" con validación de nombre + Argentina
    for label in ("Ver página", "View page"):
        try:
            btn = await page.query_selector(
                f"main a:has-text('{label}'), "
                f".search-results-container a:has-text('{label}')"
            )
            if btn:
                href = await btn.get_attribute("href")
                cand = _normalize_company_url(href)
                if cand:
                    slug  = cand.split("/company/")[1]
                    score = _name_score(query, slug.replace("-", " "))
                    if score >= _MIN_NAME_SCORE:
                        log.debug("  URL via 'Ver página' (score=%.2f): %s", score, cand)
                        return cand
                    log.debug("  'Ver página' rechazado (score=%.2f): %s", score, cand)
        except Exception:
            continue

    url = await _best_company_url_from_page(page, query)
    if url:
        log.debug("  URL via all-search: %s", url)
        return url

    # No encontrada → retornar None, el caller saltea a la siguiente empresa
    return None


async def find_company_linkedin_url(page: Page, company_name: str) -> Optional[str]:
    """Busca la empresa en LinkedIn y retorna la URL de su página."""
    url = await _do_company_search(page, company_name)
    if url:
        log.info("  Empresa encontrada: %s", url)
        return url

    log.warning("  No se encontró página LinkedIn para: %s", company_name)
    # Screenshot de debug si está activado
    if _DEBUG_SCREENSHOTS:
        shot_name = re.sub(r'[\\/:*?"<>|]', "_", company_name)[:40]
        path = OUTPUT_DIR / f"debug_{shot_name}.png"
        try:
            await page.screenshot(path=str(path))
            log.debug("  Screenshot guardado: %s", path)
        except Exception:
            pass
    return None


# ─────────────────────────────────────────────
# FILTRO ARGENTINA
# ─────────────────────────────────────────────

async def apply_argentina_filter(page: Page) -> bool:
    """
    Aplica el filtro 'Argentina' en la sección 'Dónde viven' de la pestaña Personas.

    Flujo:
      1. Click en '+ Añadir' junto a 'Dónde viven'
      2. Escribir 'Argentina' en el input del typeahead
      3. Seleccionar 'Argentina' del desplegable

    NOTE: no se usa clic directo en la lista porque Argentina no siempre
    aparece pre-listada (empresas pequeñas o con pocos empleados geolocalizados).
    """
    # Esperar que la página de personas termine de cargar
    try:
        await page.wait_for_selector(
            "input[placeholder*='empleados' i], "
            "input[placeholder*='employees' i], "
            "input[placeholder*='institución' i], "
            ".org-people-bar",
            timeout=10_000,
        )
    except Exception:
        pass

    await async_human_delay(DELAY_SHORT)
    await page.evaluate("window.scrollTo(0, 200)")
    await async_human_delay(DELAY_SHORT)

    # ── Paso 1: click en '+ Añadir' de la sección 'Dónde viven' ──────────
    # JS: sube desde el nodo de texto "Dónde viven" hasta encontrar
    # el botón/link que contiene "Añadir" / "Add".
    added = await page.evaluate("""
        () => {
            const walker = document.createTreeWalker(
                document.body, NodeFilter.SHOW_TEXT, null
            );
            let n;
            while ((n = walker.nextNode())) {
                const t = n.textContent.trim();
                if (t !== 'Dónde viven' && t !== 'Where they live') continue;
                let el = n.parentElement;
                for (let i = 0; i < 8; i++) {
                    if (!el) break;
                    for (const c of el.querySelectorAll('button, a')) {
                        if (c.textContent.includes('Añadir') || c.textContent.includes('Add')) {
                            c.click();
                            return true;
                        }
                    }
                    el = el.parentElement;
                }
            }
            return false;
        }
    """)

    if not added:
        for sel in ["button:has-text('Añadir')", "button:has-text('Add')",
                    "a:has-text('+ Añadir')"]:
            try:
                btn = await page.query_selector(sel)
                if btn and await btn.is_visible():
                    await btn.click()
                    added = True
                    break
            except Exception:
                continue

    if not added:
        log.warning("  No se encontró botón Añadir en Dónde viven.")
        return False

    await async_human_delay(DELAY_SHORT)

    # Escribir "Argentina" en el input del typeahead
    location_input = None
    for sel in [
        "input[placeholder*='ubicación' i]",
        "input[placeholder*='location' i]",
        "input[placeholder*='Agregar' i]",
        "input[placeholder*='Add' i]",
        "div[role='dialog'] input",
        ".search-basic-typeahead input",
        ".artdeco-typeahead input",
        "input[role='combobox']",
    ]:
        try:
            el = await page.wait_for_selector(sel, timeout=3_000)
            if el and await el.is_visible():
                location_input = el
                break
        except Exception:
            continue

    if not location_input:
        log.warning("  No apareció el input de ubicación para el typeahead.")
        return False

    await location_input.fill("Argentina")
    await async_human_delay(DELAY_SHORT)

    for sel in [
        "[role='option']:has-text('Argentina')",
        "li:has-text('Argentina')",
        ".basic-typeahead__selectable:has-text('Argentina')",
        ".artdeco-typeahead__option:has-text('Argentina')",
        "div[role='listbox'] li:has-text('Argentina')",
    ]:
        try:
            opt = await page.wait_for_selector(sel, timeout=5_000)
            if opt and await opt.is_visible():
                await opt.click()
                # Esperar que LinkedIn actualice la URL con facetGeoRegion
                try:
                    await page.wait_for_url(
                        lambda u: "facetGeoRegion" in u, timeout=8_000
                    )
                except Exception:
                    await async_human_delay(DELAY_MEDIUM)
                log.info("  Filtro Argentina aplicado.")
                return True
        except Exception:
            continue

    log.warning("  No se pudo seleccionar Argentina del desplegable.")
    return False


# ─────────────────────────────────────────────
# BUSCAR CARGO EN EL BUSCADOR DE EMPLEADOS
# ─────────────────────────────────────────────

async def _fill_search_input(page: Page, element, title: str) -> None:
    """Limpia y tipea el cargo en el input de búsqueda dado (Locator o ElementHandle)."""
    await element.click()
    await element.press("Control+a")
    await element.fill("")
    await async_human_delay((0.3, 0.7))
    # type() no está en Locator — usar fill() + dispatchEvent para simular escritura
    await element.fill(title)
    await async_human_delay(DELAY_SHORT)
    await element.press("Enter")
    await async_human_delay(DELAY_MEDIUM)


async def search_employees_by_title(page: Page, title: str) -> bool:
    """
    Escribe el cargo en el buscador de empleados y presiona Enter.
    Usa tres estrategias en cascada para encontrar el input:
      1. Playwright Locator API (get_by_placeholder — regex flexible)
      2. wait_for_selector con selectores CSS clásicos (timeout reducido)
      3. JS evaluate_handle — encuentra cualquier input visible que no sea
         el filtro de ubicación.
    Retorna True si pudo usar el buscador, False si no.
    """
    # Volver al tope de la página para que el buscador sea visible
    await page.evaluate("window.scrollTo(0, 0)")
    await async_human_delay((0.5, 1.0))

    # ── Estrategia 1: Playwright Locator API — placeholder específico ────
    # El input visible en el screenshot tiene placeholder:
    # "Buscar empleados por cargo, palabra clave o institución educativa"
    try:
        loc = page.get_by_placeholder(re.compile(
            r"Buscar empleados|Search employees|empleados por cargo"
            r"|palabra clave|institución|keyword",
            re.IGNORECASE,
        ))
        await loc.first.wait_for(state="visible", timeout=8_000)
        ph = (await loc.first.get_attribute("placeholder") or "").lower()
        if "ubicación" not in ph and "location" not in ph and "ciudad" not in ph:
            log.debug("  Buscador encontrado vía get_by_placeholder (ph='%s')", ph)
            await _fill_search_input(page, loc.first, title)
            return True
    except Exception as e:
        log.debug("  get_by_placeholder falló: %s", e)

    # ── Estrategia 2: selectores CSS clásicos (timeout 2s cada uno) ──────
    search_el = None
    for sel in [
        "input[placeholder*='Buscar empleados' i]",
        "input[placeholder*='empleados por cargo' i]",
        "input[placeholder*='Search employees' i]",
        "input[placeholder*='institución' i]",
        "input[placeholder*='palabra clave' i]",
        "input[placeholder*='keyword' i]",
        ".org-people-bar input[type='text']",
        ".org-people-bar input",
        ".org-people__container input[type='text']",
        ".org-people__container input",
        "input[aria-label*='empleados' i]",
        "input[aria-label*='Search employees' i]",
        ".org-people input[type='text']",
        ".org-people input",
    ]:
        try:
            el = await page.wait_for_selector(sel, timeout=2_000)
            if el and await el.is_visible():
                ph = (await el.get_attribute("placeholder") or "").lower()
                if "ubicación" not in ph and "location" not in ph and "ciudad" not in ph:
                    # Confirmar que no es la barra de búsqueda global de LinkedIn
                    box = await el.bounding_box()
                    if box and box["y"] < 100:
                        log.debug("  Selector '%s' apunta a navbar — ignorando", sel)
                        continue
                    search_el = el
                    log.debug("  Buscador encontrado vía CSS '%s' (ph='%s')", sel, ph)
                    break
        except Exception:
            continue

    if search_el:
        await _fill_search_input(page, search_el, title)
        return True

    # ── Estrategia 3: JS evaluate_handle — excluye navbar y filtros ───────
    # El buscador de empleados está DEBAJO de "miembros asociados" (y > ~100px).
    # El input del navbar de LinkedIn está en y < ~70px → lo descartamos.
    try:
        handle = await page.evaluate_handle("""
            () => {
                const LOCATION_WORDS = [
                    'ubicación', 'location', 'ciudad', 'city', 'país', 'country'
                ];
                const inputs = Array.from(document.querySelectorAll(
                    'input[type="text"], input[type="search"], input:not([type])'
                ));
                for (const inp of inputs) {
                    if (!inp.offsetParent) continue;
                    const style = window.getComputedStyle(inp);
                    if (style.display === 'none' || style.visibility === 'hidden') continue;

                    const rect = inp.getBoundingClientRect();
                    // Descartar navbar (parte superior de la página)
                    if (rect.top < 100) continue;
                    // Debe tener dimensiones razonables
                    if (rect.width < 80 || rect.height < 16) continue;

                    // Descartar inputs de filtro de ubicación
                    const ph  = (inp.placeholder  || '').toLowerCase();
                    const lbl = (inp.getAttribute('aria-label') || '').toLowerCase();
                    if (LOCATION_WORDS.some(w => (ph + lbl).includes(w))) continue;

                    return inp;
                }
                return null;
            }
        """)
        el = handle.as_element()
        if el:
            ph_raw = await el.get_attribute("placeholder") or ""
            box    = await el.bounding_box()
            log.debug(
                "  Buscador vía JS evaluate_handle — ph='%s' pos=(%s,%s)",
                ph_raw,
                round(box["x"]) if box else "?",
                round(box["y"]) if box else "?",
            )
            await _fill_search_input(page, el, title)
            return True
    except Exception as e:
        log.debug("  JS evaluate_handle falló: %s", e)

    log.warning("  No se encontró el buscador de empleados.")
    return False


# ─────────────────────────────────────────────
# EXTRACCIÓN DE PERFILES
# ─────────────────────────────────────────────

_NAME_NOISE = [
    "Conectar", "Connect", "Seguir", "Follow",
    "Mensaje", "Message", "Ver perfil", "View profile",
    "Pendiente", "Pending", "Retirar",
]


def _clean_name(text: str) -> str:
    for noise in _NAME_NOISE:
        text = text.replace(noise, "")
    return text.strip()


async def extract_people_from_page(page: Page) -> list[dict]:
    """
    Extrae los perfiles visibles en la pestaña Personas (grilla de tarjetas).
    También maneja el botón 'Mostrar más resultados' para cargar más.
    """
    profiles: list[dict] = []
    seen_urls: set[str] = set()

    # Scroll progresivo para activar lazy-loading
    for scroll_y in (400, 800, 1200, 1600):
        await page.evaluate(f"window.scrollTo(0, {scroll_y})")
        await async_human_delay((0.6, 1.2))

    # Intentar cargar más resultados con el botón si existe
    for btn_sel in [
        "button:has-text('Mostrar más')",
        "button:has-text('Show more')",
        "button:has-text('Ver más')",
    ]:
        try:
            btn = await page.query_selector(btn_sel)
            if btn and await btn.is_visible():
                await btn.click()
                await async_human_delay(DELAY_MEDIUM)
                # Scroll extra post-carga
                await page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
                await async_human_delay(DELAY_SHORT)
                break
        except Exception:
            pass

    # ── Selectores de tarjeta (grilla de Personas de empresa) ─────────────
    # La grilla usa li dentro de un ul; cada li contiene la tarjeta de persona
    card_selectors = [
        # Formato grilla empresa (más común 2024-2025)
        "li.org-people__item",
        "li[class*='org-people']",
        # Formato tarjeta clásico
        "li.org-people-profile-card",
        ".org-people-profile-card",
        "[data-view-name='org-people-profile-card']",
        # Fallback genérico
        "li.artdeco-list__item",
        "li.reusable-search__result-container",
    ]

    cards = []
    used_sel = ""
    for sel in card_selectors:
        found = await page.query_selector_all(sel)
        if found:
            cards = found
            used_sel = sel
            break

    log.debug("  Selector '%s': %d tarjetas", used_sel, len(cards))

    if cards:
        for card in cards:
            try:
                p = await _extract_card(card)
                if p and p["profile_url"] not in seen_urls:
                    seen_urls.add(p["profile_url"])
                    profiles.append(p)
            except Exception as e:
                log.debug("  Error en tarjeta: %s", e)
        return profiles

    # ── Fallback: extraer desde todos los links /in/ de la página ─────────
    log.debug("  Sin tarjetas detectadas. Fallback por links /in/")
    links = await page.query_selector_all("a[href*='/in/']")
    for link in links:
        try:
            href = await link.get_attribute("href")
            if not href or "/in/" not in href:
                continue
            url = (("https://www.linkedin.com" + href)
                   if href.startswith("/") else href)
            url = url.split("?")[0].rstrip("/")
            if url in seen_urls:
                continue
            raw = _clean_name((await link.inner_text()).strip())
            if not raw or raw.lower() in ("miembro de linkedin", "linkedin member"):
                continue
            seen_urls.add(url)
            profiles.append({"full_name": raw, "cargo": "", "profile_url": url})
        except Exception:
            continue

    return profiles


# Detecta indicadores de grado de conexión: "3er", "• 2º", "3rd", "3er grado", etc.
_DEGREE_RE = re.compile(
    r"^[•·\-]?\s*\d+\s*(er|do|rd|nd|st|th)?\s*(grado|degree|°|º)?\s*$",
    re.IGNORECASE,
)


def _is_degree_indicator(text: str) -> bool:
    return bool(_DEGREE_RE.match(text.strip()))


async def _extract_card(card) -> Optional[dict]:
    """
    Extrae nombre, cargo y URL de una tarjeta de persona.
    Cubre tanto el formato grilla (3 col) como el formato lista.
    """
    # ── URL del perfil (primero, para descartar tarjetas inválidas) ───────
    link_el = await card.query_selector("a[href*='/in/']")
    if not link_el:
        return None
    href = await link_el.get_attribute("href")
    if not href or "/in/" not in href:
        return None
    profile_url = (("https://www.linkedin.com" + href)
                   if href.startswith("/") else href)
    profile_url = profile_url.split("?")[0].rstrip("/")

    # ── Nombre ────────────────────────────────────────────────────────────
    # LinkedIn coloca en la misma tarjeta VARIOS span[aria-hidden='true']:
    # uno para el nombre real y otro para el grado de conexión ("3er", "2º").
    # query_selector devuelve el PRIMERO del DOM, que puede ser el grado.
    # Usamos query_selector_all e iteramos filtrando los indicadores de grado.
    name_text = ""
    _SKIP = {"miembro de linkedin", "linkedin member"}

    # Intento 1: spans DENTRO del link de perfil (los más seguros)
    link_spans = await card.query_selector_all("a[href*='/in/'] span[aria-hidden='true']")
    for span in link_spans:
        raw = _clean_name((await span.inner_text()).strip())
        if not raw or len(raw) < 3:
            continue
        if _is_degree_indicator(raw) or raw.lower() in _SKIP:
            continue
        name_text = raw
        break

    # Intento 2: todos los spans aria-hidden de la tarjeta, filtrando grado
    if not name_text:
        all_spans = await card.query_selector_all("span[aria-hidden='true']")
        for span in all_spans:
            raw = _clean_name((await span.inner_text()).strip())
            if not raw or len(raw) < 3:
                continue
            if _is_degree_indicator(raw) or raw.lower() in _SKIP:
                continue
            name_text = raw
            break

    # Intento 3: selectores alternativos de título
    if not name_text:
        for sel in [
            ".artdeco-entity-lockup__title",
            ".org-people-profile-card__profile-title",
            "a[href*='/in/'] span",
        ]:
            el = await card.query_selector(sel)
            if el:
                raw = _clean_name((await el.inner_text()).strip())
                # Limpiar indicador de grado pegado al final del texto
                raw = re.sub(
                    r"[\s•·]*\d+\s*(er|do|°|º|rd|nd|st|th|grado|degree)\s*$",
                    "", raw, flags=re.IGNORECASE,
                ).strip()
                if raw and len(raw) > 3 and raw.lower() not in _SKIP:
                    name_text = raw
                    break

    # Intento 4: inner_text del link de perfil (puede tener ruido, _clean_name lo filtra)
    if not name_text:
        raw = _clean_name((await link_el.inner_text()).strip())
        raw = re.sub(
            r"[\s•·]*\d+\s*(er|do|°|º|rd|nd|st|th|grado|degree)\s*$",
            "", raw, flags=re.IGNORECASE,
        ).strip()
        if raw and raw.lower() not in _SKIP:
            name_text = raw

    if not name_text:
        return None

    # ── Cargo ─────────────────────────────────────────────────────────────
    # En la grilla el cargo está como texto justo debajo del nombre,
    # generalmente en un div hermano del link de nombre.
    cargo = ""
    for sel in [
        ".artdeco-entity-lockup__subtitle",
        ".org-people-profile-card__profile-position",
        ".artdeco-entity-lockup__caption",
        ".t-14.t-black--light.t-normal",
        ".t-14.t-black.t-normal",
        # Genérico: primer div de texto que no sea el nombre
        "div.t-14",
        "span.t-14",
    ]:
        el = await card.query_selector(sel)
        if el:
            text = (await el.inner_text()).strip()
            if text and text != name_text:
                cargo = text
                break

    return {"full_name": name_text, "cargo": cargo, "profile_url": profile_url}


# ─────────────────────────────────────────────
# VALIDACIÓN DE CARGO
# ─────────────────────────────────────────────

# Palabras clave que deben aparecer en el cargo para considerarlo válido
_TITLE_KEYWORDS = [
    "recursos humanos", "rrhh", "rrh",
    "capital humano", "desarrollo humano",
    "talento humano",
    "people manager", "people operations", "head of people",
    "vp of people", "vp people",
    "director de talento",
    "talent acquisition", "talent",
    "employee experience",
    "culture manager",
    "chief human resources", "chro",
    "chief people", "cpo",
    "human resources",
    "hr manager", "hr director", "hr business", "hrbp",
]

# Patrón para detectar "HR" como palabra completa (evita falsos positivos como "chair")
_HR_STANDALONE_RE = re.compile(r"\bhr\b", re.IGNORECASE)


def _extract_facet_geo(url: str) -> Optional[str]:
    """Extrae el valor de facetGeoRegion de la URL actual, si existe."""
    try:
        params = urllib.parse.parse_qs(urllib.parse.urlparse(url).query)
        vals = params.get("facetGeoRegion", [])
        return vals[0] if vals else None
    except Exception:
        return None


async def _clear_keyword_filter(page: Page) -> bool:
    """
    Elimina el chip del filtro de cargo/keyword activo para que reaparezca
    la barra de búsqueda. NO elimina el filtro de Argentina.
    Se usa como fallback cuando no hay facetGeoRegion en la URL.
    """
    await page.evaluate("window.scrollTo(0, 0)")

    # JS: recorre todos los elementos con aspecto de chip/pill que NO sean Argentina
    # y hace clic en el botón × de descarte.
    removed = await page.evaluate("""
        () => {
            const SKIP = /\\bargentina\\b/i;

            // Selectores habituales de chips de filtro en LinkedIn org-people
            const chipSelectors = [
                '[class*="pill"]',
                '[class*="filter-pill"]',
                '[class*="artdeco-pill"]',
                '[class*="search-filter"]',
            ];

            for (const sel of chipSelectors) {
                const chips = document.querySelectorAll(sel);
                for (const chip of chips) {
                    if (!chip.offsetParent) continue;
                    const chipText = (chip.innerText || chip.textContent || '')
                        .replace(/\\s+/g, ' ').trim();
                    if (!chipText || SKIP.test(chipText)) continue;

                    // Buscar el botón × dentro del chip
                    const xBtn = chip.querySelector(
                        '[aria-label*="Quitar" i], [aria-label*="Remove" i], '
                        '[aria-label*="Dismiss" i], [aria-label*="cerrar" i], '
                        'button[class*="delete"], button[class*="remove"], '
                        'button[class*="dismiss"], button[class*="close"]'
                    );
                    if (xBtn) { xBtn.click(); return true; }

                    // Si el chip mismo es el botón de descarte (contiene ×)
                    if (/[×✕x]/.test(chipText) && chip.tagName === 'BUTTON') {
                        chip.click(); return true;
                    }
                }
            }

            // Fallback: botones sueltos con × que estén cerca de texto no-Argentina
            for (const btn of document.querySelectorAll('button')) {
                const label = (btn.getAttribute('aria-label') || btn.innerText || '').trim();
                if (!/^[×✕]$/.test(label) && !/quitar|remove|dismiss/i.test(label)) continue;
                const area = btn.closest('[class*="filter"], [class*="pill"], [class*="chip"]');
                if (!area) continue;
                if (SKIP.test(area.innerText || '')) continue;
                btn.click();
                return true;
            }
            return false;
        }
    """)

    if removed:
        await async_human_delay(DELAY_SHORT)
        # Esperar que reaparezca la barra de búsqueda
        try:
            await page.wait_for_selector(
                "input[placeholder*='empleados' i], input[placeholder*='employees' i], "
                "input[placeholder*='cargo' i], input[placeholder*='keyword' i]",
                timeout=6_000,
            )
        except Exception:
            pass
        return True

    log.debug("  _clear_keyword_filter: no se encontró chip de cargo activo.")
    return False


def _cargo_es_valido(cargo: str) -> bool:
    """
    True si el cargo del perfil coincide con algún título objetivo.
    Si el cargo está vacío se acepta igual (puede ser que LinkedIn no lo muestre
    en la tarjeta pero la persona fue encontrada por el título buscado).
    """
    if not cargo:
        return True   # sin información de cargo → incluir (fue encontrado por título)
    cargo_lower = cargo.lower()
    if any(kw in cargo_lower for kw in _TITLE_KEYWORDS):
        return True
    # "HR" como palabra completa: "HR", "HR Manager", "Head of HR", etc.
    if _HR_STANDALONE_RE.search(cargo):
        return True
    return False


# ─────────────────────────────────────────────
# SCRAPING POR EMPRESA
# ─────────────────────────────────────────────

async def scrape_company(
    page: Page,
    company_name: str,
    cuit: str,
    razon_social: str,
) -> list[dict]:
    """
    Flujo completo por empresa:
      1. Buscar empresa → URL de su página LinkedIn
      2. Navegar a /people/
      3. Click '+ Añadir' de 'Dónde viven' → seleccionar Argentina
      4. Por cada cargo: escribir en buscador de empleados → extraer tarjetas
         Solo se agregan al output perfiles cuyo cargo coincida con los targets.
    """
    log.info("━" * 55)
    log.info("Empresa: %s | CUIT: %s", company_name, cuit)
    log.info("━" * 55)

    all_results: list[dict] = []
    seen_urls:   set[str]   = set()

    # ── 1. Encontrar URL de la empresa ──────────────────────────────────
    company_url = await find_company_linkedin_url(page, company_name)
    if not company_url:
        return all_results

    # ── 2. Navegar a pestaña Personas ───────────────────────────────────
    people_url = company_url.rstrip("/") + "/people/"
    log.info("  Personas: %s", people_url)
    try:
        await page.goto(people_url, wait_until="domcontentloaded", timeout=30_000)
    except Exception as e:
        log.warning("  Error navegando a pestaña Personas: %s", e)
        return all_results

    await async_human_delay(DELAY_MEDIUM)

    if "login" in page.url or "authwall" in page.url:
        raise RuntimeError("Sesión expirada — se requiere re-autenticación.")

    # ── 3. Aplicar filtro Argentina (una sola vez) ───────────────────────
    arg_ok = await apply_argentina_filter(page)
    if not arg_ok:
        log.warning("  Continuando sin filtro de Argentina.")

    await async_human_delay(DELAY_SHORT)
    await page.evaluate("window.scrollTo(0, 0)")
    await async_human_delay(DELAY_SHORT)

    # Capturar facetGeoRegion de la URL actual.
    # Si LinkedIn lo expone en la URL podemos hacer una navegación directa
    # por cargo sin tocar la barra de búsqueda ni chips intermedios.
    facet_geo = _extract_facet_geo(page.url)
    if facet_geo:
        log.info("  facetGeoRegion=%s — navegaré por URL en cada cargo.", facet_geo)
    else:
        log.info("  Sin facetGeoRegion — usaré buscador UI + limpieza de chips.")

    # ── 4. Iterar por cargo ──────────────────────────────────────────────
    for title in TARGET_TITLES:
        if len(all_results) >= MAX_RESULTS_PER_COMPANY:
            log.info("  Límite %d alcanzado.", MAX_RESULTS_PER_COMPANY)
            break

        log.info("  Buscando: [%s]", title)

        if facet_geo:
            # ── Estrategia A: navegación directa por URL ─────────────────
            # Combina filtro Argentina + keyword de cargo en una sola URL.
            # No requiere manipular la barra de búsqueda ni chips.
            title_url = (
                people_url.rstrip("/") + "/?"
                + urllib.parse.urlencode({
                    "facetGeoRegion": facet_geo,
                    "keywords"      : title,
                })
            )
            try:
                await page.goto(title_url, wait_until="domcontentloaded", timeout=30_000)
                await async_human_delay(DELAY_MEDIUM)
            except Exception as e:
                log.warning("  Error navegando a cargo [%s]: %s", title, e)
                continue
        else:
            # ── Estrategia B: UI → limpiar chip anterior → tipear cargo ──
            # Necesaria cuando LinkedIn no expone facetGeoRegion en la URL.
            # Después de la primera búsqueda la barra desaparece, así que
            # hay que eliminar el chip del cargo anterior para que vuelva.
            await _clear_keyword_filter(page)
            await async_human_delay(DELAY_SHORT)
            search_ok = await search_employees_by_title(page, title)
            if not search_ok:
                log.warning("  Buscador no disponible para [%s], saltando.", title)
                continue

        profiles = await extract_people_from_page(page)
        new = 0
        skipped = 0
        for p in profiles:
            url   = p.get("profile_url", "")
            cargo = p.get("cargo", "")
            if not url or url in seen_urls:
                continue
            # Descartar perfiles cuyo cargo no coincide con los targets
            if not _cargo_es_valido(cargo):
                skipped += 1
                continue
            seen_urls.add(url)
            nombre, apellido = parse_full_name(p["full_name"])
            all_results.append({
                "cuit"            : cuit,
                "razon_social"    : razon_social,
                "nombre"          : nombre,
                "apellido"        : apellido,
                "puesto"          : cargo or title,
                "perfil_url"      : url,
                "titulo_busqueda" : title,
                "fecha_extraccion": datetime.now().strftime("%Y-%m-%d %H:%M"),
            })
            new += 1

        log.info("  [%s] → %d incluidos, %d descartados por cargo", title, new, skipped)
        await async_human_delay(DELAY_MEDIUM)

    log.info("  Total para %s: %d decisores", company_name, len(all_results))
    return all_results


# ─────────────────────────────────────────────
# OUTPUT EXCEL
# ─────────────────────────────────────────────

def create_excel_output(results: list[dict], filepath: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Decisores RRHH"

    hdr_font  = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    hdr_fill  = PatternFill("solid", start_color="0A5C3F")
    hdr_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    dat_font  = Font(name="Arial", size=10)
    dat_align = Alignment(vertical="center")
    lnk_font  = Font(name="Arial", size=10, color="185FA5", underline="single")
    alt_fill  = PatternFill("solid", start_color="F1F5F2")
    border    = Border(bottom=Side(style="thin", color="D3D1C7"))

    headers    = ["CUIT", "Razón Social", "Nombre", "Apellido", "Puesto", "Link Perfil LinkedIn"]
    col_widths = [18, 38, 20, 20, 38, 55]

    for i, (h, w) in enumerate(zip(headers, col_widths), 1):
        cell = ws.cell(row=1, column=i, value=h)
        cell.font      = hdr_font
        cell.fill      = hdr_fill
        cell.alignment = hdr_align
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "A2"

    for row_i, rec in enumerate(results, 2):
        alt = row_i % 2 == 0
        row = [
            rec.get("cuit", ""),
            rec.get("razon_social", ""),
            rec.get("nombre", ""),
            rec.get("apellido", ""),
            rec.get("puesto", ""),
            rec.get("perfil_url", ""),
        ]
        for col_i, val in enumerate(row, 1):
            cell = ws.cell(row=row_i, column=col_i, value=val)
            cell.border = border
            if col_i == 6 and val:
                cell.hyperlink = val
                cell.font      = lnk_font
                cell.alignment = dat_align
            else:
                cell.font      = dat_font
                cell.alignment = dat_align
                if alt:
                    cell.fill = alt_fill

    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{ws.max_row}"

    # Hoja resumen
    ws2 = wb.create_sheet("Resumen")
    for i, h in enumerate(["Razón Social", "CUIT", "Decisores encontrados"], 1):
        c = ws2.cell(row=1, column=i, value=h)
        c.font      = Font(name="Arial", bold=True, color="FFFFFF")
        c.fill      = PatternFill("solid", start_color="0A5C3F")
        c.alignment = Alignment(horizontal="center")

    counter  = Counter(r["razon_social"] for r in results)
    cuit_map = {r["razon_social"]: r["cuit"] for r in results}
    for ri, (emp, cnt) in enumerate(sorted(counter.items()), 2):
        ws2.cell(row=ri, column=1, value=emp).font   = Font(name="Arial", size=10)
        ws2.cell(row=ri, column=2, value=cuit_map.get(emp, "")).font = Font(name="Arial", size=10)
        ws2.cell(row=ri, column=3, value=cnt).font   = Font(name="Arial", size=10)

    ws2.column_dimensions["A"].width = 38
    ws2.column_dimensions["B"].width = 18
    ws2.column_dimensions["C"].width = 22

    wb.save(filepath)
    log.info("Excel guardado: %s (%d registros)", filepath, len(results))


# ─────────────────────────────────────────────
# LECTURA DEL INPUT
# ─────────────────────────────────────────────

def load_companies_from_excel(filepath: str) -> list[dict]:
    import pandas as pd
    df = pd.read_excel(filepath, dtype=str)
    df.columns = [c.strip().upper() for c in df.columns]

    col_map = {}
    for col in df.columns:
        clean = col.replace(" ", "_").replace("Ó", "O").replace("Á", "A")
        if "CUIT" in clean:
            col_map[col] = "cuit"
        elif "RAZ" in clean or "SOCIAL" in clean or "EMPRESA" in clean:
            col_map[col] = "razon_social"

    df = df.rename(columns=col_map)

    if "cuit" not in df.columns or "razon_social" not in df.columns:
        raise ValueError(
            f"El Excel debe tener columnas CUIT y Razón Social. "
            f"Columnas encontradas: {list(df.columns)}"
        )

    companies = []
    for _, row in df.iterrows():
        cuit  = str(row["cuit"]).strip()
        razon = str(row["razon_social"]).strip()
        if cuit and razon and razon.lower() != "nan" and cuit.lower() != "nan":
            companies.append({"cuit": cuit, "razon_social": razon})

    log.info("Empresas cargadas: %d", len(companies))
    return companies


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────

async def main():
    import argparse

    parser = argparse.ArgumentParser(
        description="LinkedIn HR Scraper — Decisores RRHH en Argentina"
    )
    parser.add_argument("--login",    action="store_true",
                        help="Login manual y guardado de cookies.")
    parser.add_argument("--input",    type=str, default=None,
                        help="Ruta al Excel con columnas CUIT y RAZON_SOCIAL.")
    parser.add_argument("--output",   type=str, default=None,
                        help="Nombre del archivo Excel de salida.")
    parser.add_argument("--headless", action="store_true",
                        help="Modo sin interfaz gráfica.")
    parser.add_argument("--test",     action="store_true",
                        help="Prueba: procesa solo las primeras 2 empresas.")
    parser.add_argument("--debug",    action="store_true",
                        help="Guardar capturas de pantalla cuando no se encuentra una empresa.")
    args = parser.parse_args()

    global _DEBUG_SCREENSHOTS
    _DEBUG_SCREENSHOTS = args.debug

    # Cargar empresas
    if args.input:
        companies = load_companies_from_excel(args.input)
    else:
        companies = [
            {"cuit": "30000000000", "razon_social": "EMPRESA EJEMPLO S.A."},
            {"cuit": "30000000001", "razon_social": "OTRA EMPRESA S.R.L."},
        ]
        log.info("Sin --input: usando 2 empresas de prueba por defecto.")

    if args.test:
        companies = companies[:2]
        log.info("Modo TEST: solo %d empresa(s).", len(companies))

    if not companies:
        log.error("No hay empresas para procesar.")
        return

    timestamp       = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_filename = args.output or f"linkedin_rrhh_{timestamp}.xlsx"
    output_path     = OUTPUT_DIR / output_filename

    all_results: list[dict] = []

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(
            headless=args.headless,
            args=[
                "--disable-blink-features=AutomationControlled",
                "--disable-dev-shm-usage",
                "--no-sandbox",
                "--lang=es-AR",
            ],
        )
        context = await browser.new_context(
            viewport    ={"width": 1366, "height": 768},
            user_agent  =(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
            locale      ="es-AR",
            timezone_id ="America/Argentina/Buenos_Aires",
        )
        page = await context.new_page()

        # Modo login
        if args.login:
            await manual_login(page)
            await save_cookies(context)
            log.info("Login OK. Ejecutá sin --login para hacer scraping.")
            await browser.close()
            return

        # Cargar sesión guardada
        cookies_loaded = await load_cookies(context)
        logged = await is_logged_in(page) if cookies_loaded else False

        if not logged:
            log.warning("Sin sesión activa. Iniciando login manual...")
            await manual_login(page)
            await save_cookies(context)

        log.info("=" * 55)
        log.info("INICIO — %d empresa(s)", len(companies))
        log.info("=" * 55)

        for idx, company in enumerate(companies, 1):
            log.info("\n[%d/%d] %s", idx, len(companies), company["razon_social"])
            try:
                results = await scrape_company(
                    page        =page,
                    company_name=company["razon_social"],
                    cuit        =company["cuit"],
                    razon_social=company["razon_social"],
                )
                all_results.extend(results)

                # Guardado parcial cada 3 empresas
                if idx % 3 == 0 and all_results:
                    partial = OUTPUT_DIR / f"partial_{timestamp}.xlsx"
                    create_excel_output(all_results, partial)
                    log.info("Guardado parcial: %s", partial)

            except RuntimeError as e:
                log.error("Sesión expirada (%s): %s", company["razon_social"], e)
                await manual_login(page)
                await save_cookies(context)
            except Exception as e:
                log.error("Error en %s: %s", company["razon_social"], e, exc_info=True)

            if idx < len(companies):
                wait = random.uniform(10, 20)
                log.info("Pausa %.1fs...", wait)
                await asyncio.sleep(wait)

        await browser.close()

    if all_results:
        create_excel_output(all_results, output_path)
        log.info("\n✓ Scraping completo. Total: %d decisores. Archivo: %s",
                 len(all_results), output_path)
    else:
        log.warning("No se encontraron resultados.")


if __name__ == "__main__":
    asyncio.run(main())
