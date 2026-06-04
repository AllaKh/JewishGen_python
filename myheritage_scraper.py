#!/usr/bin/env python3
"""
myheritage_scraper.py  —  v6
─────────────────────────────
Logic:
1. Open the URL chosen in GUI (e.g. https://www.myheritage.co.il/  or
   https://www.myheritage.com/?lang=RU  etc.)
2. If the page is .co.il → accept cookie banner ("לקבל הכל").
3. Click "Вход" / "כניסה" / "Log in" link to open the login form.
4. Fill email + password, click submit ("היכנס" / "Авторизация" / "Log in").
5. If a 2-FA code dialog appears → GUI shows a modal asking user to type
   the code; scraper waits for it via an asyncio.Event, then enters the code.
6. After login → navigate to the search URL for the chosen domain, fill
   the search form, collect results, open each qualifying record, save.

IMPORTANT: a PERSISTENT browser profile (.mh_profile) is used so cookies and
session survive between runs — after the first successful login + email 2FA,
later runs reuse the session and MyHeritage stops flagging logins as suspicious.
"""

import asyncio, difflib, imaplib, email as _email_lib
import io, os, re, sys, time
from pathlib import Path

if getattr(sys, "frozen", False):
    bd = Path(sys.executable).resolve().parent / "ms-playwright"
    if bd.exists():
        os.environ["PLAYWRIGHT_BROWSERS_PATH"] = str(bd)

try:
    from playwright.async_api import async_playwright
except ImportError:
    sys.stderr.write("pip install playwright && playwright install chromium\n")
    sys.exit(1)

try:
    from docx import Document
    from docx.shared import Mm, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    _DOCX_OK = True
except ImportError:
    _DOCX_OK = False

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter
    _OPENPYXL_OK = True
except ImportError:
    _OPENPYXL_OK = False

try:
    from PIL import Image
    _PIL_OK = True
except ImportError:
    _PIL_OK = False


def _to_png(data: bytes):
    """Return PNG bytes for any image (WEBP/JPEG/…) so python-docx can embed
    it. MyHeritage serves photos as WEBP which docx can't read directly."""
    if not data:
        return None
    # already a docx-friendly format? (JPEG/PNG/GIF/BMP magic)
    if (data[:3] == b"\xff\xd8\xff" or data[:8] == b"\x89PNG\r\n\x1a\n"
            or data[:4] == b"GIF8" or data[:2] == b"BM"):
        return data
    if _PIL_OK:
        try:
            im = Image.open(io.BytesIO(data)).convert("RGB")
            out = io.BytesIO()
            im.save(out, format="PNG")
            return out.getvalue()
        except Exception:
            return None
    return None

# ── Constants ─────────────────────────────────────────────────────────────── #
MIN_MATCH_PCT = 80
HYPERLINK_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"

# Persistent Chromium profile — keeps MyHeritage cookies/session between runs
# so the anti-bot stops flagging the login as suspicious.
MH_PROFILE_DIR = Path(__file__).resolve().parent / ".mh_profile"

# Site presets — key → (login_url, search_url, has_cookie_banner)
SITE_PRESETS = {
    "Israel (.co.il)":        ("https://www.myheritage.co.il/login",
                               "https://www.myheritage.co.il/research/search/all/all",
                               True),
    "English (.com EN)":      ("https://www.myheritage.com/login?lang=EN",
                               "https://www.myheritage.com/research/search/all/all?lang=EN",
                               True),
    "Russian (.com RU)":      ("https://www.myheritage.com/login?lang=RU",
                               "https://www.myheritage.com/research/search/all/all?lang=RU",
                               True),
    "Hebrew (.com HE)":       ("https://www.myheritage.com/login?lang=HE",
                               "https://www.myheritage.com/research/search/all/all?lang=HE",
                               True),
    "French (.com FR)":       ("https://www.myheritage.com/login?lang=FR",
                               "https://www.myheritage.com/research/search/all/all?lang=FR",
                               True),
    "German (.com DE)":       ("https://www.myheritage.com/login?lang=DE",
                               "https://www.myheritage.com/research/search/all/all?lang=DE",
                               True),
    "Spanish (.com ES)":      ("https://www.myheritage.com/login?lang=ES",
                               "https://www.myheritage.com/research/search/all/all?lang=ES",
                               True),
    "Portuguese (.com PT)":   ("https://www.myheritage.com/login?lang=PT",
                               "https://www.myheritage.com/research/search/all/all?lang=PT",
                               True),
}

FILTER_OPTIONS = ["All Records", "Historical Records", "Family Trees"]

# ── Multi-language UI labels ─────────────────────────────────────────────────
# MyHeritage renders the research form in the SITE language (the lang= in the
# URL, picked by the GUI "Site / Language" selector). Every place we locate a
# pill / section / button BY ITS VISIBLE TEXT must therefore use the label in
# the active site language. The GUI itself is always English and sends canonical
# English values; the scraper translates them here.
#
# We pass ALL supported-language variants to the text matchers: only the variant
# matching the loaded site exists on the page, and the data-automations verify-id
# confirms the correct popup actually opened — so the same code drives every
# site. Supported: en, ru, he, fr, de, es, pt. (Non-EN/RU labels are best-effort
# — but the data-automations IDs do the real field/apply work and are language-
# independent, so a wrong word only affects finding a pill by text, never input.)
#
# NOTE: the user's MyHeritage account UI language may override the URL lang after
# the page loads (e.g. it reverts to Russian). That's fine: _ui_labels returns
# EVERY language's variant, so whatever language the page ends up in, its label
# is in the list and still matches.
def _site_lang(site):
    """Map a Site/Language preset to a language code (en/ru/he/fr/de/es/pt)."""
    s = (site or "").lower()
    if "russ" in s or "(.com ru" in s or " ru)" in s:
        return "ru"
    if "hebrew" in s or "co.il" in s or "israel" in s or "(.com he" in s:
        return "he"
    if "french" in s or "(.com fr" in s:
        return "fr"
    if "german" in s or "(.com de" in s:
        return "de"
    if "spanish" in s or "(.com es" in s:
        return "es"
    if "portug" in s or "(.com pt" in s:
        return "pt"
    return "en"

_UI_I18N = {
    "search":      {"en": "Search",      "ru": "Поиск",            "he": "חיפוש",
                    "fr": "Rechercher",  "de": "Suchen",           "es": "Buscar",
                    "pt": "Pesquisar"},
    "more":        {"en": "More",        "ru": "Больше",           "he": "עוד",
                    "fr": "Plus",        "de": "Mehr",             "es": "Más",
                    "pt": "Mais"},
    "father":      {"en": "Father",      "ru": "Отец",             "he": "אב",
                    "fr": "Père",        "de": "Vater",            "es": "Padre",
                    "pt": "Pai"},
    "mother":      {"en": "Mother",      "ru": "Мать",             "he": "אם",
                    "fr": "Mère",        "de": "Mutter",           "es": "Madre",
                    "pt": "Mãe"},
    "spouse":      {"en": "Spouse",      "ru": ["Супруг(-а)", "Супруг"],
                    "he": ["בן/בת זוג", "בן זוג"],
                    "fr": ["Conjoint(e)", "Conjoint"], "de": ["Ehepartner", "Partner"],
                    "es": "Cónyuge",     "pt": "Cônjuge"},
    "death":       {"en": "Death",       "ru": "Смерть",           "he": "פטירה",
                    "fr": "Décès",       "de": "Tod",              "es": ["Defunción", "Fallecimiento"],
                    "pt": ["Falecimento", "Morte"]},
    "residence":   {"en": "Residence",   "ru": "Местожительство",  "he": "מגורים",
                    "fr": "Résidence",   "de": ["Wohnsitz", "Wohnort"], "es": "Residencia",
                    "pt": "Residência"},
    "military":    {"en": "Military",    "ru": "Вооруженные силы", "he": "צבא",
                    "fr": ["Service militaire", "Militaire"], "de": ["Militär", "Militärdienst"],
                    "es": ["Servicio militar", "Militar"], "pt": ["Serviço militar", "Militar"]},
    "immigration": {"en": "Immigration", "ru": "Иммиграция",       "he": "הגירה",
                    "fr": "Immigration", "de": "Einwanderung",     "es": "Inmigración",
                    "pt": "Imigração"},
    "keywords":    {"en": "Keywords",    "ru": "Ключевые слова",   "he": "מילות מפתח",
                    "fr": "Mots-clés",   "de": ["Schlüsselwörter", "Stichwörter"],
                    "es": "Palabras clave", "pt": "Palavras-chave"},
    "gender":      {"en": "Gender",      "ru": "Пол",              "he": "מין",
                    "fr": "Sexe",        "de": "Geschlecht",       "es": ["Sexo", "Género"],
                    "pt": ["Sexo", "Gênero"]},
    "apply":       {"en": "Apply",       "ru": "Применить",        "he": "החל",
                    "fr": "Appliquer",   "de": ["Anwenden", "Übernehmen"], "es": "Aplicar",
                    "pt": "Aplicar"},
    "male":        {"en": "Male",        "ru": "Мужчина",          "he": "זכר",
                    "fr": ["Homme", "Masculin"], "de": ["Männlich", "Mann"],
                    "es": ["Hombre", "Masculino"], "pt": ["Masculino", "Homem"]},
    "female":      {"en": "Female",      "ru": "Женщина",          "he": "נקבה",
                    "fr": ["Femme", "Féminin"], "de": ["Weiblich", "Frau"],
                    "es": ["Mujer", "Femenino"], "pt": ["Feminino", "Mulher"]},
}

# Stable language order used when no specific site-language is given.
_LANGS = ["en", "ru", "he", "fr", "de", "es", "pt"]

def _ui_labels(concept, lang=None):
    """All language variants of a concept's UI label, site-lang first when known.
    Accepts dict values that are a single string or a list of strings. Iterates
    every language present, so adding a language to _UI_I18N just works."""
    d = _UI_I18N.get(concept, {})
    out, seen = [], set()
    for k in ([lang] if lang else []) + _LANGS + list(d.keys()):
        vals = d.get(k) or []
        if isinstance(vals, str):
            vals = [vals]
        for v in vals:
            if v and v not in seen:
                seen.add(v); out.append(v)
    return out

def _search_words(lang=None):
    """Every language's word for the search-submit button (site-lang first)."""
    return _ui_labels("search", lang)

# ── Helpers ───────────────────────────────────────────────────────────────── #
def _name_sim(a, b):
    a, b = a.strip().lower(), b.strip().lower()
    return difflib.SequenceMatcher(None, a, b).ratio() * 100 if a and b else 0.0

def safe_fn(s, n=80):
    return re.sub(r'\s+', '_', re.sub(r'[\\/*?:"<>|]', '_', s.strip()))[:n] or "result"

# ── Cookie banner ─────────────────────────────────────────────────────────── #
async def _accept_cookies(page, log):
    """Accept cookie banner — only appears on .co.il."""
    await asyncio.sleep(1.5)
    # Exact button texts seen on myheritage.co.il
    for text in ["לקבל הכל", "לקבל רק מה שהכרחי",
                 "Accept all", "Accept All", "Accept", "OK",
                 "Принять все", "Принять"]:
        try:
            btn = page.get_by_role("button", name=re.compile(re.escape(text), re.I))
            if await btn.count():
                await btn.first.click(timeout=4000)
                log(f"  ✓ Cookie banner accepted ('{text}')")
                await asyncio.sleep(0.8)
                return
        except Exception:
            pass
    # CSS fallback
    for sel in ["#onetrust-accept-btn-handler", ".onetrust-accept-btn-handler",
                '[class*="cookie"] button', '[id*="cookie"] button',
                '[class*="consent"] button']:
        try:
            el = page.locator(sel).first
            if await el.count():
                await el.click(timeout=3000)
                log(f"  ✓ Cookie banner (selector: {sel})")
                await asyncio.sleep(0.8)
                return
        except Exception:
            pass

# ── Fill input field ─────────────────────────────────────────────────────── #
async def _fill(page, selectors, value, label, log):
    """
    Fill a form field reliably.

    MyHeritage uses  type="text" autocomplete="off"  for the email field and
    a normal type="password" for the password field.  Playwright's .fill()
    sometimes gets cleared by the site's JS on fields with autocomplete="off".
    The safest approach is:
      1. Click the element to give it real focus
      2. Select-all + Delete to clear any existing value
      3. Use page.keyboard.type() — simulates real keystrokes, never blocked
      4. Verify the value stuck; if not, try JS direct assignment as last resort
    """
    if not value:
        return True
    for sel in selectors:
        try:
            el = page.locator(sel).first
            if not await el.count():
                continue
            await el.scroll_into_view_if_needed(timeout=4000)

            # Give the field real browser focus via JS (more reliable than click)
            await page.evaluate(
                "sel => { const el = document.querySelector(sel); "
                "if (el) { el.focus(); } }",
                sel,
            )
            await asyncio.sleep(0.15)

            # Clear existing value
            await el.click(timeout=3000)
            await page.keyboard.press("Control+a")
            await page.keyboard.press("Delete")
            await asyncio.sleep(0.1)

            # Type character by character — works with autocomplete="off"
            await page.keyboard.type(value, delay=40)
            await asyncio.sleep(0.2)

            typed = await el.input_value(timeout=2000)
            if typed.strip():
                log(f"  ✓ {label}: OK ({len(typed)} chars)")
                return True

            # Last resort: JS direct value set + events
            await page.evaluate("""([s, v]) => {
                const el = document.querySelector(s);
                if (!el) return;
                const setter = Object.getOwnPropertyDescriptor(
                    window.HTMLInputElement.prototype, 'value').set;
                setter.call(el, v);
                el.dispatchEvent(new Event('input',  {bubbles: true}));
                el.dispatchEvent(new Event('change', {bubbles: true}));
                el.dispatchEvent(new KeyboardEvent('keyup', {bubbles: true}));
            }""", [sel, value])
            typed = await el.input_value(timeout=2000)
            if typed.strip():
                log(f"  ✓ {label}: JS OK ({len(typed)} chars)")
                return True

        except Exception as exc:
            log(f"  (selector {sel} failed: {exc})")
            continue

    log(f"  !! Field not found: {label}")
    return False

# ── Open login form ───────────────────────────────────────────────────────── #
async def _open_login_form(page, log):
    """
    Ensure the email/password form is visible.

    IMPORTANT: the login_url we navigate to is already MyHeritage's /login
    page, so the form is normally present immediately. We must NOT click any
    random "Log in" / "Войти" text — on the home/footer that can match a
    Facebook social link and navigate the tab to facebook.com (the bug Alla
    hit). So:
      1. If we're already on a /login or /signin URL → do nothing.
      2. If the email field is already present → do nothing.
      3. Only as a last resort click a login control that is a real
         MyHeritage button/link (never a social link).
    """
    # 1) Already on the login page → form is there.
    if "/login" in page.url.lower() or "/signin" in page.url.lower():
        return True

    # 2) Email field already on the page → nothing to open.
    for sel in ('#registrationEmail', 'input[name="registrationEmail"]',
                'input[type="email"]'):
        try:
            if await page.locator(sel).first.count():
                return True
        except Exception:
            pass

    # 3) Last resort: click a real MyHeritage login control, scoped to
    #    same-site hrefs only (skip facebook/google/twitter/etc).
    LOGIN_TEXTS = ["כניסה", "Вход", "Войти", "Авторизация",
                   "Log in", "Login", "Sign in"]
    for text in LOGIN_TEXTS:
        try:
            el = page.get_by_role(
                "link", name=re.compile(rf"^{re.escape(text)}$", re.I)).first
            if await el.count():
                href = (await el.get_attribute("href") or "").lower()
                # Skip social / external links
                if any(bad in href for bad in
                       ("facebook", "google", "twitter", "apple",
                        "instagram", "linkedin")):
                    continue
                await el.click(timeout=5000)
                await asyncio.sleep(1.5)
                log(f"  ✓ Opened login form (clicked '{text}')")
                return True
        except Exception:
            pass
    log("  (login form already visible / no nav link needed)")
    return True

# ── YANDEX 2FA CODE READER ────────────────────────────────────────────────── #

def _imap_server(email_addr: str) -> tuple:
    """Return (host, port) for IMAP based on the email domain."""
    domain = email_addr.split("@")[-1].lower() if "@" in email_addr else ""
    if "gmail" in domain:
        return ("imap.gmail.com", 993)
    if any(x in domain for x in ("yandex", "ya.ru")):
        return ("imap.yandex.ru", 993)
    if any(x in domain for x in ("outlook", "hotmail", "live", "msn")):
        return ("outlook.office365.com", 993)
    if "mail.ru" in domain:
        return ("imap.mail.ru", 993)
    if "yahoo" in domain:
        return ("imap.mail.yahoo.com", 993)
    # generic fallback
    return (f"imap.{domain}", 993)


def _imap_read_mh_code(imap_email: str, imap_password: str,
                       max_wait_sec: int = 90) -> str | None:
    """
    Read MyHeritage verification code from the user's mailbox via IMAP.
    Follows the exact pattern from the project's test suite:
      - Connect via IMAP4_SSL
      - Login with full email + password
      - Search FROM "@myheritage.com"
      - Fetch last matching message
      - Extract 6-digit code with regex \\b(\\d{6})\\b
      - Poll every 5 seconds up to max_wait_sec
    """
    host, port = _imap_server(imap_email)
    deadline = time.time() + max_wait_sec
    while time.time() < deadline:
        try:
            mail = imaplib.IMAP4_SSL(host, port)
            mail.login(imap_email, imap_password)
            mail.select("INBOX")
            _, msg_ids = mail.search(None, '(FROM "@myheritage.com")')
            ids = msg_ids[0].split() if msg_ids[0] else []
            if not ids:
                # Also try broader search
                _, msg_ids2 = mail.search(None, '(FROM "myheritage")')
                ids = msg_ids2[0].split() if msg_ids2[0] else []
            if ids:
                _, msg_data = mail.fetch(ids[-1], "(RFC822)")
                msg = _email_lib.message_from_bytes(msg_data[0][1])
                body = ""
                if msg.is_multipart():
                    for part in msg.walk():
                        if part.get_content_type() in ("text/plain", "text/html"):
                            body = part.get_payload(decode=True).decode(
                                "utf-8", errors="ignore")
                            break
                else:
                    body = msg.get_payload(decode=True).decode(
                        "utf-8", errors="ignore")
                mail.close(); mail.logout()
                m = re.search(r"\b(\d{6})\b", body)
                if m:
                    return m.group(1)
            else:
                mail.close(); mail.logout()
        except Exception as exc:
            pass  # retry after sleep
        time.sleep(5)
    return None


async def _browser_read_yandex_code(ctx, mail_email: str, mail_password: str,
                                    log, timeout: int = 120) -> str | None:
    """
    Open a NEW TAB at https://mail.yandex.ru/, log in with password, open the
    topmost email and read the 6-digit MyHeritage code. Then close the tab.
    Uses the EXACT selectors from the live Yandex login flow.
    """
    async def _click_any(sels, label, wait=8000):
        """Wait for and click the first matching selector. Returns True/False."""
        for sel in sels:
            try:
                el = page.locator(sel).first
                await el.wait_for(state="visible", timeout=wait)
                await el.click(timeout=5000)
                log(f"  2FA: {label}  ({sel})")
                return True
            except Exception:
                continue
        log(f"  2FA: !! не нашёл — {label}")
        return False

    async def _fill_any(sels, value, label, wait=8000):
        for sel in sels:
            try:
                el = page.locator(sel).first
                await el.wait_for(state="visible", timeout=wait)
                await el.click(timeout=4000)
                await el.fill(value)
                await asyncio.sleep(0.4)
                if (await el.input_value()).strip():
                    log(f"  2FA: {label} заполнено  ({sel})")
                    return True
            except Exception:
                continue
        log(f"  2FA: !! поле не найдено — {label}")
        return False

    # Protect the mail tab from the junk-tab auto-closer (_on_new_page): without
    # this it closes the tab after ~3s for "not being MyHeritage", which killed
    # the 2FA flow on fresh logins (notably the .co.il site, whose cookies don't
    # carry over from .com so it always needs a real login + email code).
    try:
        ctx._mh_pause_autoclose = True
    except Exception:
        pass
    page = await ctx.new_page()
    try:
        page._mh_protected = True
    except Exception:
        pass
    try:
        await page.bring_to_front()
        log("  2FA: открываю новую вкладку mail.yandex.ru …")
        await page.goto("https://mail.yandex.ru/", wait_until="domcontentloaded",
                        timeout=30000)
        await asyncio.sleep(2)

        # The PERSISTENT profile usually keeps the Yandex session, so the inbox
        # opens straight away. In that case there is NO login form — running the
        # login dance clicks a stray "Войти", navigates AWAY from the inbox and
        # breaks the tab. So detect "already logged in" and skip login entirely.
        async def _mail_logged_in():
            try:
                return await page.evaluate(r"""() => {
                    if (document.querySelector('#header-login-button')) return false;
                    const txt = e => (e.textContent || '').trim();
                    const compose = Array.from(
                        document.querySelectorAll('a,button,span,div'))
                        .some(e => /^(Написать|Compose|Написати|כתוב)$/i.test(txt(e)));
                    const rows = document.querySelectorAll(
                        '[class*="MessageSnippet"], .mail-MessageSnippet, '
                        + 'li[data-id]').length;
                    return compose || rows > 0
                           || /[#/](inbox|message)/i.test(location.href);
                }""")
            except Exception:
                return False

        # Poll up to ~10s for the logged-in state — the persistent session can
        # briefly show the landing page before redirecting into the inbox.
        logged_in_mail = False
        for _ in range(10):
            if await _mail_logged_in():
                logged_in_mail = True
                break
            await asyncio.sleep(1)

        if logged_in_mail:
            log("  2FA: уже залогинен в Яндекс (профиль) — вход пропускаю")
        else:
            # 1) Click "Войти" — ONLY the specific header button (a loose
            #    text match grabs a stray «Войти» ad/footer link → breaks flow).
            await _click_any(['#header-login-button'], "нажал «Войти»")
            await asyncio.sleep(2)

            # 2) Enter email/username (wait for passport page to render)
            await _fill_any(
                ['input[data-testid="text-field-input"][autocomplete="username"]',
                 'input[autocomplete="username"]',
                 'input[name="login"]', '#passp-field-login'],
                mail_email, f"логин {mail_email}", wait=15000)

            # 3) Click "Next"
            await _click_any(['button[data-testid="add-user-next"]',
                              'button:has-text("Next")', 'button:has-text("Далее")'],
                             "нажал Next")
            await asyncio.sleep(2)

            # 4) Choose "Log in with your password" (skip the one-time-code step)
            await _click_any(['span:has-text("Log in with your password")',
                              'button:has-text("Log in with your password")',
                              'span:has-text("Войти с паролем")',
                              'button:has-text("Войти с паролем")'],
                             "выбрал «Log in with your password»")
            await asyncio.sleep(1.5)

            # 5) Enter password
            filled_pw = await _fill_any(
                ['input[data-testid="text-field-input"][autocomplete="current-password"]',
                 'input[autocomplete="current-password"]',
                 'input[type="password"]', '#passp-field-passwd'],
                mail_password, "пароль почты", wait=15000)
            if filled_pw:
                await page.keyboard.press("Enter")
                await asyncio.sleep(1)
                await _click_any(['button[data-testid="add-user-next"]',
                                  'button:has-text("Sign in")',
                                  'button:has-text("Войти")'],
                                 "подтвердил вход", wait=4000)

            # 6) Wait for the inbox to load after login
            try:
                await page.wait_for_url(
                    lambda u: "mail.yandex" in u and "passport" not in u,
                    timeout=25000)
            except Exception:
                await asyncio.sleep(5)
        log("  2FA: в почте, ищу письма с кодами …")

        # MyHeritage sends TWO emails with DIFFERENT codes:
        #   B) "Ваш код подтверждения для входа в MyHeritage" / "verification
        #      code" — this is the code the «Введите верификационный код» field
        #      actually wants (priority).
        #   A) "Confirm a login attempt on MyHeritage" — the anti-fraud
        #      challenge ("flagged as suspicious"); its code is the fallback.
        # We read BOTH from the inbox snippets and return them in priority
        # order so the caller can try B first, then A.
        async def _scan_codes():
            # Read the code STRAIGHT FROM THE INBOX LIST (the email's row header
            # + preview), WITHOUT opening the message — that is how it worked.
            # CRUCIAL: match WHOLE message rows, not leaf elements. The 6-digit
            # code sits in the PREVIEW while the «код подтверждения для входа»
            # phrase is in the SUBJECT; with broad 'a,li,div' selectors they land
            # in DIFFERENT elements, so neither matches both → nothing found.
            # The MessageSnippet/li[data-id] container holds the whole row.
            return await page.evaluate(r"""() => {
                const rows = Array.from(document.querySelectorAll(
                    '[class*="MessageSnippet" i], .mail-MessageSnippet, '
                    + '[data-test-id="message-snippet"], li[data-id], '
                    + 'div[role="listitem"], a[href*="#message"]'));
                let verify = '', attempt = '';
                for (const r of rows) {
                    const t = (r.textContent || '');
                    if (!/myheritage/i.test(t)) continue;
                    const m = t.match(/\b(\d{6})\b/);
                    if (!m) continue;                 // row with no 6-digit → skip
                    const code = m[1];
                    if (!verify && /(подтверждения для входа|верификацион|verification code|confirmation code)/i.test(t))
                        verify = code;                // B — the one the field wants
                    else if (!attempt && /(login attempt|confirm a login|попытк)/i.test(t))
                        attempt = code;               // A — anti-fraud challenge
                    else if (!verify)
                        verify = code;                // MyHeritage + code, no clear phrase → treat as B
                }
                return {verify, attempt};
            }""")

        async def _open_and_read():
            """Snippet scan can miss the code (truncated/changed DOM) — so OPEN
            the newest MyHeritage *verification* email and read the code from its
            body. This is the step that used to work and was lost in a refactor.
            We match MyHeritage + a verification phrase, so the hh.ru ad at the
            very top is skipped (the lesson's «не открывать верхнее письмо»)."""
            try:
                marked = await page.evaluate(r"""() => {
                    const norm = s => (s || '').toLowerCase();
                    document.querySelectorAll('[data-pw-mail]').forEach(
                        e => e.removeAttribute('data-pw-mail'));
                    const rows = Array.from(document.querySelectorAll(
                        '[class*="MessageSnippet"], .mail-MessageSnippet, '
                        + '[data-test-id="message-snippet"], li[data-id], a'));
                    for (const r of rows) {                 // DOM order = newest first
                        const t = norm(r.textContent);
                        if (t.includes('myheritage') &&
                            /(код подтверждения для входа|верификацион|verification code|confirmation code)/.test(t)) {
                            r.setAttribute('data-pw-mail', '1');
                            return true;
                        }
                    }
                    return false;
                }""")
                if not marked:
                    return ""
                log("  2FA: открываю письмо MyHeritage с кодом …")
                await page.locator('[data-pw-mail="1"]').first.click(timeout=6000)
                await asyncio.sleep(2.5)
                # Prefer the OPENED message's body container (so we don't pick an
                # older code still shown in the inbox list); fall back to body.
                body = await page.evaluate(r"""() => {
                    const el = document.querySelector(
                        '[class*="MessageBody" i], .mail-Message-Body, '
                        + '[class*="message-body" i], [class*="msgBody" i], '
                        + '[class*="MessageViewer" i]');
                    return (el && el.innerText) || document.body.innerText || '';
                }""")
                m = (re.search(
                        r"(?:странице входа в MyHeritage|login screen|"
                        r"verification code|код подтверждения)[^\d]{0,40}(\d{6})",
                        body or "", re.I)
                     or re.search(r"\b(\d{6})\b", body or ""))
                code = m.group(1) if m else ""
                if code:
                    log(f"  2FA: код из ОТКРЫТОГО письма: {code}")
                # Return to the inbox so the next reload/scan keeps working.
                try:
                    await page.goto("https://mail.yandex.ru/",
                                    wait_until="domcontentloaded", timeout=15000)
                except Exception:
                    pass
                return code
            except Exception as e:
                log(f"  2FA: !! не смог открыть письмо: {e}")
                return ""

        deadline = asyncio.get_event_loop().time() + timeout
        last = {"verify": "", "attempt": ""}
        while asyncio.get_event_loop().time() < deadline:
            # Stop if the mail tab was closed (don't spin on a dead page).
            if page.is_closed():
                log("  2FA: вкладка почты закрыта — прекращаю чтение")
                break
            # Let the freshly-sent emails arrive, then reload to get them on top
            await asyncio.sleep(5)
            try:
                await page.reload(wait_until="domcontentloaded", timeout=15000)
                await asyncio.sleep(2)
            except Exception:
                pass
            try:
                last = await _scan_codes()
            except Exception:
                last = {"verify": "", "attempt": ""}
            # Snippet scan missed the verification code (B) → OPEN the email and
            # read the code from its body (the reliable path that used to work).
            if not last.get("verify"):
                opened = await _open_and_read()
                if opened:
                    last["verify"] = opened
            # We have at least the verification code (B) — good to go
            if last.get("verify") or last.get("attempt"):
                codes = []
                if last.get("verify"):
                    codes.append(last["verify"])
                    log(f"  2FA: код «подтверждения для входа» (B): {last['verify']}")
                if last.get("attempt") and last["attempt"] not in codes:
                    codes.append(last["attempt"])
                    log(f"  2FA: код «login attempt» (A): {last['attempt']}")
                # Give the second email a couple seconds in case only one is in yet
                if not last.get("verify"):
                    await asyncio.sleep(4)
                    try:
                        again = await _scan_codes()
                        if again.get("verify") and again["verify"] not in codes:
                            codes.insert(0, again["verify"])
                            log(f"  2FA: код «подтверждения для входа» (B): {again['verify']}")
                    except Exception:
                        pass
                return codes

        log("  2FA: коды в почте не найдены за отведённое время")
        return []
    except Exception as e:
        log(f"  2FA browser error: {e}")
        return []
    finally:
        # Close the mail tab regardless of outcome, then re-enable the auto-closer
        try:
            await page.close()
        except Exception:
            pass
        try:
            ctx._mh_pause_autoclose = False
        except Exception:
            pass


async def _get_2fa_codes(ctx, mail_email: str, mail_password: str,
                         log, ask_2fa_code=None) -> list:
    """
    Return a PRIORITY LIST of candidate 2FA codes (MyHeritage sends two
    different emails). The caller tries them in order.
      1. Browser: open a new tab to Yandex mail, read both codes.
      2. IMAP fallback for non-Yandex providers.
      3. GUI dialog as a last resort.
    `mail_email` is the SAME email used for MyHeritage login.
    """
    if mail_email and mail_password:
        codes = await _browser_read_yandex_code(ctx, mail_email, mail_password, log)
        if codes:
            return codes
        log("  2FA: пробую IMAP как запасной вариант …")
        code = await asyncio.to_thread(
            _imap_read_mh_code, mail_email, mail_password, 60)
        if code:
            log("  2FA: код получен по IMAP ✓")
            return [code]

    if ask_2fa_code:
        log("  2FA: запрашиваю код у пользователя …")
        c = ask_2fa_code()
        return [c] if c else []
    return []


# ── LOGIN ─────────────────────────────────────────────────────────────────── #
async def _login(page, login_url, has_cookies,
                 email, password, log,
                 ask_2fa_code=None,
                 imap_password=None) -> bool:
    """
    Full login flow:
    1. Navigate to login_url
    2. Accept cookie banner if has_cookies
    3. Open login form if needed
    4. Fill email + password
    5. Handle 2-FA code dialog if it appears
    """
    log(f"  → Navigating to {login_url} …")
    try:
        await page.goto(login_url, wait_until="domcontentloaded", timeout=35000)
    except Exception as exc:
        log(f"  !! Cannot load login page: {exc}")
        return False

    await asyncio.sleep(2)
    log(f"     Landed: {page.url}")

    # Step 1: cookies
    if has_cookies:
        await _accept_cookies(page, log)
        await asyncio.sleep(0.5)

    # Step 0: ALREADY logged in? With the persistent profile MyHeritage
    # redirects away from /login (e.g. to the family-site home). If we are not
    # on a login/signin page, there's nothing to do — skip the whole form.
    cur = page.url.lower()
    if "login" not in cur and "signin" not in cur:
        log("  ✓ Уже авторизован (сессия из профиля) — логин не нужен")
        # Close any social popup tab that may have opened
        await _close_social_tabs(page.context, page, log)
        return True

    # Step 2: open login form (click nav link if needed)
    await _open_login_form(page, log)

    # Step 3: wait for email field
    # Exact selectors from real page HTML (RU site, 2025):
    #   <input type="text" name="registrationEmail" id="registrationEmail" autocomplete="off">
    #   <input type="password" name="registrationLoginPassword" id="registrationLoginPassword">
    EMAIL_SELS = [
        '#registrationEmail',
        'input[name="registrationEmail"]',
        'input[data-automations="login_email_block_input"]',
        '#email',
        'input[name="email"]',
        'input[type="email"]',
        'input[placeholder*="эл. почты" i]',
        'input[placeholder*="email" i]',
        'input[placeholder*="mail" i]',
        'input[placeholder*="דוא" i]',
        'input[placeholder*="почт" i]',
    ]
    PASS_SELS = [
        '#registrationLoginPassword',
        'input[name="registrationLoginPassword"]',
        '#password',
        'input[name="password"]',
        'input[type="password"]',
        'input[placeholder*="Пароль" i]',
        'input[placeholder*="пароль" i]',
        'input[placeholder*="assword" i]',
        'input[placeholder*="סיסמ" i]',
    ]
    # Wait up to 10s for the email field to appear
    email_visible = False
    for sel in EMAIL_SELS:
        try:
            await page.locator(sel).first.wait_for(state="visible", timeout=6000)
            email_visible = True
            log(f"  ✓ Email field visible: {sel}")
            break
        except Exception:
            continue

    if not email_visible:
        log("  !! Email field not found — saving screenshot for debugging…")
        try:
            ss = Path(__file__).resolve().parent / "debug_login.png"
            await page.screenshot(path=str(ss), full_page=True)
            log(f"     Screenshot saved: {ss}")
        except Exception:
            pass
        return False

    # Step 4: fill form
    ok_e = await _fill(page, EMAIL_SELS, email,    "Email",    log)
    ok_p = await _fill(page, PASS_SELS,  password, "Password", log)
    if not ok_e or not ok_p:
        return False

    # Step 5: submit
    SUBMIT_SELS = [
        'button:has-text("היכנס")',        # Hebrew submit
        'button:has-text("Авторизация")',  # Russian submit (seen in screenshot)
        'button:has-text("Войти")',
        'button:has-text("Log in")',
        'button:has-text("Sign in")',
        'button:has-text("Login")',
        'button[type="submit"]',
        'input[type="submit"]',
    ]
    submitted = False
    for sel in SUBMIT_SELS:
        try:
            el = page.locator(sel).first
            if await el.count():
                await el.click(timeout=6000)
                submitted = True
                log(f"  ✓ Clicked submit ({sel})")
                break
        except Exception:
            continue
    if not submitted:
        await page.keyboard.press("Enter")

    # ── Step 6: wait for EITHER login success OR a 2FA code field ──────── #
    # MyHeritage often shows a "verify it's you" code screen on the SAME
    # /login URL after submit, and it can take >5s to appear. So we POLL.
    TFA_SELS = [
        'input[placeholder*="код" i]',
        'input[placeholder*="code" i]',
        'input[placeholder*="verification" i]',
        'input[placeholder*="верификац" i]',
        'input[type="number"][maxlength]',
        'input[maxlength="6"]',
        'input[autocomplete="one-time-code"]',
        'input[autocomplete*="one-time"]',
        'input[name*="code"]',
        'input[name*="otp"]',
        'input[name*="token"]',
        'input[inputmode="numeric"]',
    ]

    async def _logged_in() -> bool:
        u = page.url.lower()
        if any(x in u for x in ("login", "signin", "verify", "auth")):
            return False
        # Guard: if a password or code field is still visible we are NOT in yet
        for sel in ('input[type="password"]', 'input[autocomplete="one-time-code"]',
                    'input[maxlength="6"]'):
            try:
                if await page.locator(sel).first.is_visible(timeout=500):
                    return False
            except Exception:
                pass
        return True

    async def _find_tfa():
        for sel in TFA_SELS:
            try:
                el = page.locator(sel).first
                if await el.count() and await el.is_visible():
                    return sel
            except Exception:
                pass
        return None

    tfa_sel = None
    for _ in range(30):                       # poll up to ~30s
        await asyncio.sleep(1)
        # Check for the 2FA code field FIRST — MH may sit on a non-/login URL
        # while showing the verification screen, so a premature "logged in"
        # check would wrongly skip it.
        tfa_sel = await _find_tfa()
        if tfa_sel:
            log(f"  ⚠  2FA code field detected ({tfa_sel})")
            break
        if await _logged_in():
            log(f"  ✓ Logged in (no 2FA). URL: {page.url}")
            return True

    if tfa_sel:
        # MyHeritage sends two emails with two different codes — get both.
        codes = await _get_2fa_codes(page.context, email, imap_password,
                                     log, ask_2fa_code)
        if not codes:
            log("  !! 2FA: код не получен — прерываю.")
            return False

        async def _digits_in_field() -> str:
            try:
                return await page.evaluate("""() => {
                    const els = document.querySelectorAll(
                        'input[inputmode="numeric"], input[autocomplete="one-time-code"], '
                        + 'input[maxlength="6"], input[name*="code"], input[type="number"]');
                    let s = '';
                    els.forEach(e => { if (e.offsetParent !== null) s += (e.value || ''); });
                    return s.replace(/\\D/g, '');
                }""")
            except Exception:
                return ""

        async def _clear_code():
            """Clear all visible code boxes before entering a new code."""
            try:
                boxes = page.locator(
                    'input[inputmode="numeric"], input[autocomplete="one-time-code"], '
                    'input[maxlength="6"], input[name*="code"]')
                n = await boxes.count()
                for i in range(max(n, 1)):
                    b = boxes.nth(i) if n else page.locator(tfa_sel).first
                    try:
                        await b.click(timeout=1500)
                        await page.keyboard.press("Control+a")
                        await page.keyboard.press("Delete")
                    except Exception:
                        pass
            except Exception:
                pass

        async def _enter_code(code: str) -> bool:
            for _ in range(3):
                try:
                    fld = page.locator(tfa_sel).first
                    await fld.click(timeout=4000)
                    await page.keyboard.press("Control+a")
                    await page.keyboard.press("Delete")
                    await page.keyboard.type(code, delay=140)
                    await asyncio.sleep(0.6)
                except Exception:
                    pass
                if code in (await _digits_in_field()):
                    return True
                # segmented field — one digit per box
                try:
                    boxes = page.locator(
                        'input[inputmode="numeric"], input[autocomplete="one-time-code"]')
                    if await boxes.count() >= len(code):
                        for i, ch in enumerate(code):
                            b = boxes.nth(i)
                            await b.click(timeout=2000); await b.fill(ch)
                            await asyncio.sleep(0.1)
                        await asyncio.sleep(0.5)
                except Exception:
                    pass
                if code in (await _digits_in_field()):
                    return True
                await asyncio.sleep(0.4)
            return False

        async def _submit_once() -> bool:
            # The confirm button lives INSIDE the verification MODAL (next to the
            # code field). On the page behind it there is ALSO an "Авторизация"
            # button (the login form) — we must NOT click that one. So: find the
            # code field, climb to its modal/dialog/form container, locate the
            # action button inside it, tag it, and click THAT with a real click.
            tagged = await page.evaluate(r"""() => {
                document.querySelectorAll('[data-pw-2fa]').forEach(
                    e => e.removeAttribute('data-pw-2fa'));
                const norm = s => (s || '').replace(/\s+/g, ' ').trim();
                const fld = document.querySelector(
                    'input[inputmode="numeric"], input[autocomplete="one-time-code"], '
                    + 'input[maxlength="6"], input[name*="code"]');
                if (!fld) return false;
                // climb to a reasonable container (dialog/modal/form)
                let box = fld.closest('[role=dialog], .modal, form') || fld.parentElement;
                for (let i = 0; i < 5 && box && box.parentElement; i++) {
                    const btns = box.querySelectorAll('button, [role=button]');
                    // a submit-ish button whose text is Отправить/Авторизация/Continue
                    const re = /^(Отправить|Авторизация|Continue|Продолжить|Verify|Подтвердить|Submit|Send)$/i;
                    let hit = Array.from(btns).find(b => re.test(norm(b.textContent)));
                    if (!hit && btns.length) hit = btns[btns.length - 1];
                    if (hit) { hit.setAttribute('data-pw-2fa', '1'); return true; }
                    box = box.parentElement;
                }
                return false;
            }""")
            if tagged:
                try:
                    btn = page.locator('[data-pw-2fa="1"]').first
                    await btn.scroll_into_view_if_needed(timeout=3000)
                    await btn.click(timeout=5000)
                    log("  ✓ 2FA: нажал кнопку подтверждения в окне кода")
                    return True
                except Exception as e:
                    log(f"  2FA: реальный клик не вышел ({e}); пробую JS-клик")
                    try:
                        await page.evaluate(
                            "() => { const b=document.querySelector('[data-pw-2fa=\"1\"]');"
                            " if (b) b.click(); }")
                        log("  ✓ 2FA: нажал кнопку (JS)")
                        return True
                    except Exception:
                        pass
            # Last resort: a single Enter
            await page.keyboard.press("Enter")
            log("  2FA: Enter в поле кода")
            return False

        # Bring MH tab back to front (mail tab was focused) so input lands here
        try:
            await page.bring_to_front()
            await asyncio.sleep(0.4)
        except Exception:
            pass

        # Try each candidate code: enter → submit ONCE → wait for success.
        for ci, code in enumerate(codes, 1):
            log(f"  → Пробую код {ci}/{len(codes)}: {code}")
            await _clear_code()
            ok = await _enter_code(code)
            log("  ✓ 2FA: код введён" if ok
                else "  !! 2FA: код не подтвердился в поле — всё равно отправляю")
            await asyncio.sleep(0.4)
            await _submit_once()
            log("  2FA: отправил, жду результат…")
            for _ in range(18):
                await asyncio.sleep(1)
                if await _logged_in():
                    log(f"  ✓ Logged in after 2FA. URL: {page.url}")
                    return True
            log(f"  !! 2FA: код {code} не подошёл, пробую следующий…")
            await asyncio.sleep(1)

        log("  !! 2FA: ни один код не подошёл.")

    # Neither success nor 2FA → report any error message
    cur = page.url
    try:
        err = (await page.locator('[class*="error" i],[role="alert"]')
               .first.text_content(timeout=2000) or "").strip()
        if err:
            log(f"  !! Login error: {err!r}")
    except Exception:
        pass
    log(f"  !! Login failed. URL: {cur}")
    return False

# ── SELECT FAMILY SITE (FP/select-site.php) ───────────────────────────────── #
async def _handle_select_site(page, family_site, log):
    """
    Determine the research search URL for the chosen family site.
      - On FP/select-site.php → click the chosen site (by name, else first).
      - Already on a family-sites/<slug>/<ID> page → extract <ID> from the URL.
    Returns the research URL (research?s=<id>) or None.
    """
    cur = page.url
    # Case B: we're already inside a family site → derive id from the URL
    m = re.search(r"/family-sites/[^/]+/([A-Z0-9]+)", cur)
    if m and "select-site" not in cur.lower():
        sid = m.group(1)
        lang = "RU"
        lm = re.search(r"[?&]lang=([A-Za-z]+)", cur)
        if lm:
            lang = lm.group(1)
        log(f"  → Уже на семейном сайте (id {sid}) — иду в поиск")
        return f"https://www.myheritage.com/research?s={sid}&lang={lang}"

    if "select-site" not in cur.lower():
        return None
    log("  → Страница выбора семейного сайта")
    await asyncio.sleep(1.5)

    # Each site link:  <a ... onclick="goToSiteClicked('SITEID',
    #   'https://www.myheritage.com/family-sites/<slug>/<SITEID>?lang=..')">Name</a>
    target = await page.evaluate(r"""(wanted) => {
        const links = Array.from(document.querySelectorAll('a[onclick*="goToSiteClicked"]'));
        const parse = (a) => {
            const m = (a.getAttribute('onclick') || '')
                .match(/goToSiteClicked\(\s*'([^']+)'\s*,\s*'([^']+)'/);
            return m ? {id: m[1], url: m[2], name: (a.textContent||'').trim()} : null;
        };
        const all = links.map(parse).filter(Boolean);
        if (!all.length) return null;
        if (wanted) {
            const hit = all.find(s => s.name.toLowerCase()
                                       .includes(wanted.toLowerCase()));
            if (hit) return hit;
        }
        return all[0];   // default: first (usually the admin's own site)
    }""", family_site or "")

    if not target:
        log("  !! Сайты не найдены на странице выбора")
        return None
    log(f"  → Перехожу на сайт: {target['name']}")
    await page.goto(target["url"], wait_until="domcontentloaded", timeout=35000)
    await asyncio.sleep(2)
    # Build the research search URL for this site: /research?s=<id>&lang=..
    lang = "RU"
    m = re.search(r"[?&]lang=([A-Za-z]+)", target["url"])
    if m:
        lang = m.group(1)
    return f"https://www.myheritage.com/research?s={target['id']}&lang={lang}"


# ── close stray social popup tabs (Facebook/Google) ───────────────────────── #
async def _close_social_tabs(ctx, keep, log):
    for p in list(ctx.pages):
        if p is keep:
            continue
        try:
            u = (p.url or "").lower()
            if any(s in u for s in ("facebook.com", "accounts.google", "apple.com")):
                await p.close()
                log("  ✓ Закрыл лишнюю вкладку соцсети")
        except Exception:
            pass


# ── SEARCH FORM ───────────────────────────────────────────────────────────── #
FIRST_NAME_SELS = [
    'input[data-automations="research-family_first_name"]',
    'input[placeholder="Имя и отчество"]',
    'input[placeholder*="Имя" i]', 'input[placeholder*="first" i]',
]
LAST_NAME_SELS = [
    'input[data-automations="research-family_last_name"]',
    'input[placeholder="Фамилия"]',
    'input[placeholder*="Фамилия" i]', 'input[placeholder*="last" i]',
]


async def _find_form_root(page, sels, log, secs=25):
    """
    Return the frame (main or iframe) that contains the research search form,
    polling up to `secs`. Detect by the first-name field OR the «Поиск» submit
    button (the form can be in the main frame or an iframe).
    """
    _sw = _search_words()
    btn_sels = ([f'button:has(span.button_content:has-text("{w}"))' for w in _sw]
                + [f'button:has-text("{w}")' for w in _sw])
    for _t in range(secs):
        await asyncio.sleep(1)
        for fr in page.frames:
            for sel in list(sels) + btn_sels:
                try:
                    if await fr.locator(sel).first.count():
                        log(f"  ✓ Форма поиска готова ({_t+1}с"
                            f"{', iframe' if fr is not page.main_frame else ''})")
                        return fr
                except Exception:
                    continue
    return None


async def _fill_research_basic(root, page, params, log) -> bool:
    """
    Robustly fill the basic research fields. Selectors for the inputs vary, but
    the «Поиск» button is reliable — so we walk up from it to the form
    container and TAG its text inputs (by data-automations / placeholder /
    order), then fill the tagged fields. Works in main frame or iframe.
    """
    tagged = await root.evaluate(r"""(searchWords) => {
        const norm = s => (s || '').replace(/\s+/g, ' ').trim();
        const vis = el => el && el.offsetParent !== null;
        const isSearch = t => searchWords.some(w => norm(t).toLowerCase() === w.toLowerCase());
        // find the search submit button (in the active site language)
        let btn = Array.from(document.querySelectorAll('button, [role=button], span'))
            .find(b => isSearch(b.textContent));
        // climb to a container that holds >= 2 visible text inputs (the form)
        let form = null, n = btn;
        for (let i = 0; i < 9 && n; i++) {
            const ins = Array.from(n.querySelectorAll('input'))
                .filter(x => vis(x) && /^(text|search|number|)$/i.test(x.type || ''));
            if (ins.length >= 2) { form = n; break; }
            n = n.parentElement;
        }
        if (!form) form = document.querySelector('form') || document.body;
        const inputs = Array.from(form.querySelectorAll('input'))
            .filter(x => vis(x) && /^(text|search|number|)$/i.test(x.type || ''));
        document.querySelectorAll('[data-pw-rf]').forEach(e => e.removeAttribute('data-pw-rf'));
        const byAuto = sub => inputs.find(i =>
            ((i.getAttribute('data-automations') || '').toLowerCase()).includes(sub));
        const byPh = re => inputs.find(i => re.test(i.getAttribute('placeholder') || ''));
        const first = byAuto('first_name') || byPh(/Имя|first/i) || inputs[0];
        const last  = byAuto('last_name')  || byPh(/Фамилия|last|surname/i) || inputs[1];
        const year  = byAuto('birth')      || byPh(/Год рождения|birth\s*year/i);
        const place = byAuto('place')      || byPh(/Насел|place/i);
        const tag = (el, name) => { if (el) el.setAttribute('data-pw-rf', name); };
        tag(first, 'first'); tag(last, 'last'); tag(year, 'year'); tag(place, 'place');
        return {first: !!first, last: !!last, year: !!year, place: !!place,
                count: inputs.length};
    }""", _search_words(params.get("lang")))
    log(f"  → Поля формы: {tagged}")

    async def _type(name, value, label):
        if not value:
            return
        try:
            el = root.locator(f'[data-pw-rf="{name}"]').first
            if not await el.count():
                log(f"  !! Field not found: {label}")
                return
            await el.scroll_into_view_if_needed(timeout=4000)
            await el.click(timeout=3000)
            await page.keyboard.press("Control+a")
            await page.keyboard.press("Delete")
            await page.keyboard.type(value, delay=40)
            await asyncio.sleep(0.2)
            log(f"  ✓ {label}: OK")
        except Exception as e:
            log(f"  !! {label}: {e}")

    await _type("first", params.get("first_name", ""), "First name")
    await _type("last",  params.get("surname", ""),    "Surname")
    await _type("year",  params.get("birth_year", ""), "Birth year")
    await _type("place", params.get("birth_place", ""), "Birth place")
    return bool(tagged and (tagged.get("first") or tagged.get("last")))


async def _fill_root(root, page, sels, value, label, log):
    """Fill a field inside `root` (page or frame); type via page.keyboard so
    it works whether the field is in the main frame or an iframe."""
    if not value:
        return True
    for sel in sels:
        try:
            el = root.locator(sel).first
            if not await el.count():
                continue
            await el.scroll_into_view_if_needed(timeout=4000)
            await el.click(timeout=3000)
            await page.keyboard.press("Control+a")
            await page.keyboard.press("Delete")
            await page.keyboard.type(value, delay=40)
            await asyncio.sleep(0.2)
            if (await el.input_value(timeout=2000)).strip():
                log(f"  ✓ {label}: OK")
                return True
        except Exception:
            continue
    log(f"  !! Field not found: {label}")
    return False


# ── ADVANCED SEARCH (pills → popup with data-automations → «Применить») ─────── #
async def _fill_advanced(root, page, params, log):
    """
    Each advanced detail is a tag/pill (div.tag_label «Отец»/«Мать»/«Супруг(-а)»/
    «Смерть»). Clicking it opens a popup whose inputs have exact data-automations
    ids; after filling we MUST click the «Применить» (apply) button.
    "+ Больше" opens extra event/relative/other tags + «Точное совпадение всех
    параметров».

    On the EN/HE sites these pills carry English/Hebrew text — the site language
    comes from params["lang"] and drives _ui_labels(concept, lang).
    """
    lang = params.get("lang")
    async def _open_pill(texts, verify_auto_id):
        """REAL mouse-click the visible chip (JS .click() does NOT open the
        React popup), then confirm the popup opened by waiting for its input
        (verify_auto_id) to become visible. Returns True only if it opened."""
        for fr in page.frames:
            for t in texts:
                # tag via JS so we can click the right element via Playwright
                try:
                    tagged = await fr.evaluate(r"""(label) => {
                        const norm = s => (s || '').replace(/\s+/g, ' ').trim();
                        document.querySelectorAll('[data-pw-pill]').forEach(
                            e => e.removeAttribute('data-pw-pill'));
                        const els = Array.from(document.querySelectorAll(
                            'div, span, button, [role=button], [class*="tag" i], [class*="chip" i]'));
                        for (const e of els) {
                            const txt = norm(e.textContent);
                            if (!txt || txt.length > 40 || e.offsetParent === null) continue;
                            if (txt === label || txt.startsWith(label + ':')
                                || txt.startsWith(label + ' ')) {
                                e.setAttribute('data-pw-pill', '1');
                                return true;
                            }
                        }
                        return false;
                    }""", t)
                    if not tagged:
                        continue
                    chip = fr.locator('[data-pw-pill="1"]').first
                    # real mouse click (bubbles → React handler fires)
                    await chip.scroll_into_view_if_needed(timeout=3000)
                    await chip.click(timeout=4000)
                    if not verify_auto_id:
                        await asyncio.sleep(0.8)
                        return True
                    # confirm the popup's input is now visible
                    for _ in range(8):
                        await asyncio.sleep(0.4)
                        v_fr, v_el = await _find_auto(verify_auto_id, want_visible=True,
                                                      wait_secs=1)
                        if v_el:
                            return True
                    # popup didn't open — try clicking the closest clickable parent
                    try:
                        await chip.evaluate(
                            "e => { const b = e.closest('button,[role=button]'); "
                            "if (b && b!==e) b.click(); }")
                        for _ in range(6):
                            await asyncio.sleep(0.4)
                            v_fr, v_el = await _find_auto(verify_auto_id,
                                                          want_visible=True, wait_secs=1)
                            if v_el:
                                return True
                    except Exception:
                        pass
                except Exception:
                    continue
        return False

    async def _find_auto(auto_id, want_visible=True, wait_secs=5):
        """Return (frame, locator) for the VISIBLE data-automations element.
        MyHeritage keeps hidden template duplicates, so we must skip those and
        wait for the popup's real (visible) one to appear."""
        for _ in range(wait_secs * 2):
            for fr in page.frames:
                try:
                    loc = fr.locator(f'[data-automations="{auto_id}"]')
                    n = await loc.count()
                    for i in range(n):
                        el = loc.nth(i)
                        if not want_visible or await el.is_visible():
                            return fr, el
                except Exception:
                    continue
            if not want_visible:
                break
            await asyncio.sleep(0.5)
        return None, None

    async def _fill_auto(auto_id, value):
        if not value:
            return
        fr, el = await _find_auto(auto_id, want_visible=True)
        if not el:
            log(f"  !! {auto_id} не виден")
            return
        # Try a normal click+type first; fall back to a JS value-set with the
        # React native setter + input/change events (handles hidden/animating).
        try:
            await el.click(timeout=3000)
            await page.keyboard.press("Control+a")
            await page.keyboard.press("Delete")
            await page.keyboard.type(str(value), delay=40)
            await asyncio.sleep(0.2)
            if (await el.input_value()).strip():
                log(f"  ✓ {auto_id} = {value!r}")
                return
        except Exception:
            pass
        try:
            await fr.evaluate(r"""([id, val]) => {
                const el = Array.from(document.querySelectorAll('[data-automations="'+id+'"]'))
                    .find(e => e.offsetParent !== null) ||
                    document.querySelector('[data-automations="'+id+'"]');
                if (!el) return;
                const setter = Object.getOwnPropertyDescriptor(
                    window.HTMLInputElement.prototype, 'value').set;
                el.focus(); setter.call(el, val);
                el.dispatchEvent(new Event('input',  {bubbles: true}));
                el.dispatchEvent(new Event('change', {bubbles: true}));
            }""", [auto_id, str(value)])
            log(f"  ✓ {auto_id} = {value!r} (JS)")
        except Exception as e:
            log(f"  !! {auto_id}: {e}")

    async def _apply(auto_id):
        """Click the popup's «Применить» button — exact id first, else generic."""
        fr, el = await _find_auto(auto_id, want_visible=True, wait_secs=2)
        if el:
            try:
                await el.click(timeout=4000)
                await asyncio.sleep(0.6)
                return True
            except Exception:
                pass
        for fr in page.frames:
            for sel in ('button.search_form_apply_button',
                        'button:has-text("Применить")', 'button:has-text("Apply")',
                        'button:has-text("החל")'):
                try:
                    b = fr.locator(sel).first
                    if await b.count() and await b.is_visible():
                        await b.click(timeout=4000)
                        await asyncio.sleep(0.6)
                        return True
                except Exception:
                    continue
        return False

    # ── Family members: pill → two name inputs → apply ──────────────────────
    fam = [
        ("father", _ui_labels("father", lang), "father_first_name_field",
         "father_last_name_field", "father_filter_apply_button"),
        ("mother", _ui_labels("mother", lang), "mother_first_name_field",
         "mother_last_name_field", "mother_filter_apply_button"),
        ("spouse", _ui_labels("spouse", lang), "spouse_first_name_field",
         "spouse_last_name_field", "spouse_filter_apply_button"),
    ]
    for key, tags, fa, la, apply_id in fam:
        fv = params.get(key, "")
        lv = params.get(f"{key}_last", "")
        if not fv and not lv:
            continue
        if await _open_pill(tags, fa):
            log(f"  → открыл «{tags[0]}»")
            await _fill_auto(fa, fv)
            await _fill_auto(la, lv)
            await _apply(apply_id)
        else:
            log(f"  !! «{tags[0]}»: попап не открылся")

    # ── Death: pill → year + place → apply ──────────────────────────────────
    if params.get("death_year") or params.get("death_place"):
        if await _open_pill(_ui_labels("death", lang), "death_year_field"):
            log("  → открыл «Смерть»")
            await _fill_auto("death_year_field", params.get("death_year", ""))
            await _fill_auto("death_place_field", params.get("death_place", ""))
            await _apply("death_filter_apply_button")
        else:
            log("  !! «Смерть»: попап не открылся")

    # ── "+ Больше": extra events / keywords / gender / exact-all ─────────────
    need_more = (any(params.get(f) for f in
                     ("residence", "military", "immigration", "keywords"))
                 or params.get("gender", "Any") not in ("Any", "Любой")
                 or params.get("exact_match"))
    if need_more:
        # Open the "+ Больше" panel and VERIFY it opened by waiting for one of
        # its hallmark texts to appear (the exact-match checkbox / event tags).
        async def _more_open():
            # Hallmark texts of the open "+ More" panel in RU / EN / HE.
            hallmark = re.compile(
                r"Точное совпадение всех параметров|событие из жизни|"
                r"Семейное положение|"
                r"Exact match for all|life event|Marital status|"
                r"התאמה מדויקת|אירוע|מצב משפחתי|"
                r"Correspondance exacte|État civil|"          # fr
                r"Genaue Übereinstimmung|Familienstand|"      # de
                r"Coincidencia exacta|Estado civil|"          # es
                r"Correspondência exata", re.I)               # pt
            for fr in page.frames:
                try:
                    el = fr.get_by_text(hallmark).first
                    if await el.count() and await el.is_visible():
                        return True
                except Exception:
                    continue
            return False

        # "+ Больше" is a TOGGLE — clicking it twice closes it again. So check
        # whether it's already open BEFORE each click, and click only if not.
        opened_more = False
        for _attempt in range(4):
            if await _more_open():
                opened_more = True
                break
            await _open_pill(_ui_labels("more", lang), "")
            await asyncio.sleep(1.0)
        log("  → панель «+ Больше» открыта" if opened_more
            else "  !! панель «+ Больше» не открылась")

        for key, tags, place_auto, apply_id in [
            ("residence", _ui_labels("residence", lang),
             "residence_place_field", "residence_filter_apply_button"),
            ("military", _ui_labels("military", lang),
             "military_place_field", "military_filter_apply_button"),
            ("immigration", _ui_labels("immigration", lang),
             "immigration_place_field", "immigration_filter_apply_button"),
        ]:
            if params.get(key):
                if await _open_pill(tags, place_auto):
                    await _fill_auto(place_auto, params[key])
                    await _apply(apply_id)

        if params.get("keywords"):
            if await _open_pill(_ui_labels("keywords", lang), "keywords_field"):
                await _fill_auto("keywords_field", params["keywords"])
                await _apply("keywords_filter_apply_button")

        g = params.get("gender", "Any")
        if g not in ("Any", "Любой"):
            if await _open_pill(_ui_labels("gender", lang), ""):
                # Click the gender value in the site language (try every variant;
                # only the one rendered on the loaded site exists).
                concept = "female" if g in ("Female", "Женщина") else "male"
                for val in _ui_labels(concept, lang):
                    try:
                        opt = root.get_by_text(
                            re.compile(rf"^{re.escape(val)}$", re.I)).first
                        if await opt.count():
                            await opt.click(timeout=3000)
                            break
                    except Exception:
                        continue
                await _apply("gender_filter_apply_button")

        # «Точное совпадение всех параметров». aria-checked never flips by
        # plain clicks → the click is intercepted by the visual checkbox box
        # and/or aria-checked is decorative. So: tag the label; FORCE-click
        # several candidate targets (the visible checkbox box, the parent
        # control, the label) at real coordinates; verify "checked" via EITHER
        # aria-checked OR a nearby input[type=checkbox].checked, re-read fresh.
        if params.get("exact_match"):
            async def _exact_state():
                """Tag the control; return (frame, checked_bool) or (None,None)."""
                for fr in page.frames:
                    try:
                        st = await fr.evaluate(r"""() => {
                            const norm = s => (s || '').replace(/\s+/g, ' ').trim();
                            document.querySelectorAll('[data-pw-ex],[data-pw-exbox]')
                                .forEach(e => { e.removeAttribute('data-pw-ex');
                                                e.removeAttribute('data-pw-exbox'); });
                            for (const e of document.querySelectorAll('span,label,[role=checkbox]')) {
                                const t = norm(e.textContent);
                                if (t.length > 60) continue;
                                if (!/Точное совпадение всех параметров/i.test(t)) continue;
                                if (e.offsetParent === null) continue;
                                // control = nearest clickable wrapper
                                const ctrl = e.closest('label, [role=checkbox]')
                                          || e.parentElement || e;
                                ctrl.setAttribute('data-pw-ex', '1');
                                // the visual checkbox box (first small sibling/child)
                                const box = ctrl.querySelector(
                                    '[class*="checkbox" i]:not([class*="label" i]), '
                                    + 'svg, [class*="box" i], [class*="control" i]');
                                if (box) box.setAttribute('data-pw-exbox', '1');
                                // state from aria-checked or a real input
                                const inp = ctrl.querySelector('input[type=checkbox]');
                                const aria = (e.closest('[role=checkbox]')
                                    || ctrl).getAttribute('aria-checked');
                                const checked = (inp && inp.checked) || aria === 'true';
                                return checked;
                            }
                            return 'none';
                        }""")
                        if st != 'none':
                            return fr, bool(st)
                    except Exception:
                        continue
                return None, None

            done = False
            for _attempt in range(10):
                fr, checked = await _exact_state()
                if fr is None:
                    await asyncio.sleep(0.5)
                    continue
                if checked:
                    log("  ✓ Точное совпадение всех параметров")
                    done = True
                    break
                for sel in ('[data-pw-exbox="1"]', '[data-pw-ex="1"]'):
                    try:
                        tgt = fr.locator(sel).first
                        if not await tgt.count():
                            continue
                        await tgt.scroll_into_view_if_needed(timeout=1500)
                        await tgt.click(timeout=2500, force=True)
                        await asyncio.sleep(1.0)
                        _, now = await _exact_state()
                        if now:
                            done = True
                            break
                    except Exception:
                        continue
                if not done:
                    try:
                        await fr.locator('[data-pw-ex="1"]').first.focus()
                        await page.keyboard.press("Space")
                        await asyncio.sleep(1.0)
                        _, now = await _exact_state()
                        if now:
                            done = True
                    except Exception:
                        pass
                if done:
                    log("  ✓ Точное совпадение всех параметров")
                    break
                await asyncio.sleep(0.4)
            params["_exact_ok"] = done
            if not done:
                log("  !! не удалось включить «Точное совпадение всех параметров»")


async def _search(page, search_url, params, has_cookies, log):
    log(f"  → Navigating to search: {search_url}")
    try:
        await page.goto(search_url, wait_until="domcontentloaded", timeout=35000)
    except Exception as exc:
        log(f"  !! Cannot open search page: {exc}")
        return False
    if has_cookies:
        await _accept_cookies(page, log)
    await _close_social_tabs(page.context, page, log)

    # Wait for the research form (searches the main frame AND any iframe)
    log("  → Жду форму поиска…")
    root = await _find_form_root(page, FIRST_NAME_SELS, log, secs=25)
    if root is None:
        log("  !! Форма поиска не появилась — пробую главный фрейм")
        root = page.main_frame

    # Fill basic fields (JS-tag the form's inputs, then type) — robust to the
    # varying selectors; works in main frame or iframe.
    await _fill_research_basic(root, page, params, log)

    # ── Advanced search: pills → popup (data-automations) → «Применить» ──────
    try:
        await _fill_advanced(root, page, params, log)
    except Exception as _exc:
        log(f"  !! advanced search error (продолжаю поиск): {_exc}")

    # Record type filter
    rf = params.get("record_filter", "All Records")
    FILTER_MAP = {
        "Historical Records": ["Исторические записи", "Historical records",
                               "Historical Records", "רשומות היסטוריות",
                               "Documents historiques", "Historische Aufzeichnungen",
                               "Registros históricos", "Registos históricos"],
        "Family Trees":       ["Семейные деревья",    "Family trees",
                               "Family Trees",        "עצי משפחה",
                               "Arbres généalogiques", "Stammbäume",
                               "Árboles genealógicos", "Árvores genealógicas"],
    }
    if rf != "All Records":
        for label in FILTER_MAP.get(rf, [rf]):
            try:
                el = page.get_by_text(re.compile(re.escape(label), re.I)).first
                if await el.count():
                    await el.click(timeout=5000)
                    await asyncio.sleep(1.5)
                    break
            except Exception:
                pass

    # Submit search — results usually open in a NEW TAB
    ctx = page.context
    results_page = None
    _sw = _search_words(params.get("lang"))
    submit_sels = ([f'span.button_content:has-text("{w}")' for w in _sw]
                   + [f'button:has-text("{w}")' for w in _sw]
                   + ['button[type="submit"]', 'input[type="submit"]'])
    pages_before = set(ctx.pages)
    search_url_now = page.url
    clicked = False
    for sel in submit_sels:
        try:
            el = root.locator(sel).first
            if await el.count() and await el.is_visible():
                await el.click(timeout=6000)
                clicked = True
                log(f"  ✓ Нажал «Поиск» ({sel})")
                break
        except Exception:
            pass
    if not clicked:
        await page.keyboard.press("Enter")
        log("  → Поиск: Enter")

    # Results open in a NEW TAB (a MyHeritage results page). Poll up to 25s for
    # a new MH page, or for the current page to navigate to a results URL.
    results_page = None
    for _ in range(25):
        await asyncio.sleep(1)
        new_pages = [p for p in ctx.pages if p not in pages_before]
        for np in new_pages:
            try:
                u = (np.url or "").lower()
                if "myheritage" in u and "research" in u:
                    results_page = np
                    break
            except Exception:
                pass
        if results_page:
            break
        # same-tab navigation to a results page?
        if page.url != search_url_now and "research" in page.url.lower():
            results_page = page
            break
    if results_page is None:
        # fall back to the newest MH tab, else the search page
        mh = [p for p in ctx.pages if "myheritage" in (p.url or "").lower()]
        results_page = mh[-1] if mh else page

    try:
        await results_page.bring_to_front()
        await results_page.wait_for_load_state("networkidle", timeout=20000)
    except Exception:
        await asyncio.sleep(3)
    log(f"  ✓ Страница результатов: {results_page.url[:90]}")

    # EXACT search via the URL (deterministic — the «Точное совпадение всех
    # параметров» checkbox is unreliable). MyHeritage encodes exact mode as
    # exactSearch=1 in the results URL, which already carries every field
    # param. If exact was requested, set it and reload.
    if params.get("exact_match") and not params.get("_exact_ok"):
        u = results_page.url
        # Remember the plain (non-exact) results URL so run_scraper can fall
        # back to it if exactSearch turns out to be too strict (0 records).
        params["_noexact_url"] = u
        if re.search(r"[?&]exactSearch=", u):
            new_u = re.sub(r"([?&]exactSearch=)[^&]*", r"\g<1>1", u)
        else:
            new_u = u + ("&" if "?" in u else "?") + "exactSearch=1"
        if new_u != u:
            try:
                log("  → Применяю точное совпадение через URL (exactSearch=1)…")
                await results_page.goto(new_u, wait_until="domcontentloaded",
                                        timeout=30000)
                # WAIT for the exact results to actually render (the SPA may
                # show a loader first). Poll up to 30s for record links so we
                # never collect an empty page just because it was still loading.
                got = 0
                for _ in range(30):
                    await asyncio.sleep(1)
                    try:
                        got = await results_page.evaluate(
                            "() => document.querySelectorAll("
                            "'a[href*=\"showRecord\"], a[href*=\"recordTitle\"]').length")
                    except Exception:
                        got = 0
                    if got > 0:
                        break
                params["_exact_ok"] = True
                log(f"  ✓ Точный поиск ({got} ссылок на странице): "
                    f"{results_page.url[:80]}")
            except Exception as _e:
                log(f"  !! exactSearch через URL не вышел: {_e}")

    return results_page

async def _collect_one_page(page):
    """
    Extract record cards from the CURRENT results page.
    Returns {rows, stop}: rows are the EXACT matches (those appearing BEFORE
    the «расширения критериев / expanded search criteria» divider); stop=True
    once that divider is seen so pagination halts (the records below it are
    the relaxed, irrelevant matches).
    """
    return await page.evaluate(r"""() => {
        const out = [];
        const seen = new Set();
        // Find the "expanded criteria" divider — everything after it is relaxed.
        // The phrase may be split across child spans, so search the SMALLEST
        // element whose text contains both key words (not just leaf nodes).
        const norm = s => (s || '').replace(/\s+/g, ' ').trim();
        let divider = null, dlen = 1e9;
        for (const e of document.querySelectorAll('div, p, span, section, h1, h2, h3')) {
            const t = norm(e.textContent);
            if (t.length > 300 || t.length >= dlen) continue;
            const ru = /расширени/i.test(t) && /критери/i.test(t);
            const en = /expand/i.test(t) && /criteri/i.test(t);
            const he = /הרחב/.test(t) && /קריטריון|תנאי/.test(t);
            const fr = /élarg/i.test(t)  && /critère/i.test(t);
            const de = /erweiter/i.test(t) && /kriterien/i.test(t);
            const es = /ampli/i.test(t)  && /criterio/i.test(t);
            const pt = /expand|amplia/i.test(t) && /critério/i.test(t);
            if (ru || en || he || fr || de || es || pt) { divider = e; dlen = t.length; }
        }
        const beforeDivider = (el) => {
            if (!divider) return true;
            // el comes before divider in document order?
            return !!(divider.compareDocumentPosition(el)
                      & Node.DOCUMENT_POSITION_PRECEDING);
        };
        const anchors = Array.from(document.querySelectorAll(
            'a[href*="showRecord"], a[href*="recordTitle"]'));
        for (const a of anchors) {
            const href = a.href || '';
            if (!href || seen.has(href)) continue;
            if (!beforeDivider(a)) continue;          // skip relaxed matches
            seen.add(href);
            const card = a.closest('[class*="result" i], li, article, div') || a;
            let name = '';
            const h = card.querySelector('h1,h2,h3,[class*="recordTitle" i],[class*="name" i]');
            if (h) name = norm(h.textContent);
            if (!name) {
                const m = href.match(/recordTitle=([^&]+)/);
                if (m) { try { name = decodeURIComponent(m[1].replace(/\+/g, ' ')); } catch(e){} }
            }
            let score = -1;
            const sm = (card.textContent || '').match(/(\d{1,3})\s*%/);
            if (sm) score = parseInt(sm[1], 10);
            out.push({url: href, name_text: name.slice(0, 200), score});
        }
        return {rows: out, stop: !!divider};
    }""")


async def _set_results_per_page(page, log, want="50"):
    """Open the results-per-page dropdown (data-automations=
    selector_header_container) and pick `want` (e.g. 50) to reduce paging."""
    try:
        hdr = page.locator('[data-automations="selector_header_container"]').first
        if not await hdr.count():
            return
        cur = (await hdr.text_content() or "").strip()
        if cur == want:
            return
        await hdr.click(timeout=4000)
        await asyncio.sleep(0.8)
        # pick the option whose text is exactly `want`
        opt = page.get_by_text(re.compile(rf"^{want}$")).first
        if await opt.count():
            await opt.click(timeout=4000)
            await asyncio.sleep(2)
            log(f"  → Результатов на странице: {want}")
    except Exception:
        pass


async def _goto_next_results(page, log):
    """Click the exact pagination Next icon (a[data-automations='next_icon']).
    Returns True only if the result set actually changed."""
    try:
        nxt = page.locator('a[data-automations="next_icon"]').first
        if not await nxt.count():
            log("  → «Далее»: кнопка next_icon отсутствует — последняя страница")
            return False
        cls = (await nxt.get_attribute("class") or "").lower()
        aria = (await nxt.get_attribute("aria-disabled") or "").lower()
        if "disabled" in cls or aria == "true":
            log("  → «Далее»: кнопка неактивна — последняя страница")
            return False
        # capture current first-record url to detect the change
        before = await page.evaluate(
            """() => { const a = document.querySelector('a[href*=\"showRecord\"]');
                       return a ? a.href : ''; }""")
        await nxt.scroll_into_view_if_needed(timeout=3000)
        await nxt.click(timeout=5000)
        # wait until the first record link changes (AJAX page swap)
        for _ in range(15):
            await asyncio.sleep(1)
            after = await page.evaluate(
                """() => { const a = document.querySelector('a[href*=\"showRecord\"]');
                           return a ? a.href : ''; }""")
            if after and after != before:
                log("  → Перешёл на следующую страницу")
                return True
        log("  → «Далее»: страница не сменилась")
        return False
    except Exception as e:
        log(f"  → «Далее»: {e}")
        return False


# ── COLLECT RESULTS (exact matches across pages) ──────────────────────────── #
async def _diag_results(page, log):
    """One-shot dump of what the results page actually contains, so we can fix
    a 0-results run precisely instead of guessing (the live site can't be
    inspected directly)."""
    try:
        info = await page.evaluate(r"""() => {
            const q = s => { try { return document.querySelectorAll(s).length; }
                             catch(e){ return -1; } };
            const sample = Array.from(document.querySelectorAll('a[href*="record" i]'))
                .slice(0, 6).map(a => (a.getAttribute('href') || '').slice(0, 130));
            const bt = (document.body && document.body.innerText) || '';
            return {
                url: location.href,
                title: document.title,
                showRecord:  q('a[href*="showRecord"]'),
                recordTitle: q('a[href*="recordTitle"]'),
                anyRecord:   q('a[href*="record" i]'),
                resultCards: q('[class*="result" i]'),
                hasForm:     q('input[name*="first" i], input[name*="last" i]'),
                divider: ((/расширени|expand/i.test(bt) && /критери|criteri/i.test(bt))
                          || (/הרחב/.test(bt) && /קריטריון|תנאי/.test(bt))),
                noResults: /(ничего не найдено|no results|0 результатов|не дал|nothing|לא נמצאו)/i.test(bt),
                bodyLen: bt.length,
                sample
            };
        }""")
        log("  🔎 ДИАГНОСТИКА страницы результатов:")
        log(f"     url:   {info.get('url','')[:160]}")
        log(f"     title: {info.get('title','')}")
        log(f"     showRecord={info.get('showRecord')} "
            f"recordTitle={info.get('recordTitle')} "
            f"anyRecord={info.get('anyRecord')} "
            f"resultCards={info.get('resultCards')} "
            f"formInputs={info.get('hasForm')}")
        log(f"     divider≈{info.get('divider')} "
            f"noResultsText≈{info.get('noResults')} "
            f"bodyLen={info.get('bodyLen')}")
        for s in (info.get("sample") or []):
            log(f"     ↪ {s}")
    except Exception as e:
        log(f"  🔎 диагностика не вышла: {e}")


async def _collect(page, log, max_pages=30):
    # Wait for result cards to render
    results, seen_urls = [], set()
    for _t in range(25):
        await asyncio.sleep(1)
        try:
            res = await _collect_one_page(page)
        except Exception:
            res = {"rows": [], "stop": False}
        if res.get("rows"):
            break

    # Show 50 per page to reduce the number of page turns
    await _set_results_per_page(page, log, "50")
    await asyncio.sleep(1)

    page_no = 1
    while page_no <= max_pages:
        try:
            res = await _collect_one_page(page)
        except Exception:
            res = {"rows": [], "stop": False}
        rows = res.get("rows", [])
        stop = res.get("stop", False)
        new = 0
        for r in rows:
            if r["url"] not in seen_urls:
                seen_urls.add(r["url"])
                results.append(r)
                new += 1
        log(f"  → Страница {page_no}: записей {len(rows)} (новых {new}), всего {len(results)}")
        # The «expanded criteria» divider was reached → stop (rest are relaxed)
        if stop:
            log("  → Достигнут разделитель «расширения критериев» — стоп")
            break
        try:
            advanced = await _goto_next_results(page, log)
        except Exception:
            advanced = False
        if not advanced:
            break
        page_no += 1
        await asyncio.sleep(2.5)
        try:
            await page.wait_for_load_state("networkidle", timeout=15000)
        except Exception:
            await asyncio.sleep(2)
    log(f"  → Всего точных записей: {len(results)}")
    # Nothing collected → dump the page so the log tells us why (right page? any
    # record links at all? a "no results" message? still showing the form?).
    if not results:
        await _diag_results(page, log)
    return results

# ── DETAIL PAGE ───────────────────────────────────────────────────────────── #
async def _detail(page, url, has_cookies, log):
    d = {"url": url, "full_name": "", "category": "", "table_data": {},
         "profile_url": "", "source_text": "", "thumb_bytes": None,
         "is_historical": False}
    try:
        await page.goto(url, wait_until="domcontentloaded", timeout=35000)
        if has_cookies:
            await _accept_cookies(page, log)
        await asyncio.sleep(1.5)
        # nudge lazy-loaded images (the record photo) into loading
        try:
            await page.evaluate("() => window.scrollBy(0, 400)")
            await asyncio.sleep(1.2)
            await page.evaluate("() => window.scrollTo(0, 0)")
            await asyncio.sleep(0.5)
        except Exception:
            pass
    except Exception as exc:
        log(f"    !! {exc}")
        return d

    # Extract record data by WHITELISTED genealogy labels (avoids the sidebar
    # "more records"/Geni garbage). Labels are unique enough to scan the page.
    info = await page.evaluate(r"""() => {
        const norm = s => (s || '').replace(/\s+/g, ' ').trim();
        const res = {name: '', category: '', fields: [], profile: '', photo: ''};

        // category line
        let catEl = Array.from(document.querySelectorAll('*')).find(e =>
            !e.children.length &&
            /^(В категории|In category)/i.test(norm(e.textContent)));
        if (catEl) res.category = norm(catEl.textContent)
            .replace(/^В категории:?\s*/i, '').replace(/^In category:?\s*/i, '');

        // record title = the h1 nearest the category line (NOT the account name)
        const h1s = Array.from(document.querySelectorAll('h1, h2'));
        if (catEl) {
            // the title h1 sits just above the category line
            let best = null, bestDist = 1e9;
            const cy = catEl.getBoundingClientRect().top;
            for (const h of h1s) {
                const t = norm(h.textContent);
                if (!t || t.length > 120) continue;
                const dy = cy - h.getBoundingClientRect().top;
                if (dy >= 0 && dy < bestDist) { bestDist = dy; best = h; }
            }
            if (best) res.name = norm(best.textContent);
        }
        if (!res.name && h1s.length) res.name = norm(h1s[0].textContent);

        // WHITELISTED label/value pairs
        const LABELS = ['Имя','Рождение','Смерть','Брак','Крещение','Погребение',
            'Захоронение','Проживание','Местожительство','Пол','Возраст',
            'Отец','Мать','Муж','Жена','Супруг','Супруга','Родители','Дети',
            'Сын','Дочь','Родные брат/сестра','Члены семьи','Иммиграция',
            'Name','Birth','Death','Marriage','Residence','Gender','Father',
            'Mother','Husband','Wife','Spouse','Children'];
        const seen = new Set();
        const valueFor = (labelEl) => {
            // value = next sibling with text, else parent's last child
            let n = labelEl.nextElementSibling;
            while (n && !norm(n.textContent)) n = n.nextElementSibling;
            if (n && norm(n.textContent)) return norm(n.textContent);
            const par = labelEl.parentElement;
            if (par) {
                const kids = Array.from(par.children).filter(c => norm(c.textContent));
                if (kids.length === 2 && kids[0] === labelEl) return norm(kids[1].textContent);
            }
            return '';
        };
        Array.from(document.querySelectorAll('div, span, td, dt, li, p')).forEach(el => {
            if (el.children.length) return;                 // leaf only
            const t = norm(el.textContent).replace(/:$/, '');
            if (!LABELS.includes(t) || seen.has(t)) return;
            const v = valueFor(el);
            if (v && v.length < 400 && v !== t) { seen.add(t); res.fields.push([t, v]); }
        });

        // "Посмотреть полный профиль на этом сайте" — search whole page
        const prof = Array.from(document.querySelectorAll('a[href]')).find(a =>
            /полный профиль|full profile|profil complet/i.test(norm(a.textContent)));
        if (prof) res.profile = prof.href;

        // record photo — pick the LARGEST real image (the portrait), not the
        // first one (which was a 1×1 tracking pixel). Photos are lazy-loaded,
        // so use src / data-src / srcset and size by the displayed box.
        const SKIP = /avatar|icon|sprite|placeholder|logo|geni|brand|\.svg|badge|flag|blank|spacer|loading|pixel|1x1/i;
        const bestSrc = (im) => {
            const ss = im.getAttribute('srcset') || im.getAttribute('data-srcset') || '';
            if (ss) {
                const parts = ss.split(',').map(s => s.trim().split(' ')[0]).filter(Boolean);
                if (parts.length) return parts[parts.length - 1];
            }
            return im.getAttribute('src') || im.getAttribute('data-src')
                || im.getAttribute('data-original') || im.getAttribute('data-lazy') || '';
        };
        const candidates = [];
        for (const im of document.querySelectorAll('img')) {
            const s = bestSrc(im);
            if (!s || !/^https?:|^\/\//.test(s)) continue;
            if (SKIP.test(s) || SKIP.test(im.getAttribute('alt') || '')) continue;
            const r = im.getBoundingClientRect();
            const w = Math.max(r.width || 0, im.naturalWidth || 0,
                               parseInt(im.getAttribute('width') || '0', 10) || 0);
            const h = Math.max(r.height || 0, im.naturalHeight || 0,
                               parseInt(im.getAttribute('height') || '0', 10) || 0);
            const area = w * h;
            if (area < 50 * 50) continue;           // skip tiny pixels/icons
            candidates.push({src: s.startsWith('//') ? 'https:' + s : s, area});
        }
        // also CSS background-image photos (some cards use them)
        for (const el of document.querySelectorAll('[style*="background-image" i], [class*="photo" i], [class*="thumbnail" i]')) {
            const bg = (el.style && el.style.backgroundImage) || '';
            const m = bg.match(/url\((['"]?)(.*?)\1\)/i);
            if (!m) continue;
            let s = m[2];
            if (!s || SKIP.test(s)) continue;
            const r = el.getBoundingClientRect();
            const area = (r.width || 0) * (r.height || 0);
            if (area < 50 * 50) continue;
            candidates.push({src: s.startsWith('//') ? 'https:' + s : s, area});
        }
        candidates.sort((a, b) => b.area - a.area);
        res.photo = candidates.length ? candidates[0].src : '';

        // source line ("Семейные деревья MyHeritage", "Geni ...") as TEXT
        let src = Array.from(document.querySelectorAll('*')).find(e =>
            !e.children.length &&
            /^(Источник|Source|В категории|In category)/i.test(norm(e.textContent)));
        res.source = src ? norm(src.textContent) : '';

        res.historical = /историческ|historical/i.test(res.category || '');
        return res;
    }""")

    d["full_name"]  = (info.get("name") or "").strip()
    d["category"]   = (info.get("category") or "").strip()
    d["profile_url"] = info.get("profile") or ""
    d["source_text"] = (info.get("source") or "").strip()
    d["is_historical"] = bool(info.get("historical"))
    td = {}
    for pair in info.get("fields", []):
        if isinstance(pair, list) and len(pair) == 2:
            td[str(pair[0])] = str(pair[1])
    d["table_data"] = td

    # ── Photo: click the magnifier (лупа) to open the FULL-size photo ───────
    # The record card shows a small thumbnail with a zoom control
    # <div class="... main_record_image_zoom">. Clicking it opens a popup with
    # the large image. We download the FULL photo (saved to disk by the caller)
    # and embed a SMALL copy in Word.
    full_url = ""
    try:
        zoom = page.locator(
            '.main_record_image_zoom, .imageZoom, [class*="image_zoom" i], '
            '[class*="record_image" i] [class*="zoom" i]').first
        if await zoom.count():
            await zoom.scroll_into_view_if_needed(timeout=3000)
            await zoom.click(timeout=4000)
            await asyncio.sleep(1.5)
            full_url = await page.evaluate(r"""() => {
                // largest image inside the opened lightbox/modal
                const SKIP = /sprite|icon|logo|geni|brand|\.svg|avatar|badge/i;
                let best = '', area = 0;
                const scope = document.querySelector(
                    '[class*="modal" i], [role="dialog"], [class*="lightbox" i], '
                    + '[class*="overlay" i], [class*="popup" i]') || document;
                for (const im of scope.querySelectorAll('img')) {
                    const s = im.src || im.getAttribute('data-src') || '';
                    if (!s || SKIP.test(s)) continue;
                    const r = im.getBoundingClientRect();
                    const a = (r.width || im.naturalWidth || 0) *
                              (r.height || im.naturalHeight || 0);
                    if (a > area) { area = a; best = s; }
                }
                return best;
            }""")
            # close the popup
            try:
                closer = page.locator(
                    '[class*="modal" i] [class*="close" i], [role="dialog"] '
                    '[aria-label*="Close" i], [class*="lightbox" i] [class*="close" i]'
                ).first
                if await closer.count():
                    await closer.click(timeout=2000)
                else:
                    await page.keyboard.press("Escape")
            except Exception:
                await page.keyboard.press("Escape")
            await asyncio.sleep(0.4)
    except Exception:
        pass

    # download the photo (full from the zoom popup, else the card thumbnail)
    photo = full_url or info.get("photo") or ""
    if photo:
        try:
            r = await page.request.get(photo, timeout=15000)
            if r.ok:
                body = await r.body()
                if len(body) > 1000:
                    d["thumb_bytes"] = body          # bytes for Word (scaled) + disk
                    log(f"    📷 фото {len(body)//1024}KB"
                        f"{' (из лупы)' if full_url else ' (превью)'}")
                else:
                    log("    📷 фото слишком маленькое — пропуск")
            else:
                log(f"    !! фото HTTP {r.status}")
        except Exception as e:
            log(f"    !! фото не скачалось: {e}")
    else:
        log("    (фото на странице не найдено)")
    return d

# ── OUTPUT ───────────────────────────────────────────────────────────────── #
def _hyperlink(para, text, url):
    part = para.part
    rid  = part.relate_to(url, HYPERLINK_REL, is_external=True)
    hl   = OxmlElement("w:hyperlink"); hl.set(qn("r:id"), rid)
    run  = OxmlElement("w:r"); rPr = OxmlElement("w:rPr")
    c    = OxmlElement("w:color"); c.set(qn("w:val"), "0563C1")
    u    = OxmlElement("w:u");     u.set(qn("w:val"), "single")
    rPr.append(c); rPr.append(u); run.append(rPr)
    t    = OxmlElement("w:t"); t.text = text or url
    t.set(qn("xml:space"), "preserve")
    run.append(t); hl.append(run); para._p.append(hl)

def _docx_add_record(doc, i, rec):
    """Render ONE record into an open Document (shared by fresh + append)."""
    doc.add_heading(f"{i}. {rec.get('full_name','—')}", level=2)
    p = doc.add_paragraph()
    p.add_run("Category: ").bold = True
    p.add_run(rec.get("category", "—"))
    p2 = doc.add_paragraph()
    p2.add_run("Match: ").bold = True
    p2.add_run(f"{rec.get('score','?')}%")
    # Source as TEXT (e.g. "Семейные деревья MyHeritage / Geni") — never a logo
    if rec.get("source_text"):
        ps = doc.add_paragraph()
        ps.add_run("Источник: ").bold = True
        ps.add_run(rec["source_text"])
    if rec.get("url"):
        p3 = doc.add_paragraph()
        p3.add_run("Ссылка: ").bold = True
        _hyperlink(p3, rec["url"], rec["url"])
    td = rec.get("table_data", {})
    if td:
        tbl = doc.add_table(rows=1, cols=2)
        tbl.style = "Table Grid"
        hdr = tbl.rows[0].cells
        hdr[0].text = "Field"; hdr[1].text = "Value"
        for cell in hdr:
            for run in cell.paragraphs[0].runs:
                run.bold = True
        for f, v in td.items():
            row = tbl.add_row().cells
            row[0].text = str(f); row[1].text = str(v)
    # Photo thumbnail (small in Word; full-size saved to disk separately).
    # Convert WEBP→PNG so python-docx can embed it.
    tb = rec.get("thumb_bytes")
    if tb:
        png = _to_png(tb)
        if png:
            try:
                doc.add_picture(io.BytesIO(png), width=Inches(2.2))
            except Exception:
                pass
    # "View full profile on this site" link
    if rec.get("profile_url"):
        pp = doc.add_paragraph()
        pp.add_run("Полный профиль: ").bold = True
        _hyperlink(pp, "Посмотреть полный профиль на этом сайте",
                   rec["profile_url"])
    doc.add_paragraph("")


def write_docx(path, records, qlines, append=False):
    if not _DOCX_OK:
        raise RuntimeError("python-docx not installed")
    existing = append and Path(path).exists()
    if existing:
        # Open the existing document and append a clearly-marked new batch.
        doc = Document(str(path))
        doc.add_page_break()
        sep = doc.add_heading(
            f"➕ Appended {len(records)} more record(s) — "
            f"{time.strftime('%Y-%m-%d %H:%M')}", level=1)
        sep.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph("Search parameters:")
        for ln in qlines:
            doc.add_paragraph(ln, style="List Bullet")
        doc.add_paragraph("")
    else:
        doc = Document()
        sec = doc.sections[0]
        sec.page_width = Mm(297); sec.page_height = Mm(210)
        sec.left_margin = sec.right_margin = Mm(18)
        sec.top_margin  = sec.bottom_margin = Mm(15)
        h = doc.add_heading("MyHeritage Search Results", 0)
        h.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph("Search parameters:")
        for ln in qlines:
            doc.add_paragraph(ln, style="List Bullet")
        doc.add_paragraph(f"Records saved: {len(records)}  (match ≥ {MIN_MATCH_PCT}%)")
        doc.add_paragraph("")
    for i, rec in enumerate(records, 1):
        _docx_add_record(doc, i, rec)
    doc.save(path)

def write_xlsx(path, records, qlines, append=False):
    if not _OPENPYXL_OK:
        raise RuntimeError("openpyxl not installed")
    HF = PatternFill("solid", fgColor="2A4A7F")
    HN = Font(bold=True, color="FFFFFF", size=11)
    TS = Side(style="thin", color="B0B8C8")
    T  = Border(left=TS, right=TS, top=TS, bottom=TS)
    # Genealogy fields present across THESE records.
    aff = []
    for rec in records:
        for k in rec.get("table_data", {}):
            if k not in aff:
                aff.append(k)
    base_cols = ["#", "Full Name", "Category", "Match %", "URL"]

    existing = append and Path(path).exists()
    if existing:
        # Append rows to the existing sheet, continuing the # numbering and
        # adding columns for any genealogy fields not already present.
        wb = load_workbook(str(path))
        ws = wb.active
        header = [ws.cell(row=1, column=c).value
                  for c in range(1, ws.max_column + 1)]
        for name in base_cols + aff:
            if name not in header:
                header.append(name)
                c = ws.cell(row=1, column=len(header), value=name)
                c.font = HN; c.fill = HF; c.border = T
                c.alignment = Alignment(horizontal="center",
                                        vertical="center", wrap_text=True)
        col_idx = {name: i + 1 for i, name in enumerate(header)}
        last_n = 0
        for r in range(2, ws.max_row + 1):
            try:
                last_n = max(last_n, int(ws.cell(row=r, column=1).value))
            except Exception:
                pass
        start_row = ws.max_row + 1
        for off, rec in enumerate(records):
            td = rec.get("table_data", {})
            rowdata = {"#": last_n + off + 1,
                       "Full Name": rec.get("full_name", ""),
                       "Category":  rec.get("category", ""),
                       "Match %":   rec.get("score", ""),
                       "URL":       rec.get("url", "")}
            for f in aff:
                rowdata[f] = td.get(f, "")
            for name, val in rowdata.items():
                ci = col_idx.get(name)
                if ci:
                    c = ws.cell(row=start_row + off, column=ci, value=val)
                    c.border = T
                    c.alignment = Alignment(wrap_text=True, vertical="top")
        cols = header
    else:
        wb = Workbook(); ws = wb.active; ws.title = "MyHeritage"
        cols = base_cols + aff
        for ci, cn in enumerate(cols, 1):
            c = ws.cell(row=1, column=ci, value=cn)
            c.font = HN; c.fill = HF; c.border = T
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for ri, rec in enumerate(records, 2):
            td   = rec.get("table_data", {})
            vals = [ri-1, rec.get("full_name",""), rec.get("category",""),
                    rec.get("score",""), rec.get("url","")] + [td.get(f,"") for f in aff]
            for ci, val in enumerate(vals, 1):
                c = ws.cell(row=ri, column=ci, value=val)
                c.border = T
                c.alignment = Alignment(wrap_text=True, vertical="top")
    # Auto-size every column.
    for ci in range(1, len(cols) + 1):
        letter = get_column_letter(ci)
        mw = max(8, *(len(str(ws.cell(row=r, column=ci).value or ""))
                      for r in range(1, ws.max_row + 1)))
        ws.column_dimensions[letter].width = min(mw + 4, 60)
    wb.save(path)

# ── MAIN ENTRY POINT ─────────────────────────────────────────────────────── #
async def run_scraper(*,
    site_preset    = "Israel (.co.il)",
    first_name     = "", surname       = "",
    first_strict   = False, last_strict = False,
    birth_year     = "", birth_place   = "",
    father         = "", father_last   = "",
    mother         = "", mother_last   = "",
    spouse         = "", spouse_last   = "",
    death_year     = "",
    death_place    = "", residence     = "",
    military       = "", immigration   = "",
    keywords       = "", gender        = "Any",
    exact_match    = False,
    record_filter  = "All Records",
    record_type    = "All records",
    category       = "All collections",
    output_format  = "both",
    output_folder  = Path("."),
    email          = None, password    = None,
    log            = print,
    progress       = None,
    cancel_event   = None,
    ask_2fa_code   = None,   # callable() → str, fallback if IMAP auto-read fails
    imap_password  = None,   # mail password for MH account email (to auto-read 2FA code)
    family_site    = "",     # which family site to enter on select-site.php (name substring)
    ask_file_conflict = None, # callable(list[str]) → "overwrite"/"append"/"skip"
) -> dict:

    def _prog(pct, txt):
        log(txt)
        if progress:
            progress(pct, txt)

    def _done():
        return bool(cancel_event and cancel_event.is_set())

    login_url, search_url, has_cookies = SITE_PRESETS.get(
        site_preset, SITE_PRESETS["Israel (.co.il)"]
    )

    want_docx = output_format in ("docx", "both")
    want_xlsx = output_format in ("xlsx", "both")
    output_folder = Path(output_folder)
    output_folder.mkdir(parents=True, exist_ok=True)

    qname  = " ".join(p for p in (first_name, surname) if p)
    # Map the GUI record-type onto the internal record_filter values. Accept the
    # site-language label the GUI actually sends — Russian now; other languages
    # to be added tomorrow. (English aliases kept so nothing regresses.)
    _RT_MAP = {"Все записи": "All Records",
               "Исторические записи": "Historical Records",
               "Семейные деревья": "Family Trees",
               "All records": "All Records",
               "Historical records": "Historical Records",
               "Family trees": "Family Trees"}
    if record_type in _RT_MAP:
        record_filter = _RT_MAP[record_type]
    params = dict(
        first_name=first_name, surname=surname,
        first_strict=first_strict, last_strict=last_strict,
        birth_year=birth_year, birth_place=birth_place,
        father=father, father_last=father_last,
        mother=mother, mother_last=mother_last,
        spouse=spouse, spouse_last=spouse_last,
        death_year=death_year, death_place=death_place,
        residence=residence, military=military,
        immigration=immigration, keywords=keywords,
        gender=gender, exact_match=exact_match,
        record_filter=record_filter, category=category,
        # Site language (ru/en/he) — drives which language's UI text the advanced
        # search matches when locating pills/sections/buttons on the page.
        lang=_site_lang(site_preset),
    )
    qlines = [f"{k}: {v}" for k, v in params.items()
              if k != "lang" and v and v not in ("Any", False, "All Records")]
    summary = {"ok": False}

    _prog(0, "Launching browser…")

    async with async_playwright() as pw:
        # PERSISTENT profile — cookies + history live in MH_PROFILE_DIR between
        # runs. Once you've logged in (and passed the email 2FA) once, the next
        # runs reuse the session and MyHeritage's anti-bot stops flagging the
        # login as suspicious (no more "Confirm a login attempt" challenge).
        MH_PROFILE_DIR.mkdir(parents=True, exist_ok=True)
        ctx = await pw.chromium.launch_persistent_context(
            str(MH_PROFILE_DIR),
            headless=False,
            accept_downloads=True,
            no_viewport=True,
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
            args=["--start-maximized",
                  "--disable-blink-features=AutomationControlled",
                  "--disable-infobars",
                  "--disable-dev-shm-usage"],
        )
        # Hide automation fingerprint — same technique that lets
        # familysearch_scraper.py bypass bot detection
        await ctx.add_init_script("""
            Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
            Object.defineProperty(navigator, 'plugins',   {get: () => [1,2,3,4,5]});
            Object.defineProperty(navigator, 'languages', {get: () => ['en-US','en']});
            window.chrome = { runtime: {} };
            // ── Nuke Facebook entirely so it can never open a tab ──────────
            const _open = window.open;
            window.open = function(u, ...rest) {
                try { if (u && /facebook|fbcdn|accounts\\.google|apple\\.com/i.test(u)) return null; }
                catch (e) {}
                return _open ? _open.call(window, u, ...rest) : null;
            };
            const killFB = () => {
                try {
                    document.querySelectorAll(
                        'a[href*=\"facebook.com\"], a[href*=\"facebook.net\"], '
                        + 'iframe[src*=\"facebook\"], [class*=\"facebook\" i], '
                        + '[data-href*=\"facebook\"], .fb-page, .fb-like, .fb_iframe_widget'
                    ).forEach(e => { try { e.remove(); } catch (x) {} });
                } catch (x) {}
            };
            try { setInterval(killFB, 400); } catch (e) {}
            if (document.addEventListener) {
                document.addEventListener('DOMContentLoaded', killFB);
            }
        """)
        # Abort the Facebook SDK + plugin requests at the network level so the
        # social widget never loads and never spawns popup tabs.
        for _pat in ("**/*facebook.com/**", "**/*facebook.net/**",
                     "**/connect.facebook.net/**", "**/*fbcdn.net/**"):
            try:
                await ctx.route(_pat, lambda r: r.abort())
            except Exception:
                pass

        # Auto-close any JUNK tab the instant it opens (Facebook/Google/blank)
        # so it never steals focus. MyHeritage result tabs (myheritage.com) are
        # kept.
        def _on_new_page(p):
            async def _maybe_close():
                try:
                    # Poll the tab's URL for up to 3s: close the instant it is
                    # (or becomes) Facebook/blank-junk; keep MyHeritage tabs.
                    for _ in range(10):
                        # NEVER close a tab the scraper opened on purpose (the
                        # 2FA mail-reading tab) — guarded by a context flag and a
                        # per-page mark set in _browser_read_yandex_code.
                        if (getattr(ctx, "_mh_pause_autoclose", False)
                                or getattr(p, "_mh_protected", False)):
                            return
                        u = (p.url or "").lower()
                        # Keep mail / login-passport tabs used to read the code.
                        if any(s in u for s in ("yandex", "passport", "mail.",
                                                "/mail", "outlook.", "office365")):
                            return
                        if any(s in u for s in ("facebook", "fbcdn",
                                                "accounts.google", "apple.com")):
                            await p.close()
                            return
                        if "myheritage" in u:
                            return                      # real results tab — keep
                        await asyncio.sleep(0.3)
                    # Still blank after 3s and not MyHeritage → junk, close it
                    if ("myheritage" not in (p.url or "").lower()
                            and not getattr(ctx, "_mh_pause_autoclose", False)
                            and not getattr(p, "_mh_protected", False)):
                        await p.close()
                except Exception:
                    pass
            try:
                asyncio.create_task(_maybe_close())
            except Exception:
                pass
        ctx.on("page", _on_new_page)

        # The persistent profile RESTORES whatever tabs were open last time
        # (e.g. old Facebook/blank tabs). Keep ONE page, close the rest.
        page = ctx.pages[0] if ctx.pages else await ctx.new_page()
        for extra in list(ctx.pages):
            if extra is not page:
                try:
                    await extra.close()
                except Exception:
                    pass

        try:
            if email and password:
                _prog(5, "Logging in…")
                logged_in = await _login(
                    page, login_url, has_cookies,
                    email, password, log,
                    ask_2fa_code=ask_2fa_code,
                    imap_password=imap_password,
                )
                if not logged_in:
                    summary["error"]   = "login_failed"
                    summary["message"] = "Login failed — check credentials or 2FA code."
                    return summary
            else:
                log("  No credentials — guest mode.")
                try:
                    await page.goto(login_url, wait_until="domcontentloaded", timeout=20000)
                    if has_cookies:
                        await _accept_cookies(page, log)
                except Exception:
                    pass

            if _done():
                return summary

            # If MyHeritage shows the family-site chooser, pick a site and get
            # the research URL for it (overrides the generic search_url).
            site_research_url = await _handle_select_site(page, family_site, log)
            if site_research_url:
                search_url = site_research_url

            _prog(15, "Opening search page…")
            results_page = await _search(page, search_url, params, has_cookies, log)
            if not results_page:
                summary["error"]   = "search_failed"
                summary["message"] = "Could not reach search results page."
                return summary
            # All result collection happens on the (possibly new) results tab
            page = results_page

            if _done():
                return summary

            _prog(30, "Collecting results…")
            raw = await _collect(page, log)
            log(f"  Candidates: {len(raw)}")

            # Fallback: exactSearch=1 narrows on ALL fields at once, so a
            # many-field query (father + spouse + death year …) can legitimately
            # return zero exact matches. Rather than report "nothing found",
            # drop back to the plain results (still bounded by the «expanded
            # criteria» divider) so the user always gets the relevant records.
            if (params.get("exact_match") and params.get("_exact_ok")
                    and not raw and params.get("_noexact_url")):
                log("  !! Точный поиск (exactSearch=1) дал 0 записей — "
                    "слишком строго; откатываюсь к обычным результатам")
                try:
                    await page.goto(params["_noexact_url"],
                                    wait_until="domcontentloaded", timeout=30000)
                    await page.wait_for_load_state("networkidle", timeout=15000)
                    params["_exact_ok"] = False
                    # Only the first page — MyHeritage sorts by relevance, so the
                    # top results are the right person; deeper pages are noise.
                    raw = await _collect(page, log, max_pages=1)
                    log(f"  Candidates (без exactSearch, 1 стр.): {len(raw)}")
                except Exception as _e:
                    log(f"  !! откат к обычным результатам не вышел: {_e}")

            # Results collected are MyHeritage's matches (everything BEFORE the
            # «expanded criteria» divider). Keep them all: use the card's match
            # % when shown, otherwise treat as an exact MyHeritage match (100%).
            # Do NOT re-filter by name similarity — that wrongly drops correct
            # multi-word names (e.g. "Тамара Хананновна Рогинская (Рубина)").
            qualified = []
            for r in raw:
                s = r["score"]
                if s < 0:
                    s = 100.0
                if s >= MIN_MATCH_PCT:
                    r["score"] = round(s, 1)
                    qualified.append(r)
            log(f"  Подходящих записей: {len(qualified)}")

            # Safety: if exact match was requested but couldn't be enabled, the
            # site returns a huge fuzzy set — don't grind through 500 records.
            if (params.get("exact_match") and not params.get("_exact_ok")
                    and len(qualified) > 40):
                log(f"  !! «Точное совпадение» не включилось — обрабатываю первые "
                    f"40 из {len(qualified)} (самые релевантные сверху), "
                    f"чтобы не ждать часами.")
                qualified = qualified[:40]

            if not qualified:
                _prog(100, "No results above threshold.")
                summary.update({"ok": True, "docx_count": 0, "xlsx_path": None,
                                "n_records": 0,
                                "message": f"No records with match ≥{MIN_MATCH_PCT}%."})
                return summary

            # Folder for the full-size photos
            images_dir = output_folder / "images" / (safe_fn(qname) or "myheritage")

            records = []
            n = len(qualified)
            for i, r in enumerate(qualified, 1):
                if _done():
                    break
                _prog(35 + int(50 * i / n), f"[{i}/{n}] Reading record…")
                dp = None
                try:
                    dp = await ctx.new_page()
                    det = await _detail(dp, r["url"], has_cookies, log)
                    det["score"] = r["score"]
                    if not det.get("full_name"):
                        det["full_name"] = r["name_text"]
                    # Save the FULL photo to disk (Word keeps a small copy)
                    if det.get("thumb_bytes"):
                        try:
                            images_dir.mkdir(parents=True, exist_ok=True)
                            fn = safe_fn(det.get("full_name") or f"record_{i}") + ".jpg"
                            (images_dir / fn).write_bytes(det["thumb_bytes"])
                        except Exception:
                            pass
                    records.append(det)
                    log(f"    ✓ {det['full_name']} — {det['score']}%")
                except Exception as _exc:
                    log(f"    !! запись пропущена: {_exc}")
                    if "closed" in str(_exc).lower():
                        break          # context/browser gone — stop gracefully
                finally:
                    if dp is not None:
                        try:
                            await dp.close()
                        except Exception:
                            pass
                await asyncio.sleep(0.7)

            _prog(88, "Saving files…")
            base   = safe_fn(f"myheritage_{qname}") or "myheritage_results"
            docx_p = output_folder / f"{base}.docx"
            xlsx_p = output_folder / f"{base}.xlsx"

            # If output files already exist, ask the user what to do: overwrite,
            # append the new records, or skip saving. Default (no callback / no
            # conflict) is a plain overwrite — the previous behaviour.
            existing_names = [p.name for p, want in
                              ((docx_p, want_docx), (xlsx_p, want_xlsx))
                              if want and records and p.exists()]
            decision = "overwrite"
            if existing_names and ask_file_conflict:
                try:
                    decision = (ask_file_conflict(existing_names)
                                or "overwrite").lower()
                except Exception as _e:
                    log(f"  !! file-conflict dialog error: {_e}")
                    decision = "overwrite"
                log(f"  → Файл(ы) уже существуют {existing_names} → выбор: {decision}")
            append = (decision == "append")

            sd = sx = False
            if decision == "skip":
                log("  → Сохранение пропущено по выбору пользователя "
                    "(существующие файлы не тронуты).")
            else:
                if want_docx and records:
                    write_docx(docx_p, records, qlines, append=append)
                    sd = True
                    log(f"  → Word: {docx_p.name}"
                        f"{' (дополнен)' if append and docx_p.name in existing_names else ''}")
                if want_xlsx and records:
                    write_xlsx(xlsx_p, records, qlines, append=append)
                    sx = True
                    log(f"  → Excel: {xlsx_p.name}"
                        f"{' (дополнен)' if append and xlsx_p.name in existing_names else ''}")

            _prog(100, f"Done — {len(records)} record(s).")
            summary.update({
                "ok":            True,
                "docx_count":    1 if sd else 0,
                "xlsx_path":     str(xlsx_p) if sx else None,
                "output_folder": str(output_folder),
                "n_records":     len(records),
            })

        except Exception as exc:
            summary.update({"error": "exception",
                            "message": f"{type(exc).__name__}: {exc}"})
            log(f"  !! {exc}")
        finally:
            # Persistent context — closing it saves cookies/session for next run
            try:
                await ctx.close()
            except Exception:
                pass

    return summary
