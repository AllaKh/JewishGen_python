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

IMPORTANT: NO persistent browser profile is used — fresh context every run
so there are zero stale cookies from any previous session.
"""

import asyncio, difflib, imaplib, email as _email_lib
import os, re, sys, time
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
    from docx.shared import Mm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    _DOCX_OK = True
except ImportError:
    _DOCX_OK = False

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter
    _OPENPYXL_OK = True
except ImportError:
    _OPENPYXL_OK = False

# ── Constants ─────────────────────────────────────────────────────────────── #
MIN_MATCH_PCT = 80
HYPERLINK_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"

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

    page = await ctx.new_page()
    try:
        await page.bring_to_front()
        log("  2FA: открываю новую вкладку mail.yandex.ru …")
        await page.goto("https://mail.yandex.ru/", wait_until="domcontentloaded",
                        timeout=30000)
        await asyncio.sleep(2)

        # 1) Click "Войти" (header login button) → goes to passport.yandex
        await _click_any(['#header-login-button',
                          'a:has-text("Войти")', 'a:has-text("Log in")'],
                         "нажал «Войти»")
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

        # 6) Wait for inbox to load
        try:
            await page.wait_for_url(
                lambda u: "mail.yandex" in u and "passport" not in u,
                timeout=25000)
        except Exception:
            await asyncio.sleep(5)
        log("  2FA: вошёл в почту, ищу письмо с кодом …")

        # 7) Find the MyHeritage email (NOT the ad at the very top!) and
        #    read the 6-digit code. The code is visible right in the snippet
        #    preview: "...confirmation code in the login screen: 284383 ...".
        deadline = asyncio.get_event_loop().time() + timeout
        while asyncio.get_event_loop().time() < deadline:
            # (a) Scan the whole inbox text for the MyHeritage code phrase.
            #     This skips the ad row entirely and needs no clicking.
            try:
                code = await page.evaluate(r"""() => {
                    const body = document.body.innerText || '';
                    const pats = [
                        /confirmation code in the login screen[:\s]*?(\d{6})/i,
                        /confirmation code[^\d]{0,40}(\d{6})/i,
                        /code in the login screen[:\s]*?(\d{6})/i,
                        /код[^\d]{0,40}(\d{6})/i,
                    ];
                    for (const p of pats) {
                        const m = body.match(p);
                        if (m) return m[1];
                    }
                    return '';
                }""")
                if code:
                    log(f"  2FA: код найден в превью письма MyHeritage: {code}")
                    return code
            except Exception:
                pass

            # (b) Snippet truncated before the code → click the MyHeritage row
            #     (the one whose text mentions MyHeritage + login attempt),
            #     never the topmost ad row.
            try:
                clicked = await page.evaluate(r"""() => {
                    const rows = Array.from(document.querySelectorAll(
                        'a, li, div[role="listitem"], [class*="messageSnippet" i], '
                        + '[class*="MessageSnippet"]'));
                    // smallest element that mentions MyHeritage + login/confirm
                    const cands = rows.filter(el => {
                        const t = (el.textContent || '');
                        return /myheritage/i.test(t) &&
                               /(login attempt|confirm a login|confirmation code|подтвержд|код)/i.test(t) &&
                               t.length < 600;
                    });
                    cands.sort((a, b) => a.textContent.length - b.textContent.length);
                    if (cands.length) {
                        const el = cands[0].closest('a') || cands[0];
                        el.scrollIntoView();
                        el.click();
                        return true;
                    }
                    return false;
                }""")
                if clicked:
                    log("  2FA: открыл письмо MyHeritage")
                    await asyncio.sleep(2.5)
                    body = await page.evaluate("() => document.body.innerText")
                    m = (re.search(r"login screen[:\s]*?(\d{6})", body or "", re.I)
                         or re.search(r"\b(\d{6})\b", body or ""))
                    if m:
                        log(f"  2FA: код из открытого письма: {m.group(1)}")
                        return m.group(1)
            except Exception:
                pass

            # Wait a bit for the email to arrive, then re-check (no full reload)
            await asyncio.sleep(3)

        log("  2FA: код в почте не найден за отведённое время")
        return None
    except Exception as e:
        log(f"  2FA browser error: {e}")
        return None
    finally:
        # Close the mail tab regardless of outcome
        try:
            await page.close()
        except Exception:
            pass


async def _get_2fa_code(ctx, mail_email: str, mail_password: str,
                        log, ask_2fa_code=None) -> str | None:
    """
    Get the MyHeritage 2FA code:
      1. Browser: open a new tab to Yandex mail, log in, read the code.
      2. IMAP fallback (if browser fails) for non-Yandex providers.
      3. GUI dialog as a last resort.
    `mail_email` is the SAME email used for MyHeritage login.
    """
    if mail_email and mail_password:
        # Primary: browser-based Yandex reading (exact flow Alla described)
        code = await _browser_read_yandex_code(ctx, mail_email, mail_password, log)
        if code:
            return code
        # Secondary: IMAP (works for non-Yandex providers with IMAP enabled)
        log("  2FA: пробую IMAP как запасной вариант …")
        code = await asyncio.to_thread(
            _imap_read_mh_code, mail_email, mail_password, 60)
        if code:
            log("  2FA: код получен по IMAP ✓")
            return code

    # Last resort: ask the user
    if ask_2fa_code:
        log("  2FA: запрашиваю код у пользователя …")
        return ask_2fa_code()
    return None


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
        # Get the code: open a new tab to the mail, read it (same email as MH)
        code = await _get_2fa_code(page.context, email, imap_password,
                                   log, ask_2fa_code)
        if not code:
            log("  !! 2FA: код не получен — прерываю.")
            return False
        log(f"  → Ввожу 2FA код: {code}")
        # CRITICAL: the mail tab was brought to front; bring the MH tab back
        # so keyboard input actually lands on this page.
        try:
            await page.bring_to_front()
            await asyncio.sleep(0.4)
        except Exception:
            pass

        async def _code_entered() -> bool:
            """True if the visible numeric inputs together hold the full code."""
            try:
                digits = await page.evaluate("""() => {
                    const els = document.querySelectorAll(
                        'input[inputmode="numeric"], input[autocomplete="one-time-code"], '
                        + 'input[maxlength="6"], input[name*="code"], input[type="number"]');
                    let s = '';
                    els.forEach(e => { if (e.offsetParent !== null) s += (e.value || ''); });
                    return s.replace(/\\D/g, '');
                }""")
                return code in (digits or "")
            except Exception:
                return False

        # Strategy 1: focus the single field and type the whole code
        for attempt in range(3):
            try:
                fld = page.locator(tfa_sel).first
                await fld.click(timeout=4000)
                await page.keyboard.press("Control+a")
                await page.keyboard.press("Delete")
                await page.keyboard.type(code, delay=140)
                await asyncio.sleep(0.6)
            except Exception:
                pass
            if await _code_entered():
                break
            # Strategy 2: segmented field — type one digit per visible box
            try:
                boxes = page.locator(
                    'input[inputmode="numeric"], input[autocomplete="one-time-code"]')
                nboxes = await boxes.count()
                if nboxes >= len(code):
                    for i, ch in enumerate(code):
                        b = boxes.nth(i)
                        await b.click(timeout=2000)
                        await b.fill(ch)
                        await asyncio.sleep(0.1)
                    await asyncio.sleep(0.5)
            except Exception:
                pass
            if await _code_entered():
                break
            await asyncio.sleep(0.5)

        if await _code_entered():
            log("  ✓ 2FA: код введён в поле")
        else:
            log("  !! 2FA: код не подтверждён в поле — всё равно пробую отправить")
        await asyncio.sleep(0.5)

        # Submit. The confirm button is "Отправить". Try MANY ways and after
        # each one check whether we left the code screen — keep going until
        # one of them actually submits.
        async def _submit_done() -> bool:
            await asyncio.sleep(2)
            if await _logged_in():
                return True
            # also "done" if the code field disappeared (moved to next step)
            return (await _find_tfa()) is None

        # 1) Enter key in the focused code field (most reliable for OTP forms)
        try:
            await page.keyboard.press("Enter")
            log("  2FA: Enter в поле кода")
            if await _submit_done():
                log(f"  ✓ Logged in after 2FA. URL: {page.url}")
                return True
        except Exception:
            pass

        # 2) Click the real <button> ancestor of the "Отправить" span
        for getter in (
            lambda: page.get_by_role("button", name=re.compile(
                r"^\s*Отправить\s*$|^\s*Send\s*$|^\s*Continue\s*$", re.I)).first,
            lambda: page.locator(
                'button:has(span.button_content:has-text("Отправить"))').first,
            lambda: page.locator('button:has-text("Отправить")').last,
        ):
            try:
                el = getter()
                if await el.count() and await el.is_visible():
                    await el.click(timeout=5000)
                    log("  ✓ 2FA: клик по кнопке «Отправить»")
                    if await _submit_done():
                        log(f"  ✓ Logged in after 2FA. URL: {page.url}")
                        return True
            except Exception:
                continue

        # 3) JS: locate the exact "Отправить" button, climb to <button>, click
        try:
            ok = await page.evaluate(r"""() => {
                const norm = s => (s || '').replace(/\s+/g, ' ').trim();
                // prefer a span.button_content whose text is exactly Отправить
                let span = Array.from(document.querySelectorAll('span.button_content, span'))
                    .find(s => /^Отправить$/i.test(norm(s.textContent)));
                let btn = span ? (span.closest('button') || span.parentElement) : null;
                if (!btn) {
                    btn = Array.from(document.querySelectorAll('button,[role=button]'))
                        .find(b => /^Отправить$/i.test(norm(b.textContent)));
                }
                if (btn) {
                    btn.scrollIntoView();
                    btn.click();
                    return true;
                }
                return false;
            }""")
            if ok:
                log("  ✓ 2FA: клик по «Отправить» (JS)")
                if await _submit_done():
                    log(f"  ✓ Logged in after 2FA. URL: {page.url}")
                    return True
        except Exception:
            pass

        # Final wait for login to complete
        for _ in range(20):
            await asyncio.sleep(1)
            if await _logged_in():
                log(f"  ✓ Logged in after 2FA. URL: {page.url}")
                return True

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
    After login MyHeritage may show FP/select-site.php asking which family
    site to enter. Click the chosen site (by name, default = first / admin).
    Returns the research base URL (with the site id) if known.
    """
    if "select-site" not in page.url.lower():
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


# ── SEARCH FORM ───────────────────────────────────────────────────────────── #
async def _search(page, search_url, params, has_cookies, log):
    log(f"  → Navigating to search: {search_url}")
    try:
        await page.goto(search_url, wait_until="domcontentloaded", timeout=35000)
    except Exception as exc:
        log(f"  !! Cannot open search page: {exc}")
        return False
    if has_cookies:
        await _accept_cookies(page, log)
    await asyncio.sleep(1.5)

    # Exact field selectors from the live research form (data-automations)
    await _fill(page, [
        'input[data-automations="research-family_first_name"]',
        'input[placeholder*="Имя" i]', 'input[placeholder*="first" i]',
    ], params.get("first_name", ""), "First name", log)
    await _fill(page, [
        'input[data-automations="research-family_last_name"]',
        'input[placeholder*="Фамилия" i]', 'input[placeholder*="last" i]',
    ], params.get("surname", ""), "Surname", log)

    # Birth year
    await _fill(page, [
        'input[placeholder*="Год рождения" i]', 'input[placeholder*="birth year" i]',
        'input[name*="birthYear" i]', 'input[id*="birthYear" i]',
    ], params.get("birth_year", ""), "Birth year", log)

    # Birth place
    await _fill(page, [
        'input[placeholder*="Населенный пункт" i]', 'input[placeholder*="birth place" i]',
        'input[name*="birthPlace" i]', 'input[id*="birthPlace" i]',
    ], params.get("birth_place", ""), "Birth place", log)

    # Extended fields via pill buttons
    async def pill(ru, en, he=""):
        for txt in [ru, en, he]:
            if not txt:
                continue
            try:
                el = page.get_by_text(re.compile(rf"^{re.escape(txt)}$", re.I)).first
                if await el.count():
                    await el.click(timeout=4000)
                    await asyncio.sleep(0.6)
                    return True
            except Exception:
                pass
        return False

    if params.get("father"):
        await pill("Отец", "Father", "אב")
        await _fill(page, ['input[name*="father" i]', 'input[placeholder*="father" i]',
                           'input[placeholder*="Отец" i]'],
                   params["father"], "Father", log)

    if params.get("mother"):
        await pill("Мать", "Mother", "אם")
        await _fill(page, ['input[name*="mother" i]', 'input[placeholder*="mother" i]',
                           'input[placeholder*="Мать" i]'],
                   params["mother"], "Mother", log)

    if params.get("spouse"):
        await pill("Супруг(-а)", "Spouse", "בן/בת זוג")
        await _fill(page, ['input[name*="spouse" i]', 'input[placeholder*="spouse" i]',
                           'input[placeholder*="Супруг" i]'],
                   params["spouse"], "Spouse", log)

    if params.get("death_year") or params.get("death_place"):
        await pill("Смерть", "Death", "פטירה")
        if params.get("death_year"):
            await _fill(page, ['input[name*="deathYear" i]'],
                       params["death_year"], "Death year", log)
        if params.get("death_place"):
            await _fill(page, ['input[name*="deathPlace" i]'],
                       params["death_place"], "Death place", log)

    # More panel
    for txt in ["+ Больше", "+ More", "+ עוד"]:
        try:
            el = page.get_by_text(re.compile(re.escape(txt), re.I)).first
            if await el.count():
                await el.click(timeout=4000)
                await asyncio.sleep(0.8)
                break
        except Exception:
            pass

    for field, labels in [
        ("residence",   ["Местожительство", "Residence",  "מגורים"]),
        ("military",    ["Вооруженные силы", "Military",  "צבא"]),
        ("immigration", ["Иммиграция",       "Immigration","הגירה"]),
        ("keywords",    ["Ключевые слова",   "Keywords",  "מילות מפתח"]),
    ]:
        if params.get(field):
            await pill(*labels)
            await _fill(page,
                       [f'input[name*="{field}" i]', f'input[placeholder*="{labels[0]}" i]'],
                       params[field], labels[0], log)

    if params.get("gender", "Any") not in ("Any", "Любой", "כל"):
        await pill("Пол", "Gender", "מין")
        try:
            g = page.get_by_text(re.compile(re.escape(params["gender"]), re.I)).first
            if await g.count():
                await g.click(timeout=3000)
        except Exception:
            pass

    if params.get("exact_match"):
        try:
            cb = page.get_by_text(
                re.compile(r"Точное совпадение|Exact match|התאמה מדויקת", re.I)
            ).first
            if await cb.count():
                await cb.click(timeout=3000)
        except Exception:
            pass

    # Record type filter
    rf = params.get("record_filter", "All Records")
    FILTER_MAP = {
        "Historical Records": ["Исторические записи", "Historical records",
                               "Historical Records", "רשומות היסטוריות"],
        "Family Trees":       ["Семейные деревья",    "Family trees",
                               "Family Trees",        "עצי משפחה"],
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
    submit_sels = ['span.button_content:has-text("Поиск")',
                   'button:has-text("Поиск")', 'button:has-text("Search")',
                   'button:has-text("חיפוש")', 'button[type="submit"]',
                   'input[type="submit"]']
    try:
        async with ctx.expect_page(timeout=15000) as new_page_info:
            for sel in submit_sels:
                try:
                    el = page.locator(sel).first
                    if await el.count():
                        await el.click(timeout=6000)
                        break
                except Exception:
                    pass
        results_page = await new_page_info.value
        await results_page.wait_for_load_state("domcontentloaded", timeout=30000)
        log("  ✓ Результаты открылись в новом табе")
    except Exception:
        # No new tab — results loaded in the same page
        results_page = page
        try:
            await page.wait_for_load_state("networkidle", timeout=30000)
        except Exception:
            await asyncio.sleep(4)

    try:
        await results_page.wait_for_load_state("networkidle", timeout=20000)
    except Exception:
        await asyncio.sleep(3)
    return results_page

# ── COLLECT RESULTS ───────────────────────────────────────────────────────── #
async def _collect(page, log):
    await asyncio.sleep(2)
    links = []
    for sel in ["a.results_result_link", 'a[class*="result" i]',
                ".result-item a", ".search-result a",
                'li[class*="result" i] a', '[class*="ResultItem" i] a']:
        try:
            els = await page.query_selector_all(sel)
            if els:
                for el in els:
                    href  = await el.get_attribute("href") or ""
                    text  = (await el.text_content() or "").strip()
                    score = -1.0
                    m = re.search(r"(\d{1,3})\s*%", text)
                    if m:
                        score = float(m.group(1))
                    if href.startswith("http"):
                        links.append({"url": href, "name_text": text, "score": score})
                if links:
                    break
        except Exception:
            pass
    if not links:
        try:
            for el in await page.query_selector_all("a[href]"):
                href = await el.get_attribute("href") or ""
                text = (await el.text_content() or "").strip()
                if any(kw in href for kw in ["/record/", "/person/", "/family-site/"]):
                    if href not in [r["url"] for r in links]:
                        links.append({"url": href, "name_text": text, "score": -1.0})
        except Exception:
            pass
    return links

# ── DETAIL PAGE ───────────────────────────────────────────────────────────── #
async def _detail(page, url, has_cookies, log):
    d = {"url": url, "full_name": "", "category": "", "table_data": {}}
    try:
        await page.goto(url, wait_until="domcontentloaded", timeout=35000)
        if has_cookies:
            await _accept_cookies(page, log)
        await asyncio.sleep(1.5)
    except Exception as exc:
        log(f"    !! {exc}")
        return d
    for sel in ["h1.person-name", "h1", ".record-title", ".full-name",
                '[class*="name" i]', '[class*="title" i]']:
        try:
            name = (await page.locator(sel).first.text_content(timeout=3000) or "").strip()
            if name and len(name) < 200:
                d["full_name"] = name
                break
        except Exception:
            pass
    for sel in [".record-type", ".category", ".collection-name",
                '[class*="category" i]', '[class*="collection" i]']:
        try:
            cat = (await page.locator(sel).first.text_content(timeout=3000) or "").strip()
            if cat and len(cat) < 300:
                d["category"] = cat
                break
        except Exception:
            pass
    td = {}
    try:
        for row in await page.query_selector_all("table tr"):
            cells = await row.query_selector_all("td,th")
            if len(cells) >= 2:
                k = (await cells[0].text_content() or "").strip().rstrip(":")
                v = (await cells[1].text_content() or "").strip()
                if k and v:
                    td[k] = v
    except Exception:
        pass
    if not td:
        try:
            dts = await page.query_selector_all("dt")
            dds = await page.query_selector_all("dd")
            for dt, dd in zip(dts, dds):
                k = (await dt.text_content() or "").strip().rstrip(":")
                v = (await dd.text_content() or "").strip()
                if k and v:
                    td[k] = v
        except Exception:
            pass
    d["table_data"] = td
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

def write_docx(path, records, qlines):
    if not _DOCX_OK:
        raise RuntimeError("python-docx not installed")
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
        doc.add_heading(f"{i}. {rec.get('full_name','—')}", level=2)
        p = doc.add_paragraph()
        p.add_run("Category: ").bold = True
        p.add_run(rec.get("category", "—"))
        p2 = doc.add_paragraph()
        p2.add_run("Match: ").bold = True
        p2.add_run(f"{rec.get('score','?')}%")
        if rec.get("url"):
            p3 = doc.add_paragraph()
            p3.add_run("Source: ").bold = True
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
        doc.add_paragraph("")
    doc.save(path)

def write_xlsx(path, records, qlines):
    if not _OPENPYXL_OK:
        raise RuntimeError("openpyxl not installed")
    wb = Workbook(); ws = wb.active; ws.title = "MyHeritage"
    HF = PatternFill("solid", fgColor="2A4A7F")
    HN = Font(bold=True, color="FFFFFF", size=11)
    TS = Side(style="thin", color="B0B8C8")
    T  = Border(left=TS, right=TS, top=TS, bottom=TS)
    aff = []
    for rec in records:
        for k in rec.get("table_data", {}):
            if k not in aff:
                aff.append(k)
    cols = ["#", "Full Name", "Category", "Match %", "URL"] + aff
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
    for ci, cn in enumerate(cols, 1):
        letter = get_column_letter(ci)
        mw = max(len(str(cn)),
                 *(len(str(ws.cell(row=r, column=ci).value or ""))
                   for r in range(2, ws.max_row+1)), 8)
        ws.column_dimensions[letter].width = min(mw+4, 60)
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
    record_type    = "Все записи",
    category       = "Все коллекции",
    output_format  = "both",
    output_folder  = Path("."),
    email          = None, password    = None,
    log            = print,
    progress       = None,
    cancel_event   = None,
    ask_2fa_code   = None,   # callable() → str, fallback if IMAP auto-read fails
    imap_password  = None,   # mail password for MH account email (to auto-read 2FA code)
    family_site    = "",     # which family site to enter on select-site.php (name substring)
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
    # Map the GUI record-type (Russian) onto the internal record_filter values
    _RT_MAP = {"Исторические записи": "Historical Records",
               "Семейные деревья": "Family Trees", "Все записи": "All Records"}
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
    )
    qlines = [f"{k}: {v}" for k, v in params.items()
              if v and v not in ("Any", False, "All Records")]
    summary = {"ok": False}

    _prog(0, "Launching browser…")

    async with async_playwright() as pw:
        # FRESH context every run — no stale cookies from any previous session
        browser = await pw.chromium.launch(
            headless=False,
            args=["--start-maximized",
                  "--disable-blink-features=AutomationControlled",
                  "--disable-infobars",
                  "--disable-dev-shm-usage"],
        )
        ctx = await browser.new_context(
            no_viewport=True,
            accept_downloads=True,
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
        )
        # Hide automation fingerprint — same technique that lets
        # familysearch_scraper.py bypass bot detection
        await ctx.add_init_script("""
            Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
            Object.defineProperty(navigator, 'plugins',   {get: () => [1,2,3,4,5]});
            Object.defineProperty(navigator, 'languages', {get: () => ['en-US','en']});
            window.chrome = { runtime: {} };
        """)
        page = await ctx.new_page()

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

            qualified = []
            for r in raw:
                s = r["score"]
                if s < 0:
                    s = _name_sim(qname, r["name_text"])
                if s >= MIN_MATCH_PCT:
                    r["score"] = round(s, 1)
                    qualified.append(r)
            log(f"  Qualified (≥{MIN_MATCH_PCT}%): {len(qualified)}")

            if not qualified:
                _prog(100, "No results above threshold.")
                summary.update({"ok": True, "docx_count": 0, "xlsx_path": None,
                                "n_records": 0,
                                "message": f"No records with match ≥{MIN_MATCH_PCT}%."})
                return summary

            records = []
            n = len(qualified)
            for i, r in enumerate(qualified, 1):
                if _done():
                    break
                _prog(35 + int(50 * i / n), f"[{i}/{n}] Reading record…")
                dp = await ctx.new_page()
                try:
                    det = await _detail(dp, r["url"], has_cookies, log)
                    det["score"] = r["score"]
                    if not det.get("full_name"):
                        det["full_name"] = r["name_text"]
                    records.append(det)
                    log(f"    ✓ {det['full_name']} — {det['score']}%")
                finally:
                    await dp.close()
                await asyncio.sleep(0.7)

            _prog(88, "Saving files…")
            base   = safe_fn(f"myheritage_{qname}") or "myheritage_results"
            docx_p = output_folder / f"{base}.docx"
            xlsx_p = output_folder / f"{base}.xlsx"
            sd = sx = False
            if want_docx and records:
                write_docx(docx_p, records, qlines)
                sd = True
                log(f"  → Word: {docx_p.name}")
            if want_xlsx and records:
                write_xlsx(xlsx_p, records, qlines)
                sx = True
                log(f"  → Excel: {xlsx_p.name}")

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
            try:
                await ctx.close()
            except Exception:
                pass
            try:
                await browser.close()
            except Exception:
                pass

    return summary
