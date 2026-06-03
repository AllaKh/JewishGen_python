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
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter
    _OPENPYXL_OK = True
except ImportError:
    _OPENPYXL_OK = False

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
        log("  2FA: вошёл в почту, ищу письма с кодами …")

        # MyHeritage sends TWO emails with DIFFERENT codes:
        #   B) "Ваш код подтверждения для входа в MyHeritage" / "verification
        #      code" — this is the code the «Введите верификационный код» field
        #      actually wants (priority).
        #   A) "Confirm a login attempt on MyHeritage" — the anti-fraud
        #      challenge ("flagged as suspicious"); its code is the fallback.
        # We read BOTH from the inbox snippets and return them in priority
        # order so the caller can try B first, then A.
        async def _scan_codes():
            return await page.evaluate(r"""() => {
                const rows = Array.from(document.querySelectorAll(
                    'a, li, div[role="listitem"], [class*="MessageSnippet"], '
                    + '[class*="messageSnippet" i]'));
                let verify = '', attempt = '';
                for (const el of rows) {
                    const t = (el.textContent || '');
                    if (!/myheritage/i.test(t)) continue;
                    const m = t.match(/\b(\d{6})\b/);
                    if (!m) continue;
                    const code = m[1];
                    if (!verify && /(код подтверждения для входа|верификацион|verification code)/i.test(t))
                        verify = code;
                    else if (!attempt && /(confirm a login attempt|confirmation code in the login screen|login attempt|попытк)/i.test(t))
                        attempt = code;
                }
                return {verify, attempt};
            }""")

        deadline = asyncio.get_event_loop().time() + timeout
        last = {"verify": "", "attempt": ""}
        while asyncio.get_event_loop().time() < deadline:
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
        # Close the mail tab regardless of outcome
        try:
            await page.close()
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
    btn_sels = ['button:has(span.button_content:has-text("Поиск"))',
                'button:has-text("Поиск")', 'button:has-text("Search")']
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
    tagged = await root.evaluate(r"""() => {
        const norm = s => (s || '').replace(/\s+/g, ' ').trim();
        const vis = el => el && el.offsetParent !== null;
        // find the search submit button
        let btn = Array.from(document.querySelectorAll('button, [role=button], span'))
            .find(b => /^(Поиск|Search|חיפוש)$/i.test(norm(b.textContent)));
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
    }""")
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
    """
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
                        'button:has-text("Применить")', 'button:has-text("Apply")'):
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
        ("father", ["Отец", "Father"], "father_first_name_field",
         "father_last_name_field", "father_filter_apply_button"),
        ("mother", ["Мать", "Mother"], "mother_first_name_field",
         "mother_last_name_field", "mother_filter_apply_button"),
        ("spouse", ["Супруг(-а)", "Супруг", "Spouse"], "spouse_first_name_field",
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
        if await _open_pill(["Смерть", "Death"], "death_year_field"):
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
            for fr in page.frames:
                try:
                    el = fr.get_by_text(re.compile(
                        r"Точное совпадение всех параметров|событие из жизни|"
                        r"Семейное положение", re.I)).first
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
            await _open_pill(["Больше", "More"], "")
            await asyncio.sleep(1.0)
        log("  → панель «+ Больше» открыта" if opened_more
            else "  !! панель «+ Больше» не открылась")

        for key, tags, place_auto, apply_id in [
            ("residence", ["Местожительство", "Residence"],
             "residence_place_field", "residence_filter_apply_button"),
            ("military", ["Вооруженные силы", "Military"],
             "military_place_field", "military_filter_apply_button"),
            ("immigration", ["Иммиграция", "Immigration"],
             "immigration_place_field", "immigration_filter_apply_button"),
        ]:
            if params.get(key):
                if await _open_pill(tags, place_auto):
                    await _fill_auto(place_auto, params[key])
                    await _apply(apply_id)

        if params.get("keywords"):
            if await _open_pill(["Ключевые слова", "Keywords"], "keywords_field"):
                await _fill_auto("keywords_field", params["keywords"])
                await _apply("keywords_filter_apply_button")

        g = params.get("gender", "Any")
        if g not in ("Any", "Любой"):
            if await _open_pill(["Пол", "Gender"], ""):
                ru = {"Male": "Мужчина", "Female": "Женщина"}.get(g, g)
                try:
                    opt = root.get_by_text(re.compile(rf"^{re.escape(ru)}$", re.I)).first
                    if await opt.count():
                        await opt.click(timeout=3000)
                except Exception:
                    pass
                await _apply("gender_filter_apply_button")

        # «Точное совпадение всех параметров». The checkbox text/role lives in
        # the "+ Больше" panel; it may render slightly late and/or sit below a
        # scroll. Poll up to 6s across all frames, tag the checkbox via JS,
        # scroll it in and REAL-click it; verify aria-checked flips to true.
        if params.get("exact_match"):
            done = False
            for _attempt in range(12):
                seen_count = 0
                for fr in page.frames:
                    try:
                        info = await fr.evaluate(r"""() => {
                            const norm = s => (s || '').replace(/\s+/g, ' ').trim();
                            document.querySelectorAll('[data-pw-exact]').forEach(
                                e => e.removeAttribute('data-pw-exact'));
                            const els = Array.from(document.querySelectorAll(
                                '[role=checkbox], span, label, div'));
                            let found = 0, checked = null;
                            for (const e of els) {
                                const t = norm(e.textContent);
                                if (t.length > 60) continue;
                                if (!/Точное совпадение/i.test(t)) continue;
                                found++;
                                if (e.offsetParent === null) continue;   // hidden
                                const cb = e.closest('[role=checkbox]') || e;
                                cb.setAttribute('data-pw-exact', '1');
                                checked = cb.getAttribute('aria-checked');
                                return {found, visible: true, checked};
                            }
                            return {found, visible: false, checked};
                        }""")
                        seen_count += info.get("found", 0)
                        if info.get("visible"):
                            cb = fr.locator('[data-pw-exact="1"]').first
                            try:
                                await cb.scroll_into_view_if_needed(timeout=2000)
                            except Exception:
                                pass
                            if info.get("checked") != "true":
                                try:
                                    await cb.click(timeout=3000)
                                except Exception:
                                    await fr.evaluate(
                                        "() => { const e=document.querySelector("
                                        "'[data-pw-exact=\"1\"]'); if(e) e.click(); }")
                                await asyncio.sleep(0.4)
                            log("  ✓ Точное совпадение всех параметров")
                            done = True
                            break
                    except Exception:
                        continue
                if done:
                    break
                await asyncio.sleep(0.5)
            if not done:
                log(f"  !! чекбокс «Точное совпадение» не найден "
                    f"(в DOM элементов с текстом: {seen_count})")


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
    log(f"  ✓ Страница результатов: {results_page.url[:80]}")
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
            if (ru || en) { divider = e; dlen = t.length; }
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
        await asyncio.sleep(2)
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

        // record photo — a REAL person/document image only. Skip brand logos
        // (Geni), avatars, icons, sprites. NOTE: do NOT skip "myheritage" — the
        // real person photos are served from the MyHeritage CDN!
        const SKIP = /avatar|icon|sprite|placeholder|logo|geni|brand|\.svg|badge|flag|blank/i;
        const img = Array.from(document.querySelectorAll('img[src]')).find(i =>
            (i.naturalWidth || i.width || 0) >= 100 &&
            (i.naturalHeight || i.height || 0) >= 100 &&
            !SKIP.test(i.src || '') && !SKIP.test(i.alt || ''));
        if (img) res.photo = img.src;

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

    # Photo thumbnail for Word (download the preview image bytes)
    photo = info.get("photo") or ""
    if photo:
        try:
            ip = await page.context.new_page()
            try:
                r = await ip.goto(photo, timeout=15000)
                if r and r.ok:
                    body = await r.body()
                    if len(body) > 1000:
                        d["thumb_bytes"] = body
            finally:
                await ip.close()
        except Exception:
            pass
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
        # Photo thumbnail (historical records: preview only — full needs Omni)
        tb = rec.get("thumb_bytes")
        if tb:
            try:
                doc.add_picture(io.BytesIO(tb), width=Inches(2.2))
            except Exception:
                pass
        # "View full profile on this site" link
        if rec.get("profile_url"):
            pp = doc.add_paragraph()
            pp.add_run("Полный профиль: ").bold = True
            _hyperlink(pp, "Посмотреть полный профиль на этом сайте",
                       rec["profile_url"])
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
                        u = (p.url or "").lower()
                        if any(s in u for s in ("facebook", "fbcdn",
                                                "accounts.google", "apple.com")):
                            await p.close()
                            return
                        if "myheritage" in u:
                            return                      # real results tab — keep
                        await asyncio.sleep(0.3)
                    # Still blank after 3s and not MyHeritage → junk, close it
                    if "myheritage" not in (p.url or "").lower():
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
                dp = None
                try:
                    dp = await ctx.new_page()
                    det = await _detail(dp, r["url"], has_cookies, log)
                    det["score"] = r["score"]
                    if not det.get("full_name"):
                        det["full_name"] = r["name_text"]
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
            # Persistent context — closing it saves cookies/session for next run
            try:
                await ctx.close()
            except Exception:
                pass

    return summary
