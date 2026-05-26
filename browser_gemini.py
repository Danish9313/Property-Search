"""
browser_gemini.py
=================
Playwright-based Gemini caller with Web Search enabled.
Replaces the direct API call — all other pipeline logic stays unchanged.

First run : a browser window opens. Log in to your Google account once.
            The session is saved in browser_profile/ and reused automatically.
Subsequent: browser opens, logs in automatically, web search is toggled on,
            prompt is submitted, response is returned as plain text.
"""

from playwright.sync_api import sync_playwright, Playwright, BrowserContext, Page
from pathlib import Path
import time
import logging
import pyperclip
import config

logger = logging.getLogger(__name__)

PROFILE_DIR = str(Path(__file__).parent / "browser_profile")
GEMINI_URL  = "https://gemini.google.com/app"

# ---------------------------------------------------------------------------
# Singleton — browser launched once, reused for every row
# ---------------------------------------------------------------------------
_playwright: Playwright | None = None
_context: BrowserContext | None = None
_page: Page | None = None


def _start_browser() -> Page:
    global _playwright, _context, _page

    if _page is not None:
        return _page

    _playwright = sync_playwright().start()

    # Use system Chrome so Google accepts the session and saves login properly
    launch_kwargs = dict(
        user_data_dir=PROFILE_DIR,
        headless=config.BROWSER_HEADLESS,
        args=[
            "--start-maximized",
            "--disable-blink-features=AutomationControlled",
        ],
        viewport=None,
        permissions=["clipboard-read", "clipboard-write"],
        ignore_default_args=["--enable-automation"],
    )
    try:
        _context = _playwright.chromium.launch_persistent_context(
            channel="chrome", **launch_kwargs
        )
        logger.info("Using system Chrome for Gemini.")
    except Exception:
        logger.warning("System Chrome not found — falling back to Playwright Chromium.")
        _context = _playwright.chromium.launch_persistent_context(**launch_kwargs)

    _page = _context.new_page()
    _page.add_init_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    _page.goto(GEMINI_URL, wait_until="domcontentloaded", timeout=30000)

    # Wait for page to settle then check login state
    time.sleep(4)

    def _is_logged_in() -> bool:
        try:
            url = _page.url
            if "accounts.google.com" in url or "signin" in url.lower():
                return False
            # Check for Sign In button in page
            btns = _page.evaluate("""
                () => Array.from(document.querySelectorAll('a,button'))
                         .map(b => b.innerText.trim().toLowerCase())
            """)
            return not any(t in ("sign in", "log in", "signin") for t in btns)
        except Exception:
            return True

    if not _is_logged_in():
        logger.info("=" * 60)
        logger.info("ACTION REQUIRED: Log in to your Google account in the browser.")
        logger.info("Session will be saved automatically — you will not need to log in again.")
        logger.info("=" * 60)
        deadline = time.time() + 180
        while time.time() < deadline:
            time.sleep(3)
            if _is_logged_in():
                logger.info("Login detected — continuing.")
                time.sleep(3)
                break
        else:
            logger.warning("Login wait timed out — proceeding anyway.")

    logger.info("Gemini browser session ready.")
    return _page


def close_browser():
    """Call at end of pipeline run to cleanly close the browser."""
    global _playwright, _context, _page
    if _context:
        try:
            _context.close()
        except Exception:
            pass
    if _playwright:
        try:
            _playwright.stop()
        except Exception:
            pass
    _playwright = _context = _page = None


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _new_chat(page: Page):
    """Always navigate to a fresh URL to guarantee a clean conversation."""
    page.goto(GEMINI_URL, wait_until="domcontentloaded", timeout=30000)
    # Wait for input field to appear before proceeding
    try:
        page.wait_for_selector(
            'rich-textarea [contenteditable="true"], div[contenteditable="true"], textarea',
            timeout=15000,
        )
    except Exception:
        pass
    time.sleep(2)


def _enable_web_search(page: Page):
    """Toggle on the Web Search tool in Gemini's input toolbar."""
    selectors = [
        '[aria-label*="Web search"]',
        '[aria-label*="Search the web"]',
        '[aria-label*="Google Search"]',
        '[title*="Web search"]',
        '[title*="Search"]',
        'button:has-text("Search")',
        '[data-tool-id*="search"]',
    ]
    for sel in selectors:
        try:
            btn = page.locator(sel).first
            if btn.is_visible(timeout=1500):
                pressed = btn.get_attribute("aria-pressed")
                selected = btn.get_attribute("aria-selected")
                if pressed != "true" and selected != "true":
                    btn.click()
                    time.sleep(0.5)
                logger.info("Web search enabled.")
                return
        except Exception:
            continue
    logger.warning("Web search toggle not found — proceeding without it.")


def _type_and_send(page: Page, prompt: str):
    """Paste the prompt into Gemini's input box via clipboard and send it."""
    input_selectors = [
        'rich-textarea [contenteditable="true"]',
        '[contenteditable="true"][aria-label*="message"]',
        '[contenteditable="true"][aria-label*="prompt"]',
        'div[contenteditable="true"]',
        'textarea',
    ]
    input_el = None
    for sel in input_selectors:
        try:
            el = page.locator(sel).last
            el.wait_for(state="visible", timeout=10000)
            input_el = el
            logger.info(f"Gemini input found: {sel}")
            break
        except Exception:
            continue

    if input_el is None:
        # Save screenshot to help diagnose
        try:
            page.screenshot(path=str(Path(__file__).parent / "gemini_debug.png"))
        except Exception:
            pass
        raise RuntimeError("Could not locate Gemini input field.")

    # Copy full prompt to OS clipboard (no size limit, works with any length)
    pyperclip.copy(prompt)
    logger.info(f"Prompt copied to clipboard: {len(prompt)} chars")

    # Click input to focus it
    input_el.click()
    time.sleep(0.5)

    # Clear existing content
    page.keyboard.press("Control+a")
    page.keyboard.press("Delete")
    time.sleep(0.3)

    # Paste from OS clipboard — guaranteed to paste the full text
    page.keyboard.press("Control+v")
    time.sleep(2)

    # Verify how many chars are actually in the input field after paste
    pasted_len = page.evaluate("""
        () => {
            const el = document.querySelector('rich-textarea [contenteditable="true"]')
                    || document.querySelector('div[contenteditable="true"]');
            return el ? el.innerText.length : 0;
        }
    """)
    logger.info(f"Chars in prompt: {len(prompt)} | Chars pasted into Gemini: {pasted_len}")

    # Try send button first, then Enter
    send_selectors = [
        'button[aria-label*="Send"]',
        'button[aria-label*="submit"]',
        'button[data-test-id="send-button"]',
    ]
    for sel in send_selectors:
        try:
            btn = page.locator(sel).first
            if btn.is_visible(timeout=1000):
                btn.click()
                return
        except Exception:
            continue
    page.keyboard.press("Enter")


def _wait_and_extract(page: Page, timeout: int = 180) -> str:
    """Poll until Gemini stops generating, then return the LAST model response text only."""
    response_selectors = [
        'model-response',
        '.response-container',
        '[data-message-author-role="model"]',
        '.model-response-text',
        'message-content',
    ]

    def _get_current_text() -> str:
        for sel in response_selectors:
            try:
                els = page.locator(sel).all()
                if els:
                    text = els[-1].inner_text(timeout=2000).strip()
                    if text and len(text) < 15000:  # ignore huge page dumps
                        return text
            except Exception:
                continue
        return ""  # never fall back to main — avoids grabbing full page

    # Wait for large prompts to start generating
    time.sleep(10)

    prev_text = ""
    stable_count = 0
    deadline = time.time() + timeout

    while time.time() < deadline:
        current_text = _get_current_text()

        if current_text and current_text == prev_text:
            stable_count += 1
            if stable_count >= 6 and current_text.rstrip().endswith("}"):
                logger.info("Response stable — extraction complete.")
                return current_text
            elif stable_count >= 20:
                logger.warning("Response stable but no closing } — returning anyway.")
                return current_text
        else:
            stable_count = 0

        prev_text = current_text
        time.sleep(1)

    logger.warning("Timed out waiting for stable response — returning what we have.")
    return prev_text


# Required output field names — used to detect wrong schema
_REQUIRED_FIELDS = [
    "Verified Equity", "Verified Liens", "Homestead Applied", "Homestead State",
    "Final Collateral", "Collateral Calculation", "Collateralization",
    "Principal Balance", "Collectibility Judgment", "Recovery Summary",
    "Criminal records:", "Other Assets:", "Professional Licenses:",
    "Other Owned Businesses:", "notes",
]

_REFORMAT_MSG = (
    "Your response used the wrong field names. "
    "Now output ONLY this JSON with EXACTLY these field names filled from your analysis. "
    "No explanation. Start with { end with }:\n\n"
    "{\n"
    + "\n".join(f'  "{f}": "",' for f in _REQUIRED_FIELDS)
    + "\n}"
)


def _uses_correct_schema(text: str) -> bool:
    return all(f'"{f}"' in text for f in _REQUIRED_FIELDS)


# ---------------------------------------------------------------------------
# Public entry point (called by main.py in place of the API)
# ---------------------------------------------------------------------------

def call_gemini_browser(prompt: str) -> str:
    """
    Submit prompt to Gemini via browser with Web Search enabled.
    If Gemini returns the wrong schema, sends a follow-up to reformat.
    Returns the raw response text.
    """
    page = _start_browser()
    try:
        _new_chat(page)
        _enable_web_search(page)
        _type_and_send(page, prompt)
        response = _wait_and_extract(page)

        logger.info(f"Gemini response length: {len(response)} chars")
        if response:
            logger.info(f"Gemini response preview: {response[:300]}")
        else:
            logger.warning("Gemini returned empty response.")
            return response

        # If wrong schema — send reformat follow-up in the same chat
        if not _uses_correct_schema(response):
            logger.info("Wrong schema — sending reformat follow-up.")
            _type_and_send(page, _REFORMAT_MSG)
            response = _wait_and_extract(page)
            logger.info(f"Reformat response: {len(response)} chars | preview: {response[:200]}")

        return response
    except Exception as e:
        logger.error(f"Gemini browser error: {e}")
        return ""
