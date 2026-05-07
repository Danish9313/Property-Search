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
    _context = _playwright.chromium.launch_persistent_context(
        user_data_dir=PROFILE_DIR,
        headless=config.BROWSER_HEADLESS,
        args=["--start-maximized"],
        viewport=None,
        permissions=["clipboard-read", "clipboard-write"],
    )
    _page = _context.new_page()
    _page.goto(GEMINI_URL, wait_until="domcontentloaded", timeout=30000)

    # If not logged in — wait for user to log in manually (up to 3 minutes)
    if "accounts.google.com" in _page.url or "signin" in _page.url.lower():
        logger.info("=" * 60)
        logger.info("ACTION REQUIRED: Log in to your Google account in the browser window.")
        logger.info("The session will be saved automatically after login.")
        logger.info("=" * 60)
        _page.wait_for_url("*gemini.google.com*", timeout=180000)
        time.sleep(3)

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
    """Navigate to a fresh conversation so rows don't bleed into each other."""
    try:
        for selector in [
            'a[href="/app"]',
            'button[aria-label*="New chat"]',
            'button[aria-label*="New conversation"]',
        ]:
            btn = page.locator(selector)
            if btn.count() > 0 and btn.first.is_visible(timeout=2000):
                btn.first.click()
                time.sleep(1.5)
                return
    except Exception:
        pass
    # Fallback: reload the app URL
    page.goto(GEMINI_URL, wait_until="domcontentloaded", timeout=20000)
    time.sleep(1.5)


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
    ]
    input_el = None
    for sel in input_selectors:
        try:
            el = page.locator(sel).last
            if el.is_visible(timeout=2000):
                input_el = el
                break
        except Exception:
            continue

    if input_el is None:
        raise RuntimeError("Could not locate Gemini input field.")

    # Click input to focus it
    input_el.click()
    time.sleep(0.5)

    # Clear existing content
    page.keyboard.press("Control+a")
    page.keyboard.press("Delete")
    time.sleep(0.3)

    # Insert text via execCommand — works with Gemini's Angular contenteditable
    page.evaluate(
        """(text) => {
            const el = document.querySelector('rich-textarea [contenteditable="true"]')
                    || document.querySelector('div[contenteditable="true"]');
            if (el) {
                el.focus();
                document.execCommand('selectAll');
                document.execCommand('insertText', false, text);
            }
        }""",
        prompt,
    )
    time.sleep(1)

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
    """Poll until Gemini stops generating, then return the complete response text."""
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
                    if text:
                        return text
            except Exception:
                continue
        try:
            return page.locator("main").inner_text(timeout=2000).strip()
        except Exception:
            return ""

    # Wait for response to start appearing
    time.sleep(4)

    prev_text = ""
    stable_count = 0
    deadline = time.time() + timeout

    while time.time() < deadline:
        current_text = _get_current_text()

        if current_text and current_text == prev_text:
            stable_count += 1
            if stable_count >= 4:   # unchanged for 4 seconds = generation done
                logger.info("Response stable — extraction complete.")
                return current_text
        else:
            stable_count = 0

        prev_text = current_text
        time.sleep(1)

    logger.warning("Timed out waiting for stable response — returning what we have.")
    return prev_text


# ---------------------------------------------------------------------------
# Public entry point (called by main.py in place of the API)
# ---------------------------------------------------------------------------

def call_gemini_browser(prompt: str) -> str:
    """
    Submit prompt to Gemini via browser with Web Search enabled.
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
    except Exception as e:
        logger.error(f"Gemini browser error: {e}")
        return ""
