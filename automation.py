# automation_print_to_pdf.py
import base64
import os
import time
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from pathlib import Path

import win32com.client
from playwright.sync_api import sync_playwright, expect, TimeoutError as PWTimeoutError

PRES_URL = os.environ["PRES_URL"]
RECIPIENT = os.environ["RECIPIENT"]
SUBJECT = os.environ["SUBJECT"]

MAX_RETRIES = int(os.environ["MAX_RETRIES"])
RETRY_DELAY_SECONDS = int(os.environ["RETRY_DELAY_SECONDS"])


def yesterday_str_il(fmt: str = "%d/%m/%Y") -> str:
    y = datetime.now(ZoneInfo("Asia/Jerusalem")).date() - timedelta(days=1)
    return y.strftime(fmt)

def set_date_input(page, selector: str, value: str) -> None:
    # set value + dispatch events so site commits it
    page.evaluate(
        """([sel, val]) => {
            const el = document.querySelector(sel);
            if (!el) throw new Error("Date input not found: " + sel);
            el.focus();
            el.value = val;
            el.dispatchEvent(new Event('input', { bubbles: true }));
            el.dispatchEvent(new Event('change', { bubbles: true }));
            el.blur();
        }""",
        [selector, value],
    )


def send_via_outlook(subject: str, body: str, to_email: str, attachments: list[str]) -> None:
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = to_email
    mail.Subject = subject
    mail.Body = body
    for file_path in attachments:
        mail.Attachments.Add(str(file_path))
    mail.Send()
    print("Email sent via Outlook")


def open_report_view(page) -> None:
    page.locator("#report-criteria").get_by_role("link", name=" הצג דוח").click()
    page.locator("#report-view-back-btn").wait_for(state="visible", timeout=180_000)


def _extract_pdf_from_frame_html(target) -> bytes | None:
    """
    Stimulsoft writes print-ready HTML directly into stiPrintReportFrame
    (src=about:blank). Extract that HTML, load it into a clean temporary page
    with no viewer chrome, and generate a proper PDF from it.
    """
    sti_frame = target.frame(name="stiPrintReportFrame")
    if sti_frame is None:
        el = target.query_selector("#stiPrintReportFrame")
        if el:
            sti_frame = el.content_frame()

    if sti_frame is None:
        print("  stiPrintReportFrame not found")
        return None

    html = sti_frame.content()
    print(f"  frame HTML: {len(html):,} chars")

    temp_page = target.context.new_page()
    try:
        temp_page.set_content(html, wait_until="domcontentloaded")
        temp_page.wait_for_timeout(500)
        pdf_bytes = temp_page.pdf(
            print_background=False,
            prefer_css_page_size=True,
        )
        print(f"  PDF from frame HTML: {len(pdf_bytes):,} bytes")
        return pdf_bytes
    finally:
        temp_page.close()


def print_then_save_pdf(context, page, save_path: Path) -> None:
    """
    Click 'Print' to open the Stimulsoft viewer popup, extract the PDF from
    the blob URL the viewer creates, and save it directly.
    Falls back to page.pdf() if blob extraction fails.
    """
    popup = None
    try:
        with context.expect_page(timeout=5_000) as popup_info:
            page.get_by_role("cell", name="Print").nth(2).click()
        popup = popup_info.value
    except PWTimeoutError:
        try:
            page.get_by_text("Print", exact=True).click(timeout=2_000)
        except Exception:
            pass

    target = popup if popup else page

    try:
        target.wait_for_load_state("domcontentloaded", timeout=60_000)
    except PWTimeoutError:
        pass
    try:
        target.wait_for_load_state("networkidle", timeout=60_000)
    except PWTimeoutError:
        pass
    # Extra time for Stimulsoft to finish generating the PDF blob
    target.wait_for_timeout(3_000)

    # Attempt 1: extract clean HTML from stiPrintReportFrame and render to PDF
    pdf_bytes = _extract_pdf_from_frame_html(target)
    if pdf_bytes:
        save_path.write_bytes(pdf_bytes)
        print(f"PDF saved from frame HTML ({len(pdf_bytes):,} bytes)")
    else:
        # Fallback: page.pdf() — captures viewer chrome but always produces output
        print("Frame extraction failed, falling back to page.pdf()")
        target.pdf(path=str(save_path), print_background=True, prefer_css_page_size=True)

    if popup:
        popup.close()

    if not save_path.exists() or save_path.stat().st_size < 5_000:
        Path("debug").mkdir(exist_ok=True)
        page.screenshot(path="debug/print_pdf_too_small_or_missing.png", full_page=True)
        raise RuntimeError(f"PDF missing/too small after print->pdf: {save_path}")


def run() -> None:
    # ---- creds from ENV (do not hardcode) ----
    pres_code = os.environ["PRES_POS_CODE"]
    pres_username = os.environ["NLC_USER"]
    pres_password = os.environ["NLC_PASSWORD"]

    out_dir = Path(os.environ.get("PRES_OUT_DIR", "downloads")).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    file_meznon = out_dir / "מזנון.pdf"
    file_kupot = out_dir / "קופות.pdf"

    date_str = yesterday_str_il("%d/%m/%Y")

    with sync_playwright() as p:
        # MUST be headless for PDF generation
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            viewport={"width": 1400, "height": 900},
            accept_downloads=True,
        )
        page = context.new_page()

        # login
        page.goto(PRES_URL, wait_until="domcontentloaded")
        page.get_by_role("textbox", name="קוד").fill(pres_code)
        page.get_by_role("textbox", name="שם משתמש").fill(pres_username)
        page.get_by_role("textbox", name="סיסמא").fill(pres_password)
        page.get_by_role("button", name="היכנס").click()
        expect(page.get_by_text("הנהלת חשבונות", exact=True)).to_be_visible(timeout=30_000)

        # navigate
        page.get_by_text("הנהלת חשבונות", exact=True).click()
        page.get_by_role("link", name="דוחות ").click()
        page.get_by_role("link", name="דוח יומי").click()
        expect(page.locator("#report-criteria")).to_be_visible(timeout=30_000)

        # filters
        page.locator('select[name="filterLocations"]').select_option(
            ["1181", "1178", "1170", "1350", "1174", "1175", "1176", "1173"]
        )

        set_date_input(page, 'input[name="startDatePicker"]', date_str)
        set_date_input(page, 'input[name="endDatePicker"]', date_str)

        print("start:", page.locator('input[name="startDatePicker"]').input_value())
        print("end:", page.locator('input[name="endDatePicker"]').input_value())

        # ---- report 1: מזנון (posClasses=3) ----
        page.locator('select[name="posClasses"]').select_option("3")
        open_report_view(page)
        print_then_save_pdf(context, page, file_meznon)

        # back
        page.locator("#report-view-back-btn").click()
        expect(page.locator("#report-criteria")).to_be_visible(timeout=30_000)

        # ---- report 2: קופות (posClasses=[1,4,2,5]) ----
        page.locator('select[name="posClasses"]').select_option(["1", "4", "2", "5"])
        open_report_view(page)
        print_then_save_pdf(context, page, file_kupot)

        context.close()
        browser.close()

    # send via Outlook (desktop must be installed/logged-in)
    send_via_outlook(
        subject=SUBJECT,
        body="Attached are the two reports: מזנון and קופות.",
        to_email=RECIPIENT,
        attachments=[str(file_meznon), str(file_kupot)],
    )

    print(f"Done:\n- {file_meznon}\n- {file_kupot}")


if __name__ == "__main__":
    last_error = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            print(f"Attempt {attempt}/{MAX_RETRIES}...")
            run()
            print("Success.")
            break
        except Exception as e:
            last_error = e
            print(f"Attempt {attempt}/{MAX_RETRIES} failed: {e}")
            if attempt < MAX_RETRIES:
                print(f"Retrying in {RETRY_DELAY_SECONDS} seconds...")
                time.sleep(RETRY_DELAY_SECONDS)
    else:
        raise RuntimeError(
            f"Script failed after {MAX_RETRIES} attempts. Last error: {last_error}"
        ) from last_error
