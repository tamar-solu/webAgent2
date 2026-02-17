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


def print_then_save_pdf(context, page, save_path: Path) -> None:
    """
    Click 'Print' in the Stimulsoft viewer.  Stimulsoft creates a hidden
    iframe (#stiPrintReportFrame) with the report as full HTML, then calls
    window.print().  We block window.print(), extract that HTML, inject a
    <base> tag so resources resolve, and use CDP Page.printToPDF on a fresh
    page to produce a PDF identical to Chrome's "Print → Save as PDF".
    """

    # Block window.print() so the native dialog never opens
    page.evaluate("""() => {
        window.__printCalled = false;
        window.print = function() { window.__printCalled = true; };
    }""")

    # Click the Print toolbar button
    page.get_by_role("cell", name="Print").nth(2).click()

    # Wait for Stimulsoft to build the print iframe
    page.locator("#stiPrintReportFrame").wait_for(state="attached", timeout=30_000)
    page.wait_for_timeout(2_000)  # let it finish rendering

    # Get the base URL so relative resources in the iframe HTML resolve
    base_url = page.evaluate("() => window.location.origin")

    # Extract the full HTML from the print iframe and inject <base>
    print_html = page.evaluate("""([baseUrl]) => {
        const iframe = document.getElementById('stiPrintReportFrame');
        if (!iframe || !iframe.contentDocument) return null;
        const doc = iframe.contentDocument;
        // Inject <base> tag so relative URLs resolve against the original site
        if (!doc.querySelector('base')) {
            const base = doc.createElement('base');
            base.href = baseUrl;
            doc.head.prepend(base);
        }
        return doc.documentElement.outerHTML;
    }""", [base_url])

    if not print_html:
        Path("debug").mkdir(exist_ok=True)
        page.screenshot(path="debug/print_iframe_missing.png", full_page=True)
        raise RuntimeError("Could not extract HTML from #stiPrintReportFrame")

    # Open a fresh page, load the print HTML, and generate the PDF via CDP
    print_page = context.new_page()
    print_page.set_content(print_html, wait_until="networkidle")
    print_page.wait_for_timeout(2_000)

    cdp = context.new_cdp_session(print_page)
    result = cdp.send("Page.printToPDF", {
        "printBackground": False,
        "preferCSSPageSize": True,
    })
    cdp.detach()
    print_page.close()

    pdf_bytes = base64.b64decode(result["data"])
    save_path.write_bytes(pdf_bytes)

    # Clean up the print iframe from the original page
    page.evaluate("""() => {
        const f = document.getElementById('stiPrintReportFrame');
        if (f) f.remove();
    }""")

    if not save_path.exists() or save_path.stat().st_size < 5_000:
        Path("debug").mkdir(exist_ok=True)
        page.screenshot(path="debug/print_pdf_too_small_or_missing.png", full_page=True)
        raise RuntimeError(f"PDF missing/too small after print: {save_path}")


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