from playwright.sync_api import sync_playwright
import time
import re
import sys
import pandas as pd
from datetime import datetime
#import ctypes

sys.stdout.reconfigure(encoding='utf-8')

# ---------- CONFIG ----------
SAP_URL_LOGIN = "https://my416346.s4hana.cloud.sap/ui#Shell-home"
SAP_URL_INVOICES = "https://my416346.s4hana.cloud.sap/ui?sap-ushell-config=lean#CentralInvoice-manageSupplierInvoice"

USERNAME = "Your_Username"
PASSWORD = "Password"

INPUT_EXCEL = r"C:\Users\Input.xlsx"   # Sheet1 contains Invoice IDs in column A
OUTPUT_EXCEL = r"C:\Users\Output.xlsx" # Extracted data will be written here


def extract_invoice_log(page, invoice_id):
    """Extracts activity log data for given invoice"""
    data = []

    try:
        # 6️⃣ Apply Invoice Filter
        invoice_filter = page.locator(
            "#application-SupplierInvoice-manageCentrally-component---InvoiceList--filterbar_filteritem_filterCimId-inner"
        )
        invoice_filter.fill(invoice_id)
        time.sleep(1)
        invoice_filter.press("Enter")
        print(f"[INFO] Filter applied for Invoice: {invoice_id}")
        time.sleep(5)

        try:
            Invoice_No = page.locator(".sapMObjectIdentifierTopRow").inner_text().strip()
        except Exception:
            Invoice_No = ""

        # 7️⃣ Click on Invoice Row
        page.locator(
            "#application-SupplierInvoice-manageCentrally-component---InvoiceList--table-tblBody"
        ).click()
        time.sleep(5)

        # 8️⃣ Click on Log Activity Button
        page.locator(
            "#application-SupplierInvoice-manageCentrally-component---InvoiceDetail--ObjectPageHeader_ObjectPageHeaderOpenActivityLogButton-BDI-content"
        ).click()
        print(f"[INFO] Opened Log Activity for Invoice: {invoice_id}")
        time.sleep(8)

        try:
            page.click("#__mbox-btn-0", timeout=1000)
        except Exception:
            "No Error Popup Appears"

        while True:
            
            # 9️⃣ Extract Log Items
            items = page.locator(".sapSuiteUiCommonsTimelineItemShell")
            count = items.count()

            item = items.nth(count - 1)         
            item.locator(".sapSuiteUiCommonsTimelineItemShellUser").click()
            time.sleep(1)
            items = page.locator(".sapSuiteUiCommonsTimelineItemShell")
            Newcount = items.count()

            if count == Newcount:
                "Scrolled Till End"
                break
            else:
                "Scrolling Untill Reach to End"

            count = Newcount

        for i in range(count):
            item = items.nth(i)
            user = item.locator(".sapSuiteUiCommonsTimelineItemShellUser").inner_text().strip()

            # Extract action
            action_text = item.locator(".sapSuiteUiCommonsTimelineItemShellHdr").inner_text().strip()
            action_match = re.search(r"^\S+", action_text)
            action = action_match.group(0).strip() if action_match else ""

            # Extract field name (everything after action word)
            field_match = re.search(rf"(?<=\b{action}\s).*", action_text) if action else None
            field_name = field_match.group(0).strip() if field_match else ""

            timestamp = item.locator(".sapSuiteUiCommonsTimelineItemShellDateTime").inner_text().replace("\u202f", " ")

            try:
                details = item.locator(".sapSuiteUiCommonsTimelineItemTextWrapper span").inner_text().strip()

                new_value_match = re.search(r"(?<=New value:\s).*?(?=\s*Previous value:)", details)
                new_value = new_value_match.group(0).strip() if new_value_match else ""

                prev_value_match = re.search(r"(?<=Previous value:\s).*", details)
                prev_value = prev_value_match.group(0).strip() if prev_value_match else ""

            except Exception:
                new_value, prev_value = "", ""

            bot_time = datetime.now().strftime("%d %b %Y %H:%M:%S")

            data.append({
                "InvoiceID": invoice_id,
                "Invoice Number": Invoice_No,
                "User": user,
                "Action": action,
                "Field Name": field_name,
                "Time": timestamp,
                "New Value": new_value,
                "Previous Value": prev_value,
                "BotExecutionTime": bot_time
            })

        print(f"[SUCCESS] Extracted {len(data)} log entries for Invoice {invoice_id}")

    except Exception as e:
        print(f"[ERROR] Failed for Invoice {invoice_id}: {e}")

    return data


with sync_playwright() as p:

    browser = p.chromium.launch(headless=False, args=["--start-maximized"])
    context = browser.new_context(no_viewport=True)   # 👈 disables viewport size limit
    page = context.new_page()

    # 1️⃣ Open SAP Login Page
    page.goto(SAP_URL_LOGIN, timeout=60000)

    # 2️⃣ Handle iframe if present
    login_frame = None
    try:
        frame_locator = page.frame_locator("iframe")
        if frame_locator.locator("#j_username").count() > 0:
            login_frame = frame_locator
    except:
        pass

    if not login_frame:
        login_frame = page  # fallback

    # 3️⃣ Perform Login
    try:
        login_frame.locator("#j_username").fill(USERNAME)
        login_frame.locator("#j_password").fill(PASSWORD)
        login_frame.locator("#logOnFormSubmit").click()
        print("✅ SAP S4H Login Successful!")
    except Exception as e:
        print("❌ Login Failed:", e)

    # 4️⃣ Wait for Launchpad
    page.wait_for_url(re.compile(r".*Shell-home.*"), timeout=60000)

    # 5️⃣ Navigate to "Manage Supplier Invoices"
    page.goto(SAP_URL_INVOICES, timeout=60000)
    page.wait_for_url(re.compile(r".*SupplierInvoice-manageCentrally.*"), timeout=60000)

    # Wait until filter field is available
    page.wait_for_selector(
        "#application-SupplierInvoice-manageCentrally-component---InvoiceList--filterbar_filteritem_filterCimId-inner",
        timeout=60000
    )

    # ✅ use openpyxl engine explicitly
    df_input = pd.read_excel(INPUT_EXCEL, sheet_name="Sheet1", engine="openpyxl")
    invoice_ids = df_input.iloc[:, 0].dropna().astype(str).tolist()

    all_data = []

    # Process each Invoice
    for inv in invoice_ids:
        logs = extract_invoice_log(page, inv)
        all_data.extend(logs)

        # 🔄 Go back to invoice list after processing
        page.goto(SAP_URL_INVOICES, timeout=60000)
        page.wait_for_selector(
            "#application-SupplierInvoice-manageCentrally-component---InvoiceList--filterbar_filteritem_filterCimId-inner",
            timeout=60000
        )
        time.sleep(2)

    # Save results to Excel (Sheet2)
    if all_data:
        df_output = pd.DataFrame(all_data)
        with pd.ExcelWriter(OUTPUT_EXCEL, engine="openpyxl", mode="w") as writer:
            df_input.to_excel(writer, sheet_name="Sheet1", index=False)  # keep original IDs
            df_output.to_excel(writer, sheet_name="Sheet2", index=False)
        print(f"✅ Extracted data saved to {OUTPUT_EXCEL}")
    else:
        print("⚠ No data extracted!")

    # Get handle of the active (foreground) window
    #hWnd = ctypes.windll.user32.GetForegroundWindow()

    # Show message box attached to active window
    #ctypes.windll.user32.MessageBoxW(hWnd, "All Logs has been Exported!!!", "Export Complete", 0)

    browser.close()
