# SAP Logs Export Automation

This project automates the extraction of SAP logs using **Python + Playwright**.

## Features
- Launches SAP web portal in full screen
- Handles popups automatically
- Extracts logs and exports to CSV/Excel
- Displays completion message via Windows MessageBox

## Setup
1. Clone repository:
   ```bash
   git clone https://github.com/your-username/sap-logs-export-automation.git

pip install -r requirements.txt
playwright install

python src/main.py

Notes

Make sure you have Python 3.9+ installed

Playwright will download Chromium/Firefox/Webkit on first run


---

### ✅ Git Commands to Upload  

```bash
# Initialize git
git init
git add .
git commit -m "Initial commit - SAP Logs Export Automation"

# Add remote
git branch -M main
git remote add origin https://github.com/YOUR-USERNAME/sap-logs-export-automation.git

# Push
git push -u origin main

⚠️ If chrome.exe is missing

That means the download was incomplete. In that case, two workarounds:

Manually download Chromium for Playwright:

Download ZIP for playwright build v1181 from here:
👉 https://playwright.azureedge.net/builds/chromium/1181/chromium-win.zip

Extract to:
C:\Users\Administrator\AppData\Local\ms-playwright\chromium-1181\chrome-win\

Or use system-installed Chrome instead:

chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

with sync_playwright() as p:
    browser = p.chromium.launch(
        headless=False,
        executable_path=chrome_path
    )
