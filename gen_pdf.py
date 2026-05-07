from playwright.sync_api import sync_playwright
import os

report_path = r"C:\Users\User\bosom-lite-pq-sim\report.html"
pdf_path = r"C:\Users\User\Downloads\보솜라이트_PQ시뮬레이션_보고서.pdf"

with sync_playwright() as p:
    browser = p.chromium.launch()
    page = browser.new_page()
    page.goto("file:///" + report_path.replace("\\", "/"))
    page.wait_for_timeout(2000)
    page.pdf(
        path=pdf_path,
        format="A4",
        landscape=True,
        print_background=True,
        margin={"top": "0", "right": "0", "bottom": "0", "left": "0"}
    )
    browser.close()

print(f"PDF saved: {pdf_path}")
print(f"Size: {os.path.getsize(pdf_path):,} bytes")
