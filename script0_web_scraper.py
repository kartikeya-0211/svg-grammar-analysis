"""
WEB SCRAPER - FRESH START
1. Creates 'railroad_diagrams.xlsx' with 10 empty columns.
2. Scrapes Command Name -> Column A
3. Scrapes URL -> Column B
4. Scrapes Raw SVG -> Column C
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font
import os
import re
import time

# --- CONFIGURATION ---
LINKS_FILE = 'links_cics.txt'
OUTPUT_FILE = 'railroad_diagrams.xlsx'
DRIVER_PATH = os.path.join(os.getcwd(), 'msedgedriver.exe')

def extract_command_name_from_url(url):
    """Extracts command name from URL."""
    filename = url.split('/')[-1].replace('.html', '')
    command = filename.replace('dfhp4_', '').replace('dfhp4-', '')
    words = re.sub(r'([a-z])([A-Z])', r'\1 \2', command).upper()
    return words 

def setup_driver():
    """Setup Edge driver."""
    edge_options = Options()
    edge_options.add_argument('--headless')
    edge_options.add_argument('--enable-unsafe-swiftshader')
    edge_options.add_argument("--log-level=3")
    edge_options.add_experimental_option('excludeSwitches', ['enable-logging'])
    edge_options.add_argument('--window-size=1920,1080')
    
    if not os.path.exists(DRIVER_PATH):
        raise Exception(f"❌ msedgedriver.exe not found at: {DRIVER_PATH}")
    
    service = Service(DRIVER_PATH)
    service.creation_flags = 0x08000000 
    
    driver = webdriver.Edge(service=service, options=edge_options)
    return driver

def create_excel_structure(filename):
    """Creates the specific 10-column layout."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Railroad Data"
    
    # THE 10 COLUMNS
    headers = [
        "Command",                  # A (Script 0)
        "URL",                      # B (Script 0)
        "Raw SVG Code",             # C (Script 0)
        "Raw SVG Image",            # D (Script 1)
        "Simplified SVG Code",      # E (Script 1)
        "Simplified SVG Image",     # F (Script 1)
        "Connected SVG Code",       # G (Script 1)
        "Connected SVG Image",      # H (Script 1)
        "Unoptimized Grammar",      # I (Script 2)
        "Optimized Regex"           # J (Script 3)
    ]
    
    # Apply Headers
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx)
        cell.value = header
        cell.font = Font(bold=True, size=11)
        cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # Set Widths for visibility
    ws.column_dimensions['A'].width = 25  # Command
    ws.column_dimensions['B'].width = 40  # URL
    ws.column_dimensions['C'].width = 40  # Raw Code
    ws.column_dimensions['D'].width = 50  # Raw Img
    ws.column_dimensions['E'].width = 25  # Simp Code
    ws.column_dimensions['F'].width = 50  # Simp Img
    ws.column_dimensions['G'].width = 25  # Conn Code
    ws.column_dimensions['H'].width = 50  # Conn Img
    ws.column_dimensions['I'].width = 40  # Grammar
    ws.column_dimensions['J'].width = 50  # Regex
    
    ws.row_dimensions[1].height = 30
    
    return wb, ws

def extract_svg(driver, url):
    """Loads page and grabs SVG HTML."""
    try:
        driver.get(url)
        # Fast wait: 2 seconds max
        try:
            svg = WebDriverWait(driver, 2).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'svg.syntaxdiagram'))
            )
        except:
            try:
                svg = WebDriverWait(driver, 1).until(
                    EC.presence_of_element_located((By.TAG_NAME, 'svg'))
                )
            except:
                return None
        return svg.get_attribute('outerHTML')
    except:
        return None

def main():
    print("=" * 70)
    print("STEP 1: FRESH WEB SCRAPER (10-COLUMN SETUP)")
    print("=" * 70)
    
    # 1. Read Links
    if not os.path.exists(LINKS_FILE):
        print(f"❌ Error: '{LINKS_FILE}' not found.")
        return

    with open(LINKS_FILE, 'r') as f:
        urls = [line.strip() for line in f if line.strip()]

    print(f"📄 Found {len(urls)} URLs.")
    
    # 2. Ask Mode
    choice = input("Run (t)est [first 3] or (f)ull [all]? ").strip().lower()
    if choice == 't':
        urls = urls[:3]
        print("🧪 Test Mode: Processing 3 URLs...")
    
    # 3. Create File
    print(f"💾 Creating {OUTPUT_FILE} with 10 columns...")
    wb, ws = create_excel_structure(OUTPUT_FILE)
    
    # 4. Scrape
    driver = setup_driver()
    print("✅ Driver ready. Starting scrape...\n")
    
    row = 2
    success = 0
    
    try:
        for i, url in enumerate(urls, 1):
            cmd = extract_command_name_from_url(url)
            print(f"[{i}] {cmd}...", end=" ")
            
            svg_content = extract_svg(driver, url)
            
            # Write Data
            ws.cell(row=row, column=1, value=cmd)      # Col A
            ws.cell(row=row, column=2, value=url)      # Col B
            
            if svg_content:
                cell_c = ws.cell(row=row, column=3, value=svg_content) # Col C
                cell_c.alignment = Alignment(wrap_text=True, vertical='top')
                ws.row_dimensions[row].height = 40
                print("✅ SVG Saved")
                success += 1
            else:
                print("⚠️ No SVG found")
                
            row += 1
            # Save every 5 rows to be safe
            if i % 5 == 0: wb.save(OUTPUT_FILE)
            
    finally:
        driver.quit()
        wb.save(OUTPUT_FILE)
        print("-" * 70)
        print(f"✅ Finished. Scraped {success}/{len(urls)}.")
        print(f"📂 Output: {OUTPUT_FILE}")

if __name__ == "__main__":
    main()