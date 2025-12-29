"""
WEB SCRAPER - Creates proper Excel with all headers
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

def extract_command_name_from_url(url):
    """Extract command name from URL."""
    filename = url.split('/')[-1].replace('.html', '')
    command = filename.replace('dfhp4_', '').replace('dfhp4-', '')
    words = re.sub(r'([a-z])([A-Z])', r'\1 \2', command).upper()
    return words 


def setup_driver():
    """Setup Edge driver."""
    edge_options = Options()
    edge_options.add_argument('--headless')
    edge_options.add_argument('--disable-gpu')
    edge_options.add_argument('--no-sandbox')
    edge_options.add_argument('--disable-dev-shm-usage')
    edge_options.add_argument('--window-size=1920,1080')
    
    driver_path = os.path.join(os.getcwd(), 'msedgedriver.exe')
    if not os.path.exists(driver_path):
        raise Exception(f"msedgedriver.exe not found in: {os.getcwd()}")
    
    service = Service(driver_path)
    driver = webdriver.Edge(service=service, options=edge_options)
    return driver


def extract_svg_from_page(driver, url, timeout=10):
    """Load page and extract SVG."""
    try:
        print(f"    Loading: {url}...", end=" ")
        driver.get(url)
        time.sleep(2)
        
        svg_element = None
        try:
            svg_element = WebDriverWait(driver, timeout).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'svg.syntaxdiagram'))
            )
        except:
            try:
                svg_element = WebDriverWait(driver, 3).until(
                    EC.presence_of_element_located((By.TAG_NAME, 'svg'))
                )
            except:
                pass
        
        if svg_element:
            svg_html = svg_element.get_attribute('outerHTML')
            print("✅ SVG Captured!")
            return svg_html
        else:
            print("⚠️  No SVG found")
            return None
            
    except Exception as e:
        print(f"❌ Error: {e}")
        return None


def create_excel_with_headers(filename):
    """Create new Excel with all proper headers."""
    wb = Workbook()
    ws = wb.active
    
    headers = [
        "Command",
        "Original SVG",
        "Original Formatted SVG",
        "Original SVG Image",
        "Simplified Formatted SVG",
        "Simplified SVG Image"
    ]
    
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx)
        cell.value = header
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.font = Font(bold=True)
    
    ws.column_dimensions['A'].width = 25
    ws.column_dimensions['B'].width = 15
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 55
    ws.column_dimensions['E'].width = 15
    ws.column_dimensions['F'].width = 55
    
    return wb, ws


def scrape_cics_commands(links_file='links_cics.txt', output_excel='railroad_diagrams.xlsx', max_commands=None):
    """Main scraper."""
    
    print("=" * 70)
    print("CICS SVG WEB SCRAPER")
    print("=" * 70)
    print(f"📂 Reading links from: {links_file}")
    print(f"💾 Output Excel file: {output_excel}")
    if max_commands:
        print(f"⚠️  Test mode: Processing only first {max_commands} commands")
    print("-" * 70)
    
    # Read URLs
    try:
        with open(links_file, 'r', encoding='utf-8') as f:
            urls = [line.strip() for line in f if line.strip()]
    except FileNotFoundError:
        print(f"❌ ERROR: File '{links_file}' not found!")
        return
    
    if max_commands:
        urls = urls[:max_commands]
    
    print(f"✅ Found {len(urls)} URL(s) to process")
    
    # Check if file exists - ask what to do
    if os.path.exists(output_excel):
        print(f"\n⚠️  File '{output_excel}' already exists!")
        choice = input("(a)ppend new data or (o)verwrite file? [a/o]: ").strip().lower()
        
        if choice == 'o':
            print("🗑️  Creating new file...")
            wb, ws = create_excel_with_headers(output_excel)
            row_num = 2
        else:
            print("📂 Loading existing file...")
            wb = load_workbook(output_excel)
            ws = wb.active
            
            # Find next empty row
            row_num = 2
            while ws.cell(row=row_num, column=1).value:
                row_num += 1
            print(f"   Will append starting at row {row_num}")
    else:
        print("📄 Creating new file...")
        wb, ws = create_excel_with_headers(output_excel)
        row_num = 2
    
    print("🔧 Setting up Edge driver...")
    
    # Setup driver
    try:
        driver = setup_driver()
        print("✅ Edge driver ready\n")
    except Exception as e:
        print(f"❌ Driver setup failed: {e}")
        return
    
    # Scrape URLs
    success_count = 0
    failed_count = 0
    
    print("Scraping web pages...\n")
    
    try:
        for idx, url in enumerate(urls, 1):
            print(f"[{idx}/{len(urls)}] Processing...")
            
            command_name = extract_command_name_from_url(url)
            print(f"    Command: {command_name}")
            
            svg_content = extract_svg_from_page(driver, url)
            
            if svg_content:
                ws.cell(row=row_num, column=1, value=command_name)
                
                cell_b = ws.cell(row=row_num, column=2, value=svg_content)
                cell_b.alignment = Alignment(wrap_text=True, vertical='top')
                
                ws.row_dimensions[row_num].height = 100
                
                success_count += 1
                row_num += 1
            else:
                failed_count += 1
            
            print()
            time.sleep(0.5)
            
    finally:
        driver.quit()
        print("🔧 Edge driver closed")
    
    # Save
    try:
        wb.save(output_excel)
        print("-" * 70)
        print(f"✅ Successfully scraped: {success_count} command(s)")
        if failed_count > 0:
            print(f"⚠️  Failed to scrape:    {failed_count} command(s)")
        print(f"💾 Excel saved: {output_excel}")
        print("=" * 70)
        
        if success_count > 0:
            print("\n✅ NEXT STEP:")
            print("   Run: python simplified_svg.py")
            print("=" * 70)
    except Exception as e:
        print(f"\n❌ Save Error: {e}")


if __name__ == "__main__":
    print("\n" + "=" * 70)
    print("STEP 1: WEB SCRAPER")
    print("=" * 70)
    print("\nScrapes SVG diagrams from IBM docs")
    print("Creates Excel with proper headers")
    print("\n⚠️  Requires: msedgedriver.exe in same folder")
    print("=" * 70 + "\n")
    
    choice = input("Process ALL or TEST (first 3)? [all/test]: ").strip().lower()
    
    if choice == 'test':
        print("\n🧪 TEST MODE: First 3 commands\n")
        scrape_cics_commands(max_commands=3)
    else:
        print("\n🚀 FULL MODE: All commands\n")
        print("⚠️  This will take 10-20 minutes...")
        input("Press Enter to continue...")
        print()
        scrape_cics_commands()
    
    print("\nPress Enter to exit...")
    input()