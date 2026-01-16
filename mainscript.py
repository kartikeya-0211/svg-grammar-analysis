"""
MAIN ORCHESTRATOR - INTEGRATED V3
=========================================================
1. FETCH (Calls Script 0) -> Memory
2. SCREENSHOT (Calls Script 1) -> Column D
3. SIMPLIFY (Calls Script 1) -> Column E
4. GRAMMAR (Calls Script 2) -> Column F

[A] Command  [B] URL
[C] Raw Code [D] Raw Image (Physically Resized)
[E] Simp Code
[F] Unoptimized Grammar

DEPENDENCIES:
- pip install openpyxl selenium pillow
"""

import os
import sys
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.drawing.image import Image as ExcelImage
from PIL import Image as PILImage

# --- IMPORTS ---
# We use try/except to give you clear errors if files are missing
try:
    import script1_simplified_svg as script1
    import script2_unoptimzed_grammar as script2
except ImportError as e:
    print(f"❌ CRITICAL IMPORT ERROR: {e}")
    print("Ensure 'script1_simplified_svg.py' and 'script2_unoptimzed_grammar.py' are in this folder.")
    sys.exit()

# Try importing Script 0 (Scraper). If missing, we define a dummy placeholder.
try:
    import script0_web_scraper as script0
except ImportError:
    print("⚠️  'script0_web_scraper.py' not found. Assuming you will provide logic or file.")
    script0 = None

# --- CONFIG ---
LINKS_FILE = 'links_cics.txt'
OUTPUT_EXCEL = 'railroad_diagrams_complete.xlsx'

def resize_img(path):
    """
    Physically resizes the image file on disk to height=100px using Pillow.
    """
    try:
        img = PILImage.open(path)
        w, h = img.size
        if h > 100:
            scale = 100/h
            img = img.resize((int(w*scale), 100), PILImage.Resampling.LANCZOS)
            img.save(path)
    except Exception as e: 
        print(f"  ⚠️ Resize warning: {e}")

def cleanup(path):
    if os.path.exists(path): 
        try: os.remove(path)
        except: pass

def main():
    print("=" * 70)
    print("🚀 MASTER PIPELINE: GRAMMAR EXTRACTION (MEMORY OPTIMIZED)")
    print("=" * 70)
    
    if not os.path.exists(LINKS_FILE): 
        print(f"❌ Error: '{LINKS_FILE}' missing. Please create it with URLs.")
        return

    # --- 1. USER INPUT ---
    # Determine how many links to process
    user_limit = input("Process how many commands? (Enter number or 'all'): ").strip().lower()
    
    with open(LINKS_FILE, 'r') as f: 
        all_urls = [l.strip() for l in f if l.strip()]
    
    if user_limit == 'all':
        urls = all_urls
        print(f"📋 Processing ALL {len(urls)} commands.")
    else:
        try:
            limit = int(user_limit)
            urls = all_urls[:limit]
            print(f"🧪 TEST MODE: Processing first {len(urls)} commands.")
        except ValueError:
            print("❌ Invalid input. Defaulting to 3.")
            urls = all_urls[:3]

    print("-" * 70)

    # --- 2. SETUP DRIVER ---
    # We use script1's setup because it's configured for the screenshots
    driver = script1.setup_driver()
    if not driver:
        print("❌ Failed to initialize WebDriver. Check 'msedgedriver.exe'.")
        return

    # --- 3. EXCEL SETUP ---
    wb = Workbook()
    ws = wb.active
    
    headers = ["Command", "URL", "Raw SVG Code", "Raw SVG Image", "Simplified SVG Code", "Unoptimized Grammar"]
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=i, value=h)
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal='center', vertical='center')
    
    # Adjust widths for readability
    widths = [20, 30, 30, 80, 30, 60]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[chr(64+i)].width = w
    
    temp_files = [] 
    row = 2
    
    try:
        for idx, url in enumerate(urls, 1):
            # Safe extraction of command name (fallback to 'Unknown' if logic fails)
            try:
                cmd = script0.extract_command_name_from_url(url) if script0 else f"CMD_{idx}"
            except: 
                cmd = f"CMD_{idx}"
                
            print(f"[{idx}/{len(urls)}] {cmd:<15} ...", end=" ")
            
            # ---------------------------------------------------------
            # STEP A: FETCH (Calls Script 0)
            # ---------------------------------------------------------
            raw_svg = ""
            if script0:
                try:
                    raw_svg = script0.extract_svg_from_page(driver, url)
                except Exception as e:
                    print(f"❌ Script0 Error: {e}")
            
            if not raw_svg:
                print("⚠️ No SVG found -> Skipped")
                ws.cell(row=row, column=1, value=cmd)
                ws.cell(row=row, column=2, value=url)
                ws.cell(row=row, column=3, value="[NO SVG DETECTED]")
                row += 1
                continue
            
            # ---------------------------------------------------------
            # STEP B: SCREENSHOT (Calls Script 1)
            # ---------------------------------------------------------
            img_raw_path = f"temp_raw_{row}.png"
            has_raw_img = script1.svg_to_image(driver, raw_svg, img_raw_path)
            if has_raw_img: 
                temp_files.append(img_raw_path)
            
            # ---------------------------------------------------------
            # STEP C: SIMPLIFY (Calls Script 1)
            # ---------------------------------------------------------
            # MEMORY FIX: We pass 'raw_svg' variable directly. No reading from Excel.
            simp_svg = script1.simplify_railroad_svg(raw_svg)
            
            # ---------------------------------------------------------
            # STEP D: GRAMMAR (Calls Script 2)
            # ---------------------------------------------------------
            # MEMORY FIX: We pass 'simp_svg' variable directly.
            try:
                grammar_rules = script2.extract_grammar_from_svg(simp_svg)
                grammar_text = "\n".join(grammar_rules) if grammar_rules else "No grammar found"
            except Exception as e:
                grammar_text = f"Error generating grammar: {str(e)}"

            # ---------------------------------------------------------
            # STEP E: WRITE TO EXCEL
            # ---------------------------------------------------------
            ws.cell(row=row, column=1, value=cmd).alignment = Alignment(vertical='top')
            ws.cell(row=row, column=2, value=url).alignment = Alignment(vertical='top')
            
            # TRUNCATION FIX: We slice [:32000] ONLY for writing to the sheet.
            # The logic above used the full strings.
            ws.cell(row=row, column=3, value=script1.prettify_xml(raw_svg)[:32000]).alignment = Alignment(wrap_text=True, vertical='top')
            
            # Insert Image
            if has_raw_img and os.path.exists(img_raw_path):
                resize_img(img_raw_path) # Helper function to shrink file size
                img = ExcelImage(img_raw_path)
                ws.add_image(img, f"D{row}")
            
            ws.cell(row=row, column=5, value=script1.prettify_xml(simp_svg)[:32000]).alignment = Alignment(wrap_text=True, vertical='top')
            ws.cell(row=row, column=6, value=grammar_text[:32000]).alignment = Alignment(wrap_text=True, vertical='top')
            
            # Formatting
            ws.row_dimensions[row].height = 120
            print("✅ Success")
            
            row += 1
            # Auto-save every 5 rows
            if idx % 5 == 0: wb.save(OUTPUT_EXCEL)
            
    except KeyboardInterrupt:
        print("\n🛑 Stopped by user.")
    except Exception as e:
        print(f"\n❌ Unexpected Error: {e}")
    finally:
        # Cleanup
        print("💾 Saving final Excel file...")
        wb.save(OUTPUT_EXCEL)
        driver.quit()
        
        # Remove temp files
        if os.path.exists("temp_canvas.html"): os.remove("temp_canvas.html")
        for f in temp_files: cleanup(f)
        
        print(f"✅ Finished. Output: {OUTPUT_EXCEL}")

if __name__ == "__main__":
    main()