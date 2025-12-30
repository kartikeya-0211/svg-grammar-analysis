"""
script1_simplified_svg.py
------------------------------------------------------------------------
1. Reads 'Original SVG' from Column B.
2. Formats it -> Column C.
3. Generates Image -> Column D (Conditionally Resized).
4. SIMPLIFIES -> Column E.
5. Generates Simplified Image -> Column F (Conditionally Resized).
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import Alignment
import xml.etree.ElementTree as ET
import os
import re
import time

# --- FILES ---
INPUT_FILE = 'railroad_diagrams.xlsx'
DRIVER_FILENAME = "msedgedriver.exe"


def setup_driver():
    cwd = os.getcwd()
    driver_path = os.path.join(cwd, DRIVER_FILENAME)
    if not os.path.exists(driver_path):
        print(f"❌ ERROR: '{DRIVER_FILENAME}' missing.")
        return None

    options = Options()
    options.use_chromium = True
    options.add_argument('--headless')
    options.add_argument("--log-level=3") 
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    options.add_argument('--disable-gpu')
    options.add_argument('--force-device-scale-factor=1')
    options.add_argument('--window-size=2000,2000') 
    
    service = Service(executable_path=driver_path)
    service.creation_flags = 0x08000000
    
    return webdriver.Edge(service=service, options=options)


def svg_to_image(driver, svg_code, output_path):
    if not svg_code: return False
    try:
        html = f'<html><body style="margin: 0; padding: 20px; background: white;">{svg_code}</body></html>'
        temp_html = os.path.abspath("temp_canvas.html")
        with open(temp_html, "w", encoding="utf-8") as f:
            f.write(html)
        driver.get(f"file:///{temp_html}")
        time.sleep(0.5) 
        try:
            driver.find_element(By.TAG_NAME, "svg").screenshot(output_path)
            return True
        except: return False
    except: return False


def simplify_railroad_svg(svg_string):
    ET.register_namespace('', 'http://www.w3.org/2000/svg')
    try:
        root = ET.fromstring(svg_string)
    except ET.ParseError:
        return svg_string

    # 1. Extract <defs> and Main Group
    defs_element = None
    main_group = None
    for child in root:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag == 'defs': defs_element = child
        elif tag == 'g' and 'transform' in child.attrib: main_group = child
    
    new_root = ET.Element('svg')
    for k, v in root.attrib.items(): new_root.set(k, v)
    if defs_element is not None: new_root.append(defs_element)

    # 2. Flatten
    flat_elements = []
    if main_group is not None:
        process_group_recursive(main_group, 0.0, 0.0, flat_elements)
    
    for elem in flat_elements: new_root.append(elem)

    # 3. Clean up
    remove_polygons(new_root)
    merge_text_nodes(new_root)
    round_all_coordinates(new_root)

    xml_str = ET.tostring(new_root, encoding='unicode')
    xml_str = re.sub(r'<ns\d+:', '<', xml_str)
    xml_str = re.sub(r'</ns\d+:', '</', xml_str)
    xml_str = re.sub(r'xmlns:ns\d+="[^"]*"', '', xml_str)
    
    return prettify_xml(xml_str)


def process_group_recursive(element, acc_x, acc_y, collector):
    tag = element.tag.split('}')[-1] if '}' in element.tag else element.tag
    curr_x = acc_x
    curr_y = acc_y
    
    if 'transform' in element.attrib:
        t_str = element.get('transform')
        match = re.search(r'translate\(\s*([-+]?\d*\.?\d+)\s*(?:[,\s]\s*([-+]?\d*\.?\d+))?\s*\)', t_str)
        if match:
            curr_x += float(match.group(1))
            curr_y += float(match.group(2)) if match.group(2) else 0.0
    
    if tag == 'g':
        for child in list(element):
            process_group_recursive(child, curr_x, curr_y, collector)
    else:
        import copy
        new_elem = copy.deepcopy(element)
        apply_offset(new_elem, curr_x, curr_y)
        collector.append(new_elem)


def apply_offset(element, dx, dy):
    tag = element.tag.split('}')[-1] if '}' in element.tag else element.tag
    if tag == 'path':
        d = element.get('d', '')
        def shift_coords(match):
            nums = [float(n) for n in re.findall(r'[-+]?\d*\.?\d+', match.group(0))]
            shifted = []
            for i in range(0, len(nums), 2):
                if i+1 < len(nums):
                    shifted.append(f"{round(nums[i] + dx, 2)},{round(nums[i+1] + dy, 2)}")
            return " ".join(shifted)
        element.set('d', re.sub(r'[-+]?\d*\.?\d+[\s,]+[-+]?\d*\.?\d+', shift_coords, d))
    elif tag == 'rect':
        element.set('x', str(round(float(element.get('x', '0')) + dx, 2)))
        element.set('y', str(round(float(element.get('y', '0')) + dy, 2)))
    elif tag == 'line':
        element.set('x1', str(round(float(element.get('x1', '0')) + dx, 2)))
        element.set('y1', str(round(float(element.get('y1', '0')) + dy, 2)))
        element.set('x2', str(round(float(element.get('x2', '0')) + dx, 2)))
        element.set('y2', str(round(float(element.get('y2', '0')) + dy, 2)))
    elif tag == 'text':
        element.set('x', str(round(float(element.get('x', '0')) + dx, 2)))
        element.set('y', str(round(float(element.get('y', '0')) + dy, 2)))
    
    if 'transform' in element.attrib: del element.attrib['transform']


def remove_polygons(root):
    for parent in root.iter():
        for child in list(parent):
            if 'polygon' in child.tag: parent.remove(child)

def merge_text_nodes(root):
    children = list(root)
    i = 0
    while i < len(children) - 1:
        curr, next_node = children[i], children[i+1]
        if 'text' in curr.tag and 'text' in next_node.tag:
            try:
                if abs(float(curr.get('y', 0)) - float(next_node.get('y', 0))) < 5:
                    curr.text = (curr.text or "") + (next_node.text or "")
                    root.remove(next_node)
                    children.remove(next_node)
                    continue
            except: pass
        i += 1

def round_all_coordinates(root):
    for elem in root.iter():
        for attr in ['x','y','x1','y1','x2','y2','width','height']:
            if attr in elem.attrib:
                try: elem.set(attr, str(round(float(elem.get(attr)), 2)))
                except: pass

def prettify_xml(xml_str):
    xml_str = re.sub(r'>\s+<', '><', xml_str)
    lines = []
    indent = 0
    for part in re.split(r'(<[^>]+>)', xml_str):
        if not part.strip(): continue
        if '</' in part: indent -= 1
        lines.append('  ' * indent + part)
        if '<' in part and '</' not in part and '/>' not in part and '<?' not in part: indent += 1
    return '\n'.join(lines)


def main():
    print("=" * 60)
    print("      SCRIPT 1: FLATTENER & SIMPLIFIER")
    print("=" * 60)

    if not os.path.exists(INPUT_FILE):
        print(f"❌ Error: {INPUT_FILE} not found.")
        return

    driver = setup_driver()
    if not driver: return

    print(f"📂 Opening {INPUT_FILE}...")
    wb = load_workbook(INPUT_FILE)
    ws = wb.active
    
    # 🧹 Clear old images to prevent duplication
    ws._images = []

    row = 2
    try:
        while True:
            raw_svg = ws.cell(row=row, column=2).value
            if not raw_svg: break
            
            print(f"Processing Row {row}...", end="")

            # C: Format Original
            ws.cell(row=row, column=3, value=prettify_xml(raw_svg)).alignment = Alignment(wrap_text=True, vertical='top')

            # D: Image Original (CONDITIONAL RESIZE)
            img1 = f"temp_orig_{row}.png"
            if svg_to_image(driver, raw_svg, img1):
                img = ExcelImage(img1)
                # If image is taller than 100px, shrink it. Otherwise, keep original size.
                if img.height > 100:
                    img.height = 100 
                ws.add_image(img, f"D{row}")

            # E: FLATTEN
            flat_svg = simplify_railroad_svg(raw_svg)
            ws.cell(row=row, column=5, value=flat_svg).alignment = Alignment(wrap_text=True, vertical='top')

            # F: Image Simplified (CONDITIONAL RESIZE)
            img2 = f"temp_simp_{row}.png"
            if svg_to_image(driver, flat_svg, img2):
                img = ExcelImage(img2)
                # If image is taller than 100px, shrink it. Otherwise, keep original size.
                if img.height > 100:
                    img.height = 100
                ws.add_image(img, f"F{row}")

            ws.row_dimensions[row].height = 120
            print(" ✅")
            row += 1

    finally:
        driver.quit()
        wb.save(INPUT_FILE)
        
        # Cleanup
        if os.path.exists("temp_canvas.html"): os.remove("temp_canvas.html")
        for f in os.listdir():
            if f.startswith("temp_") and f.endswith(".png"):
                try: os.remove(f)
                except: pass

        print("\n✅ Done.")

if __name__ == "__main__":
    main()