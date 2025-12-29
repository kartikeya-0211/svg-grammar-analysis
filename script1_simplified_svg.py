"""
script1_simplified_svg.py
------------------------------------------------------------------------
1. Reads 'Original SVG' from Column B.
2. Formats it -> Column C.
3. Generates Image of Original -> Column D.
4. SIMPLIFIES -> Column E:
   - FLATTENS TRANSFORMS (Removes <g translate> and adds values to children)
   - Removes polygons
   - Merges text
   - Rounds to 2 decimals
5. Generates Image of Simplified -> Column F.
6. DELETES temporary images.
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


# 1. SETUP BROWSER
def setup_driver():
    cwd = os.getcwd()
    driver_path = os.path.join(cwd, DRIVER_FILENAME)
    if not os.path.exists(driver_path):
        print(f"❌ ERROR: '{DRIVER_FILENAME}' missing.")
        return None

    options = Options()
    options.use_chromium = True
    options.add_argument('--headless')
    
    # --- ADD THESE LINES TO SILENCE ERRORS ---
    options.add_argument("--log-level=3") # Fatal errors only
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    # -----------------------------------------

    options.add_argument('--disable-gpu')
    options.add_argument('--force-device-scale-factor=1')
    options.add_argument('--window-size=2000,2000') 
    
    service = Service(executable_path=driver_path)
    
    # --- OPTIONAL: KEEPS CONSOLE CLEAN ---
    service.creation_flags = 0x08000000
    # -------------------------------------
    
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


# 2. FLATTENING & SIMPLIFICATION LOGIC
# ==================================================
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
        if tag == 'defs':
            defs_element = child
        elif tag == 'g' and 'transform' in child.attrib:
            main_group = child
    
    # Create a fresh root for the output
    new_root = ET.Element('svg')
    # Copy attributes from original root (width, height, viewBox, etc.)
    for k, v in root.attrib.items():
        new_root.set(k, v)

    if defs_element is not None:
        new_root.append(defs_element)

    # 2. FLATTEN: Recursively process the main group
    # This list will hold all the lines, paths, rects, text with UPDATED coords
    flat_elements = []
    
    if main_group is not None:
        # Start recursion with 0,0 offset
        process_group_recursive(main_group, 0.0, 0.0, flat_elements)
    
    # 3. Add flattened elements to new root
    for elem in flat_elements:
        new_root.append(elem)

    # 4. Post-Process: Merge Text, Remove Polygons, Round Coords
    remove_polygons(new_root)
    merge_text_nodes(new_root)
    round_all_coordinates(new_root)

    # Output Formatting
    xml_str = ET.tostring(new_root, encoding='unicode')
    xml_str = re.sub(r'<ns\d+:', '<', xml_str)
    xml_str = re.sub(r'</ns\d+:', '</', xml_str)
    xml_str = re.sub(r'xmlns:ns\d+="[^"]*"', '', xml_str)
    
    return prettify_xml(xml_str)


def process_group_recursive(element, acc_x, acc_y, collector):
    """
    Recursively drills down into <g> tags.
    - If it finds a <g transform="translate(x,y)">, it adds x,y to accumulator.
    - If it finds a shape/text, it applies the accumulated x,y to the shape's coords and saves it.
    """
    tag = element.tag.split('}')[-1] if '}' in element.tag else element.tag

    # Calculate new offset if this element is a group with transform
    curr_x = acc_x
    curr_y = acc_y
    
    if 'transform' in element.attrib:
        # Parse "translate(10, 20)" or "translate(10)"
        t_str = element.get('transform')
        match = re.search(r'translate\(\s*([-+]?\d*\.?\d+)\s*(?:[,\s]\s*([-+]?\d*\.?\d+))?\s*\)', t_str)
        if match:
            dx = float(match.group(1))
            dy = float(match.group(2)) if match.group(2) else 0.0
            curr_x += dx
            curr_y += dy
    
    # If it's a Group, process its children
    if tag == 'g':
        # Preserve class attribute (like class="groupseq") if needed, 
        # but we are stripping the structure, so usually we just want contents.
        # Exception: "text" groups (wrappers for keyword/var). We might want to keep the wrapper 
        # but move it? Actually, easier to flatten content OF the group.
        
        # NOTE: For text merging to work later, we usually need the <g class="text"> wrapper.
        # But Vini asked to remove <g translate>. 
        # Strategy: If it's a structural group, peel it. If it's a text wrapper, keep it but remove transform.
        
        is_text_wrapper = (element.get('class') == 'text') or (element.get('class') == 'boxed syntaxkwd') # heuristic
        
        # Iterate children
        for child in list(element):
            process_group_recursive(child, curr_x, curr_y, collector)
            
    else:
        # It is a Leaf Node (path, rect, line, text, polygon)
        # We need to CLONE it so we don't mess up the iterator or original references
        import copy
        new_elem = copy.deepcopy(element)
        
        # Apply the accumulated transform to this element's coordinates
        apply_offset(new_elem, curr_x, curr_y)
        
        # Add to collection
        collector.append(new_elem)


def apply_offset(element, dx, dy):
    """Adds dx, dy to the specific coordinate attributes of the element."""
    tag = element.tag.split('}')[-1] if '}' in element.tag else element.tag
    
    # 1. PATH (The complex one)
    if tag == 'path':
        d = element.get('d', '')
        # Regex to find all coordinate pairs
        # We look for number pairs. This handles M, L, Q, C, etc. assuming absolute coordinates.
        def shift_coords(match):
            # match.group(0) is like "10 20" or "10,20"
            nums = [float(n) for n in re.findall(r'[-+]?\d*\.?\d+', match.group(0))]
            # Shift pairs
            shifted = []
            for i in range(0, len(nums), 2):
                if i+1 < len(nums):
                    sx = nums[i] + dx
                    sy = nums[i+1] + dy
                    shifted.append(f"{round(sx, 2)},{round(sy, 2)}")
            return " ".join(shifted)

        # Replace sequences of coordinates in the d-string
        # This regex looks for patterns that look like coordinate pairs
        new_d = re.sub(r'[-+]?\d*\.?\d+[\s,]+[-+]?\d*\.?\d+', shift_coords, d)
        element.set('d', new_d)

    # 2. RECT
    elif tag == 'rect':
        x = float(element.get('x', '0'))
        y = float(element.get('y', '0'))
        element.set('x', str(round(x + dx, 2)))
        element.set('y', str(round(y + dy, 2)))
    
    # 3. LINE
    elif tag == 'line':
        for attr_x in ['x1', 'x2']:
            val = float(element.get(attr_x, '0'))
            element.set(attr_x, str(round(val + dx, 2)))
        for attr_y in ['y1', 'y2']:
            val = float(element.get(attr_y, '0'))
            element.set(attr_y, str(round(val + dy, 2)))
            
    # 4. TEXT
    elif tag == 'text':
        # Text sometimes has x,y. If not, it defaults to 0,0 relative to parent.
        x = float(element.get('x', '0'))
        y = float(element.get('y', '0'))
        element.set('x', str(round(x + dx, 2)))
        element.set('y', str(round(y + dy, 2)))

    # Remove the 'transform' attribute if it exists on the leaf itself 
    # (We already processed the group transforms, but sometimes leaves have them too)
    if 'transform' in element.attrib:
        del element.attrib['transform']


def remove_polygons(root):
    for parent in root.iter():
        for child in list(parent):
            if 'polygon' in child.tag:
                parent.remove(child)

def merge_text_nodes(root):
    # Since we flattened everything, text nodes might just be siblings now.
    # We look for adjacent text nodes and merge them.
    # Note: Logic is simpler now that groups are gone.
    children = list(root)
    i = 0
    while i < len(children) - 1:
        curr = children[i]
        next_node = children[i+1]
        
        # Check if both are text
        if 'text' in curr.tag and 'text' in next_node.tag:
            # Simple heuristic: if they are close in Y position (same line)
            try:
                y1 = float(curr.get('y', 0))
                y2 = float(next_node.get('y', 0))
                if abs(y1 - y2) < 5: # Threshold for "same line"
                    curr.text = (curr.text or "") + (next_node.text or "")
                    root.remove(next_node)
                    children.remove(next_node) # Update local list
                    continue
            except: pass
        i += 1

def round_all_coordinates(root):
    for elem in root.iter():
        for attr in ['x','y','x1','y1','x2','y2','width','height','rx','ry']:
            if attr in elem.attrib:
                try:
                    val = float(elem.get(attr))
                    elem.set(attr, str(round(val, 2)))
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


# 3. MAIN EXECUTION
# ==========================================
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

    row = 2
    count = 0
    
    try:
        while True:
            raw_svg = ws.cell(row=row, column=2).value
            if not raw_svg: break
            
            print(f"Processing Row {row}...", end="")

            # C: Format Original
            ws.cell(row=row, column=3, value=prettify_xml(raw_svg)).alignment = Alignment(wrap_text=True, vertical='top')

            # D: Image Original
            img1 = f"temp_orig_{row}.png"
            if svg_to_image(driver, raw_svg, img1):
                ws.add_image(ExcelImage(img1), f"D{row}")

            # E: FLATTEN & SIMPLIFY
            flat_svg = simplify_railroad_svg(raw_svg)
            ws.cell(row=row, column=5, value=flat_svg).alignment = Alignment(wrap_text=True, vertical='top')

            # F: Image Simplified
            img2 = f"temp_simp_{row}.png"
            if svg_to_image(driver, flat_svg, img2):
                ws.add_image(ExcelImage(img2), f"F{row}")

            ws.row_dimensions[row].height = 120
            print(" ✅")
            row += 1
            count += 1

    finally:
        driver.quit()
        wb.save(INPUT_FILE)
        
        # Cleanup
        if os.path.exists("temp_canvas.html"): os.remove("temp_canvas.html")
        for f in os.listdir():
            if f.startswith("temp_") and f.endswith(".png"):
                try: os.remove(f)
                except: pass

        print(f"\n✅ Done. Processed {count} rows.")

if __name__ == "__main__":
    main()