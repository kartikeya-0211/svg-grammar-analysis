"""
MAIN ORCHESTRATOR - FINAL (HIGH DPI + COLORED EDGES)
=========================================================
1. FETCH (Script 0) -> Memory
2. RAW IMAGE (Script 1 File-Based) -> Column D
3. SIMPLIFY (Script 1) -> Column E (Code Only, No Image)
4. PARSE (Script 2 + Patch 85px) -> Column F (Textual categories)
5. VISUALIZE (Script 3 Logic) -> Column G (High DPI, Colored Edges, No Blur)

[A] Command  [B] URL
[C] Raw Code [D] Raw Image (Physically Resized)
[E] Simp Code
[F] Text Graph
[G] Visual Graph (High Res File, Visually Scaled in Excel)
"""

import os
import re
import time
import copy
import math
import xml.etree.ElementTree as ET
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.drawing.image import Image as ExcelImage
from bs4 import BeautifulSoup
import graphviz
from PIL import Image as PILImage

# External Dependencies
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options

# --- CONFIG ---
LINKS_FILE = 'links_cics.txt'
OUTPUT_EXCEL = 'railroad_diagrams_complete.xlsx'
DRIVER_FILENAME = 'msedgedriver.exe'
GRAPHS_DIR = 'graphs_output'

if not os.path.exists(GRAPHS_DIR):
    os.makedirs(GRAPHS_DIR)

# ============================================================================
# DRIVER SETUP
# ============================================================================
def setup_driver():
    edge_options = Options()
    edge_options.use_chromium = True
    edge_options.add_argument('--headless')
    edge_options.add_argument('--enable-unsafe-swiftshader')
    edge_options.add_argument("--log-level=3")
    edge_options.add_experimental_option('excludeSwitches', ['enable-logging'])
    edge_options.add_argument('--disable-gpu')
    edge_options.add_argument('--force-device-scale-factor=1')
    edge_options.add_argument('--window-size=2500,2500')
    
    driver_path = os.path.join(os.getcwd(), DRIVER_FILENAME)
    if not os.path.exists(driver_path):
        raise FileNotFoundError(f"{DRIVER_FILENAME} not found.")
    
    service = Service(driver_path)
    service.creation_flags = 0x08000000
    return webdriver.Edge(service=service, options=edge_options)

# ============================================================================
# STAGE 1: FETCH & RAW IMAGE
# ============================================================================
def extract_cmd_name(url):
    filename = url.split('/')[-1].replace('.html', '')
    command = filename.replace('dfhp4_', '').replace('dfhp4-', '')
    return re.sub(r'([a-z])([A-Z])', r'\1 \2', command).upper()

def scrape_raw_svg(driver, url):
    try:
        driver.get(url)
        try:
            svg = WebDriverWait(driver, 6).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'svg.syntaxdiagram'))
            )
        except:
            svg = WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.TAG_NAME, 'svg'))
            )
        return svg.get_attribute('outerHTML')
    except: return None

def svg_to_image_file(driver, svg_code, output_path):
    """File-based screenshot method (Matches Script 1)."""
    if not svg_code: return False
    try:
        html = f'<html><body style="margin: 0; padding: 20px; background: white;">{svg_code}</body></html>'
        temp = os.path.abspath("temp_canvas.html")
        with open(temp, "w", encoding="utf-8") as f: f.write(html)
        
        driver.get(f"file:///{temp}")
        time.sleep(0.5)
        
        driver.find_element(By.TAG_NAME, "svg").screenshot(output_path)
        return True
    except: return False

# ============================================================================
# STAGE 2: SIMPLIFY (Memory Only - Code Saved to Col E)
# ============================================================================
def prettify_xml(xml_str):
    xml_str = re.sub(r'>\s+<', '><', xml_str)
    lines, indent = [], 0
    for part in re.split(r'(<[^>]+>)', xml_str):
        if not part.strip(): continue
        if '</' in part: indent -= 1
        lines.append('  ' * indent + part)
        if '<' in part and '</' not in part and '/>' not in part and '<?' not in part: indent += 1
    return '\n'.join(lines)

def simplify_svg_memory(svg_string):
    try:
        soup = BeautifulSoup(svg_string, "xml") 
        clean = str(soup).replace('&nbsp;', ' ')
        clean = re.sub(r'(xmlns:?\w*="[^"]+")', '', clean)
        clean = re.sub(r'\w+:', '', clean)
        root = ET.fromstring(clean)
    except: return svg_string

    defs, main = None, None
    for child in root:
        if child.tag == 'defs': defs = child
        elif child.tag == 'g' and 'transform' in child.attrib: main = child
    
    new = ET.Element('svg')
    for k, v in root.attrib.items(): new.set(k, v)
    if defs: new.append(defs)

    flat = []
    if main: _process_group(main, 0.0, 0.0, flat)
    for elem in flat: new.append(elem)

    _clean_nodes(new)
    return prettify_xml(ET.tostring(new, encoding='unicode'))

def _process_group(elem, ax, ay, col):
    cx, cy = ax, ay
    if 'transform' in elem.attrib:
        m = re.search(r'translate\(\s*([-+]?\d*\.?\d+)\s*(?:[,\s]\s*([-+]?\d*\.?\d+))?\s*\)', elem.get('transform'))
        if m:
            cx += float(m.group(1))
            cy += float(m.group(2)) if m.group(2) else 0.0
    
    if elem.tag == 'g':
        for child in list(elem): _process_group(child, cx, cy, col)
    else:
        new = copy.deepcopy(elem)
        _offset(new, cx, cy)
        col.append(new)

def _offset(elem, dx, dy):
    tag = elem.tag
    if tag == 'path':
        d = elem.get('d', '')
        def shift(m):
            nums = [float(n) for n in re.findall(r'[-+]?\d*\.?\d+', m.group(0))]
            s = []
            for i in range(0, len(nums), 2):
                if i+1 < len(nums): s.append(f"{round(nums[i]+dx,2)},{round(nums[i+1]+dy,2)}")
            return " ".join(s)
        elem.set('d', re.sub(r'[-+]?\d*\.?\d+[\s,]+[-+]?\d*\.?\d+', shift, d))
    elif tag in ['rect', 'text']:
        elem.set('x', str(round(float(elem.get('x', '0')) + dx, 2)))
        elem.set('y', str(round(float(elem.get('y', '0')) + dy, 2)))
    elif tag == 'line':
        elem.set('x1', str(round(float(elem.get('x1', '0')) + dx, 2)))
        elem.set('y1', str(round(float(elem.get('y1', '0')) + dy, 2)))
        elem.set('x2', str(round(float(elem.get('x2', '0')) + dx, 2)))
        elem.set('y2', str(round(float(elem.get('y2', '0')) + dy, 2)))
    if 'transform' in elem.attrib: del elem.attrib['transform']

def _clean_nodes(root):
    for p in root.iter():
        for c in list(p):
            if c.tag == 'polygon': p.remove(c)
    
    children = list(root)
    i = 0
    while i < len(children) - 1:
        c, n = children[i], children[i+1]
        if c.tag == 'text' and n.tag == 'text':
            try:
                if abs(float(c.get('y', 0)) - float(n.get('y', 0))) < 5:
                    c.text = (c.text or "") + (n.text or "")
                    root.remove(n)
                    children.remove(n)
                    continue
            except: pass
        i += 1

# ============================================================================
# STAGE 3: PARSE (Script 2 + Patch)
# ============================================================================
class Node:
    def __init__(self, id, text, x, y, width=0, is_rect=False):
        self.id, self.text = id, text
        self.x, self.y = float(x), float(y)
        self.width, self.is_rect = float(width), is_rect
        self.left, self.right = self.x, self.x + (self.width if width > 0 else 30)
    def __repr__(self): return f"{self.id}({self.text})"

class Edge:
    def __init__(self, start, end, type_):
        self.start, self.end, self.type = start, end, type_
    def __eq__(self, o): return (self.start, self.end, self.type) == (o.start, o.end, o.type)
    def __hash__(self): return hash((self.start, self.end, self.type))
    def __lt__(self, o): return (self.start, self.end) < (o.start, o.end)
    def __repr__(self): return f"{self.start}->{self.end} ({self.type})"

def parse_graph(svg_string):
    if not svg_string: return [], []
    try: root = ET.fromstring(svg_string)
    except: return [], []

    nodes = _get_nodes(root)
    if not nodes: return [], []
    
    ys = [int(n.y) for n in nodes]
    main_y = max(set(ys), key=ys.count) if ys else nodes[0].y
    
    edges = []
    paths = [e for e in root.iter() if e.tag in ['path', 'line']]

    for p in paths:
        pts = []
        if p.tag == 'path':
            d = p.get('d', '')
            ns = [float(n) for n in re.findall(r'[-+]?\d*\.?\d+', d)]
            for k in range(0, len(ns), 2):
                if k+1 < len(ns): pts.append((ns[k], ns[k+1]))
        elif p.tag == 'line':
            try: pts = [(float(p.get('x1')), float(p.get('y1'))), (float(p.get('x2')), float(p.get('y2')))]
            except: continue
            
        if len(pts) < 2: continue

        start, end = pts[0], pts[-1]
        e_type = _classify_edge(start, end, pts, main_y)
        src = _closest_node(start[0], start[1], nodes, True, main_y)
        dst = _closest_node(end[0], end[1], nodes, False, main_y)
        
        sid = src.id if src else "START"
        did = dst.id if dst else "END"
        
        if sid != did or e_type == "Loopback":
            edges.append(Edge(sid, did, e_type))

    return nodes, sorted(list(set(edges)))

def _get_nodes(root):
    rects = []
    for r in root.findall('.//rect'):
        try: rects.append({'x': float(r.get('x')), 'y': float(r.get('y')), 'w': float(r.get('width')), 'h': float(r.get('height', 20))})
        except: pass
    
    valid = [r for r in rects if r['w'] < 2000 and r['h'] < 85]
    valid.sort(key=lambda r: (r['y'], r['x']))

    texts = []
    for t in root.iter():
        if t.tag == 'text':
            txt = (t.text or "").strip()
            if txt: texts.append({'txt': txt, 'x': float(t.get('x',0)), 'y': float(t.get('y',0)), 'obj': t})

    nodes = []
    consumed = []
    for r in valid:
        content = [t for t in texts if (r['x']-5 <= t['x'] <= r['x']+r['w']+10) and (r['y']-10 <= t['y'] <= r['y']+r['h']+15)]
        if content:
            content.sort(key=lambda k: k['x'])
            full = " ".join([tc['txt'] for tc in content]).replace(" (", "(").replace("( ", "(").replace(" )", ")")
            nodes.append(Node("", full, r['x'], r['y'], r['w'], True))
            for t in content: consumed.append(t['obj'])

    bare = [t for t in texts if t['obj'] not in consumed]
    bare.sort(key=lambda t: t['x'])
    if bare:
        buf, bx, by = [bare[0]['txt']], bare[0]['x'], bare[0]['y']
        last = bare[0]['x'] + (len(bare[0]['txt']) * 8)
        for i in range(1, len(bare)):
            c = bare[i]
            if abs(c['y']-by) < 5 and (c['x']-last) < 15:
                buf.append(c['txt'])
                last = c['x'] + (len(c['txt']) * 8)
            else:
                nodes.append(Node("", "".join(buf), bx, by, last-bx))
                buf, bx, by = [c['txt']], c['x'], c['y']
                last = c['x'] + (len(c['txt']) * 8)
        if buf: nodes.append(Node("", "".join(buf), bx, by, last-bx))

    clean = []
    idx = 1
    nodes.sort(key=lambda n: (n.y, n.x))
    for n in nodes:
        if re.search(r'[a-zA-Z0-9]', n.text):
            n.id = f"n{idx}"
            clean.append(n)
            idx += 1
    return clean

def _closest_node(x, y, nodes, is_src, main_y):
    if not nodes: return None
    best, min_d = None, float('inf')
    for n in nodes:
        tx = n.right if is_src else n.left
        ty = n.y + (10 if n.is_rect else 0)
        d = math.sqrt((x-tx)**2 + (y-ty)**2)
        if d < min_d: min_d, best = d, n
    return best if min_d < 300 else None

def _classify_edge(start, end, pts, main_y):
    if end[0] < (start[0] - 20): return "Loopback"
    ys = [p[1] for p in pts]
    if max(ys) > (main_y + 25): return "Alternative"
    if min(ys) < (main_y - 25): return "Default"
    return "Mainline"

def format_text(nodes, edges):
    out = "Nodes:\n"
    for n in nodes: out += f"  {n}\n"
    for lbl, t in [("Mainline edges", "Mainline"), ("Default edges", "Default"), ("Alternative edges", "Alternative"), ("Loopback edges", "Loopback")]:
        grp = [str(e) for e in edges if e.type == t]
        out += f"\n{lbl}:\n"
        if grp: out += "  " + "\n  ".join(grp) + "\n"
    return out

# ============================================================================
# STAGE 4: VISUALIZE (HIGH DPI + COLORED EDGES)
# ============================================================================
def generate_demo_style_graph(cmd, nodes, edges):
    """
    Generates graph at HIGH DPI (300) with color-coded edges.
    Does NOT shrink the file.
    """
    try:
        # dpi='300' ensures the raw file is crisp
        dot = graphviz.Digraph(comment=cmd, format='png', strict=True)
        dot.attr(rankdir='LR', dpi='300') 
        
        # Node styling
        dot.attr('node', shape='box', style='rounded,filled', fillcolor='#f9f9f9', fontname='Arial', fontsize='10')

        # Add Nodes
        for n in nodes:
            if n.text in ["START", "END"]:
                 dot.node(n.id, n.text, shape='point', width='0', style='invis')
            else:
                 dot.node(n.id, n.text)

        # Add Edges with Colors based on type
        # Color Legend:
        # Mainline: Dark Gray (Standard)
        # Default: Blue (Automatic path)
        # Alternative: Orange (Optional branch)
        # Loopback: Red (Repeating path)
        edge_colors = {
            'Mainline': '#333333',    
            'Default': '#0066CC',     
            'Alternative': '#FF6600', 
            'Loopback': '#CC0000'     
        }

        for e in edges:
            # Get color, default to mainline gray if type somehow fails
            color = edge_colors.get(e.type, '#333333')
            # Add edge with slightly thicker penwidth for visibility
            dot.edge(e.start, e.end, color=color, penwidth='2.0')

        clean = re.sub(r'[\\/*?:"<>|]', "_", cmd)
        path = os.path.join(GRAPHS_DIR, clean)
        dot.render(path, cleanup=True)
        return path + ".png"
    except: return None

def resize_img(path):
    """
    PHYSICAL RESIZE: Only used for RAW screenshot (Col D).
    Does NOT use this for Visual Graph (Col G).
    """
    try:
        img = PILImage.open(path)
        w, h = img.size
        if h > 100:
            scale = 100/h
            img = img.resize((int(w*scale), 100), PILImage.Resampling.LANCZOS)
            img.save(path)
    except: pass

def cleanup(path):
    if os.path.exists(path): 
        try: os.remove(path)
        except: pass

# ============================================================================
# MAIN
# ============================================================================
def main():
    print("=" * 70)
    print("🚀 MASTER PIPELINE: FINAL (HIGH DPI + COLORED EDGES)")
    print("=" * 70)
    
    if not os.path.exists(LINKS_FILE): return
    with open(LINKS_FILE, 'r') as f: urls = [l.strip() for l in f if l.strip()]
    
    driver = setup_driver()
    wb = Workbook()
    ws = wb.active
    
    # Headers
    headers = ["Command", "URL", "Raw SVG Code", "Raw SVG Image", "Simplified SVG Code", "Textual Graph", "Visual Graph"]
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=i, value=h)
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal='center', vertical='center')
    
    # Widths
    ws.column_dimensions['A'].width = 22
    ws.column_dimensions['B'].width = 35
    ws.column_dimensions['C'].width = 35
    ws.column_dimensions['D'].width = 90
    ws.column_dimensions['E'].width = 35
    ws.column_dimensions['F'].width = 50
    ws.column_dimensions['G'].width = 90
    
    temp_files = [] 
    row = 2
    
    try:
        for idx, url in enumerate(urls, 1):
            cmd = extract_cmd_name(url)
            print(f"[{idx}] {cmd}...", end=" ")
            
            # 1. Fetch
            raw_svg = scrape_raw_svg(driver, url)
            if not raw_svg:
                print("❌ Failed")
                ws.cell(row=row, column=1, value=cmd)
                row += 1
                continue
            
            # 2. Raw Image
            img_raw = f"temp_raw_{row}.png"
            has_raw = svg_to_image_file(driver, raw_svg, img_raw)
            if has_raw: temp_files.append(img_raw)
            
            # 3. Simplify (Memory)
            simp_svg = simplify_svg_memory(raw_svg)
            
            # 4. Parse
            nodes, edges = parse_graph(simp_svg)
            txt_graph = format_text(nodes, edges) if nodes else "No nodes"
            
            # 5. Visual Graph (Generates High DPI, Colored Edge file on disk)
            viz_path = generate_demo_style_graph(cmd, nodes, edges) if nodes else None
            
            # --- WRITE ---
            ws.cell(row=row, column=1, value=cmd)
            ws.cell(row=row, column=2, value=url)
            ws.cell(row=row, column=3, value=prettify_xml(raw_svg)[:32000]).alignment = Alignment(wrap_text=True, vertical='top')
            ws.cell(row=row, column=5, value=prettify_xml(simp_svg)[:32000]).alignment = Alignment(wrap_text=True, vertical='top')
            ws.cell(row=row, column=6, value=txt_graph).alignment = Alignment(wrap_text=True, vertical='top')
            
            # COL D: Physical Resize (Save disk space for raw screenshots)
            if has_raw and os.path.exists(img_raw):
                resize_img(img_raw)
                ws.add_image(ExcelImage(img_raw), f"D{row}")
                
            # COL G: VISUAL SCALE ONLY (Preserve High Def File)
            if viz_path and os.path.exists(viz_path):
                # We do NOT call resize_img() on viz_path.
                # The file on disk stays HUGE (300 DPI).
                
                img = ExcelImage(viz_path)
                
                # We tell Excel to DISPLAY it small (Thumbnail size)
                # Row height is 120, so target_h=100 fits nicely inside.
                curr_w, curr_h = img.width, img.height
                target_h = 100 
                if curr_h > 0:
                    scale = target_h / curr_h
                    img.height = target_h
                    img.width = curr_w * scale
                
                ws.add_image(img, f"G{row}")

            ws.row_dimensions[row].height = 120
            print("✅ Done")
            row += 1
            if idx % 5 == 0: wb.save(OUTPUT_EXCEL)
            
    finally:
        driver.quit()
        wb.save(OUTPUT_EXCEL)
        cleanup("temp_canvas.html")
        for f in temp_files: cleanup(f)
        print(f"\n✅ Finished! File: {OUTPUT_EXCEL}")

if __name__ == "__main__":
    main()