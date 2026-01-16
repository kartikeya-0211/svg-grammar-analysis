"""
script2_unoptimzed_grammar.py
------------------------------------------------------------------------
STANDALONE MODE:
Run this file directly to process 'railroad_diagrams_complete.xlsx'.
It reads simplified SVGs from Column E and writes Grammar to Column F.
"""

import xml.dom.minidom
import re
import math
import sys
import os

# Try importing openpyxl for standalone mode
try:
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment
except ImportError:
    pass

INPUT_FILE = 'railroad_diagrams_complete.xlsx'

# ==========================================
# 1. Coordinate Mapper System (Small n)
# ==========================================
class NodeMapper:
    """
    Acts as the 'Lexer'. Groups geometric points into Nodes (n0, n1...).
    Uses tolerance to connect fragmented lines.
    """
    def __init__(self, tolerance=4.0):
        self.nodes = [] 
        self.counter = 0
        self.tolerance = tolerance

    def get_node_id(self, x, y):
        for node in self.nodes:
            dist = math.sqrt((node['x'] - x)**2 + (node['y'] - y)**2)
            if dist < self.tolerance:
                return node['id']
        
        # Generates n0, n1, n2...
        new_id = f"n{self.counter}"
        self.nodes.append({'x': x, 'y': y, 'id': new_id})
        self.counter += 1
        return new_id
    
    def get_all_nodes(self):
        return [n['id'] for n in self.nodes]

# ==========================================
# 2. SVG Path Parser
# ==========================================
def parse_path_segments(d_string):
    # Normalize command spacing
    d_string = re.sub(r'([a-zA-Z])', r' \1 ', d_string)
    tokens = d_string.replace(',', ' ').split()
    
    segments = []
    current_x, current_y = 0.0, 0.0
    
    i = 0
    while i < len(tokens):
        cmd = tokens[i]
        try:
            if cmd == 'M':
                current_x = float(tokens[i+1])
                current_y = float(tokens[i+2])
                i += 3
            elif cmd == 'L':
                next_x = float(tokens[i+1])
                next_y = float(tokens[i+2])
                segments.append((current_x, current_y, next_x, next_y))
                current_x, current_y = next_x, next_y
                i += 3
            elif cmd == 'H':
                next_x = float(tokens[i+1])
                segments.append((current_x, current_y, next_x, current_y))
                current_x = next_x
                i += 2
            elif cmd == 'V':
                next_y = float(tokens[i+1])
                segments.append((current_x, current_y, current_x, next_y))
                current_y = next_y
                i += 2
            elif cmd == 'C': # Cubic Bezier - approximate to end point
                next_x = float(tokens[i+5])
                next_y = float(tokens[i+6])
                segments.append((current_x, current_y, next_x, next_y))
                current_x, current_y = next_x, next_y
                i += 7
            elif cmd == 'Q': # Quadratic Bezier - approximate to end point
                next_x = float(tokens[i+3])
                next_y = float(tokens[i+4])
                segments.append((current_x, current_y, next_x, next_y))
                current_x, current_y = next_x, next_y
                i += 5
            elif cmd in ['Z', 'z']:
                i += 1
            else:
                i += 1
        except IndexError:
            break
            
    return segments

# ==========================================
# 3. Main Grammar Extraction Logic
# ==========================================
def natural_keys(text):
    """
    Helper to sort strings like n1, n2, n10 correctly.
    """
    def atoi(text):
        return int(text) if text.isdigit() else text
    return [atoi(c) for c in re.split(r'(\d+)', text)]

def extract_node_number(node_str):
    match = re.search(r'n(\d+)', node_str)
    return int(match.group(1)) if match else 0

def clean_svg(svg_content):
    """
    Basic cleanup only. 
    HTML entities are ALREADY handled by Script 1, so we do NOT unescape here.
    """
    if not svg_content: return ""
    # Remove any weird xmlns definitions that might duplicate
    svg_content = re.sub(r'xmlns:ns\d+=""[^""]+""', '', svg_content)
    return svg_content

def extract_grammar_from_svg(svg_content):
    """
    Parses SVG XML string and returns a LIST of grammar rules.
    """
    clean_content = clean_svg(svg_content)
    
    try:
        doc = xml.dom.minidom.parseString(clean_content)
    except Exception as e:
        # Fallback: wrap in root if it's missing (rare)
        try:
            doc = xml.dom.minidom.parseString(f"<root>{clean_content}</root>")
        except:
            return [f"Error parsing SVG: {str(e)}"]

    mapper = NodeMapper(tolerance=4.0)
    rules = []
    
    # 1. Parse Lines & Paths
    raw_segments = []
    for line in doc.getElementsByTagName('line'):
        try:
            x1 = float(line.getAttribute('x1'))
            y1 = float(line.getAttribute('y1'))
            x2 = float(line.getAttribute('x2'))
            y2 = float(line.getAttribute('y2'))
            is_default = 'default' in line.getAttribute('class').lower()
            raw_segments.append({'p1':(x1,y1), 'p2':(x2,y2), 'default': is_default})
        except ValueError: continue

    for path in doc.getElementsByTagName('path'):
        d = path.getAttribute('d')
        is_default = 'default' in path.getAttribute('class').lower()
        segs = parse_path_segments(d)
        for s in segs:
            raw_segments.append({'p1':(s[0], s[1]), 'p2':(s[2], s[3]), 'default': is_default})

    # 2. Parse Text/Rects (Terminals)
    terminals = []
    for text_node in doc.getElementsByTagName('text'):
        if text_node.firstChild and text_node.firstChild.nodeType == text_node.TEXT_NODE:
            content = text_node.firstChild.nodeValue.strip()
        else:
            content = ""
        if not content: continue
        
        # Heuristic for input/output points if rect is missing
        try:
            tx = float(text_node.getAttribute('x') or 0)
            ty = float(text_node.getAttribute('y') or 0)
            
            # Look for sibling rect
            sibling = text_node.nextSibling
            rect_found = False
            while sibling:
                if sibling.nodeType == 1 and sibling.tagName == 'rect':
                    rx = float(sibling.getAttribute('x'))
                    ry = float(sibling.getAttribute('y'))
                    rw = float(sibling.getAttribute('width'))
                    rh = float(sibling.getAttribute('height'))
                    mid_y = ry + (rh / 2.0)
                    input_pt = (rx, mid_y)
                    output_pt = (rx + rw, mid_y)
                    rect_found = True
                    break
                sibling = sibling.nextSibling
            
            if not rect_found:
                # Fallback based on text length
                approx_w = len(content) * 8 
                mid_y = ty - 5
                input_pt = (tx - 5, mid_y)
                output_pt = (tx + approx_w, mid_y)
            
            terminals.append({
                'text': content,
                'in': input_pt,
                'out': output_pt
            })
        except: continue

    # 3. Build Rules
    for seg in raw_segments:
        start_id = mapper.get_node_id(*seg['p1'])
        end_id = mapper.get_node_id(*seg['p2'])
        
        suffix = " (default)" if seg['default'] else ""
        if start_id != end_id:
            rules.append(f"{start_id} -> {end_id}{suffix}")

    # Add terminals
    for term in terminals:
        in_id = mapper.get_node_id(*term['in'])
        out_id = mapper.get_node_id(*term['out'])
        rules.append(f"{in_id} -> {term['text']} {out_id}")

    # 4. Add 'null' for end nodes
    all_nodes = mapper.get_all_nodes()
    all_nodes.sort(key=lambda x: extract_node_number(x))
    
    sources = set([r.split('->')[0].strip() for r in rules])
    for node in all_nodes:
        if node not in sources:
            rules.append(f"{node} -> null")

    # Sort rules naturally (n0, n1, n2...)
    rules.sort(key=lambda x: natural_keys(x.split('->')[0]))
    
    return rules

# ==========================================
# 4. COMPATIBILITY & STANDALONE RUNNER
# ==========================================
def convert_svg_to_grammar(svg_content):
    """Wrapper for Main Script."""
    rules_list = extract_grammar_from_svg(svg_content)
    if not rules_list: return ""
    return "\n".join(rules_list)

def main_standalone():
    print("=" * 70)
    print("      SCRIPT 2: STANDALONE MODE (SMALL N)")
    print("=" * 70)
    
    if not os.path.exists(INPUT_FILE):
        print(f"❌ {INPUT_FILE} not found.")
        return

    print(f"📂 Opening {INPUT_FILE}...")
    wb = load_workbook(INPUT_FILE)
    ws = wb.active
    
    # Ensure header
    ws.cell(row=1, column=6).value = "Right Regular Grammar (Unoptimized)"
    
    row = 2
    success = 0
    
    while True:
        cmd = ws.cell(row=row, column=1).value
        if not cmd: break
        
        svg = ws.cell(row=row, column=5).value
        print(f"Row {row} ({cmd}): ", end="")
        
        if not svg:
            print("⚠️ Skipped")
            row += 1
            continue
            
        try:
            grammar = convert_svg_to_grammar(svg)
            if grammar:
                ws.cell(row=row, column=6, value=grammar).alignment = Alignment(wrap_text=True, vertical='top')
                print(f"✅ Generated ({len(grammar.splitlines())} lines)")
                success += 1
            else:
                print("⚠️ No rules")
        except Exception as e:
            print(f"❌ Error: {e}")
            
        row += 1
        
    wb.save(INPUT_FILE)
    print("=" * 70)
    print(f"✅ Done! Processed {success} rows.")
    print("💾 Saved to Excel.")

if __name__ == "__main__":
    main_standalone()