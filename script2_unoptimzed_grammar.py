import os
import xml.dom.minidom
import openpyxl
from openpyxl.styles import Alignment

# A tiny helper function to safely pull numbers out of the SVG tags.
# If an attribute (like 'x' or 'y') is missing, it returns 0.0 to prevent crashes.
def get_flt(elem, attr):
    val = elem.getAttribute(attr)
    return float(val) if val else 0.0

def generate_grammar_from_svg(svg_path):
    doc = xml.dom.minidom.parse(svg_path)
    
    # 1. FIND THE RAILROAD TRACKS (Mainlines)
    # We look at all horizontal lines where the start Y equals the end Y.
    # We save these Y-coordinates because they tell us the exact vertical level of the main flow.
    mainlines = []
    for line in doc.getElementsByTagName('line'):
        y1 = get_flt(line, 'y1')
        y2 = get_flt(line, 'y2')
        if y1 == y2 and y1 not in mainlines:
            mainlines.append(y1)
    
    # We sort them. If the diagram has no horizontal lines, we default to 0.0.
    mainlines.sort()
    if not mainlines:
        mainlines = [0.0]

    # 2. CREATE LOGICAL NODES (Bounding Boxes)
    # The SVG draws a <rect> around every logical text group.
    # We extract these boxes so we know exactly where the group's borders are.
    nodes = []
    for rect in doc.getElementsByTagName('rect'):
        x = get_flt(rect, 'x')
        y = get_flt(rect, 'y')
        w = get_flt(rect, 'width')
        h = get_flt(rect, 'height')
        nodes.append({'x': x, 'y': y, 'w': w, 'h': h, 'texts': []})

    # 3. FILL THE BOXES WITH TEXT
    # We grab every piece of text and check its coordinates.
    # If the text sits inside one of our bounding boxes, we assign it to that specific Node.
    for t in doc.getElementsByTagName('text'):
        if t.firstChild and t.firstChild.nodeType == t.TEXT_NODE:
            tx = get_flt(t, 'x')
            ty = get_flt(t, 'y')
            text_val = t.firstChild.nodeValue.strip()
            
            # We add a 5-pixel padding just in case the text spills slightly over the box edge.
            for node in nodes:
                if node['x'] - 5 <= tx <= node['x'] + node['w'] + 5 and node['y'] - 5 <= ty <= node['y'] + node['h'] + 5:
                    node['texts'].append({'x': tx, 'val': text_val})
                    break

    # 4. CLEAN UP THE NODES
    # We merge the text chunks (e.g., "ABCODE(", "name", ")") into one solid string.
    # We assign this completed Node to the closest horizontal Mainline (Row).
    valid_nodes = []
    for node in nodes:
        if node['texts']:
            node['texts'].sort(key=lambda k: k['x'])
            node['text_str'] = "".join([t['val'] for t in node['texts']])
            
            closest_main = min(mainlines, key=lambda m: abs(m - node['y']))
            node['row_y'] = closest_main
            valid_nodes.append(node)

    # Our safety net. If no valid boxes exist, we just output null.
    if not valid_nodes:
        return "n0 -> null"

    # 5. REBUILD THE GRAPH 
    # We sort primarily by the assigned Mainline Row, then horizontally by Column (X-axis).
    valid_nodes.sort(key=lambda k: (k['row_y'], k['x']))
    grammar_lines = []
    n_count = 0

    # 6. GROUP INTO ROWS AND COLUMNS
    # We group the nodes by their assigned Row.
    rows_dict = {}
    for node in valid_nodes:
        rows_dict.setdefault(node['row_y'], []).append(node)

    for row_y, row_nodes in rows_dict.items():
        cols = []
        current_col = []
        max_x = -1
        
        # We process the Columns. If nodes physically overlap on the X-axis, they are Parallel Alternatives.
        # If they don't overlap, they are Sequential Steps (a new column).
        for node in row_nodes:
            if not current_col or node['x'] <= max_x + 5:
                current_col.append(node)
                max_x = max(max_x, node['x'] + node['w'])
            else:
                cols.append(current_col)
                current_col = [node]
                max_x = node['x'] + node['w']
        if current_col:
            cols.append(current_col)

        # 7. GENERATE THE GRAMMAR
        for col in cols:
            current_n = f"n{n_count}"
            next_n = f"n{n_count + 1}"
            
            # We check if ANY box in this column is physically touched by the mainline.
            # If nothing touches the mainline, the mainline is an empty bypass path!
            intersects_main = any(node['y'] - 2 <= row_y <= node['y'] + node['h'] + 2 for node in col)
            
            if not intersects_main:
                grammar_lines.append(f"{current_n} -> {next_n}")
                
            for node in col:
                grammar_lines.append(f"{current_n} -> {node['text_str']} {next_n}")
                
            n_count += 1

    # We cap it off and return the string.
    grammar_lines.append(f"n{n_count} -> null")
    return "\n".join(grammar_lines)


# The Excel execution pipeline remains the exact same.
def process_railroad_diagrams(excel_filename, svg_folder):
    wb = openpyxl.load_workbook(excel_filename)
    sheet = wb.active
    sheet['G1'] = "Unoptimized Grammar"

    for row in range(2, sheet.max_row + 1):
        raw_val = sheet.cell(row=row, column=1).value
        if not raw_val:
            continue
            
        svg_filename = str(raw_val).strip()
        if not svg_filename.lower().endswith('.svg'):
            svg_filename += '.svg'
            
        absolute_path = os.path.abspath(os.path.join(svg_folder, svg_filename))
        
        # NEW CODE - TYPE THIS
        if os.path.exists(absolute_path):
            grammar_result = generate_grammar_from_svg(absolute_path)
            
            # We grab the specific cell we are about to write to.
            target_cell = sheet.cell(row=row, column=7)
            
            # We insert our grammar string.
            target_cell.value = grammar_result
            
            # We force Excel to enable "Wrap Text" so the \n characters stack vertically.
            target_cell.alignment = Alignment(wrap_text=True)
        else:
            sheet.cell(row=row, column=7).value = "SVG not found"

    wb.save(excel_filename)
    print("Pipeline complete. Bounding-Box grammar saved to Column G.")

if __name__ == "__main__":
    process_railroad_diagrams("railroad_diagrams.xlsx", "redlinesSVGs")