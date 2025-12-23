import xml.etree.ElementTree as ET
import re
import openpyxl
from openpyxl import load_workbook

def simplify_railroad_svg(svg_string):
    """
    Simplify an SVG railroad diagram by:
    1. Keeping only the main transform group and <defs>
    2. Removing polygons (arrowheads)
    3. Converting all curved paths to straight orthogonal lines
    4. Connecting disconnected path segments with gaps
    5. Merging consecutive syntax text nodes
    6. Rounding all coordinates to 2 decimals
    """
    
    # Register namespace to handle ns0: prefix
    ET.register_namespace('', 'http://www.w3.org/2000/svg')
    
    # Parse the SVG
    root = ET.fromstring(svg_string)
    
    # 1. Keep only <defs> and main <g transform="...">
    defs_element = None
    main_group = None
    
    for child in root:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        
        if tag == 'defs':
            defs_element = child
        elif tag == 'g' and 'transform' in child.attrib:
            main_group = child
    
    if main_group is None:
        raise ValueError("No <g> element with transform attribute found")
    
    # Clear root and add back only defs and main group
    for child in list(root):
        root.remove(child)
    
    if defs_element is not None:
        root.append(defs_element)
    root.append(main_group)
    
    # 2. Remove all polygon tags
    remove_polygons(root)
    
    # 3. Convert curved paths to orthogonal lines
    convert_paths_to_orthogonal(root)
    
    # 4. Connect disconnected path segments
    connect_path_gaps(root)
    
    # 5. Merge consecutive syntax text nodes
    merge_text_nodes(root)
    
    # 6. Round all numeric coordinates to 2 decimals
    round_all_coordinates(root)
    
    # 7. Clean up namespace prefixes and format output
    xml_str = ET.tostring(root, encoding='unicode')
    
    # Remove namespace prefixes for cleaner output
    xml_str = re.sub(r'<ns\d+:', '<', xml_str)
    xml_str = re.sub(r'</ns\d+:', '</', xml_str)
    xml_str = re.sub(r'xmlns:ns\d+="[^"]*"', '', xml_str)
    
    # Pretty print manually
    return prettify_xml(xml_str)


def remove_polygons(root):
    """Remove all polygon elements."""
    for parent in root.iter():
        for child in list(parent):
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if tag == 'polygon':
                parent.remove(child)


def convert_paths_to_orthogonal(root):
    """Convert all curved paths to straight orthogonal lines."""
    for path in root.iter():
        tag = path.tag.split('}')[-1] if '}' in path.tag else path.tag
        if tag == 'path':
            d_attr = path.get('d', '')
            if d_attr:
                new_d = make_orthogonal_path(d_attr)
                path.set('d', new_d)


def make_orthogonal_path(path_d):
    """
    Convert a path with curves to straight orthogonal lines.
    Q commands indicate corners - we need to trace the corner properly.
    """
    
    # Parse all commands
    tokens = re.findall(
        r'([MQL])\s*([-+]?\d*\.?\d+)\s+([-+]?\d*\.?\d+)(?:\s+([-+]?\d*\.?\d+)\s+([-+]?\d*\.?\d+))?',
        path_d
    )
    
    if not tokens:
        return path_d
    
    # Build path by interpreting each command
    points = []
    
    for token in tokens:
        cmd = token[0]
        if cmd == 'M':
            # Move to start point
            x, y = round(float(token[1]), 2), round(float(token[2]), 2)
            points.append((x, y))
        elif cmd == 'L':
            # Line to point
            x, y = round(float(token[1]), 2), round(float(token[2]), 2)
            points.append((x, y))
        elif cmd == 'Q':
            # Quadratic curve: control point tells us the corner direction
            cx, cy = round(float(token[1]), 2), round(float(token[2]), 2)
            x, y = round(float(token[3]), 2), round(float(token[4]), 2)
            
            # Add the corner point (where the direction changes)
            points.append((cx, cy))
            # Then add the endpoint
            points.append((x, y))
    
    if len(points) < 2:
        return path_d
    
    # Build orthogonal path
    result = [f'M{points[0][0]},{points[0][1]}']
    
    for i in range(1, len(points)):
        x, y = points[i]
        result.append(f'L{x},{y}')
    
    return ' '.join(result)


def connect_path_gaps(root):
    """
    Connect consecutive path elements that have gaps between them.
    Only connects paths that are part of loop-backs (non-zero Y movement).
    """
    
    # Collect all path elements in document order
    all_paths = []
    for elem in root.iter():
        tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
        if tag == 'path':
            all_paths.append(elem)
    
    # Process consecutive paths
    for i in range(len(all_paths) - 1):
        current = all_paths[i]
        next_path = all_paths[i + 1]
        
        current_d = current.get('d', '')
        next_d = next_path.get('d', '')
        
        if not current_d or not next_d:
            continue
        
        # Extract end point of current path (last two numbers)
        current_coords = re.findall(r'[-+]?\d*\.?\d+', current_d)
        if len(current_coords) >= 2:
            end_x = round(float(current_coords[-2]), 2)
            end_y = round(float(current_coords[-1]), 2)
        else:
            continue
        
        # Extract start point of next path (after M command)
        next_match = re.search(r'M\s*([-+]?\d*\.?\d+)\s*,?\s*([-+]?\d*\.?\d+)', next_d)
        if next_match:
            start_x = round(float(next_match.group(1)), 2)
            start_y = round(float(next_match.group(2)), 2)
        else:
            continue
        
        # Check if current path has vertical movement (loop-back indicator)
        # Loop paths go up/down (negative Y values)
        has_negative_y = any(float(c) < -5 for c in current_coords[1::2])  # Check Y coordinates
        
        # Only connect if:
        # 1. Path has vertical loop movement (negative Y)
        # 2. Paths are horizontally aligned
        # 3. There's a significant horizontal gap
        if has_negative_y and abs(end_y - start_y) < 1:  # Horizontally aligned
            gap_x = abs(end_x - start_x)
            if gap_x > 5:  # Significant gap (not just rounding)
                # Add horizontal connecting line
                current_d += f' L{start_x},{start_y}'
                current.set('d', current_d)
                print(f"✓ Connected loop gap: ({end_x},{end_y}) → ({start_x},{start_y})")


def merge_text_nodes(root):
    """
    Merge consecutive text nodes within each <g class="text"> wrapper.
    """
    
    for parent in root.iter():
        tag = parent.tag.split('}')[-1] if '}' in parent.tag else parent.tag
        
        # Look for <g class="text"> wrappers
        if tag == 'g' and parent.get('class') == 'text':
            # Get all text children
            text_children = []
            for child in parent:
                child_tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                if child_tag == 'text':
                    text_children.append(child)
            
            if len(text_children) < 2:
                continue
            
            # Merge all text nodes into the first one
            first_text = text_children[0]
            to_remove = []
            
            for i in range(1, len(text_children)):
                next_text = text_children[i]
                # Merge text content
                if next_text.text:
                    first_text.text = (first_text.text or '') + next_text.text
                to_remove.append(next_text)
            
            # Remove merged elements from parent
            for elem in to_remove:
                parent.remove(elem)


def round_all_coordinates(root):
    """Round all numeric coordinates to 2 decimal places."""
    
    for elem in root.iter():
        tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
        
        # Round line coordinates
        if tag == 'line':
            for attr in ['x1', 'y1', 'x2', 'y2']:
                if attr in elem.attrib:
                    try:
                        value = float(elem.get(attr))
                        elem.set(attr, str(round(value, 2)))
                    except (ValueError, TypeError):
                        pass
        
        # Round rect coordinates
        elif tag == 'rect':
            for attr in ['x', 'y', 'width', 'height', 'rx', 'ry']:
                if attr in elem.attrib:
                    try:
                        value = float(elem.get(attr))
                        elem.set(attr, str(round(value, 2)))
                    except (ValueError, TypeError):
                        pass
        
        # Round text coordinates
        elif tag == 'text':
            for attr in ['x', 'y']:
                if attr in elem.attrib:
                    try:
                        value = float(elem.get(attr))
                        elem.set(attr, str(round(value, 2)))
                    except (ValueError, TypeError):
                        pass
        
        # Round transform values
        if 'transform' in elem.attrib:
            transform = elem.get('transform')
            # Round numbers in transform attribute
            def round_match(match):
                return str(round(float(match.group(0)), 2))
            transform = re.sub(r'[-+]?\d*\.?\d+', round_match, transform)
            elem.set('transform', transform)


def prettify_xml(xml_str):
    """Simple XML pretty printing."""
    # Remove extra whitespace
    xml_str = re.sub(r'>\s+<', '><', xml_str)
    
    lines = []
    indent_level = 0
    
    # Split by tags
    parts = re.split(r'(<[^>]+>)', xml_str)
    
    for part in parts:
        if not part.strip():
            continue
        
        if part.startswith('</'):
            # Closing tag
            indent_level -= 1
            lines.append('  ' * indent_level + part)
        elif part.startswith('<') and not part.endswith('/>'):
            # Opening tag
            lines.append('  ' * indent_level + part)
            if not part.startswith('<?') and not part.startswith('<!'):
                indent_level += 1
        elif part.startswith('<') and part.endswith('/>'):
            # Self-closing tag
            lines.append('  ' * indent_level + part)
        else:
            # Text content
            if part.strip():
                lines.append('  ' * indent_level + part)
    
    return '\n'.join(lines)


def process_excel_file(input_file, output_file=None):
    """
    Process SVGs from Excel file.
    Reads original SVGs from column 1, writes simplified SVGs to column 2.
    
    Args:
        input_file: Path to input Excel file (.xlsx)
        output_file: Path to output Excel file (default: same as input)
    """
    
    if output_file is None:
        output_file = input_file
    
    # Load the workbook
    wb = load_workbook(input_file)
    ws = wb.active
    
    # Process each row
    row_num = 1
    while True:
        # Read original SVG from column A (column 1)
        cell = ws.cell(row=row_num, column=1)
        original_svg = cell.value
        
        # Stop if empty cell
        if not original_svg:
            break
        
        try:
            # Simplify the SVG
            simplified_svg = simplify_railroad_svg(original_svg)
            
            # Write to column B (column 2)
            ws.cell(row=row_num, column=2, value=simplified_svg)
            
            print(f"✓ Processed row {row_num}")
            
        except Exception as e:
            print(f"✗ Error processing row {row_num}: {e}")
            ws.cell(row=row_num, column=2, value=f"ERROR: {e}")
        
        row_num += 1
    
    # Save the workbook
    wb.save(output_file)
    print(f"\n✓ Saved results to: {output_file}")
    print(f"✓ Processed {row_num - 1} rows total")


# Example usage
if __name__ == "__main__":
    # Option 1: Process Excel file
    # process_excel_file('railroad_diagrams.xlsx')
    
    # Option 2: Test single SVG
    sample_svg = '''
  <svg contentScriptType="text/ecmascript" zoomAndPan="magnify" contentStyleType="text/css" version="1.0" width="274.17706500291825px" preserveAspectRatio="xMidYMid meet" viewBox="0 0 274.17706500291825 110.54166793823242" height="110.54166793823242px" class="syntaxdiagram" fill="currentColor"><defs><style type="text/css" xml:space="preserve">
.arrow, .syntaxarrow { fill: none; stroke: black; }
.arrowheadStartEnd, .arrowheadRepSep, .arrowheadRepSepReturn { stroke: black; fill: black; }
.arrowheadSeq, .arrowheadStartChoice, .arrowheadAfterChoice, .arrowheadStartRepGroup, .arrowheadEndRepGroup, .arrowheadRev { stroke: none; fill: none; }
rect { fill: none; stroke: none; }
rect.fragref,rect.syntaxfragref { fill: none; stroke: black; }
text {
fill: #000000;
fill-opacity: 1;
font-family: IBM Plex Sans,Arial Unicode MS,Arial,Helvetica;
font-style: normal;
font-weight: normal;
font-size: 8pt;
stroke: #000000;
stroke-width: 0.1;
}
text.var, text.syntaxvar {font-style:italic;}
</style></defs><g transform="translate(5,5)" class="diagram" xml:base="..//"><g transform="translate(10,26.70138931274414)" class="boxed groupcomp"><g transform="translate(4,0)" class="unboxed syntaxkwd"><g class="text" transform="translate(0,3)"><text class="syntaxkwd">HANDLE ABEND</text></g></g><rect rx="3" x="0" width="88.06121826171875" height="15.44444465637207" y="-8.44444465637207" class="syntaxgroupcomp"></rect></g><g transform="translate(113.06121826171875,26.70138931274414)" class="groupchoice"><g></g><g transform="translate(43.963547569513324,-18.128472328186035)" class="boxed syntaxkwd"><g class="text" transform="translate(4,3)"><text class="syntaxkwd">CANCEL</text></g><rect rx="3" x="0" width="48.18875160217285" height="15.70138931274414" y="-8.572916984558105" class="syntaxkwd"></rect></g><g transform="translate(20,20.10763931274414)" class="boxed groupcomp"><g transform="translate(4,0)" class="unboxed syntaxkwd"><g class="text" transform="translate(0,3)"><text class="syntaxkwd">PROGRAM(</text></g><g class="text" transform="translate(55.86165103912354,3)"><text class="syntaxvar"> name</text></g><g class="text" transform="translate(85.23928432464601,3)"><text class="syntaxdelim">)</text></g></g><rect rx="3" x="0" width="96.1158467411995" height="17.56944465637207" y="-9.10763931274414" class="syntaxgroupcomp"></rect></g><g transform="translate(32.132452917099,43.67708396911621)" class="boxed groupcomp"><g transform="translate(4,0)" class="unboxed syntaxkwd"><g class="text" transform="translate(0,3)"><text class="syntaxkwd">LABEL(</text></g><g class="text" transform="translate(35.00087985992432,3)"><text class="syntaxvar"> label</text></g><g class="text" transform="translate(60.974378490448004,3)"><text class="syntaxdelim">)</text></g></g><rect rx="3" x="0" width="71.8509409070015" height="17.56944465637207" y="-9.10763931274414" class="syntaxgroupcomp"></rect></g><g transform="translate(48.134441095590596,66.71180629730225)" class="boxed syntaxkwd"><g class="text" transform="translate(4,3)"><text class="syntaxkwd">RESET</text></g><rect rx="3" x="0" width="39.84696455001831" height="15.70138931274414" y="-8.572916984558105" class="syntaxkwd"></rect></g><line y2="0" x1="0" x2="68.05792337059975" class="syntaxarrow" y1="0"></line><line y2="0" x1="68.05792337059975" x2="136.1158467411995" class="syntaxarrow" y1="0"></line><polygon class="arrowheadAfterChoice" points="131.1158467411995,0 126.1158467411995,2.5 126.1158467411995,-2.5" transform="rotate(0,131.1158467411995,0)"></polygon><path class="syntaxarrow" d="M0 0 Q5 0 5 -5 L5 -13.128472328186035 Q5 -18.128472328186035 10 -18.128472328186035 L43.963547569513324 -18.128472328186035"></path><polygon class="arrowheadStartChoice" points="43.963547569513324,-18.128472328186035 38.963547569513324,-15.628472328186035 38.963547569513324,-20.628472328186035" transform="rotate(0,43.963547569513324,-18.128472328186035)"></polygon><path class="syntaxarrow" d="M92.15229917168618 -18.128472328186035 L126.1158467411995 -18.128472328186035 Q131.1158467411995 -18.128472328186035 131.1158467411995 -13.128472328186035 L131.1158467411995 -5 Q131.1158467411995 0 136.1158467411995 0"></path><polygon class="arrowheadAfterChoice" points="131.1158467411995,-5 126.1158467411995,-2.5 126.1158467411995,-7.5" transform="rotate(90,131.1158467411995,-5)"></polygon><path class="syntaxarrow" d="M0 0 Q5 0 5 5 L5 15.10763931274414 Q5 20.10763931274414 10 20.10763931274414 L20 20.10763931274414"></path><polygon class="arrowheadStartChoice" points="20,20.10763931274414 15,22.60763931274414 15,17.60763931274414" transform="rotate(0,20,20.10763931274414)"></polygon><path class="syntaxarrow" d="M116.1158467411995 20.10763931274414 L126.1158467411995 20.10763931274414 Q131.1158467411995 20.10763931274414 131.1158467411995 15.10763931274414 L131.1158467411995 5 Q131.1158467411995 0 136.1158467411995 0"></path><polygon class="arrowheadAfterChoice" points="131.1158467411995,5 126.1158467411995,7.5 126.1158467411995,2.5" transform="rotate(-90,131.1158467411995,5)"></polygon><path class="syntaxarrow" d="M0 0 Q5 0 5 5 L5 38.67708396911621 Q5 43.67708396911621 10 43.67708396911621 L32.132452917099 43.67708396911621"></path><polygon class="arrowheadStartChoice" points="32.132452917099,43.67708396911621 27.132452917099002,46.17708396911621 27.132452917099002,41.17708396911621" transform="rotate(0,32.132452917099,43.67708396911621)"></polygon><path class="syntaxarrow" d="M103.9833938241005 43.67708396911621 L126.1158467411995 43.67708396911621 Q131.1158467411995 43.67708396911621 131.1158467411995 38.67708396911621 L131.1158467411995 5 Q131.1158467411995 0 136.1158467411995 0"></path><polygon class="arrowheadAfterChoice" points="131.1158467411995,5 126.1158467411995,7.5 126.1158467411995,2.5" transform="rotate(-90,131.1158467411995,5)"></polygon><path class="syntaxarrow" d="M0 0 Q5 0 5 5 L5 61.711806297302246 Q5 66.71180629730225 10 66.71180629730225 L48.134441095590596 66.71180629730225"></path><polygon class="arrowheadStartChoice" points="48.134441095590596,66.71180629730225 43.134441095590596,69.21180629730225 43.134441095590596,64.21180629730225" transform="rotate(0,48.134441095590596,66.71180629730225)"></polygon><path class="syntaxarrow" d="M87.9814056456089 66.71180629730225 L126.1158467411995 66.71180629730225 Q131.1158467411995 66.71180629730225 131.1158467411995 61.711806297302246 L131.1158467411995 5 Q131.1158467411995 0 136.1158467411995 0"></path><polygon class="arrowheadAfterChoice" points="131.1158467411995,5 126.1158467411995,7.5 126.1158467411995,2.5" transform="rotate(-90,131.1158467411995,5)"></polygon></g><polygon class="arrowheadStartEnd" points="0,26.70138931274414 -5,29.20138931274414 -5,24.20138931274414" transform="rotate(0,0,26.70138931274414)"></polygon><polygon class="arrowheadStartEnd" points="5,26.70138931274414 0,29.20138931274414 0,24.20138931274414" transform="rotate(0,5,26.70138931274414)"></polygon><line y2="26.70138931274414" x1="0" x2="10" class="syntaxarrow" y1="26.70138931274414"></line><line y2="26.70138931274414" x1="98.06121826171875" x2="113.06121826171875" class="syntaxarrow" y1="26.70138931274414"></line><polygon class="arrowheadSeq" points="113.06121826171875,26.70138931274414 108.06121826171875,29.20138931274414 108.06121826171875,24.20138931274414" transform="rotate(0,113.06121826171875,26.70138931274414)"></polygon><line y2="26.70138931274414" x1="249.17706500291825" x2="259.17706500291825" class="syntaxarrow" y1="26.70138931274414"></line><polygon class="arrowheadStartEnd" points="259.17706500291825,26.70138931274414 254.17706500291825,29.20138931274414 254.17706500291825,24.20138931274414" transform="rotate(0,259.17706500291825,26.70138931274414)"></polygon><polygon class="arrowheadStartEnd" points="259.17706500291825,26.70138931274414 254.17706500291825,29.20138931274414 254.17706500291825,24.20138931274414" transform="rotate(180,259.17706500291825,26.70138931274414)"></polygon></g></svg>

   ''' 
    print(simplify_railroad_svg(sample_svg))