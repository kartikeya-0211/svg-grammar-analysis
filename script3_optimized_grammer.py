"""
script3_optimized_regex.py
------------------------------------------------------------------------
1. Opens 'railroad_diagrams_complete.xlsx' (safely, keeping images).
2. Reads 'Unoptimized Grammar' from Column F.
3. Converts it to a clean Regex (replacing 'null' with '$' and fixing spaces).
4. Writes the result to Column G ('Optimized Regex').
5. Sets Column G width to 43.
"""

import sys
import os
import re

try:
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment, Font
    from openpyxl.utils import get_column_letter
except ImportError:
    print("❌ Error: openpyxl is missing. Run: pip install openpyxl")
    sys.exit()

INPUT_FILE = 'railroad_diagrams_complete.xlsx'
COL_GRAMMAR = 6  # Column F
COL_OUTPUT = 7   # Column G

def clean_label(label):
    """
    Cleans up labels:
    1. Removes extra spaces inside parentheses: 'FILE( name )' -> 'FILE(name)'
    """
    if not label: return ""
    label = re.sub(r'\(\s+', '(', label)
    label = re.sub(r'\s+\)', ')', label)
    return label

def get_regex_from_grammar(input_data):
    if not isinstance(input_data, str) or not input_data.strip():
        return ""

    adj = {}
    all_nodes = set()
    incoming = set()

    # 1. Parse the Grammar
    lines = input_data.strip().split('\n')
    for line in lines:
        line = line.strip()
        if not line or '->' not in line: continue

        parts = line.split('->')
        src = parts[0].strip()
        rhs = parts[1].strip()
        
        all_nodes.add(src)
        rhs_tokens = rhs.split()
        target = None
        label = None
        
        # Case A: nX -> nY (Epsilon)
        if len(rhs_tokens) == 1 and rhs_tokens[0].startswith('n') and rhs_tokens[0][1:].isdigit():
            target = rhs_tokens[0]
            label = None 
        # Case B: nX -> null (End) -> Terminal '$'
        elif rhs_tokens[0] == 'null':
            target = "FINAL_STATE"
            label = "$"
        # Case C: nX -> LABEL (Terminal)
        elif len(rhs_tokens) == 1:
            target = "FINAL_STATE"
            label = clean_label(rhs_tokens[0])
        # Case D: nX -> LABEL nY (Standard)
        else:
            target = rhs_tokens[-1]
            label = clean_label(" ".join(rhs_tokens[:-1]))

        if src not in adj: adj[src] = []
        adj[src].append((target, label))
        
        if target != "FINAL_STATE":
            all_nodes.add(target)
            incoming.add(target)

    # 2. Find Start Node
    start_node = None
    for node in all_nodes:
        if node not in incoming:
            start_node = node
            break
    if not start_node and all_nodes:
        start_node = list(all_nodes)[0]

    if not start_node: return ""

    # 3. DFS Pathfinding
    paths = []
    stack = [(start_node, [], set())]
    MAX_PATHS = 1000 
    
    while stack:
        curr, path, visited = stack.pop()
        
        if curr == "FINAL_STATE":
            paths.append(path)
            if len(paths) > MAX_PATHS: break 
            continue
            
        if curr not in adj or not adj[curr]:
            if path: paths.append(path)
            continue
            
        if curr in visited: continue
        visited.add(curr)

        for target, label in adj[curr]:
            new_path = list(path)
            if label: new_path.append(label)
            stack.append((target, new_path, visited.copy()))

    # 4. Simplify / Regex Construction
    def simplify(path_list):
        if not path_list: return ""
        if len(path_list) == 1: return "".join(path_list[0])

        # Common Prefix
        prefix = []
        while path_list and all(p for p in path_list):
            if all(p[0] == path_list[0][0] for p in path_list):
                prefix.append(path_list[0][0])
                for p in path_list: p.pop(0)
            else: break
        
        # Common Suffix
        suffix = []
        while path_list and all(p for p in path_list):
            if all(p[-1] == path_list[0][-1] for p in path_list):
                suffix.insert(0, path_list[0][-1])
                for p in path_list: p.pop()
            else: break

        # Middles
        middles = []
        has_empty = False
        for p in path_list:
            if not p: has_empty = True
            else: middles.append(simplify([p]))
        
        middles = sorted(list(set(middles)))
        middle_str = ""
        if middles:
            if len(middles) > 1: middle_str = "(" + "|".join(middles) + ")"
            else: middle_str = middles[0]
            if has_empty and ("|" in middle_str or len(middles[0]) > 4):
                 middle_str = f"({middle_str})"

        if has_empty:
            if middle_str:
                if not (middle_str.startswith("(") and middle_str.endswith(")")):
                    middle_str = f"({middle_str})"
                middle_str += "?"
            
        return "".join(prefix) + middle_str + "".join(suffix)

    return simplify(paths)

def main():
    print("=" * 60)
    print("      SCRIPT 3: OPTIMIZED REGEX GENERATOR")
    print("=" * 60)

    if not os.path.exists(INPUT_FILE):
        print(f"❌ File not found: {INPUT_FILE}")
        return

    print(f"📂 Loading {INPUT_FILE}...")
    wb = load_workbook(INPUT_FILE)
    ws = wb.active

    # --- UPDATED: SET COLUMN WIDTH ---
    output_col_letter = get_column_letter(COL_OUTPUT) # Column G
    ws.column_dimensions[output_col_letter].width = 43  # <--- HERE IS YOUR WIDTH 43
    
    # Header
    header_cell = ws.cell(row=1, column=COL_OUTPUT)
    header_cell.value = "Optimized Regex"
    header_cell.font = Font(bold=True)
    header_cell.alignment = Alignment(horizontal='center')

    # Iterate rows
    row_idx = 2
    count = 0
    
    while True:
        # Stop if no more data in Column A
        if not ws.cell(row=row_idx, column=1).value:
            break
            
        # Read Grammar from Column F
        grammar_text = ws.cell(row=row_idx, column=COL_GRAMMAR).value
        
        if not grammar_text:
            row_idx += 1
            continue

        # Process
        regex = get_regex_from_grammar(grammar_text)
        
        # Write to Column G
        out_cell = ws.cell(row=row_idx, column=COL_OUTPUT)
        out_cell.value = regex
        out_cell.alignment = Alignment(wrap_text=True, vertical='top')

        # Console Progress
        print(f"Row {row_idx}: {regex[:50]}..." if regex else f"Row {row_idx}: [Empty]")
        
        row_idx += 1
        count += 1

    print(f"\n💾 Saving updates to {INPUT_FILE}...")
    wb.save(INPUT_FILE)
    print(f"✅ Done! Processed {count} rows.")

if __name__ == "__main__":
    main()