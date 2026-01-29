"""
script3_optimized_regex.py (v2)
------------------------------------------------------------------------
1. Opens 'railroad_diagrams_complete.xlsx'.
2. Reads 'Command' (Col A) and 'Unoptimized Grammar' (Col F).
3. Generates Regex.
4. PREPENDS Command Name if missing from the diagram.
5. Writes result to 'Optimized Regex' (Col G).
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
COL_COMMAND = 1  # Column A
COL_GRAMMAR = 6  # Column F
COL_OUTPUT = 7   # Column G

# Common CICS verbs to help split command names (e.g. "CHANGEPASSWORD" -> "CHANGE PASSWORD")
KNOWN_VERBS = [
    "CHANGE", "INQUIRE", "SET", "PERFORM", "DISABLE", "ENABLE", 
    "EXTRACT", "ISSUE", "RESYNC", "RETRIEVE", "SEND", "RECEIVE", 
    "START", "STOP", "SUSPEND", "WAIT", "WRITE", "REWRITE", 
    "DELETE", "FREE", "POINT", "PROCESS", "PUT", "QUERY", 
    "READ", "RELEASE", "RESET", "RESUME", "ROLLBACK", "SIGNAL", 
    "SPOOL", "TEST", "UNLOCK", "VERIFY", "CONVERSE", "CONNECT"
]

def clean_label(label):
    """Cleans up labels, removing spaces in parens."""
    if not label: return ""
    label = re.sub(r'\(\s+', '(', label)
    label = re.sub(r'\s+\)', ')', label)
    return label

def format_command_name(cmd_raw):
    """
    Tries to format 'CHANGEPASSWORD' -> 'CHANGE PASSWORD' 
    based on known verbs.
    """
    if not cmd_raw: return ""
    cmd_upper = cmd_raw.upper().strip()
    
    # Check if it starts with a known verb
    for verb in KNOWN_VERBS:
        if cmd_upper.startswith(verb) and len(cmd_upper) > len(verb):
            # If the next char is not a space, insert one
            suffix = cmd_upper[len(verb):]
            if not suffix.startswith(" "):
                return f"{verb} {suffix}"
    return cmd_upper

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

        # Prefix
        prefix = []
        while path_list and all(p for p in path_list):
            if all(p[0] == path_list[0][0] for p in path_list):
                prefix.append(path_list[0][0])
                for p in path_list: p.pop(0)
            else: break
        
        # Suffix
        suffix = []
        while path_list and all(p for p in path_list):
            if all(p[-1] == path_list[0][-1] for p in path_list):
                suffix.insert(0, path_list[0][-1])
                for p in path_list: p.pop()
            else: break

        # Middle
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
    print("      SCRIPT 3: OPTIMIZED REGEX GENERATOR (V2)")
    print("=" * 60)

    if not os.path.exists(INPUT_FILE):
        print(f"❌ File not found: {INPUT_FILE}")
        return

    print(f"📂 Loading {INPUT_FILE}...")
    wb = load_workbook(INPUT_FILE)
    ws = wb.active

    # Setup Column G
    output_col_letter = get_column_letter(COL_OUTPUT)
    ws.column_dimensions[output_col_letter].width = 43
    header_cell = ws.cell(row=1, column=COL_OUTPUT)
    header_cell.value = "Optimized Regex"
    header_cell.font = Font(bold=True)
    header_cell.alignment = Alignment(horizontal='center')

    row_idx = 2
    count = 0
    
    while True:
        # Check if row exists
        cmd_cell = ws.cell(row=row_idx, column=COL_COMMAND)
        if not cmd_cell.value: break
        
        cmd_name_raw = str(cmd_cell.value).strip()
        grammar_text = ws.cell(row=row_idx, column=COL_GRAMMAR).value
        
        if not grammar_text:
            row_idx += 1
            continue

        # 1. Generate Basic Regex
        regex = get_regex_from_grammar(grammar_text)
        
        # 2. Check for missing Command Name
        # Logic: If the regex doesn't start with the command (or the command's prefix), prepend it.
        
        # Format "CHANGEPASSWORD" -> "CHANGE PASSWORD"
        nice_cmd_name = format_command_name(cmd_name_raw)
        
        # Determine the first word of the generated regex to compare
        # (Remove optional parens like '(TCTUA...' -> 'TCTUA')
        first_token = regex.lstrip('(').split(' ')[0].split('(')[0]
        
        # Does the Regex already contain the command?
        # e.g. cmd="CHANGE", regex="CHANGE TASK..." -> Yes.
        # e.g. cmd="ADDRESS", regex="TCTUA..." -> No.
        
        # We check if the Command starts with the First Token (e.g. ALLOCATE starts with ALLOCATE)
        # OR if the First Token starts with the Command (rare, but possible)
        # If neither, we assume the command is missing.
        
        starts_with_cmd = False
        if first_token and nice_cmd_name:
            # Check overlap
            if nice_cmd_name.startswith(first_token) or first_token.startswith(nice_cmd_name.split(' ')[0]):
                starts_with_cmd = True
                
        # Special fix: "ADDRESS" vs "TCTUA" -> starts_with_cmd = False
        
        if not starts_with_cmd and regex:
            # Prepend the formatted command name
            if regex == "$": 
                 # If regex is just End of Line, it means "COMMAND_NAME" is the only thing.
                regex = f"{nice_cmd_name}$"
            else:
                regex = f"{nice_cmd_name} {regex}"

        # Write result
        out_cell = ws.cell(row=row_idx, column=COL_OUTPUT)
        out_cell.value = regex
        out_cell.alignment = Alignment(wrap_text=True, vertical='top')

        print(f"Row {row_idx}: {regex[:60]}...")
        
        row_idx += 1
        count += 1

    print(f"\n💾 Saving updates to {INPUT_FILE}...")
    wb.save(INPUT_FILE)
    print(f"✅ Done! Processed {count} rows.")

if __name__ == "__main__":
    main()
    
    