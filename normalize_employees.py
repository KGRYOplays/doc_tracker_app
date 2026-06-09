"""
Standalone script to normalize the employees column in code_lookup table.

Fixes two issues:
1. Converts old comma-delimited employee names to '|||' delimiter
2. Converts "LastName, FirstName MI" format to "FirstName MI LastName" format
3. Handles "LastName, FirstName MI AND N OTHERS" → "FirstName MI LastName AND N OTHERS"

Uses master_data as reference to distinguish delimiter commas from
commas that are part of employee names.

Usage:
    python normalize_employees.py
"""
import sqlite3
import re

DB_FILE = 'app.db'

# Pattern: "AND 9 OTHERS", "AND 12 OTHERS", etc.
AND_OTHERS_RE = re.compile(r'^(.*?)\s+AND\s+(\d+)\s+OTHERS$', re.IGNORECASE)


def get_db_connection():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    return conn


def normalize_name(name):
    """Normalize a name for comparison: lowercase, strip periods, collapse spaces."""
    s = name.strip().lower()
    s = s.replace('.', '')
    s = re.sub(r'\s+', ' ', s)
    return s


def reorder_name(name):
    """Convert 'LastName, FirstName MI' to 'FirstName MI LastName'.
    If no comma, return as-is.
    Also handles 'LastName, FirstName MI AND N OTHERS' → 'FirstName MI LastName AND N OTHERS'."""
    name = name.strip()
    if ',' not in name:
        return name

    parts = [p.strip() for p in name.split(',', 1)]
    if len(parts) != 2:
        return name

    last_name = parts[0].strip()
    rest = parts[1].strip()

    # Check for "AND N OTHERS" suffix
    m = AND_OTHERS_RE.match(rest)
    if m:
        first_part = m.group(1).strip()
        num_others = m.group(2)
        return f"{first_part} {last_name} AND {num_others} OTHERS"

    # Simple "LastName, FirstName MI" → "FirstName MI LastName"
    return f"{rest} {last_name}"


def normalize():
    conn = get_db_connection()
    cursor = conn.cursor()

    # Build a set of all known employee names from master_data (normalized)
    master_rows = cursor.execute("SELECT DISTINCT employee_name FROM master_data").fetchall()
    known_names = {}
    for r in master_rows:
        if r['employee_name']:
            norm = normalize_name(r['employee_name'])
            if norm:
                known_names[norm] = r['employee_name']
    print(f"Loaded {len(known_names)} unique employee names from master_data.")

    # Process ALL code_lookup entries (both comma-delimited and already using |||)
    all_rows = cursor.execute("SELECT code, employees FROM code_lookup").fetchall()
    print(f"Processing {len(all_rows)} code_lookup entries...\n")

    fixed_count = 0
    unresolved = []

    for row in all_rows:
        code = row['code']
        employees_str = row['employees']
        if not employees_str:
            continue

        # Step 1: Split into name parts (handle both , and ||| delimiters)
        if '|||' in employees_str:
            names = [n.strip() for n in employees_str.split('|||') if n.strip()]
            already_piped = True
        elif ',' in employees_str:
            # Old comma-delimited format — split and try to match against master_data
            parts = [p.strip() for p in employees_str.split(',') if p.strip()]
            names = _greedy_match(parts, known_names)
            already_piped = False
        else:
            # Single name, no delimiters — just reorder if needed
            reordered = reorder_name(employees_str)
            if reordered != employees_str:
                cursor.execute(
                    "UPDATE code_lookup SET employees = ? WHERE code = ?",
                    (reordered, code)
                )
                fixed_count += 1
            continue

        # Step 2: Reorder each name from "Last, First" to "First Last"
        new_names = []
        changed = False
        for name in names:
            reordered = reorder_name(name)
            if reordered != name:
                changed = True
            new_names.append(reordered)

        # Step 3: Join with ||| and update if changed
        new_employees = '|||'.join(new_names)
        if new_employees != employees_str:
            cursor.execute(
                "UPDATE code_lookup SET employees = ? WHERE code = ?",
                (new_employees, code)
            )
            fixed_count += 1

    conn.commit()
    conn.close()

    print(f"Results:")
    print(f"  Fixed:    {fixed_count} entries")


def _greedy_match(parts, known_names):
    """Greedily match comma-separated parts against known employee names."""
    names = []
    i = 0
    while i < len(parts):
        matched = False
        for length in range(min(4, len(parts) - i), 0, -1):
            candidate = ' '.join(parts[i:i + length])
            norm_candidate = normalize_name(candidate)
            if norm_candidate in known_names:
                names.append(known_names[norm_candidate])
                i += length
                matched = True
                break
        if not matched:
            names.append(parts[i])
            i += 1
    return names


if __name__ == '__main__':
    normalize()