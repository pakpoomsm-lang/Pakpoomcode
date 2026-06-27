#!/usr/bin/env python3
"""
Machine Name Validation — compare DB records vs UI labels
"""
import json, sqlite3, re
from collections import defaultdict

# Load mapping
with open(r'W:\PD\2.HEAT INDOOR\13.Suphamat P\Server_firstlot\public\data\machine_mapping.json') as f:
    mapping = json.load(f)

UI_MACHINES = {m['type']+':'+str(m['num']): m for m in mapping['machines']}

def extract_key(name):
    """Extract type:num from any machine name format (normalize leading zeros)"""
    if not name: return None
    t = name.lower()

    # Type detection
    type_map = {
        'expan|^ex\\d': 'ex',
        'finpress|fin.press': 'fp',
        'braz': 'ab',
        'oven': 'ov',
        'bender|hpb|h/p.bend': 'hp',
        'hairpin.insert|hairpininsert|h/p.ins': 'hi',
        'cut': 'ct',
    }

    mtype = None
    for pattern, typ in type_map.items():
        if re.search(pattern, t):
            mtype = typ
            break

    if not mtype: return None

    # Extract last number and normalize (strip leading zeros)
    nums = re.findall(r'\d+', t)
    if not nums: return None

    num = str(int(nums[-1]))  # strip leading zeros
    return mtype + ':' + num

def check_db(db_path, table, field, db_type):
    """Extract unique machine names from DB"""
    try:
        con = sqlite3.connect(db_path)
        rows = con.execute(f'SELECT DISTINCT {field} FROM {table}').fetchall()
        con.close()
        return [r[0] for r in rows if r[0]]
    except Exception as e:
        return []

# ─────────────────────────────────────
# Validate each API
# ─────────────────────────────────────
DBs = [
    (r'W:\PD\2.HEAT INDOOR\12.Pakpoom\mecp-python\expander_records.db',
     'records', 'machine', 'ex', 'Expander'),
    (r'W:\PD\2.HEAT INDOOR\13.Suphamat P\Server_firstlot\fp_records.db',
     'fp_records', 'mc', 'fp', 'Finpress'),
    (r'W:\PD\2.HEAT INDOOR\13.Suphamat P\Server_firstlot\fp_records.db',
     'hp_records', 'mc', 'hp', 'Hairpin Bender'),
    (r'W:\PD\2.HEAT INDOOR\13.Suphamat P\data\expander_records.db',
     'cutting_records', 'mc_line', 'ct', 'Cutting'),
]

print('\n' + '='*80 + '\n')
print('[MACHINE NAME VALIDATION]')
print('='*80)

for db_path, table, field, db_type, label in DBs:
    print(f'\n[{db_type.upper()}] {label} -- {table}.{field}')
    print('-' * 80)

    names = check_db(db_path, table, field, db_type)
    if not names:
        print(f'  No records found')
        continue

    unmatched = []
    matched = []

    for name in names:
        key = extract_key(name)
        if key and key in UI_MACHINES:
            matched.append((name, key, UI_MACHINES[key]['label']))
        else:
            unmatched.append((name, key))

    print(f'  Found {len(names)} unique values')
    print(f'  Matched: {len(matched)} | Unmatched: {len(unmatched)}')

    if matched:
        for db_name, key, ui_label in matched[:3]:
            print(f'    OK "{db_name}" -> {key} = {ui_label}')
        if len(matched) > 3:
            print(f'    ... and {len(matched)-3} more')

    if unmatched:
        print(f'\n  UNMATCHED ({len(unmatched)}):')
        for db_name, key in unmatched[:5]:
            print(f'    ?? "{db_name}" -> {key}')
        if len(unmatched) > 5:
            print(f'    ... and {len(unmatched)-5} more')

print('\n' + '='*80 + '\n')
