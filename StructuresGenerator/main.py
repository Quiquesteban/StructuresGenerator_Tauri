import pandas as pd
import json
import difflib
import os
import re
import uuid
import sys
from typing import Dict, List, Optional

CONFIG_FILE = "modbus_ai_config.json"
DEFAULT_VALUES = {
    'address': "0",
    'name': "UnnamedVariable",
    'datatype': "UINT",
    'offset': "0",
    'unit': "",
    'description': ""
}

def load_configuration() -> Dict[str, List[str]]:
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

def find_best_column(possible_aliases: List[str], columns: List[str]) -> Optional[str]:
    for alias in possible_aliases:
        best_match = difflib.get_close_matches(alias.lower(), [col.lower() for col in columns], n=1, cutoff=0.7)
        if best_match:
            return best_match[0]
    return None

def clean_variable_name(name: str) -> str:
    name = re.sub(r'[^\w ]', '', name)
    name = re.sub(r'^[^A-Za-z_]+', '', name)
    name = ''.join(word.capitalize() for word in name.split())
    return name or "InvalidName"

def validate_address(address: str) -> bool:
    try:
        int(address)
        return True
    except (ValueError, TypeError):
        return False

def process_sheet(df_raw: pd.DataFrame, sheet_name: str, excel_name: str) -> str:
    df = df_raw.iloc[1:].copy()
    df.columns = [str(col).strip().lower() for col in df_raw.iloc[0]]

    column_aliases = load_configuration()
    col_map = {key: find_best_column(aliases, df.columns) for key, aliases in column_aliases.items()}

    for field, col_name in col_map.items():
        if col_name and field in DEFAULT_VALUES:
            df[col_name] = df[col_name].fillna(DEFAULT_VALUES[field])

    if col_map['address']:
        df = df[df[col_map['address']].astype(str).apply(validate_address)]

    variable_names = set()
    variables = []

    for _, row in df.iterrows():
        name = clean_variable_name(str(row.get(col_map['name'], DEFAULT_VALUES['name'])))
        original_name = name
        counter = 1
        while name in variable_names:
            name = f"{original_name}_{counter}"
            counter += 1
        variable_names.add(name)

        address = int(row.get(col_map['address'], DEFAULT_VALUES['address']))
        scale = row.get(col_map.get('scale', ''), '1')
        unit = row.get(col_map.get('unit', ''), '')
        offset = row.get(col_map.get('offset', ''), '0')

        variables.append({
            'name': name,
            'address': address,
            'scale': str(scale).strip() or '1',
            'unit': str(unit).strip(),
            'offset': str(offset).strip() or '0'
        })

    max_name_length = max(len(v['name']) for v in variables)
    struct_lines = [
        f"    {v['name']}{' ' * (max_name_length - len(v['name']) + 4)}: UINT; "
        f"(* ModbusAddress: {v['address']};  Scale: {v['scale']};  Unit: {v['unit']};  Offset: {v['offset']}; *)"
        for v in variables
    ]

    dut_name = clean_variable_name(f"{excel_name}_{sheet_name}")
    struct_text = f"TYPE {dut_name} :\nSTRUCT\n" + "\n".join(struct_lines) + "\nEND_STRUCT\nEND_TYPE"

    script_dir = os.path.dirname(os.path.abspath(__file__))
    output_dir = os.path.join(script_dir, "generated")
    os.makedirs(output_dir, exist_ok=True)
    output_base = os.path.join(output_dir, f"{excel_name}_{sheet_name}_modbus_struct")

    with open(output_base + ".txt", "w", encoding="utf-8") as f_txt:
        f_txt.write(struct_text)

    guid = str(uuid.uuid4())
    xml_content = f"""<?xml version=\"1.0\" encoding=\"utf-8\"?>
<TcPlcObject Version=\"1.1.0.1\">
  <DUT Name=\"{dut_name}\" Id=\"{{{guid}}}\">
    <Declaration><![CDATA[
{struct_text}
]]></Declaration>
  </DUT>
</TcPlcObject>
"""

    full_path = output_base + ".TcDUT"
    with open(full_path, "w", encoding="utf-8") as f_tcdut:
        f_tcdut.write(xml_content)

    return json.dumps({
        "file": os.path.basename(full_path),
        "fullPath": os.path.abspath(full_path)
    })

def process_excel(file_path: str, sheet_name: Optional[str] = None):
    try:
        xl = pd.ExcelFile(file_path)
        if sheet_name is None:
            sheet_name = xl.sheet_names[0]
        elif sheet_name not in xl.sheet_names:
            raise ValueError(f"La hoja '{sheet_name}' no existe en el archivo.")

        df_raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
        excel_name = os.path.splitext(os.path.basename(file_path))[0]
        safe_sheet_name = re.sub(r'[^\w]', '', sheet_name)
        result = process_sheet(df_raw, safe_sheet_name, excel_name)
        print(result)
    except Exception as e:
        print(json.dumps({ "error": str(e) }), file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(json.dumps({ "error": "Uso: python main.py archivo.xlsx [hoja]" }), file=sys.stderr)
        sys.exit(1)

    excel_path = sys.argv[1]
    sheet = sys.argv[2] if len(sys.argv) > 2 else None
    process_excel(excel_path, sheet)
