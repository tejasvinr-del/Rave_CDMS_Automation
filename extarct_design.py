import xml.etree.ElementTree as ET
import pandas as pd
def extract_rave_xml_to_df(file_path, sheet_name):
    """
    Custom parser to read Medidata Rave 'XML Spreadsheet 2003' exports
    that masquerade as .xls files.
    """
    tree = ET.parse(file_path)
    root = tree.getroot()
    
    # Identify the XML namespace dynamically
    ns = ''
    if '}' in root.tag:
        ns = root.tag.split('}')[0] + '}'
        
    # 1. Find the target Worksheet
    target_ws = None
    for ws in root.findall(f'.//{ns}Worksheet'):
        for key, val in ws.attrib.items():
            if key.endswith('Name') and val == sheet_name:
                target_ws = ws
                break
        if target_ws is not None:
            break
            
    if target_ws is None:
        raise ValueError(f"Sheet '{sheet_name}' not found in the XML.")
        
    # 2. Extract Data from Rows and Cells
    table = target_ws.find(f'.//{ns}Table')
    data = []
    max_cols = 0
    
    for row in table.findall(f'.//{ns}Row'):
        row_data = []
        current_col_idx = 1
        
        for cell in row.findall(f'.//{ns}Cell'):
            # Handle XML 'ss:Index' which skips over completely empty cells
            for key, val in cell.attrib.items():
                if key.endswith('Index'):
                    current_col_idx = int(val)
                    
            # Pad row_data with None if we skipped columns
            while len(row_data) < (current_col_idx - 1):
                row_data.append(None)
                
            # Grab the actual cell data (Added safe strip to prevent NoneType errors and whitespace)
            data_elem = cell.find(f'.//{ns}Data')
            val = data_elem.text.strip() if (data_elem is not None and data_elem.text) else None
            row_data.append(val)
            
            current_col_idx += 1
            
        if len(row_data) > max_cols:
            max_cols = len(row_data)
        data.append(row_data)
        
    # 3. Ensure uniform row lengths to prevent pandas framing errors
    for r in data:
        while len(r) < max_cols:
            r.append(None)
            
    if not data:
        return pd.DataFrame()
    
        
    # Use first row as headers
    df = pd.DataFrame(data[1:], columns=data[0])
    return df