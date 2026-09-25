import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.utils import range_boundaries

# Define the directory containing the BOM files
bom_dir = r"c:\Users\Abdullah_Abuzaid\OneDrive - Dell Technologies\Desktop\IREN\IREN BOM - Sep 2026"

# Reference file path
reference_file = r"C:\Users\Abdullah_Abuzaid\OneDrive - Dell Technologies\Desktop\Dell Networking\Dell Networking components list Unified details WO Pricing.xlsx"

# Use the specific files provided by the user
excel_files = [
    "BOM - NETWORK - 2026-08-24- IREN -  512 Racks -(B300) - Mackenzie - 512 GPU Storage Testing [Locked]_UNLOCKED.xlsx",
    "BOM - NETWORK - 2026-09-04- IREN - PS - VR72 4x Test Racks [Quote]_UNLOCKED.xlsx",
    "BOM - NETWORK - 2026-09-11- IREN - 50MW 252 Racks - Sweetwater VR NVL72_SN6600-LD_CORE [Quoted]_UNLOCKED.xlsx",
    "BOM - NETWORK - 2026-09-11- IREN - 50MW 252 Racks - Sweetwater VR NVL72_SN6600-LD_DH [Quoted]_UNLOCKED.xlsx",
    "BOM - NETWORK - 2026-09-15- IREN CHLD_HORIZON_210 Racks - VR NVL72_SN6600-LD_Core-GOLDEN [Quoting]_UNLOCKED.xlsx",
    "BOM - NETWORK - 2026-09-15- IREN CHLD_HORIZON_210 Racks - VR NVL72_SN6600-LD_DH_GOLDEN [Quoting]_UNLOCKED.xlsx",
    "BOM - NETWORK - 2026-09-023 - IREN -  256 Racks - 8192-GPU (B300)_GOLDEN.xlsx",
    "BOM - NETWORK - 2026-09-23 - IREN -  256 Racks - 8192-GPU (B200).xlsx",
    "BOM - NETWORK - 2026-09-23 - IREN -  B4-5 256 Racks - 8192-GPU (B200).xlsx"
]

print(f"Using {len(excel_files)} files as specified")

# Load reference file for classification
print("Loading reference file for classification...")
reference_data = {}

try:
    # Load all sheets from reference file
    xls_ref = pd.ExcelFile(reference_file)
    
    for sheet_name in xls_ref.sheet_names:
        df_ref = pd.read_excel(reference_file, sheet_name=sheet_name, header=0)
        
        # Skip if the sheet doesn't have the expected columns
        if 'Category' not in df_ref.columns:
            continue
        
        # Build lookup dictionary
        for idx, row in df_ref.iterrows():
            category = str(row.get('Category', '')).strip() if pd.notna(row.get('Category')) else ''
            if not category:
                continue
            
            # Look up by Model/PN columns (D-H: Factory Install SKU, Cust Kit SKU, Cust Kit MOD, Model, Dell Part No)
            model_columns = ['Factory Install SKU', 'Cust Kit (APOS) SKU', 'Cust Kit (APOS) MOD', 'Model', 'Dell Part No']
            for col in model_columns:
                if col in df_ref.columns and pd.notna(row.get(col)):
                    model_value = str(row[col]).strip()
                    if model_value and model_value != 'NaN':
                        if model_value not in reference_data:
                            reference_data[model_value] = category
            
            # Look up by Description columns (J-K: Short Description, Long Description)
            desc_columns = ['Short Description', 'Long Description', 'Description']
            for col in desc_columns:
                if col in df_ref.columns and pd.notna(row.get(col)):
                    desc_value = str(row[col]).strip()
                    if desc_value and desc_value != 'NaN' and len(desc_value) > 10:
                        if desc_value not in reference_data:
                            reference_data[desc_value] = category
    
    print(f"Loaded {len(reference_data)} reference entries for classification")
    
except Exception as e:
    print(f"Warning: Could not load reference file: {e}")
    print("Will use keyword-based classification as fallback")

# Function to categorize items using reference file first, then fallback to keyword logic
def categorize_item_with_reference(model_pn, description):
    model_pn_str = str(model_pn).strip() if pd.notna(model_pn) else ""
    description_str = str(description).strip() if pd.notna(description) else ""
    
    # First try to find exact match in reference data by Model/PN
    if model_pn_str and model_pn_str in reference_data:
        return reference_data[model_pn_str]
    
    # Then try to find match by description (partial match)
    if description_str and len(description_str) > 10:
        for ref_key, category in reference_data.items():
            if len(ref_key) > 10 and ref_key in description_str or description_str in ref_key:
                return category
    
    # Fallback to keyword-based classification
    return categorize_item_refined(model_pn_str, description_str)

# Function to categorize items with refined logic (fallback)
def categorize_item_refined(model_pn, description):
    model_pn_str = str(model_pn).upper() if pd.notna(model_pn) else ""
    description_str = str(description).upper() if pd.notna(description) else ""
    combined = model_pn_str + " " + description_str
    
    # Network Switches - check for specific switch model numbers and keywords
    switch_model_patterns = [
        'SN6600', 'SN5600', 'SN4600', 'SN3800', 'SN3400', 'SN2700', 'SN2201', 'SN2100', 'SN2000',
        'SN4700', 'SN4200', 'SN4000', 'SN3700', 'SN3600', 'SN3500', 'SN3400', 'SN2700',
        'N9K', 'N3K', 'N2K', 'N5K', 'N7K', 'N9K-C', 'N3K-C',
        'VR72', 'VR', 'Mellanox', 'NVIDIA'
    ]
    
    switch_keywords = [
        'SWITCH', 'LEAF', 'SPINE', 'CORE', 'TOR', 'TOP OF RACK', 'AGGREGATOR', 
        'BORDER', 'ROUTER', 'GATEWAY', 'SPECTRUM', 'NPU', 'PORTS', 'QSFP PORTS',
        'SFP PORTS', 'OSFP PORTS', 'OPEN ETHERNET', 'CUMULUS', 'SONiC',
        '64 OSFP', '64 PORTS', '32 PORTS', '48 PORTS', '2U', '1U'
    ]
    
    # Transceivers - check for specific transceiver model patterns and keywords
    transceiver_model_patterns = [
        '980-9I', '980-', '920-9N', 'C4X6C', 'D3P20', '21RKC', '8R3F'
    ]
    
    transceiver_keywords = [
        'TRANSCEIVER', 'OPTIC', 'SR4', 'LR4', 'DR4',
        'SR8', 'LR8', 'FR4', 'CR4', 'CWDM', 'DWDM', 'OSFP', 'QSFPDD',
        'QSFP-DD', 'SFP28', 'SFP56', 'SFP+', 'XFP', 'XENPAK', 'GBIC',
        '800G', '400G', '200G', '100G', '40G', '25G', '10G', '1G',
        'LC-LC', 'LC-SC', 'SC-SC', '1310NM', '850NM', '1550NM'
    ]
    
    # Cables - check for cable-specific keywords (but not transceivers)
    cable_model_patterns = [
        '9VFNX', 'GEDEDU', 'EDGE-', 'CBL-'
    ]
    
    cable_keywords = [
        'CABLE', 'FIBER CABLE', 'COPPER CABLE', 'PATCH CABLE', 'CORD',
        'TWINAX', 'TWIN-AX', 'CAT6', 'CAT6A', 'CAT7', 'CAT8', 'CAT 6',
        'FIBER PATCH', 'COPPER PATCH', 'MPO CABLE', 'LC CABLE', 'SC CABLE',
        'OM3', 'OM4', 'OM5', 'OS2', 'MMF', 'SMF', 'APC', 'UPC',
        'MPO12', 'MPO8', 'MPO24', 'MPO16'
    ]
    
    # Racks
    rack_keywords = [
        'RACK', 'CABINET', 'ENCLOSURE', 'SHELF', 'CHASSIS', 'NtwkRack', 'NTWRACK',
        'GPU RACK', 'GPU RACKS', 'RAINWATER', 'MGX', 'IR7044', 'IR7050', 'IR9048', 'IR9148', 'IR9149', 'IR9053',
        '48U', '44U', '50U', '53H', '750MM', '1200MM'
    ]
    
    # Jumpers
    jumper_keywords = [
        'JUMPER', 'JUMPER CABLE', 'JUMPER CORD'
    ]
    
    # PDUs
    pdu_keywords = [
        'PDU', 'POWER DISTRIBUTION UNIT', 'POWER STRIP', '3PH', '415V', '60A'
    ]
    
    # Panels
    panel_keywords = [
        'PANEL', 'PATCH PANEL', 'FIBER PANEL', 'COPPER PANEL', 'FIBERPANEL',
        'FILLER PANEL', 'BLANKING PANEL'
    ]
    
    # Shuffle components
    shuffle_keywords = [
        'SHUFFLE', 'CASSETTE', 'SIDE CAR', 'EDGE SHUFFLING', 'HOUSING', 'MODULES'
    ]
    
    # Check in order of specificity to avoid misclassification
    
    # First check for shuffle components (highest priority to avoid misclassification)
    if any(keyword in description_str for keyword in shuffle_keywords):
        return 'Shuffle'
    
    # Then check for transceivers by model pattern (highest priority for 920-9N)
    if any(pattern in model_pn_str for pattern in transceiver_model_patterns):
        return 'Transceiver'
    
    # Then check for cables by model pattern
    if any(pattern in model_pn_str for pattern in cable_model_patterns):
        return 'Cable'
    
    # Then check for cables by keywords
    if any(keyword in combined for keyword in cable_keywords):
        return 'Cable'
    
    # Then check for switches by model pattern
    if any(pattern in model_pn_str for pattern in switch_model_patterns):
        return 'Switch'
    
    # Then check for switches by keywords
    if any(keyword in description_str for keyword in switch_keywords):
        return 'Switch'
    
    # Then check for transceivers by keywords
    if any(keyword in combined for keyword in transceiver_keywords):
        return 'Transceiver'
    
    # Then check for other categories
    if any(keyword in combined for keyword in rack_keywords):
        return 'Rack'
    elif any(keyword in combined for keyword in jumper_keywords):
        return 'Jumper'
    elif any(keyword in combined for keyword in pdu_keywords):
        return 'PDU'
    elif any(keyword in combined for keyword in panel_keywords):
        return 'Panel'
    else:
        return 'Other'

# Function to read Excel file with merged cell handling
def read_excel_with_merged_cells(file_path, sheet_name):
    wb = load_workbook(file_path, data_only=True)
    ws = wb[sheet_name]
    
    # Force calculation of all formulas
    ws.calculate_dimension()
    
    # Get merged cell ranges
    merged_cells = ws.merged_cells.ranges
    
    # Create a dictionary to map merged cell values
    merged_values = {}
    for merged_range in merged_cells:
        min_col, min_row, max_col, max_row = range_boundaries(merged_range.coord)
        top_left_value = ws.cell(row=min_row, column=min_col).value
        if top_left_value:
            for row in range(min_row, max_row + 1):
                for col in range(min_col, max_col + 1):
                    merged_values[(row, col)] = top_left_value
    
    # Read all data with merged cell handling
    data = []
    for row_idx, row in enumerate(ws.iter_rows(), start=1):
        row_data = []
        for col_idx, cell in enumerate(row, start=1):
            # Check if this cell is part of a merged range
            if (row_idx, col_idx) in merged_values:
                row_data.append(merged_values[(row_idx, col_idx)])
            else:
                row_data.append(cell.value)
        data.append(row_data)
    
    return pd.DataFrame(data)

# Dictionary to store all data
all_data = {}

# Process each file
for file in excel_files:
    file_path = os.path.join(bom_dir, file)
    print(f"\nProcessing: {file}")
    
    try:
        # Check for L11 BOM Ntwk tab (handle variations)
        xls = pd.ExcelFile(file_path)
        l11_tab = None
        for sheet_name in xls.sheet_names:
            if 'L11 BOM Ntwk' in sheet_name:
                l11_tab = sheet_name
                break
        
        if l11_tab is None:
            print(f"  Warning: 'L11 BOM Ntwk' tab not found in {file}")
            continue
        
        # Read with merged cell handling and formula calculation
        df = read_excel_with_merged_cells(file_path, l11_tab)
        
        # Read all data rows after header
        all_data_rows = []
        header_row_idx = None
        header_columns = None
        
        for idx, row in df.iterrows():
            row_str = ' '.join([str(val) for val in row.values[:7] if pd.notna(val)]).upper()
            
            # Check if this is the header row (contains "Model/PN")
            if 'MODEL/PN' in row_str:
                header_row_idx = idx
                header_columns = row.values.tolist()
            # Collect all data rows after header
            elif header_row_idx is not None and idx > header_row_idx:
                row_dict = {}
                for i, col_name in enumerate(header_columns):
                    if i < len(row.values):
                        row_dict[str(col_name)] = row.iloc[i]
                all_data_rows.append(row_dict)
        
        # Assign rows to sections based on their content
        sections = {
            'E-W': [],
            'N-S': [],
            'OOB': [],
            'Racks, PDUs and CDUs': [],
            'Patch Panels / Shuffle Componets Side Car': [],
            'East-West NIC-L': [],
            'North-South NIC-L': []
        }
        
        for row_dict in all_data_rows:
            item = str(row_dict.get('Item', '')).upper()
            model_pn = str(row_dict.get('Model/PN', '')).upper()
            
            # Assign to section based on Item content
            if 'E-W' in item or 'EAST-WEST' in item:
                sections['E-W'].append(row_dict)
            elif 'N-S' in item or 'NORTH-SOUTH' in item:
                sections['N-S'].append(row_dict)
            elif 'OOB' in item:
                sections['OOB'].append(row_dict)
            elif 'RACK' in item or 'PDU' in item:
                sections['Racks, PDUs and CDUs'].append(row_dict)
            elif 'PATCH' in item or 'PANEL' in item or 'SHUFFLE' in item:
                sections['Patch Panels / Shuffle Componets Side Car'].append(row_dict)
            elif 'NIC' in item:
                if 'EAST' in item or 'E-W' in item:
                    sections['East-West NIC-L'].append(row_dict)
                elif 'NORTH' in item or 'N-S' in item:
                    sections['North-South NIC-L'].append(row_dict)
                else:
                    sections['E-W'].append(row_dict)  # Default to E-W
            else:
                # Default to E-W if no section identified
                sections['E-W'].append(row_dict)
        
        print(f"  Using tab '{l11_tab}', Found sections: {list(sections.keys())}")
        
        # Process each section
        file_data = {}
        for section_name, section_list in sections.items():
            if not section_list or len(section_list) == 0:
                continue
            
            # Convert to DataFrame
            section_df = pd.DataFrame(section_list)
            
            # Filter out rows without Model/PN or quantities
            def has_valid_data(row):
                model_pn = str(row.get('Model/PN', '')).strip()
                item = str(row.get('Item', '')).strip()
                count = row.get('Count')
                subtotal = row.get('Sub-Total')
                
                # Check if it has a valid Model/PN (not a header)
                has_model_pn = model_pn and model_pn != 'nan' and model_pn != 'Model/PN' and model_pn != 'Item'
                has_item = item and item != 'nan' and item != 'Item'
                has_quantity = (pd.notna(count) and str(count).strip() != '' and str(count).strip() != '0') or (pd.notna(subtotal) and str(subtotal).strip() != '' and str(subtotal).strip() != '0')
                
                return has_model_pn or has_item or has_quantity
            
            section_df = section_df[section_df.apply(has_valid_data, axis=1)]
            
            # If Model/PN is in column A, move it to the Model/PN column
            if 'Item' in section_df.columns:
                section_df['Model/PN'] = section_df.apply(
                    lambda row: row['Item'] if (pd.isna(row['Model/PN']) or str(row['Model/PN']).strip() == '' or str(row['Model/PN']).strip() == 'Model/PN' or str(row['Model/PN']).strip() == 'nan') else row['Model/PN'],
                    axis=1
                )
            
            # Clean up Model/PN - extract actual part number if it contains DPN info
            if 'Model/PN' in section_df.columns:
                section_df['Model/PN'] = section_df['Model/PN'].apply(
                    lambda x: str(x).split('(DPN:')[0].strip() if '(DPN:' in str(x) else str(x)
                )
            
            # Check if required columns exist
            if 'Model/PN' not in section_df.columns or 'Description' not in section_df.columns:
                print(f"  Warning: Required columns not found in {section_name} section")
                continue
            
            # Filter to only rows with Model/PN or Item (allow either column)
            # Create a mask for rows that have either Model/PN or Item
            has_model_pn = section_df['Model/PN'].notna() & (section_df['Model/PN'].astype(str).str.strip() != '') & (section_df['Model/PN'].astype(str).str.strip() != 'Model/PN')
            has_item = section_df['Item'].notna() & (section_df['Item'].astype(str).str.strip() != '') & (section_df['Item'].astype(str).str.strip() != 'Item')
            
            # Keep rows that have either Model/PN or Item
            section_df = section_df[has_model_pn | has_item]
            
            # If Model/PN is empty but Item has value, use Item as Model/PN
            if 'Item' in section_df.columns:
                section_df['Model/PN'] = section_df.apply(
                    lambda row: row['Item'] if (pd.isna(row['Model/PN']) or str(row['Model/PN']).strip() == '' or str(row['Model/PN']).strip() == 'Model/PN') else row['Model/PN'],
                    axis=1
                )
            
            # Also update Description if it's empty but Item had a value
            if 'Description' in section_df.columns and 'Item' in section_df.columns:
                section_df['Description'] = section_df.apply(
                    lambda row: row['Item'] if (pd.isna(row['Description']) or str(row['Description']).strip() == '') else row['Description'],
                    axis=1
                )
            
            if section_df.empty:
                continue
            
            # Add category column with reference-based classification
            section_df['Category'] = section_df.apply(
                lambda row: categorize_item_with_reference(row.get('Model/PN'), row.get('Description')), 
                axis=1
            )
            
            # Group by Model/PN, Description, Category and sum Subtotal
            # Check both column B (Count) and column D (Sub-Total) for quantities
            if 'Sub-Total' in section_df.columns:
                # Convert Sub-Total to numeric (handle empty strings and None as 0)
                section_df['Sub-Total'] = section_df['Sub-Total'].replace('', 0)
                section_df['Sub-Total'] = section_df['Sub-Total'].fillna(0)
                section_df['Sub-Total'] = pd.to_numeric(section_df['Sub-Total'], errors='coerce').fillna(0)
                
                # If Count column exists, use it as fallback when Sub-Total is 0 or NaN
                if 'Count' in section_df.columns:
                    section_df['Count'] = section_df['Count'].replace('', 0)
                    section_df['Count'] = section_df['Count'].fillna(0)
                    section_df['Count'] = pd.to_numeric(section_df['Count'], errors='coerce').fillna(0)
                    # Use Count when Sub-Total is 0 or NaN
                    section_df['FinalQty'] = section_df.apply(
                        lambda row: row['Count'] if row['Sub-Total'] == 0 or pd.isna(row['Sub-Total']) else row['Sub-Total'],
                        axis=1
                    )
                else:
                    section_df['FinalQty'] = section_df['Sub-Total']
                
                grouped = section_df.groupby(['Model/PN', 'Category'])['FinalQty'].sum().reset_index()
                # Get the most common description for each Model/PN
                desc_map = section_df.groupby('Model/PN')['Description'].first().to_dict()
                grouped['Description'] = grouped['Model/PN'].map(desc_map)
                grouped.rename(columns={'FinalQty': 'Subtotal'}, inplace=True)
                # Don't filter out items with zero quantity - keep all items
                file_data[section_name] = grouped
            elif 'Subtotal' in section_df.columns:
                # Convert Subtotal to numeric (handle empty strings and None as 0)
                section_df['Subtotal'] = section_df['Subtotal'].replace('', 0)
                section_df['Subtotal'] = section_df['Subtotal'].fillna(0)
                section_df['Subtotal'] = pd.to_numeric(section_df['Subtotal'], errors='coerce').fillna(0)
                
                # If Count column exists, use it as fallback when Subtotal is 0 or NaN
                if 'Count' in section_df.columns:
                    section_df['Count'] = section_df['Count'].replace('', 0)
                    section_df['Count'] = section_df['Count'].fillna(0)
                    section_df['Count'] = pd.to_numeric(section_df['Count'], errors='coerce').fillna(0)
                    # Use Count when Subtotal is 0 or NaN
                    section_df['FinalQty'] = section_df.apply(
                        lambda row: row['Count'] if row['Subtotal'] == 0 or pd.isna(row['Subtotal']) else row['Subtotal'],
                        axis=1
                    )
                else:
                    section_df['FinalQty'] = section_df['Subtotal']
                
                grouped = section_df.groupby(['Model/PN', 'Category'])['FinalQty'].sum().reset_index()
                # Get the most common description for each Model/PN
                desc_map = section_df.groupby('Model/PN')['Description'].first().to_dict()
                grouped['Description'] = grouped['Model/PN'].map(desc_map)
                grouped.rename(columns={'FinalQty': 'Subtotal'}, inplace=True)
                # Don't filter out items with zero quantity - keep all items
                file_data[section_name] = grouped
            elif 'Count' in section_df.columns:
                # If only Count column exists, use it (handle empty strings and None as 0)
                section_df['Count'] = section_df['Count'].replace('', 0)
                section_df['Count'] = section_df['Count'].fillna(0)
                section_df['Count'] = pd.to_numeric(section_df['Count'], errors='coerce').fillna(0)
                grouped = section_df.groupby(['Model/PN', 'Description', 'Category'])['Count'].sum().reset_index()
                grouped.rename(columns={'Count': 'Subtotal'}, inplace=True)
                # Don't filter out items with zero quantity - keep all items
                file_data[section_name] = grouped
            else:
                # Try to extract quantity from Description if no numeric column
                if 'Description' in section_df.columns:
                    def extract_qty_from_desc(desc):
                        import re
                        if pd.isna(desc):
                            return 0
                        desc_str = str(desc)
                        # Look for patterns like "432 x 18" or "= 432"
                        match = re.search(r'(\d+)\s*[xX]\s*(\d+)', desc_str)
                        if match:
                            return int(match.group(1)) * int(match.group(2))
                        match = re.search(r'=\s*(\d+)', desc_str)
                        if match:
                            return int(match.group(1))
                        return 0
                    
                    section_df['Subtotal'] = section_df['Description'].apply(extract_qty_from_desc)
                    grouped = section_df.groupby(['Model/PN', 'Description', 'Category'])['Subtotal'].sum().reset_index()
                    # Filter out items with zero quantity
                    grouped = grouped[grouped['Subtotal'] > 0]
                    file_data[section_name] = grouped
        
        all_data[file] = file_data
        
    except Exception as e:
        print(f"  Error processing {file}: {e}")
        import traceback
        traceback.print_exc()
        continue

print(f"\nSuccessfully processed {len(all_data)} files")

if not all_data:
    print("No data was processed. Please check the file structure.")
    exit()

# Create the consolidated Excel file
output_file = os.path.join(bom_dir, "BOM_CONSOLIDATED_SUMMARY_WITH_REFERENCE.xlsx")
print(f"\nCreating consolidated file: {output_file}")

try:
    writer = pd.ExcelWriter(output_file, engine='openpyxl')
    
    # Create "Summary per sections" tab (single tab for all sections)
    sections_to_process = ['E-W', 'N-S', 'OOB', 'Racks, PDUs and CDUs', 'Patch Panels / Shuffle Componets Side Car', 'East-West NIC-L', 'North-South NIC-L']
    sheets_created = 0
    
    # Create a combined dataframe for all sections
    combined_sections = pd.DataFrame()
    
    for section in sections_to_process:
        # Check if this section exists in any file
        section_exists = any(section in file_data for file_data in all_data.values())
        if not section_exists:
            continue
        
        # Get only items that exist in this specific section across all files
        section_items = {}
        for file, file_data in all_data.items():
            if section in file_data:
                for _, row in file_data[section].iterrows():
                    key = (row['Model/PN'], row['Description'], row['Category'])
                    if key not in section_items:
                        section_items[key] = {'Model/PN': row['Model/PN'], 'Description': row['Description'], 'Category': row['Category']}
        
        if not section_items:
            continue
        
        # Create section summary
        items_list = []
        for key, item_info in section_items.items():
            item_dict = {
                'Section': section,
                'Model/PN': item_info['Model/PN'],
                'Description': item_info['Description'],
                'Category': item_info['Category']
            }
            
            # Add file columns
            for file in all_data.keys():
                if section in all_data[file]:
                    matching = all_data[file][section]
                    match = matching[
                        (matching['Model/PN'] == key[0]) & 
                        (matching['Description'] == key[1]) & 
                        (matching['Category'] == key[2])
                    ]
                    if not match.empty:
                        qty = match['Subtotal'].values[0]
                        try:
                            item_dict[file] = float(qty) if pd.notna(qty) else 0
                        except (ValueError, TypeError):
                            item_dict[file] = 0
                    else:
                        item_dict[file] = 0
                else:
                    item_dict[file] = 0
            
            items_list.append(item_dict)
        
        if items_list:
            section_df = pd.DataFrame(items_list)
            
            # Add total column per section
            file_cols = [col for col in section_df.columns if col in all_data.keys()]
            section_df['Total per Section'] = section_df[file_cols].sum(axis=1)
            
            # Filter out items with zero total per section
            section_df = section_df[section_df['Total per Section'] > 0]
            
            # Append to combined sections
            if combined_sections.empty:
                combined_sections = section_df
            else:
                combined_sections = pd.concat([combined_sections, section_df], ignore_index=True)
    
    if not combined_sections.empty:
        # Reorder columns: Section, Model/PN, Description, Category, then file columns, then Total per Section
        file_cols = [col for col in combined_sections.columns if col in all_data.keys()]
        column_order = ['Section', 'Model/PN', 'Description', 'Category'] + file_cols + ['Total per Section']
        combined_sections = combined_sections[column_order]
        
        # Add number of buildings in row #1 (index 0)
        num_buildings = len(all_data)
        # Insert a row at index 0 for number of buildings
        buildings_row = {col: '' for col in combined_sections.columns}
        buildings_row['Section'] = 'Number of Buildings'
        buildings_row['Total per Section'] = num_buildings
        combined_sections = pd.concat([
            pd.DataFrame([buildings_row]), 
            combined_sections
        ], ignore_index=True)
        
        # Write to sheet
        combined_sections.to_excel(writer, sheet_name="Summary per sections", index=False)
        sheets_created += 1
    
    # Create "Summary Total" tab
    # Get all unique items across all files and sections
    all_items = {}
    for file, file_data in all_data.items():
        for section, section_df in file_data.items():
            for _, row in section_df.iterrows():
                key = (row['Model/PN'], row['Description'], row['Category'])
                if key not in all_items:
                    all_items[key] = {'Model/PN': row['Model/PN'], 'Description': row['Description'], 'Category': row['Category']}
    
    if all_items:
        # Rebuild items list for Summary Total
        items_list = []
        for key, item_info in all_items.items():
            items_list.append(item_info)
        summary_total = pd.DataFrame(items_list)
        
        # Add a single column for total quantity per file (sum across all sections)
        for file in all_data.keys():
            file_col = []
            for key in all_items.keys():
                total_qty = 0
                for section in all_data[file].keys():
                    matching = all_data[file][section]
                    match = matching[
                        (matching['Model/PN'] == key[0]) & 
                        (matching['Description'] == key[1]) & 
                        (matching['Category'] == key[2])
                    ]
                    if not match.empty:
                        qty = match['Subtotal'].values[0]
                        try:
                            total_qty += float(qty) if pd.notna(qty) else 0
                        except (ValueError, TypeError):
                            total_qty += 0
                file_col.append(total_qty)
            summary_total[file] = file_col
        
        # Add total column
        file_cols = [col for col in summary_total.columns if col in all_data.keys()]
        for col in file_cols:
            summary_total[col] = pd.to_numeric(summary_total[col], errors='coerce').fillna(0)
        
        try:
            summary_total['Total Quantity'] = summary_total[file_cols].sum(axis=1)
            summary_total['Total Quantity'] = pd.to_numeric(summary_total['Total Quantity'], errors='coerce').fillna(0)
        except Exception as e:
            print(f"Error calculating total quantity: {e}")
        
        # Filter out items with zero total quantity
        summary_total = summary_total[summary_total['Total Quantity'] > 0]
        
        # Add number of buildings in row #1 (index 0)
        num_buildings = len(all_data)
        buildings_row = {col: '' for col in summary_total.columns}
        buildings_row['Model/PN'] = 'Number of Buildings'
        buildings_row['Total Quantity'] = num_buildings
        summary_total = pd.concat([
            pd.DataFrame([buildings_row]), 
            summary_total
        ], ignore_index=True)
        
        # Write to sheet
        summary_total.to_excel(writer, sheet_name="Summary Total", index=False)
        sheets_created += 1
    
    # Create "Summary total comparison" tab
    all_items_comparison = {}
    for file, file_data in all_data.items():
        for section, section_df in file_data.items():
            for _, row in section_df.iterrows():
                key = (row['Model/PN'], row['Description'], row['Category'])
                if key not in all_items_comparison:
                    all_items_comparison[key] = {'Model/PN': row['Model/PN'], 'Description': row['Description'], 'Category': row['Category']}
    
    if all_items_comparison:
        comparison_df = pd.DataFrame()
        
        items_list = []
        for key, item_info in all_items_comparison.items():
            item_dict = {
                'Model/PN': item_info['Model/PN'],
                'Description': item_info['Description'],
                'Category': item_info['Category']
            }
            
            for file in all_data.keys():
                total_qty = 0
                for section in all_data[file].keys():
                    matching = all_data[file][section]
                    match = matching[
                        (matching['Model/PN'] == key[0]) & 
                        (matching['Description'] == key[1]) & 
                        (matching['Category'] == key[2])
                    ]
                    if not match.empty:
                        qty = match['Subtotal'].values[0]
                        try:
                            total_qty += float(qty) if pd.notna(qty) else 0
                        except (ValueError, TypeError):
                            total_qty += 0
                item_dict[file] = total_qty
            
            items_list.append(item_dict)
        
        if items_list:
            comparison_df = pd.DataFrame(items_list)
            
            file_cols = [col for col in comparison_df.columns if col in all_data.keys()]
            for col in file_cols:
                comparison_df[col] = pd.to_numeric(comparison_df[col], errors='coerce').fillna(0)
            comparison_df['Total from File Sections'] = comparison_df[file_cols].sum(axis=1)
            
            comparison_df['Total from Summary Total Tab'] = 0.0
            comparison_df['Total from Summary Total Tab'] = comparison_df['Total from Summary Total Tab'].astype(float)
            for idx, row in comparison_df.iterrows():
                match = summary_total[
                    (summary_total['Model/PN'] == row['Model/PN']) & 
                    (summary_total['Description'] == row['Description']) &
                    (summary_total['Category'] == row['Category'])
                ]
                if not match.empty:
                    comparison_df.at[idx, 'Total from Summary Total Tab'] = float(match['Total Quantity'].values[0])
            
            comparison_df['Discrepancy'] = comparison_df['Total from File Sections'] - comparison_df['Total from Summary Total Tab']
            comparison_df['Status'] = comparison_df['Discrepancy'].apply(
                lambda x: 'MATCH' if abs(x) < 0.01 else f'DISCREPANCY: {x}'
            )
            
            column_order = ['Model/PN', 'Description', 'Category'] + file_cols + ['Total from File Sections', 'Total from Summary Total Tab', 'Discrepancy', 'Status']
            comparison_df = comparison_df[column_order]
            
            # Add number of buildings in row #1 (index 0)
            num_buildings = len(all_data)
            buildings_row = {col: '' for col in comparison_df.columns}
            buildings_row['Model/PN'] = 'Number of Files Processed'
            buildings_row['Total from File Sections'] = num_buildings
            comparison_df = pd.concat([
                pd.DataFrame([buildings_row]), 
                comparison_df
            ], ignore_index=True)
            
            comparison_df.to_excel(writer, sheet_name="Summary total comparison", index=False)
            sheets_created += 1
    
    # Tab 4: Summary table total (NEW - extracts SUMMARY TABLE section from each file)
    summary_table_data = {}
    for file in all_data.keys():
        file_path = os.path.join(bom_dir, file)
        try:
            wb = load_workbook(file_path, data_only=True)
            l11_tab = None
            for sheet_name in wb.sheetnames:
                if 'L11 BOM Ntwk' in sheet_name:
                    l11_tab = sheet_name
                    break
            
            if l11_tab:
                ws = wb[l11_tab]
                df = pd.read_excel(file_path, sheet_name=l11_tab, header=None)
                
                # Find SUMMARY TABLE section
                summary_table_start = None
                for idx, row in df.iterrows():
                    row_str = ' '.join([str(val) for val in row.values if pd.notna(val)]).upper()
                    if 'SUMMARY TABLE' in row_str:
                        summary_table_start = idx
                        break
                
                if summary_table_start is not None and summary_table_start + 1 < len(df):
                    # Read the summary table section
                    summary_data = df.iloc[summary_table_start + 1:].copy()
                    # Find the header row (look for row with Model/PN or Item)
                    header_row = None
                    for idx, row in summary_data.iterrows():
                        row_str = ' '.join([str(val) for val in row.values if pd.notna(val)]).upper()
                        if 'MODEL/PN' in row_str or 'ITEM' in row_str:
                            header_row = idx
                            break
                    
                    if header_row is not None and header_row + 1 < len(summary_data):
                        summary_data.columns = summary_data.iloc[header_row].values
                        summary_data = summary_data.iloc[header_row + 1:].reset_index(drop=True)
                        
                        # Extract Model/PN and quantity
                        for idx, row in summary_data.iterrows():
                            model_pn = row.get('Model/PN') or row.get('Item')
                            description = row.get('Description')
                            quantity = row.get('Sub-Total') or row.get('Subtotal') or row.get('Count') or row.get('Quantity')
                            
                            if model_pn and pd.notna(quantity):
                                try:
                                    qty = float(quantity)
                                    if model_pn not in summary_table_data:
                                        summary_table_data[model_pn] = {
                                            'Model/PN': model_pn,
                                            'Description': description if pd.notna(description) else ''
                                        }
                                    summary_table_data[model_pn][file] = qty
                                except (ValueError, TypeError):
                                    pass
            
            wb.close()
        except Exception as e:
            print(f"  Warning: Could not read SUMMARY TABLE from {file}: {e}")
    
    if summary_table_data:
        summary_table_df = pd.DataFrame.from_dict(summary_table_data, orient='index')
        summary_table_df.reset_index(drop=True, inplace=True)
        
        # Reorder columns
        file_cols = [col for col in summary_table_df.columns if col in all_data.keys()]
        other_cols = [col for col in summary_table_df.columns if col not in file_cols]
        column_order = other_cols + file_cols
        summary_table_df = summary_table_df[column_order]
        
        # Add total column
        summary_table_df['Total'] = summary_table_df[file_cols].sum(axis=1)
        
        # Add number of files row
        num_buildings = len(all_data)
        buildings_row = {col: '' for col in summary_table_df.columns}
        buildings_row['Model/PN'] = 'Number of Files Processed'
        buildings_row['Total'] = num_buildings
        summary_table_df = pd.concat([
            pd.DataFrame([buildings_row]), 
            summary_table_df
        ], ignore_index=True)
        
        summary_table_df.to_excel(writer, sheet_name="Summary table total", index=False)
        sheets_created += 1
    
    writer.close()
    
    print(f"Successfully created {sheets_created} summary tabs")

    print(f"\nConsolidation complete! Output saved to: {output_file}")
    print(f"Processed {len(all_data)} files with {len(all_items)} unique items")
    print(f"Created {sheets_created} summary tabs:")
    print(f"  1. Summary per sections - Item quantities per section per file")
    print(f"  2. Summary Total - Total quantities per file and grand total")
    print(f"  3. Summary total comparison - Comparison of file sections vs calculated totals")
    print(f"  4. Summary table total - Summary Total section from each file")

except Exception as e:
    print(f"Error creating Excel file: {e}")
    import traceback
    traceback.print_exc()
