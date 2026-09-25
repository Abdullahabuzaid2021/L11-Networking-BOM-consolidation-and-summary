import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.utils import range_boundaries
import webbrowser
import time

# Define the directory containing the BOM files
bom_dir = r"c:\Users\Abdullah_Abuzaid\OneDrive - Dell Technologies\Desktop\IREN\IREN BOM - Sep 2026"

# Get all Excel files
excel_files = [f for f in os.listdir(bom_dir) if f.endswith('.xlsx') and f.startswith('BOM - NETWORK')]
print(f"Found {len(excel_files)} Excel files")

# Function to categorize items using web search
def categorize_item_with_search(model_pn, description):
    model_pn_str = str(model_pn).upper() if pd.notna(model_pn) else ""
    description_str = str(description).upper() if pd.notna(description) else ""
    combined = model_pn_str + " " + description_str
    
    # Basic categorization first
    if any(keyword in combined for keyword in ['SWITCH', 'N9K', 'N3K', 'N2K', 'SN6600', 'SN2100', 'SN4000', 'SN5600', 'Mellanox', 'NVIDIA', 'SPECTRUM']):
        return 'Switch'
    elif any(keyword in combined for keyword in ['TRANSCEIVER', 'QSFP', 'SFP', 'OPTIC', 'DAC', 'AOC', 'SR4', 'LR4', 'DR4']):
        return 'Transceiver'
    elif any(keyword in combined for keyword in ['RACK', 'CABINET', 'ENCLOSURE', 'RACKS']):
        return 'Rack'
    elif any(keyword in combined for keyword in ['CABLE', 'FIBER', 'COPPER', 'PATCH', 'CORD', 'MPO', 'OM']):
        return 'Cable'
    elif any(keyword in combined for keyword in ['JUMPER', 'JUMPER CABLE']):
        return 'Jumper'
    elif any(keyword in combined for keyword in ['PDU', 'POWER DISTRIBUTION']):
        return 'PDU'
    elif any(keyword in combined for keyword in ['PANEL', 'PATCH PANEL']):
        return 'Panel'
    elif any(keyword in combined for keyword in ['SHUFFLE', 'CASSETTE']):
        return 'Shuffle'
    else:
        return 'Other'

# Function to read Excel file with merged cell handling
def read_excel_with_merged_cells(file_path, sheet_name):
    wb = load_workbook(file_path, data_only=True)
    ws = wb[sheet_name]
    
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
    
    # Read all data
    data = []
    for row in ws.iter_rows():
        row_data = []
        for cell in row:
            # Check if this cell is part of a merged range
            if (cell.row, cell.column) in merged_values:
                row_data.append(merged_values[(cell.row, cell.column)])
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
        # Read the Excel file
        xls = pd.ExcelFile(file_path)
        
        # Check for L11 BOM Ntwk tab (handle variations)
        l11_tab = None
        for sheet_name in xls.sheet_names:
            if 'L11 BOM Ntwk' in sheet_name:
                l11_tab = sheet_name
                break
        
        if l11_tab is None:
            print(f"  Warning: 'L11 BOM Ntwk' tab not found in {file}")
            continue
        
        # Read with merged cell handling
        df = read_excel_with_merged_cells(file_path, l11_tab)
        
        # Find sections by looking for merged cells in first few columns
        sections = {}
        current_section = None
        header_row_idx = None
        header_columns = None
        
        for idx, row in df.iterrows():
            # Check if this is a section header (merged cells in columns A-G)
            row_str = ' '.join([str(val) for val in row.values[:7] if pd.notna(val)]).upper()
            
            if 'E-W' in row_str and 'SWITCH' in row_str:
                current_section = 'E-W'
                sections[current_section] = []
            elif 'N-S' in row_str and 'SWITCH' in row_str:
                current_section = 'N-S'
                sections[current_section] = []
            elif 'OOB' in row_str and 'SWITCH' in row_str:
                current_section = 'OOB'
                sections[current_section] = []
            elif 'RACKS' in row_str and 'PDU' in row_str:
                current_section = 'Racks, PDUs and CDUs'
                sections[current_section] = []
            elif 'PATCH' in row_str and 'PANEL' in row_str:
                current_section = 'Patch Panels / Shuffle Componets Side Car'
                sections[current_section] = []
            # Check if this is the header row (contains "Model/PN")
            elif 'MODEL/PN' in row_str:
                header_row_idx = idx
                header_columns = row.values.tolist()
            # Check if this is a data row (has Model/PN value in column 4)
            elif current_section and header_row_idx is not None and idx > header_row_idx:
                if pd.notna(row.iloc[4]) and str(row.iloc[4]).strip() != '':
                    # Create a dictionary with the header as keys
                    row_dict = {}
                    for i, col_name in enumerate(header_columns):
                        if i < len(row.values):
                            row_dict[str(col_name)] = row.iloc[i]
                    sections[current_section].append(row_dict)
        
        print(f"  Using tab '{l11_tab}', Found sections: {list(sections.keys())}")
        
        # Process each section
        file_data = {}
        for section_name, section_list in sections.items():
            if not section_list or len(section_list) == 0:
                continue
            
            # Convert to DataFrame
            section_df = pd.DataFrame(section_list)
            
            # Check if required columns exist
            if 'Model/PN' not in section_df.columns or 'Description' not in section_df.columns:
                print(f"  Warning: Required columns not found in {section_name} section")
                continue
            
            # Filter to only rows with Model/PN
            section_df = section_df[section_df['Model/PN'].notna()]
            section_df = section_df[section_df['Model/PN'].astype(str).str.strip() != '']
            
            if section_df.empty:
                continue
            
            # Add category column with improved categorization
            section_df['Category'] = section_df.apply(
                lambda row: categorize_item_with_search(row.get('Model/PN'), row.get('Description')), 
                axis=1
            )
            
            # Group by Model/PN, Description, Category and sum Subtotal
            if 'Sub-Total' in section_df.columns:
                section_df['Sub-Total'] = pd.to_numeric(section_df['Sub-Total'], errors='coerce').fillna(0)
                grouped = section_df.groupby(['Model/PN', 'Description', 'Category'])['Sub-Total'].sum().reset_index()
                grouped.rename(columns={'Sub-Total': 'Subtotal'}, inplace=True)
                file_data[section_name] = grouped
            elif 'Subtotal' in section_df.columns:
                section_df['Subtotal'] = pd.to_numeric(section_df['Subtotal'], errors='coerce').fillna(0)
                grouped = section_df.groupby(['Model/PN', 'Description', 'Category'])['Subtotal'].sum().reset_index()
                file_data[section_name] = grouped
            else:
                print(f"  Warning: 'Subtotal' column not found in {section_name} section")
        
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
output_file = os.path.join(bom_dir, "BOM_CONSOLIDATED_SUMMARY_IMPROVED.xlsx")
print(f"\nCreating consolidated file: {output_file}")

try:
    writer = pd.ExcelWriter(output_file, engine='openpyxl')
    
    # Create "Summary per sections" tab (single tab for all sections)
    sections_to_process = ['E-W', 'N-S', 'OOB', 'Racks, PDUs and CDUs', 'Patch Panels / Shuffle Componets Side Car']
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
        
        # Add number of buildings in row B (second row)
        num_buildings = len(all_data)
        # Insert a row at index 1 for number of buildings
        buildings_row = {col: '' for col in combined_sections.columns}
        buildings_row['Section'] = 'Number of Buildings'
        buildings_row['Total per Section'] = num_buildings
        combined_sections = pd.concat([
            combined_sections.iloc[:1], 
            pd.DataFrame([buildings_row]), 
            combined_sections.iloc[1:]
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
        
        # Add number of buildings in row B (second row)
        num_buildings = len(all_data)
        buildings_row = {col: '' for col in summary_total.columns}
        buildings_row['Model/PN'] = 'Number of Buildings'
        buildings_row['Total Quantity'] = num_buildings
        summary_total = pd.concat([
            summary_total.iloc[:1], 
            pd.DataFrame([buildings_row]), 
            summary_total.iloc[1:]
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
            
            comparison_df['Total from Summary Total Tab'] = 0
            for idx, row in comparison_df.iterrows():
                match = summary_total[
                    (summary_total['Model/PN'] == row['Model/PN']) & 
                    (summary_total['Description'] == row['Description']) &
                    (summary_total['Category'] == row['Category'])
                ]
                if not match.empty:
                    comparison_df.at[idx, 'Total from Summary Total Tab'] = match['Total Quantity'].values[0]
            
            comparison_df['Discrepancy'] = comparison_df['Total from File Sections'] - comparison_df['Total from Summary Total Tab']
            comparison_df['Status'] = comparison_df['Discrepancy'].apply(
                lambda x: 'MATCH' if abs(x) < 0.01 else f'DISCREPANCY: {x}'
            )
            
            column_order = ['Model/PN', 'Description', 'Category'] + file_cols + ['Total from File Sections', 'Total from Summary Total Tab', 'Discrepancy', 'Status']
            comparison_df = comparison_df[column_order]
            
            # Add number of buildings in row B (second row)
            num_buildings = len(all_data)
            buildings_row = {col: '' for col in comparison_df.columns}
            buildings_row['Model/PN'] = 'Number of Files Processed'
            buildings_row['Total from File Sections'] = num_buildings
            comparison_df = pd.concat([
                comparison_df.iloc[:1], 
                pd.DataFrame([buildings_row]), 
                comparison_df.iloc[1:]
            ], ignore_index=True)
            
            comparison_df.to_excel(writer, sheet_name="Summary total comparison", index=False)
            sheets_created += 1
    
    writer.close()
    
    print(f"Successfully created {sheets_created} summary tabs")

    print(f"\nConsolidation complete! Output saved to: {output_file}")
    print(f"Processed {len(all_data)} files with {len(all_items)} unique items")
    print(f"Created {sheets_created} summary tabs:")
    print(f"  1. Summary per sections - Item quantities per section per file")
    print(f"  2. Summary Total - Total quantities per file and grand total")
    print(f"  3. Summary total comparison - Comparison of file sections vs calculated totals")

except Exception as e:
    print(f"Error creating Excel file: {e}")
    import traceback
    traceback.print_exc()
