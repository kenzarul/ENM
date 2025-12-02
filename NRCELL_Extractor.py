import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

def extract_parameters(input_file_path, cell_type):
    with open(input_file_path, 'r') as file:
        lines = file.readlines()

    found_section = False
    parameter_names = []

    current_parent_parameter = ""

    for line in lines:
        section_identifier = f"{cell_type}="

        if section_identifier in line:
            found_section = True
            continue

        if found_section:
            stripped_line = line.strip()

            if (
                stripped_line == "" or
                stripped_line.startswith(("===", "Total:", "INFO:")) or
                stripped_line[0].isdigit() or
                stripped_line[0] in {'X', 'G', 'E'}
            ):
                continue

            if stripped_line.startswith(">>>"):
                parts = stripped_line.split('.')
                if len(parts) > 1:
                    sub_parameter = parts[1].split('=')[0].strip()
                    parameter_names.append((current_parent_parameter,
sub_parameter))
            else:
                current_parent_parameter = stripped_line.split()[0].strip()
                parameter_names.append((None, current_parent_parameter))

    # Create DataFrame for the extracted parameters
    extracted_data = []
    for parent, sub in parameter_names:
        if parent:
            extracted_data.append({'Struct': parent, 'Parameter': sub})
        else:
            extracted_data.append({'Struct': '', 'Parameter': sub})

    return pd.DataFrame(extracted_data)

def normalize_parameter_name(param):
    """Normalize parameter name by removing extra spaces and
standardizing format"""
    if pd.isna(param):
        return ""
    # Remove extra spaces and normalize
    param = str(param).strip()
    # Replace multiple spaces with single space
    param = ' '.join(param.split())
    return param

def compare_with_excel_sheet(lab_params_df, excel_file_path,
cell_type, output_file_path):
    # Read parameters from LAB DataFrame
    lab_params = set()
    lab_params_dict = {}  # Store original parameter with struct info

    for _, row in lab_params_df.iterrows():
        struct = row['Struct'].strip() if pd.notna(row['Struct']) else ""
        param = row['Parameter'].strip() if pd.notna(row['Parameter']) else ""
        normalized_param = normalize_parameter_name(param)

        if struct:
            full_param = f"{struct}.{normalized_param}"
            lab_params.add(full_param)
            lab_params_dict[full_param] = {'struct': struct,
'subparam': normalized_param}
        else:
            lab_params.add(normalized_param)
            lab_params_dict[normalized_param] = {'struct': '',
'subparam': normalized_param}

    # Read Excel file
    SHEET_NAME = "LTE - NR parameters"
    PARAM_COL_NAME = "Parameter"
    MO_COL_NAME = "MO"

    try:
        # Read the Excel file
        df = pd.read_excel(excel_file_path, sheet_name=SHEET_NAME)

        # More precise filtering based on cell type
        if cell_type == "NRCellDU":
            # For NRCellDU, include all variations
            filtered_df = df[df[MO_COL_NAME].str.contains(r'^NRCellDU', na=False)]
        elif cell_type == "NRCellCU":
            # For NRCellCU, only exact NRCellCU (not containing commas for other MOs)
            filtered_df = df[df[MO_COL_NAME] == "NRCellCU"]
        else:
            # For other cell types, use exact match
            filtered_df = df[df[MO_COL_NAME] == cell_type]

        print(f"Found {len(filtered_df)} parameters for {cell_type} in Excel sheet")

        # Remove barred/obsolete parameters (assuming they are marked in some way)
        # Common indicators: "BAR", "OBSOLETE", "NOT USED", etc.
        barred_indicators = ["BAR", "OBSOLETE", "NOT USED", "DEPRECATED"]

        # Check if there's a column that might indicate barred status
        # If not, we'll assume parameters with certain patterns are barred
        status_column = None
        for col in df.columns:
            if any(indicator in col.upper() for indicator in ["STATUS", "REMARK", "NOTE"]):
                status_column = col
                break

        if status_column:
            # Filter out barred parameters based on status column
            original_count = len(filtered_df)
            filtered_df = filtered_df[~filtered_df[status_column].str.contains('|'.join(barred_indicators), na=False, case=False)]
            barred_count = original_count - len(filtered_df)
            print(f"Removed {barred_count} barred/obsolete parameters")
        else:
            # If no status column, check if parameter names contain barred indicators
            original_count = len(filtered_df)
            filtered_df = filtered_df[~filtered_df[PARAM_COL_NAME].str.contains('|'.join(barred_indicators), na=False, case=False)]
            barred_count = original_count - len(filtered_df)
            print(f"Removed {barred_count} barred/obsolete parameters based on parameter names")

        # Get expected parameters from Excel and normalize them
        expected_params_dict = {}
        for idx, row in filtered_df.iterrows():
            original_param = str(row[PARAM_COL_NAME]) if pd.notna(row[PARAM_COL_NAME]) else ""
            normalized_param = normalize_parameter_name(original_param)
            if normalized_param:  # Only add non-empty parameters
                expected_params_dict[normalized_param] = original_param

        expected_params = set(expected_params_dict.keys())

        # Find matching and missing parameters
        matching_params = lab_params.intersection(expected_params)
        missing_params = expected_params - lab_params

        # Create comparison DataFrame with clear column names
        comparison_data = []

        for param in sorted(matching_params):
            lab_info = lab_params_dict.get(param, {'struct': '', 'subparam': param})
            comparison_data.append({
                'Parameter_in_Excel': expected_params_dict.get(param, param),
                'Parameter_in_LAB': param,
                'Struct_in_LAB': lab_info['struct'],
                'Status': 'PRESENT'
            })

        for param in sorted(missing_params):
            comparison_data.append({
                'Parameter_in_Excel': expected_params_dict.get(param, param),
                'Parameter_in_LAB': 'NOT FOUND',
                'Struct_in_LAB': '',
                'Status': 'MISSING'
            })

        # Create DataFrame for comparison
        comparison_df = pd.DataFrame(comparison_data)

        # Reorder columns to have Status as the last column
        column_order = ['Parameter_in_Excel', 'Parameter_in_LAB', 'Struct_in_LAB', 'Status']
        comparison_df = comparison_df[column_order]

        # Create summary DataFrame
        summary_data = {
            'Category': ['Total Expected Parameters in file Excel', 'Parameters Excel Found in LAB', 'Missing Parameters'],
            'Count': [len(expected_params), len(matching_params), len(missing_params)]
        }
        summary_df = pd.DataFrame(summary_data)

        # Save all sheets to the same Excel file
        with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
            # Sheet 1: Extracted parameters from LAB
            lab_params_df.to_excel(writer, sheet_name='Extracted Parameters', index=False)

            # Sheet 2: Summary
            summary_df.to_excel(writer, sheet_name='Summary', index=False)

            # Sheet 3: Detailed comparison
            comparison_df.to_excel(writer, sheet_name='Detailed Comparison', index=False)

        # Apply formatting to highlight missing parameters in yellow
        apply_formatting_to_excel(output_file_path)

        print(f"Comparison completed successfully!")
        print(f"Total expected parameters in file Excel: {len(expected_params)}")
        print(f"Parameters Excel found in LAB: {len(matching_params)}")
        print(f"Missing parameters: {len(missing_params)}")
        print(f"All results saved to: {output_file_path}")

        return len(missing_params) == 0  # Return True if no missing parameters

    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return False

def apply_formatting_to_excel(file_path):
    """Apply yellow highlighting to missing parameters in the Detailed
Comparison sheet"""
    try:
        workbook = load_workbook(file_path)

        if 'Detailed Comparison' in workbook.sheetnames:
            sheet = workbook['Detailed Comparison']

            # Yellow fill for missing parameters
            yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')

            # Find the Status column
            header_row = 1
            status_col_idx = None

            for col_idx, cell in enumerate(sheet[header_row], 1):
                if cell.value == 'Status':
                    status_col_idx = col_idx
                    break

            if status_col_idx:
                # Apply formatting to all rows
                for row_idx in range(2, sheet.max_row + 1):  # Start from row 2 (after header)
                    status_cell = sheet.cell(row=row_idx, column=status_col_idx)
                    if status_cell.value == 'MISSING':
                        # Highlight the entire row for missing parameters
                        for col_idx in range(1, sheet.max_column + 1):
                            cell = sheet.cell(row=row_idx, column=col_idx)
                            cell.fill = yellow_fill

        workbook.save(file_path)
        print("Formatting applied successfully - missing parameters highlighted in yellow")

    except Exception as e:
        print(f"Error applying formatting: {e}")

def main():
    input_file_path = input("Please enter the path to the input LAB file: ").strip()
    cell_type = input("Please enter the cell type (NRCellDU or NRCellCU): ").strip()

    # Extract parameters from LAB file and create DataFrame
    lab_params_df = extract_parameters(input_file_path, cell_type)

    # Create single output file
    script_directory = os.path.dirname(__file__)
    output_file_name = f"Parameter_Analysis_{cell_type}.xlsx"
    output_file_path = os.path.join(script_directory, output_file_name)

    # Save extracted parameters to the Excel file first
    with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
        lab_params_df.to_excel(writer, sheet_name='Extracted Parameters in LAB', index=False)

    print(f"Parameters extracted and saved to: {output_file_path}")

    # Ask if user wants to compare with Excel sheet
    compare_choice = input("Do you want to compare with an Excel parameter sheet? (y/n): ").strip().lower()

    if compare_choice in ['y', 'yes']:
        excel_file_path = input("Please enter the path to the Excel parameter sheet: ").strip()

        # Perform comparison and add sheets to the same file
        success = compare_with_excel_sheet(lab_params_df, excel_file_path, cell_type, output_file_path)

    else:
        print(f"Only extracted parameters saved to: {output_file_path}")

if __name__ == "__main__":
    main()
