import os
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Border, Side
from openpyxl.cell.cell import MergedCell
import zipfile
import re


def try_open_excel(path):
    candidates = [path]
    if not os.path.splitext(path)[1]:
        candidates += [path + ext for ext in (".xlsx", ".xls", ".xlsm")]
    for p in candidates:
        if os.path.exists(p):
            try:
                if p.endswith(('.xlsx', '.xlsm')):
                    with zipfile.ZipFile(p, 'r') as zf:
                        pass
                return p
            except (zipfile.BadZipFile, Exception):
                continue
    return None


def is_valid_excel_file(file_path):
    """Check if the file is a valid Excel file"""
    try:
        if file_path.endswith(('.xlsx', '.xlsm')):
            with zipfile.ZipFile(file_path, 'r') as zf:
                return True
        elif file_path.endswith('.xls'):
            pd.read_excel(file_path, nrows=1, engine='xlrd')
            return True
        else:
            pd.read_excel(file_path, nrows=1)
            return True
    except Exception as e:
        print(f"❌ File validation failed for {file_path}: {e}")
        return False


def get_node_type(node_name):
    """Extract node type from NeName (E, X, G)"""
    if node_name == "NOT FOUND":
        return None
    
    if isinstance(node_name, str):
        if node_name.startswith('E'):
            return 'E'
        elif node_name.startswith('X'):
            return 'X'
        elif node_name.startswith('G'):
            return 'G'
    return None


def should_feature_be_active(activation_rule, node_type, feature_state):
    """
    Determine if feature should be active based on activation rule and node type
    Returns: True if should be active, False if should be inactive, None if rule doesn't apply
    """
    if not activation_rule or not node_type:
        return None
    
    activation_rule = str(activation_rule).lower().strip()
    
    # Rules that require activation
    activate_rules = [
        "a activer sur sites e et x éligibles à la dual-co",
        "a activer sur les bb configurées en mixed mode",
        "a activer sur les bb configurées pour supporter mode ess",
        "a activer sur site x",
        "a activer sur sites g et x",
        "a activer pour tests",
        "a activer au cas par cas",
        "a activer sur site g en crz",
        "généralisé en crz",
        "a installer",
        "a activer sur sur sites e et x en ztd",
        "a activer sur sites g et x à partir de"
    ]
    
    # Rules that require deactivation
    deactivate_rules = [
        "a désactiver sur sites e et x",
        "ne pas activer",
        "ne pas activer par défaut"
    ]
    
    # Check activation rules
    for rule in activate_rules:
        if rule in activation_rule:
            if "site x" in activation_rule and "site g" not in activation_rule:
                return node_type == 'X'
            elif "sites g et x" in activation_rule or "site g + x" in activation_rule:
                return node_type in ['G', 'X']
            elif "sites e et x" in activation_rule:
                return node_type in ['E', 'X']
            elif "site g" in activation_rule and "site x" not in activation_rule:
                return node_type == 'G'
            else:
                return True
    
    # Check deactivation rules
    for rule in deactivate_rules:
        if rule in activation_rule:
            return False
    
    # Special case for n/a
    if "n/a" in activation_rule:
        return None
    
    return None


def validate_feature_state(activation_rule, node_type, actual_feature_state):
    """
    Validate if the actual featureState matches what it should be based on activation rule
    Returns: "CORRECT", "INCORRECT", or "UNKNOWN"
    """
    expected_active = should_feature_be_active(activation_rule, node_type, actual_feature_state)
    
    if expected_active is None:
        return "UNKNOWN"
    
    # featureState: 1 = active, 0 = inactive
    if expected_active and actual_feature_state == 1:
        return "CORRECT"
    elif not expected_active and actual_feature_state == 0:
        return "CORRECT"
    else:
        return "INCORRECT"


# Get parameter file path
param_file = input("Enter path to parameter Excel file (updated file with Features + Licenses sheet): ").strip()
param_path = try_open_excel(param_file)

if param_path is None:
    print(f"❌ Parameter file not found: {param_file}")
    raise SystemExit

# Get data file path
data_file = input("Enter path to data Excel file (with featureState data): ").strip()
data_path = try_open_excel(data_file)

if data_path is None:
    print(f"❌ Data file not found: {data_file}")
    raise SystemExit

print("🔍 Validating Excel files...")
if not is_valid_excel_file(param_path):
    print(f"❌ Parameter file is not a valid Excel file: {param_path}")
    raise SystemExit

if not is_valid_excel_file(data_path):
    print(f"❌ Data file is not a valid Excel file: {data_path}")
    raise SystemExit

print(f"📁 Parameter file: {param_path}")
print(f"📁 Data file: {data_path}")

# Load parameter workbook
try:
    print("📖 Loading parameter workbook...")
    param_wb = load_workbook(param_path)
except Exception as e:
    print(f"❌ Error loading parameter workbook: {e}")
    raise SystemExit

# Check if Features + Licenses sheet exists
if "Features + Licenses" not in param_wb.sheetnames:
    print(f"❌ Sheet 'Features + Licenses' not found in parameter file")
    print(f"Available sheets: {param_wb.sheetnames}")
    raise SystemExit

# Read Features + Licenses sheet
features_ws = param_wb["Features + Licenses"]
features_header = [cell.value for cell in features_ws[1]]

print(f"📋 Columns in Features + Licenses sheet: {features_header}")

# Map column names to indices
features_col_idx_map = {name: idx + 1 for idx, name in enumerate(features_header) if name is not None}

# Check required columns in Features + Licenses file
required_features_columns = ["Feature name", "Bytel nodes", "FeatureState", "A activer ou pas pour Bytel"]

missing_columns = []
for col in required_features_columns:
    if col not in features_col_idx_map:
        missing_columns.append(col)

if missing_columns:
    print(f"❌ Missing required columns in Features + Licenses sheet: {missing_columns}")
    raise SystemExit

# Get features data
features_data = []
feature_name_col = features_col_idx_map["Feature name"]
feature_state_col = features_col_idx_map["FeatureState"]
activate_col = features_col_idx_map["A activer ou pas pour Bytel"]
nodes_bytel = features_col_idx_map["Bytel nodes"]

print(f"\n🔍 Collecting features from 'Features + Licenses' sheet...")

for row in range(2, features_ws.max_row + 1):
    feature_name = features_ws.cell(row=row, column=feature_name_col).value
    feature_state = features_ws.cell(row=row, column=feature_state_col).value
    activate = features_ws.cell(row=row, column=activate_col).value
    nodes = features_ws.cell(row=row, column=nodes_bytel).value

    if feature_name is not None and str(feature_name).strip() != "":
        feature_clean = str(feature_name).strip()

        features_data.append({
            "Feature name": feature_clean,
            "FeatureState": feature_state,
            "A activer ou pas pour Bytel": activate,
            "Bytel nodes": nodes
        })

print(f"✅ Collected {len(features_data)} features from Features + Licenses sheet")

# Load data workbook
try:
    print(f"\n📖 Loading data workbook...")
    data_df = pd.read_excel(data_path, engine="openpyxl")
except Exception as e:
    print(f"❌ Error loading data file: {e}")
    raise SystemExit

print(f"📊 Data file shape: {data_df.shape}")
print(f"📋 Data file columns: {list(data_df.columns)}")

# Check if data file has required columns
required_data_columns = ["featureStateId", "NeName", "featureState", "serviceState"]
missing_data_columns = []
for col in required_data_columns:
    if col not in data_df.columns:
        missing_data_columns.append(col)

if missing_data_columns:
    print(f"❌ Missing required columns in data file: {missing_data_columns}")
    raise SystemExit

# Create output data
output_data = []
incorrect_data = []

print(f"\n📝 Creating license validation table...")

# Define color fills
RED_FILL = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")     # Red for incorrect
YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")  # Yellow for not found

# Process each feature from the parameter file
for feature_info in features_data:
    feature_name = feature_info["Feature name"]
    feature_state_id = feature_info["FeatureState"]  # This is what we search for in data file
    activate_for_bytel = feature_info["A activer ou pas pour Bytel"]
    nodes_for_bytel = feature_info["Bytel nodes"]

    # Search for this FeatureState in the data file's featureStateId column
    matching_rows = data_df[data_df["featureStateId"] == feature_state_id]

    if len(matching_rows) > 0:
        # FeatureState found in data file
        for _, row in matching_rows.iterrows():
            site_name = row["NeName"]
            actual_feature_state = row["featureState"]
            service_state = row["serviceState"]
            
            # Get node type and validate
            node_type = get_node_type(site_name)
            validation_status = validate_feature_state(activate_for_bytel, node_type, actual_feature_state)
            
            output_row = {
                "Feature name": feature_name,
                "FeatureState": feature_state_id,
                "Bytel nodes": nodes_for_bytel,
                "A activer ou pas pour Bytel": activate_for_bytel,
                "NeName": site_name,
                "featureState": actual_feature_state,
                "serviceState": service_state,
                "Validation": validation_status
            }
            
            output_data.append(output_row)
            
            # Add to incorrect data if validation failed
            if validation_status == "INCORRECT":
                incorrect_data.append(output_row)
    else:
        # FeatureState not found in data file
        output_row = {
            "Feature name": feature_name,
            "FeatureState": feature_state_id,
            "Bytel nodes": nodes_for_bytel,
            "A activer ou pas pour Bytel": activate_for_bytel,
            "NeName": "NOT FOUND",
            "featureState": "NOT FOUND",
            "serviceState": "NOT FOUND",
            "Validation": "NOT FOUND"
        }
        output_data.append(output_row)

# Create output DataFrame
output_df = pd.DataFrame(output_data)
incorrect_df = pd.DataFrame(incorrect_data)

print(f"✅ Created output table with {len(output_df)} rows")
print(f"❌ Found {len(incorrect_df)} incorrect feature states")

# Create output file name
output_filename = "License_Validation_Report.xlsx"

# Save to new Excel file
try:
    print(f"\n💾 Saving output to: {output_filename}")

    # Create a new workbook
    wb = Workbook()

    # Remove default sheet
    wb.remove(wb.active)

    # Create main validation sheet
    ws_main = wb.create_sheet("License_Validation")

    # Write headers for main sheet
    main_headers = ["Feature name", "FeatureState", "Bytel nodes", "A activer ou pas pour Bytel",
                    "NeName", "featureState", "serviceState", "Validation"]

    for col_idx, header in enumerate(main_headers, 1):
        ws_main.cell(row=1, column=col_idx, value=header)

    # Write data and apply formatting
    current_feature = None
    merge_start_row = 2

    for row_idx, (_, row_data) in enumerate(output_df.iterrows(), 2):
        # Write row data
        for col_idx, header in enumerate(main_headers, 1):
            ws_main.cell(row=row_idx, column=col_idx, value=row_data[header])

        # Apply color coding to columns starting from NeName (column 5) for incorrect entries
        validation_status = row_data["Validation"]
        if validation_status == "INCORRECT":
            # Highlight columns from NeName to Validation (columns 5 to 8) in red
            for col in range(5, len(main_headers) + 1):  # Columns 5 to 8
                ws_main.cell(row=row_idx, column=col).fill = RED_FILL
        elif validation_status == "NOT FOUND":
            # Highlight columns from NeName to Validation (columns 5 to 8) in yellow for not found
            for col in range(5, len(main_headers) + 1):  # Columns 5 to 8
                ws_main.cell(row=row_idx, column=col).fill = YELLOW_FILL

        # Check if this is a new feature group
        if row_data["Feature name"] != current_feature:
            # If we were tracking a previous feature, merge its cells
            if current_feature is not None and merge_start_row < row_idx - 1:
                for col in range(1, 4):  # Merge columns A to C (Feature name, FeatureState, A activer ou pas)
                    ws_main.merge_cells(start_row=merge_start_row, start_column=col,
                                        end_row=row_idx - 1, end_column=col)

            # Start tracking new feature
            current_feature = row_data["Feature name"]
            merge_start_row = row_idx

    # Merge the last feature group
    if current_feature is not None and merge_start_row < len(output_df) + 1:
        for col in range(1, 4):  # Merge columns A to C
            ws_main.merge_cells(start_row=merge_start_row, start_column=col,
                                end_row=len(output_df) + 1, end_column=col)

    # Create incorrect entries sheet
    ws_incorrect = wb.create_sheet("Incorrect_Entries")
    
    if len(incorrect_df) > 0:
        # Write headers for incorrect sheet
        incorrect_headers = ["Feature name", "FeatureState", "Bytel nodes", "A activer ou pas pour Bytel",
                            "NeName", "featureState", "serviceState", "Validation"]
        
        for col_idx, header in enumerate(incorrect_headers, 1):
            ws_incorrect.cell(row=1, column=col_idx, value=header)
        
        # Write incorrect data
        for row_idx, (_, row_data) in enumerate(incorrect_df.iterrows(), 2):
            for col_idx, header in enumerate(incorrect_headers, 1):
                ws_incorrect.cell(row=row_idx, column=col_idx, value=row_data[header])
            
            # Highlight columns from NeName to Validation (columns 5 to 8) in red for incorrect entries
            for col in range(5, len(incorrect_headers) + 1):  # Columns 5 to 8
                ws_incorrect.cell(row=row_idx, column=col).fill = RED_FILL
    else:
        ws_incorrect.cell(row=1, column=1, value="No incorrect entries found!")
    
    # Create summary sheet
    ws_summary = wb.create_sheet("Summary")

    # Calculate summary statistics
    total_features = len(features_data)
    total_entries = len(output_df)

    features_found = len(output_df[output_df["NeName"] != "NOT FOUND"])
    features_not_found = len(output_df[output_df["NeName"] == "NOT FOUND"])

    correct_entries = len(output_df[output_df["Validation"] == "CORRECT"])
    incorrect_entries = len(output_df[output_df["Validation"] == "INCORRECT"])
    unknown_entries = len(output_df[output_df["Validation"] == "UNKNOWN"])
    not_found_entries = len(output_df[output_df["Validation"] == "NOT FOUND"])
    
    unique_sites = output_df[output_df["NeName"] != "NOT FOUND"]["NeName"].nunique()

    # Write summary
    ws_summary.cell(row=1, column=1, value="License Validation Summary").font = Font(bold=True, size=14)
    ws_summary.merge_cells('A1:B1')

    summary_data = [
        ("Total Features in Parameter File", total_features),
        ("Total Validation Entries", total_entries),
        ("Unique Sites Found", unique_sites),
        ("Features Found in Data File", features_found),
        ("Features Not Found in Data File", features_not_found),
        ("✅ CORRECT Feature States", correct_entries),
        ("❌ INCORRECT Feature States", incorrect_entries),
        ("⚪ UNKNOWN Feature States", unknown_entries),
        ("🔍 NOT FOUND Features", not_found_entries),
    ]

    for i, (label, value) in enumerate(summary_data, 3):
        ws_summary.cell(row=i, column=1, value=label)
        ws_summary.cell(row=i, column=2, value=value)

    # Add features not found section
    start_row = len(summary_data) + 5
    ws_summary.cell(row=start_row, column=1, value="Features Not Found in Data File").font = Font(bold=True)
    ws_summary.merge_cells(f'A{start_row}:B{start_row}')

    not_found_features = []
    for feature_info in features_data:
        feature_state_id = feature_info["FeatureState"]
        if feature_state_id not in data_df["featureStateId"].values:
            not_found_features.append(f"{feature_info['Feature name']} (FeatureState: {feature_state_id})")

    if not_found_features:
        for i, feature in enumerate(not_found_features, start_row + 1):
            ws_summary.cell(row=i, column=1, value=feature)
    else:
        ws_summary.cell(row=start_row + 1, column=1, value="All features were found in data file")

    # Auto-adjust column widths for all sheets
    sheets_to_adjust = [ws_main, ws_incorrect, ws_summary]

    for ws in sheets_to_adjust:
        for column_cells in ws.columns:
            # Skip if it's a MergedCell (which doesn't have column_letter attribute)
            if isinstance(column_cells[0], MergedCell):
                continue

            max_length = 0
            column_letter = column_cells[0].column_letter
            for cell in column_cells:
                try:
                    # Skip MergedCell objects
                    if isinstance(cell, MergedCell):
                        continue
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            ws.column_dimensions[column_letter].width = adjusted_width

    # Add borders and formatting to all sheets
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                         top=Side(style='thin'), bottom=Side(style='thin'))

    for ws in sheets_to_adjust:
        if ws == ws_main:
            max_row = len(output_df) + 1
            max_col = len(main_headers)
        elif ws == ws_incorrect:
            max_row = len(incorrect_df) + 1 if len(incorrect_df) > 0 else 1
            max_col = len(incorrect_headers) if len(incorrect_df) > 0 else 1
        elif ws == ws_summary:
            max_row = start_row + len(not_found_features) if not_found_features else start_row + 1
            max_col = 2
        else:
            continue

        for row in ws.iter_rows(min_row=1, max_row=max_row, max_col=max_col):
            for cell in row:
                # Only apply border to non-merged cells
                if not isinstance(cell, MergedCell):
                    cell.border = thin_border

        # Style header row (skip merged cells)
        header_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
        for cell in ws[1]:
            if not isinstance(cell, MergedCell):
                cell.fill = header_fill
                cell.font = Font(bold=True)

    # Add legend to main sheet
    legend_row = len(output_df) + 3
    ws_main.cell(row=legend_row, column=1, value="LEGEND:").font = Font(bold=True)
    ws_main.cell(row=legend_row + 1, column=1, value="❌ RED highlighting").fill = RED_FILL
    ws_main.cell(row=legend_row + 1, column=2, value="= Incorrect feature state (from NeName to Validation)")
    ws_main.cell(row=legend_row + 2, column=1, value="🔍 YELLOW highlighting").fill = YELLOW_FILL
    ws_main.cell(row=legend_row + 2, column=2, value="= Feature not found in data file (from NeName to Validation)")

    wb.save(output_filename)

    print(f"🎉 SUCCESS!")
    print(f"📁 Output file created: {output_filename}")
    print(f"📊 Sheets created:")
    print(f"   - License_Validation: Main validation results (NeName to Validation highlighted for issues)")
    print(f"   - Incorrect_Entries: List of all incorrect feature states")
    print(f"   - Summary: Overall validation summary")
    print(f"📊 Final Summary:")
    print(f"   📋 Total Features: {total_features}")
    print(f"   🏢 Unique Sites: {unique_sites}")
    print(f"   ✅ CORRECT Feature States: {correct_entries}")
    print(f"   ❌ INCORRECT Feature States: {incorrect_entries}")
    print(f"   ⚪ UNKNOWN Feature States: {unknown_entries}")
    print(f"   🔍 NOT FOUND Features: {not_found_entries}")

    # Print incorrect features if any
    if incorrect_entries > 0:
        print(f"\n❌ INCORRECT FEATURE STATES FOUND:")
        for _, row in incorrect_df.iterrows():
            print(f"   - {row['Feature name']} on {row['NeName']} (State: {row['featureState']}, Expected: {row['A activer ou pas pour Bytel']})")

except Exception as e:
    print(f"❌ Error saving output file: {e}")
    import traceback
    traceback.print_exc()

print(f"\n💡 How it works:")
print(f"   • Reads 'Feature name', 'FeatureState', 'A activer ou pas pour Bytel' from parameter file")
print(f"   • Searches for 'FeatureState' value in data file's 'featureStateId' column")
print(f"   • Validates featureState based on activation rules and node type (E, X, G)")
print(f"   • Highlights columns from NeName to Validation: RED=incorrect, YELLOW=not found")
print(f"   • Creates separate sheet with all incorrect entries")
