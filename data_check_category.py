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


def get_cell_type(cellname):
    """Determine cell type based on cellname prefix"""
    if pd.isna(cellname):
        return "unknown"

    cellname_str = str(cellname).strip().upper()

    if cellname_str.startswith('Y'):
        return "FDD"
    elif cellname_str.startswith('Q'):
        return "TDD"
    else:
        return "unknown"


def get_node_type(nename):
    """Determine node type based on NeName"""
    if pd.isna(nename):
        return "unknown"
    
    nename_str = str(nename).strip().upper()
    
    # LIST ALL YOUR SFR NENAMES HERE
    sfr_nenames = [
        # Add all your actual SFR NeNames here...
        # Example: "NODE_SFR_1", "NODE_SFR_2", etc.
    ]
    
    # Check if it's SFR
    for sfr_nename in sfr_nenames:
        if sfr_nename.upper() in nename_str:
            return "SFR"
    
    # Check for TDD+FDD co-nodes (separate category)
    tdd_fdd_nodes = ['X90240', 'X90299', 'X90296', 'X90400']
    for node in tdd_fdd_nodes:
        if node in nename_str:
            return "TDD+FDD"
    
    # Check for ZTD nodes (separate category)
    if 'X90295' in nename_str:
        return "ZTD"
    
    # Check for CRZ nodes (separate category) 
    if 'X90260' in nename_str:
        return "CRZ"
    
    # Check for Ran4 nodes (separate category)
    if 'X90252' in nename_str:
        return "Ran4"
    
    # Everything else is BYT (separate category)
    return "BYT"


def convert_boolean_display(value):
    """Convert French boolean values to English for display"""
    if pd.isna(value) or value == "":
        return value
        
    # Handle both string and other types
    value_str = str(value).strip().lower()
    
    # Comprehensive list of French boolean representations
    if value_str in ['vrai', 'oui', 'true', '1', 'yes', 'on', 'activé', 'activée', 'activer']:
        return "true"
    elif value_str in ['faux', 'non', 'false', '0', 'no', 'off', 'désactivé', 'désactivée', 'désactiver']:
        return "false"
    else:
        # Return original value if it's not a boolean
        return value


def normalize_boolean_value(value):
    """Convert various boolean representations to standardized true/false"""
    if pd.isna(value) or value == "":
        return None

    value_str = str(value).strip().lower()

    # Handle French boolean values and convert to English true/false
    if value_str in ['true', '1', 'yes', 'oui', 'on', 'vrai']:
        return "true"
    elif value_str in ['false', '0', 'no', 'non', 'off', 'faux']:
        return "false"
    else:
        return value_str  # Return as-is for non-boolean values


def extract_main_value(value):
    """Extract main value before space or = (e.g., '14' from '14 = 14ms', '-10' from '-10 dBm')"""
    if pd.isna(value) or value == "":
        return None

    value_str = str(value).strip()

    # Handle negative numbers and regular numbers
    # Match numbers with optional negative sign, including decimals
    number_pattern = r'^-?\d+\.?\d*'
    number_match = re.match(number_pattern, value_str)

    if number_match:
        return number_match.group()

    # If value contains space or =, extract the first part
    if ' ' in value_str or '=' in value_str:
        # Split by space or = and take the first part
        parts = re.split(r'[\s=]', value_str)
        main_value = parts[0].strip()

        # Check if it's a number (including negative) or boolean
        if main_value and (re.match(number_pattern, main_value) or main_value.lower() in ['true', 'false']):
            return main_value

    return value_str


def is_na_value(value):
    """Check if value is N/A, null, empty, or similar"""
    if pd.isna(value) or value == "":
        return True

    value_str = str(value).strip().lower()
    na_values = ['n/a', 'null', 'none', 'nan', 'empty', 'vide', '-']
    return value_str in na_values


def normalize_parameter_key(key):
    """Normalize parameter key by removing common prefixes and converting to lowercase"""
    if pd.isna(key):
        return ""

    key_str = str(key).strip().lower()

    # Remove common telecom parameter prefixes
    prefixes = [
        'vsdata', 'vs', 'data', 'nr', 'lte', 'cell', 'param', 'parameter',
        'eutran', 'geran', 'uran', 'wcdma', 'gsm', 'umts', 'hspa'
    ]

    # Remove prefixes
    for prefix in prefixes:
        if key_str.startswith(prefix):
            key_str = key_str[len(prefix):]
            break

    return key_str.strip()


def fuzzy_key_match(expected_key, actual_key):
    """Check if keys match approximately (handles prefixes like vsData)"""
    if pd.isna(expected_key) or pd.isna(actual_key):
        return False

    expected_clean = normalize_parameter_key(expected_key)
    actual_clean = normalize_parameter_key(actual_key)

    # Exact match after normalization
    if expected_clean == actual_clean:
        return True

    # Check if one is contained in the other (for partial matches)
    if expected_clean in actual_clean or actual_clean in expected_clean:
        return True

    # Split by common separators and check word overlap
    expected_words = set(re.findall(r'[a-z]+', expected_clean))
    actual_words = set(re.findall(r'[a-z]+', actual_clean))

    if expected_words and actual_words:
        common_words = expected_words & actual_words
        # If most words match, consider it a match
        if len(common_words) >= min(len(expected_words), len(actual_words)):
            return True

    return False


def parse_key_value_pairs(value_str):
    """Parse key=value pairs from a string with robust error handling"""
    if pd.isna(value_str) or value_str == "":
        return {}

    pairs = {}
    try:
        # Split by comma and parse each pair
        for pair in str(value_str).split(','):
            pair = pair.strip()
            if '=' in pair:
                key, value = pair.split('=', 1)
                pairs[key.strip()] = value.strip()
    except Exception as e:
        print(f"⚠️  Error parsing key-value pairs: {e}")

    return pairs


def find_key_value_in_string(search_key, search_value, long_string):
    """Search for a specific key=value pair within a longer string containing multiple pairs"""
    if pd.isna(long_string) or long_string == "":
        return False

    long_str = str(long_string).strip()

    # NEW: Handle the case where the expected value is a simple key=value pair
    # and the actual value is a longer path containing that key=value pair
    if "=" in long_str and "," in long_str:
        # This looks like a comma-separated list of key=value pairs
        all_pairs = parse_key_value_pairs(long_str)

        # Search for a matching key (with fuzzy matching) and exact value match
        for actual_key, actual_value in all_pairs.items():
            if fuzzy_key_match(search_key, actual_key) and actual_value == search_value:
                return True
    else:
        # Handle case where the actual value might be a single key=value pair or path
        # Check if the search_key=search_value appears anywhere in the long string
        expected_pair = f"{search_key}={search_value}"
        if expected_pair in long_str:
            return True

        # Also try with fuzzy key matching
        # Split the long string by commas and check each part
        parts = [part.strip() for part in long_str.split(',')]
        for part in parts:
            if '=' in part:
                actual_key, actual_value = part.split('=', 1)
                actual_key = actual_key.strip()
                actual_value = actual_value.strip()
                if fuzzy_key_match(search_key, actual_key) and actual_value == search_value:
                    return True

    return False


def detect_validation_pattern(expected_value):
    """Detect what type of validation pattern the expected value represents"""
    if pd.isna(expected_value) or is_na_value(expected_value):
        return "no_expected_value"

    expected_str = str(expected_value).strip()

    # Pattern 1: Value with explanation (e.g., "14 = 14ms", "0 = DEACTIVATED", "-10 = -10 dBm")
    if (" = " in expected_str or " " in expected_str) and len(expected_str.split()) >= 2:
        # Check if first part is a number (including negative) or boolean
        first_part = expected_str.split()[0]
        number_pattern = r'^-?\d+\.?\d*$'
        if re.match(number_pattern, first_part) or first_part.lower() in ['true', 'false']:
            return "value_with_explanation"

    # Pattern 2: Contains multiple key=value pairs separated by commas
    if "," in expected_str and "=" in expected_str:
        pairs = [pair.strip() for pair in expected_str.split(",")]
        if all("=" in pair for pair in pairs):
            return "key_value_pairs"

    # Pattern 3: Comma-separated list of values (no key=value)
    if "," in expected_str and "=" not in expected_str:
        return "value_list"

    # Pattern 4: Single key=value pair (NEW pattern for cases like "CgSwitch=Default")
    if "=" in expected_str and "," not in expected_str:
        return "single_key_value"

    # Pattern 5: Contains specific keywords that indicate it's a partial match
    partial_keywords = ['enabled', 'disabled', 'active', 'inactive', 'on', 'off', 'yes', 'no']
    if any(keyword in expected_str.lower() for keyword in partial_keywords):
        return "partial_match"

    # Pattern 6: Numeric range (e.g., "0-100", "1..10", "-10-10")
    if re.match(r'^-?\d+\s*-\s*-?\d+$', expected_str) or re.match(r'^-?\d+\s*\.\.\s*-?\d+$', expected_str):
        return "numeric_range"

    # Pattern 7: Multiple options separated by | or /
    if "|" in expected_str or "/" in expected_str:
        return "multiple_options"
    
    # Pattern 8: Node-specific values (e.g., "20 = 20slots en ZTD")
    node_types = ["ZTD", "CRZ", "Ran4", "SFR", "BYT", "TDD+FDD"]
    if any(node_type in expected_str.upper() for node_type in node_types):
        return "node_specific"

    return "exact_match"


def extract_node_specific_value(expected_value, node_type):
    """Extract value for specific node type from expected value string"""
    if pd.isna(expected_value) or pd.isna(node_type):
        return None
        
    expected_str = str(expected_value).strip()
    node_type_upper = node_type.upper()
    
    # Check if the expected value contains node-specific information
    if node_type_upper in expected_str.upper():
        # Extract the value before the node type specification
        # Example: "20 = 20slots en ZTD" -> extract "20"
        parts = expected_str.split()
        if parts and len(parts) > 0:
            # Try to extract the first numeric value or boolean
            main_value = extract_main_value(parts[0])
            if main_value:
                return main_value
                
    return None


def validate_tdd_fdd_co_node_value(actual_value, expected_co_node_value, cell_type):
    """Validate TDD+FDD co-node values with Profile parameter"""
    if pd.isna(expected_co_node_value) or is_na_value(expected_co_node_value):
        return False
        
    expected_str = str(expected_co_node_value).strip()
    actual_str = str(actual_value).strip()
    
    # Check for Profile pattern (Profile=0 for TDD, Profile=1 for FDD)
    if "Profile=" in expected_str:
        # Extract the expected profile value based on cell type
        expected_profile = "0" if cell_type == "TDD" else "1"
        expected_profile_pair = f"Profile={expected_profile}"
        
        # Check if the actual value contains the expected profile
        if expected_profile_pair in actual_str:
            return True
            
        # Also try fuzzy matching for profile
        if "Profile=" in actual_str:
            # Extract profile from actual value
            profile_match = re.search(r'Profile=(\d+)', actual_str)
            if profile_match:
                actual_profile = profile_match.group(1)
                return actual_profile == expected_profile
    
    return False


def apply_special_validation(expected_value, actual_value, pattern_type, node_type=None, expected_co_node_value=None, cell_type=None):
    """Apply special validation based on the detected pattern"""
    if pd.isna(actual_value) or pd.isna(expected_value) or is_na_value(expected_value):
        return False

    expected_str = str(expected_value).strip()
    actual_str = str(actual_value).strip()

    if pattern_type == "value_with_explanation":
        # Extract the main value (e.g., "14" from "14 = 14ms", "-10" from "-10 dBm")
        main_value = extract_main_value(expected_str)
        return actual_str == main_value

    elif pattern_type == "key_value_pairs":
        # Parse expected key-value pairs
        expected_pairs = parse_key_value_pairs(expected_str)

        # Debug output for key-value pairs
        if expected_pairs:
            print(f"🔍 Key-Value Comparison:")
            print(f"   Expected: {expected_pairs}")
            print(f"   Actual string: '{actual_str}'")

        # Search for each expected key-value pair within the actual string
        for exp_key, exp_value in expected_pairs.items():
            if not find_key_value_in_string(exp_key, exp_value, actual_str):
                print(f"   ❌ Not found in actual: '{exp_key}'='{exp_value}'")
                return False
            else:
                print(f"   ✅ Found in actual: '{exp_key}'='{exp_value}'")

        return True

    elif pattern_type == "single_key_value":  # NEW pattern handler
        # Handle cases like expected: "CgSwitch=Default", actual: "SubNetwork=NR_lte,...,vsDataCgSwitch=Default"
        if "=" in expected_str:
            exp_key, exp_value = expected_str.split("=", 1)
            exp_key = exp_key.strip()
            exp_value = exp_value.strip()

            print(f"🔍 Single Key-Value Validation:")
            print(f"   Expected: '{exp_key}'='{exp_value}'")
            print(f"   Actual: '{actual_str}'")

            # Use the enhanced find_key_value_in_string function
            result = find_key_value_in_string(exp_key, exp_value, actual_str)
            print(f"   Result: {'✅ MATCH' if result else '❌ NO MATCH'}")
            return result

        return False

    elif pattern_type == "value_list":
        # Check if all expected values are in the actual list
        expected_items = [item.strip() for item in expected_str.split(",")]
        actual_items = [item.strip() for item in actual_str.split(",")]
        return all(item in actual_items for item in expected_items)

    elif pattern_type == "partial_match":
        # Check if expected string is contained within actual string
        return expected_str.lower() in actual_str.lower()

    elif pattern_type == "numeric_range":
        # Extract numeric range and check if actual value is within range
        numbers = re.findall(r'-?\d+', expected_str)
        if len(numbers) == 2:
            min_val, max_val = map(int, numbers)
            try:
                actual_num = float(actual_str)
                return min_val <= actual_num <= max_val
            except ValueError:
                return False
        return False

    elif pattern_type == "multiple_options":
        # Check if actual value matches any of the options
        options = re.split(r'[|/]', expected_str)
        options = [opt.strip() for opt in options]
        return actual_str in options
        
    elif pattern_type == "node_specific":
        # Handle node-specific values (e.g., "20 = 20slots en ZTD")
        if node_type == "TDD+FDD" and expected_co_node_value is not None and cell_type is not None:
            # First try TDD+FDD co-node validation
            if validate_tdd_fdd_co_node_value(actual_value, expected_co_node_value, cell_type):
                return True
            
        if node_type:
            # Try to extract node-specific value
            node_specific_value = extract_node_specific_value(expected_str, node_type)
            if node_specific_value:
                return actual_str == node_specific_value
            
        # If no node-specific value found or node_type doesn't match, extract main value
        main_value = extract_main_value(expected_str)
        return actual_str == main_value

    return False


def validate_parameter_value(actual_value, expected_tdd_value, expected_fdd_value, expected_default_value, expected_co_node_value, cell_type,
                             parameter_name, node_type):
    """Validate if the actual value matches the expected value based on cell type and node type"""

    # Skip validation for administrativeState parameter
    if parameter_name and "administrativestate" in parameter_name.lower():
        return "skipped"

    if parameter_name and "nrtac" in parameter_name.lower():
        return "skipped"

    if pd.isna(actual_value) or actual_value == "":
        return "no_data"

    # For TDD+FDD co-nodes, prioritize the co-node value
    if node_type == "TDD+FDD" and not pd.isna(expected_co_node_value) and not is_na_value(expected_co_node_value):
        expected_value = expected_co_node_value
        # Use special TDD+FDD validation
        pattern_type = detect_validation_pattern(expected_value)
        if pattern_type != "exact_match" and pattern_type != "no_expected_value":
            if apply_special_validation(expected_value, actual_value, pattern_type, node_type, expected_co_node_value, cell_type):
                return "correct_fuzzy"
    else:
        # Determine which expected value to use based on cell type
        expected_value = None
        if cell_type == "TDD" and not pd.isna(expected_tdd_value) and not is_na_value(expected_tdd_value):
            expected_value = expected_tdd_value
        elif cell_type == "FDD" and not pd.isna(expected_fdd_value) and not is_na_value(expected_fdd_value):
            expected_value = expected_fdd_value
        elif not pd.isna(expected_default_value) and not is_na_value(expected_default_value):
            expected_value = expected_default_value

        # If no valid expected value found, can't validate
        if expected_value is None:
            return "no_expected_value"

    actual_str = str(actual_value).strip()
    
    # Convert French boolean values to English for display
    if actual_str.lower() in ['vrai', 'oui']:
        actual_str = "true"
    elif actual_str.lower() in ['faux', 'non']:
        actual_str = "false"

    expected_str = str(expected_value).strip()

    # First, try special validation patterns
    pattern_type = detect_validation_pattern(expected_value)
    if pattern_type != "exact_match" and pattern_type != "no_expected_value":
        if apply_special_validation(expected_value, actual_value, pattern_type, node_type, expected_co_node_value, cell_type):
            return "correct_fuzzy"

    # Try extracting main value for comparison (e.g., "14" from "14 = 14ms", "-10" from "-10 dBm")
    main_expected = extract_main_value(expected_value)
    if main_expected and main_expected != expected_str:
        if actual_str == main_expected:
            return "correct_extracted"

    # Then try normalized boolean comparison
    actual_normalized = normalize_boolean_value(actual_value)
    expected_normalized = normalize_boolean_value(expected_value)

    if actual_normalized is not None and expected_normalized is not None:
        if actual_normalized == expected_normalized:
            return "correct"

    # Try numeric comparison for numbers (including negative numbers)
    try:
        # Check if both values can be converted to numbers
        actual_num = float(actual_str)
        expected_num = float(expected_str)
        if actual_num == expected_num:
            return "correct_numeric"
    except (ValueError, TypeError):
        pass

    # Finally, try exact string comparison
    if actual_str == expected_str:
        return "correct"

    return "incorrect"


# Get user choice
choice = input("Select sheet type (NRCellCU / NRCellDU): ").strip()
if choice not in ["NRCellCU", "NRCellDU"]:
    print("❌ Invalid choice. Please choose NRCellCU or NRCellDU.")
    raise SystemExit

SHEET_NAME = choice

# Get parameter file path
param_file = input("Enter path to parameter Excel file (updated file): ").strip()
param_path = try_open_excel(param_file)

if param_path is None:
    print(f"❌ Parameter file not found: {param_file}")
    raise SystemExit

# Get data file path
data_file = input("Enter path to data Excel file (with node data): ").strip()
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

# Check if selected sheet exists
if SHEET_NAME not in param_wb.sheetnames:
    print(f"❌ Sheet '{SHEET_NAME}' not found in parameter file")
    print(f"Available sheets: {param_wb.sheetnames}")
    raise SystemExit

# Read parameter sheet
param_ws = param_wb[SHEET_NAME]
param_header = [cell.value for cell in param_ws[1]]

print(f"📋 Columns in parameter sheet: {param_header}")

# Map column names to indices
param_col_idx_map = {name: idx + 1 for idx, name in enumerate(param_header) if name is not None}

# Check required columns in parameter file
required_param_columns = ["Parameter", "Valeur par défaut RBS", "Valeur Bytel TDD MidBand",
                          "Valeur Bytel FDD ESS 15MHz", "Valeur Bytel TDD HigBand"]

# Add TDD+FDD co-node column for NRCellCU
if SHEET_NAME == "NRCellCU":
    required_param_columns.append("Valeur Bytel TDD+FDD co-node\nAppliquer la valeur commune si valeur TDD et FDD sont même, sinon appliquer la valeur spécifiée dans cette colonne.")

missing_columns = []
for col in required_param_columns:
    if col not in param_col_idx_map:
        missing_columns.append(col)

if missing_columns:
    print(f"❌ Missing required columns in parameter file: {missing_columns}")
    raise SystemExit

# Get parameters (excluding read-only ones and administrativeState)
parameters_to_include = []
parameter_data = {}  # Store parameter info
missing_parameters_in_data = []  # Track parameters not found in data file

param_col_idx = param_col_idx_map["Parameter"]
readonly_col_idx = param_col_idx_map["Valeur par défaut RBS"]

print(f"\n🔍 Collecting parameters from '{SHEET_NAME}' sheet...")

for row in range(2, param_ws.max_row + 1):
    param_value = param_ws.cell(row=row, column=param_col_idx).value
    readonly_value = param_ws.cell(row=row, column=readonly_col_idx).value

    if param_value is not None and str(param_value).strip() != "":
        param_clean = str(param_value).strip()

        # Skip administrativeState parameter
        if "administrativestate" in param_clean.lower():
            print(f"⚠️  Skipping administrativeState parameter: {param_clean}")
            continue

        if "nrtac" in param_clean.lower():
            print(f"⚠️  Skipping nRTAC parameter: {param_clean}")
            continue

        # Check if parameter is NOT read-only
        if readonly_value is None or str(readonly_value).strip().lower() != "read-only":
            parameters_to_include.append(param_clean)

            # Store parameter data for later use
            parameter_data[param_clean] = {
                "Valeur par défaut RBS": param_ws.cell(row=row,
                                                       column=param_col_idx_map["Valeur par défaut RBS"]).value,
                "Valeur Bytel TDD MidBand": param_ws.cell(row=row,
                                                          column=param_col_idx_map["Valeur Bytel TDD MidBand"]).value,
                "Valeur Bytel FDD ESS 15MHz": param_ws.cell(row=row, column=param_col_idx_map[
                    "Valeur Bytel FDD ESS 15MHz"]).value,
                "Valeur Bytel TDD HigBand": param_ws.cell(row=row,
                                                          column=param_col_idx_map["Valeur Bytel TDD HigBand"]).value
            }
            
            # Add TDD+FDD co-node value for NRCellCU
            if SHEET_NAME == "NRCellCU":
                parameter_data[param_clean]["Valeur Bytel TDD+FDD co-node"] = param_ws.cell(row=row,
                    column=param_col_idx_map["Valeur Bytel TDD+FDD co-node\nAppliquer la valeur commune si valeur TDD et FDD sont même, sinon appliquer la valeur spécifiée dans cette colonne."]).value
        else:
            print(f"⚠️  Skipping read-only parameter: {param_clean}")

print(f"✅ Collected {len(parameters_to_include)} parameters (excluding read-only and administrativeState and nRTAC)")

# Load data workbook and convert French booleans
try:
    print(f"\n📖 Loading data workbook...")
    data_df = pd.read_excel(data_path, engine="openpyxl")
    
    # Convert French boolean values to English in the entire dataframe
    print("🔄 Converting French boolean values to English...")
    for column in data_df.columns:
        if data_df[column].dtype == 'object':  # Only check string columns
            # Convert all values in this column
            data_df[column] = data_df[column].apply(convert_boolean_display)
                
except Exception as e:
    print(f"❌ Error loading data file: {e}")
    raise SystemExit

print(f"📊 Data file shape: {data_df.shape}")
print(f"📋 Data file columns: {list(data_df.columns)}")

# Check if data file has required columns
if "CellName" not in data_df.columns:
    print("❌ 'CellName' column not found in data file")
    raise SystemExit

# Check if NeName column exists for node type categorization
if "NeName" not in data_df.columns:
    print("❌ 'NeName' column not found in data file - needed for node type categorization")
    raise SystemExit

# Find which parameters from our list exist in the data file
available_parameters_in_data = []
for param in parameters_to_include:
    if param in data_df.columns:
        available_parameters_in_data.append(param)
    else:
        print(f"⚠️  Parameter '{param}' not found in data file columns")
        missing_parameters_in_data.append(param)

print(f"🔍 Found {len(available_parameters_in_data)} parameters in data file")
print(f"❌ {len(missing_parameters_in_data)} parameters not found in data file")

if not available_parameters_in_data:
    print("❌ No matching parameters found between parameter file and data file")
    raise SystemExit

# Create DataFrames for output
main_output_data = []
wrong_parameters_data = []  # For incorrect values
missing_data_records = []  # For records with no data

print(f"\n📝 Creating structured output table with validation...")

# Define color fills
GREEN_FILL = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")  # Light Green
YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")  # Yellow
VIOLET_FILL = PatternFill(start_color="CBC3E3", end_color="CBC3E3", fill_type="solid")  # Light Purple/Violet
BLUE_FILL = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")  # Light Blue for special rules

validation_stats = {
    "correct": 0,
    "correct_fuzzy": 0,
    "correct_special": 0,
    "correct_extracted": 0,
    "correct_numeric": 0,
    "incorrect": 0,
    "no_data": 0,
    "no_expected_value": 0,
    "skipped": 0
}

print(f"🔧 Analyzing validation patterns...")

for param in available_parameters_in_data:
    # Get parameter info from parameter file
    param_info = parameter_data[param]

    # Get all unique value-cellname-nename pairs for this parameter
    value_cellname_nename_pairs = []

    for idx, row in data_df.iterrows():
        value = row[param]
        cellname = row['CellName']
        nename = row['NeName']
        if pd.notna(value) or pd.notna(cellname) or pd.notna(nename):  # Include if any has data
            value_cellname_nename_pairs.append((value, cellname, nename))

    if value_cellname_nename_pairs:
        # For the first value, include all parameter info
        first_value, first_cellname, first_nename = value_cellname_nename_pairs[0]
        cell_type = get_cell_type(first_cellname)
        node_type = get_node_type(first_nename)  # Use NeName for node type
        
        # Convert display value for output (double conversion to be safe)
        display_value = convert_boolean_display(first_value)
        
        # Get co-node value for NRCellCU
        expected_co_node_value = None
        if SHEET_NAME == "NRCellCU":
            expected_co_node_value = param_info.get("Valeur Bytel TDD+FDD co-node")
        
        validation = validate_parameter_value(
            first_value,
            param_info["Valeur Bytel TDD MidBand"] or param_info["Valeur Bytel TDD HigBand"],
            param_info["Valeur Bytel FDD ESS 15MHz"],
            param_info["Valeur par défaut RBS"],
            expected_co_node_value,
            cell_type,
            param,
            node_type
        )

        validation_stats[validation] += 1

        # Track missing data
        if validation == "no_data":
            missing_data_records.append({
                "Parameter": param,
                "CellName": first_cellname,
                "NeName": first_nename,
                "CellType": cell_type,
                "NodeType": node_type,
                "Valeur par défaut RBS": param_info["Valeur par défaut RBS"],
                "Valeur Bytel TDD MidBand": param_info["Valeur Bytel TDD MidBand"],
                "Valeur Bytel FDD ESS 15MHz": param_info["Valeur Bytel FDD ESS 15MHz"],
                "Valeur Bytel TDD HigBand": param_info["Valeur Bytel TDD HigBand"]
            })

        # Add to main output
        output_row = {
            "Parameter": param,
            "Valeur par défaut RBS": param_info["Valeur par défaut RBS"],
            "Valeur Bytel TDD MidBand": param_info["Valeur Bytel TDD MidBand"],
            "Valeur Bytel FDD ESS 15MHz": param_info["Valeur Bytel FDD ESS 15MHz"],
            "Valeur Bytel TDD HigBand": param_info["Valeur Bytel TDD HigBand"],
            "Value": display_value,  # Use converted value
            "CellName": first_cellname,
            "NeName": first_nename,
            "CellType": cell_type,
            "NodeType": node_type,
            "Validation": validation
        }
        
        # Add co-node value for NRCellCU
        if SHEET_NAME == "NRCellCU":
            output_row["Valeur Bytel TDD+FDD co-node"] = expected_co_node_value
            
        main_output_data.append(output_row)

        # Add to wrong parameters sheet if incorrect
        if validation == "incorrect":
            wrong_row = {
                "Parameter": param,
                "Valeur par défaut RBS": param_info["Valeur par défaut RBS"],
                "Valeur Bytel TDD MidBand": param_info["Valeur Bytel TDD MidBand"],
                "Valeur Bytel FDD ESS 15MHz": param_info["Valeur Bytel FDD ESS 15MHz"],
                "Valeur Bytel TDD HigBand": param_info["Valeur Bytel TDD HigBand"],
                "Actual_Value": display_value,  # Use converted value
                "Expected_Value": param_info["Valeur Bytel FDD ESS 15MHz"] if cell_type == "FDD" else
                param_info["Valeur Bytel TDD MidBand"] or param_info["Valeur Bytel TDD HigBand"],
                "CellName": first_cellname,
                "NeName": first_nename,
                "CellType": cell_type,
                "NodeType": node_type
            }
            
            # Add co-node value for NRCellCU
            if SHEET_NAME == "NRCellCU":
                wrong_row["Valeur Bytel TDD+FDD co-node"] = expected_co_node_value
                wrong_row["Expected_Value"] = expected_co_node_value if node_type == "TDD+FDD" else wrong_row["Expected_Value"]
                
            wrong_parameters_data.append(wrong_row)

        # For subsequent values, keep parameter info blank (will be merged in Excel)
        for value, cellname, nename in value_cellname_nename_pairs[1:]:
            cell_type = get_cell_type(cellname)
            node_type = get_node_type(nename)  # Use NeName for node type
            
            # Convert display value for output (double conversion to be safe)
            display_value = convert_boolean_display(value)
            
            # Get co-node value for NRCellCU
            expected_co_node_value = None
            if SHEET_NAME == "NRCellCU":
                expected_co_node_value = param_info.get("Valeur Bytel TDD+FDD co-node")
            
            validation = validate_parameter_value(
                value,
                param_info["Valeur Bytel TDD MidBand"] or param_info["Valeur Bytel TDD HigBand"],
                param_info["Valeur Bytel FDD ESS 15MHz"],
                param_info["Valeur par défaut RBS"],
                expected_co_node_value,
                cell_type,
                param,
                node_type
            )

            validation_stats[validation] += 1

            # Track missing data
            if validation == "no_data":
                missing_data_records.append({
                    "Parameter": param,
                    "CellName": cellname,
                    "NeName": nename,
                    "CellType": cell_type,
                    "NodeType": node_type,
                    "Valeur par défaut RBS": param_info["Valeur par défaut RBS"],
                    "Valeur Bytel TDD MidBand": param_info["Valeur Bytel TDD MidBand"],
                    "Valeur Bytel FDD ESS 15MHz": param_info["Valeur Bytel FDD ESS 15MHz"],
                    "Valeur Bytel TDD HigBand": param_info["Valeur Bytel TDD HigBand"]
                })

            # Add to main output
            output_row = {
                "Parameter": "",  # Blank for merging
                "Valeur par défaut RBS": "",  # Blank for merging
                "Valeur Bytel TDD MidBand": "",  # Blank for merging
                "Valeur Bytel FDD ESS 15MHz": "",  # Blank for merging
                "Valeur Bytel TDD HigBand": "",  # Blank for merging
                "Value": display_value,  # Use converted value
                "CellName": cellname,
                "NeName": nename,
                "CellType": cell_type,
                "NodeType": node_type,
                "Validation": validation
            }
            
            # Add co-node value for NRCellCU
            if SHEET_NAME == "NRCellCU":
                output_row["Valeur Bytel TDD+FDD co-node"] = ""
                
            main_output_data.append(output_row)

            # Add to wrong parameters sheet if incorrect
            if validation == "incorrect":
                wrong_row = {
                    "Parameter": param,
                    "Valeur par défaut RBS": param_info["Valeur par défaut RBS"],
                    "Valeur Bytel TDD MidBand": param_info["Valeur Bytel TDD MidBand"],
                    "Valeur Bytel FDD ESS 15MHz": param_info["Valeur Bytel FDD ESS 15MHz"],
                    "Valeur Bytel TDD HigBand": param_info["Valeur Bytel TDD HigBand"],
                    "Actual_Value": display_value,  # Use converted value
                    "Expected_Value": param_info["Valeur Bytel FDD ESS 15MHz"] if cell_type == "FDD" else
                    param_info["Valeur Bytel TDD MidBand"] or param_info["Valeur Bytel TDD HigBand"],
                    "CellName": cellname,
                    "NeName": nename,
                    "CellType": cell_type,
                    "NodeType": node_type
                }
                
                # Add co-node value for NRCellCU
                if SHEET_NAME == "NRCellCU":
                    wrong_row["Valeur Bytel TDD+FDD co-node"] = expected_co_node_value
                    wrong_row["Expected_Value"] = expected_co_node_value if node_type == "TDD+FDD" else wrong_row["Expected_Value"]
                    
                wrong_parameters_data.append(wrong_row)

# Create output DataFrames
main_output_df = pd.DataFrame(main_output_data)
wrong_parameters_df = pd.DataFrame(wrong_parameters_data)
missing_data_df = pd.DataFrame(missing_data_records)

# Create summary data for the Summary sheet
summary_data = {
    "Category": [
        "Total Parameters in Parameter File",
        "Parameters Found in Data File",
        "Parameters Not Found in Data File",
        "Total Validations Performed",
        "Correct Values (Exact Match)",
        "Correct Values (Fuzzy Match)",
        "Correct Values (Special Rules)",
        "Correct Values (Extracted)",
        "Correct Values (Numeric)",
        "Incorrect Values",
        "Missing Data",
        "No Expected Value",
        "Skipped (administrativeState and nRTAC)"
    ],
    "Count": [
        len(parameters_to_include),
        len(available_parameters_in_data),
        len(missing_parameters_in_data),
        sum(validation_stats.values()),
        validation_stats["correct"],
        validation_stats["correct_fuzzy"],
        validation_stats["correct_special"],
        validation_stats["correct_extracted"],
        validation_stats["correct_numeric"],
        validation_stats["incorrect"],
        validation_stats["no_data"],
        validation_stats["no_expected_value"],
        validation_stats["skipped"]
    ]
}

summary_df = pd.DataFrame(summary_data)

print(f"✅ Created main output table with {len(main_output_df)} rows")
print(f"✅ Created wrong parameters table with {len(wrong_parameters_df)} rows")
print(f"✅ Created missing data table with {len(missing_data_df)} rows")
print(f"📊 Validation Statistics:")
print(f"   ✅ Correct values (exact match): {validation_stats['correct']}")
print(f"   🤖 Correct values (fuzzy match): {validation_stats['correct_fuzzy']}")
print(f"   🔵 Correct values (special rules): {validation_stats['correct_special']}")
print(f"   🔷 Correct values (extracted): {validation_stats['correct_extracted']}")
print(f"   🔢 Correct values (numeric): {validation_stats['correct_numeric']}")
print(f"   ⚠️  Incorrect values: {validation_stats['incorrect']}")
print(f"   💜 No data: {validation_stats['no_data']}")
print(f"   ❓ No expected value: {validation_stats['no_expected_value']}")
print(f"   ⏭️  Skipped (administrativeState and nRTAC): {validation_stats['skipped']}")

# Show parameters not found in data file
if missing_parameters_in_data:
    print(f"\n❌ Parameters not found in data file ({len(missing_parameters_in_data)}):")
    for param in missing_parameters_in_data:
        print(f"   - {param}")

# Create output file name
output_filename = f"{SHEET_NAME}_Parameter_Validation.xlsx"

# Save to new Excel file with multiple sheets
try:
    print(f"\n💾 Saving output to: {output_filename}")

    # Create a new workbook
    wb = Workbook()

    # Remove default sheet
    wb.remove(wb.active)

    # Create main validation sheet
    ws_main = wb.create_sheet("Parameter_Validation")

    # Write headers for main sheet
    main_headers = ["Parameter", "Valeur par défaut RBS", "Valeur Bytel TDD MidBand",
                    "Valeur Bytel FDD ESS 15MHz", "Valeur Bytel TDD HigBand", "Value",
                    "CellName", "NeName", "CellType", "NodeType", "Validation"]
    
    # Add TDD+FDD co-node column for NRCellCU
    if SHEET_NAME == "NRCellCU":
        main_headers.insert(5, "Valeur Bytel TDD+FDD co-node")

    for col_idx, header in enumerate(main_headers, 1):
        ws_main.cell(row=1, column=col_idx, value=header)

    # Write main data and apply formatting
    current_param = None
    merge_start_row = 2

    for row_idx, (_, row_data) in enumerate(main_output_df.iterrows(), 2):
        # Write row data
        for col_idx, header in enumerate(main_headers, 1):
            ws_main.cell(row=row_idx, column=col_idx, value=row_data[header])

        # Apply color based on validation
        validation = row_data["Validation"]
        if validation in ["correct", "correct_fuzzy", "correct_special", "correct_extracted", "correct_numeric"]:
            fill_color = GREEN_FILL
        elif validation == "incorrect":
            fill_color = YELLOW_FILL
        elif validation == "no_data":
            fill_color = VIOLET_FILL
        elif validation == "skipped":
            fill_color = PatternFill(start_color="FFB6C1", end_color="FFB6C1", fill_type="solid")  # Light Pink
        else:
            fill_color = None

        if fill_color:
            # Apply color to Value and CellName columns
            value_col = 6 if SHEET_NAME == "NRCellDU" else 7  # Adjust column index for co-node
            cellname_col = value_col + 1
            ws_main.cell(row=row_idx, column=value_col).fill = fill_color
            ws_main.cell(row=row_idx, column=cellname_col).fill = fill_color

        # Check if this is a new parameter group
        if row_data["Parameter"] != "":
            # If we were tracking a previous parameter, merge its cells
            if current_param is not None and merge_start_row < row_idx - 1:
                for col in range(1, 6):  # Merge columns A to E
                    ws_main.merge_cells(start_row=merge_start_row, start_column=col,
                                        end_row=row_idx - 1, end_column=col)

            # Start tracking new parameter
            current_param = row_data["Parameter"]
            merge_start_row = row_idx

    # Merge the last parameter group
    if current_param is not None and merge_start_row < len(main_output_df) + 1:
        for col in range(1, 6):  # Merge columns A to E
            ws_main.merge_cells(start_row=merge_start_row, start_column=col,
                                end_row=len(main_output_df) + 1, end_column=col)

    # Create wrong parameters sheet
    if len(wrong_parameters_df) > 0:
        ws_wrong = wb.create_sheet("Wrong_Parameters")

        # Write headers for wrong parameters sheet
        wrong_headers = ["Parameter", "Valeur par défaut RBS", "Valeur Bytel TDD MidBand",
                         "Valeur Bytel FDD ESS 15MHz", "Valeur Bytel TDD HigBand",
                         "Actual_Value", "Expected_Value", "CellName", "NeName", "CellType", "NodeType"]
        
        # Add TDD+FDD co-node column for NRCellCU
        if SHEET_NAME == "NRCellCU":
            wrong_headers.insert(5, "Valeur Bytel TDD+FDD co-node")

        for col_idx, header in enumerate(wrong_headers, 1):
            ws_wrong.cell(row=1, column=col_idx, value=header)

        # Write wrong parameters data
        for row_idx, (_, row_data) in enumerate(wrong_parameters_df.iterrows(), 2):
            for col_idx, header in enumerate(wrong_headers, 1):
                ws_wrong.cell(row=row_idx, column=col_idx, value=row_data[header])

            # Highlight the incorrect values in yellow
            actual_value_col = 6 if SHEET_NAME == "NRCellDU" else 7
            expected_value_col = actual_value_col + 1
            ws_wrong.cell(row=row_idx, column=actual_value_col).fill = YELLOW_FILL
            ws_wrong.cell(row=row_idx, column=expected_value_col).fill = YELLOW_FILL

    # Create missing data sheet
    if len(missing_data_df) > 0:
        ws_missing = wb.create_sheet("Missing_Data")

        # Write headers for missing data sheet
        missing_headers = ["Parameter", "CellName", "NeName", "CellType", "NodeType", "Valeur par défaut RBS",
                           "Valeur Bytel TDD MidBand", "Valeur Bytel FDD ESS 15MHz", "Valeur Bytel TDD HigBand"]

        for col_idx, header in enumerate(missing_headers, 1):
            ws_missing.cell(row=1, column=col_idx, value=header)

        # Write missing data
        for row_idx, (_, row_data) in enumerate(missing_data_df.iterrows(), 2):
            for col_idx, header in enumerate(missing_headers, 1):
                ws_missing.cell(row=row_idx, column=col_idx, value=row_data[header])

            # Highlight in violet to indicate missing data
            for col in range(1, len(missing_headers) + 1):
                ws_missing.cell(row=row_idx, column=col).fill = VIOLET_FILL

    # Create summary sheet
    ws_summary = wb.create_sheet("Summary")

    # Write summary headers
    ws_summary.cell(row=1, column=1, value="Validation Summary").font = Font(bold=True, size=14)
    ws_summary.merge_cells('A1:B1')

    # Write summary data
    for row_idx, (_, row_data) in enumerate(summary_df.iterrows(), 3):
        ws_summary.cell(row=row_idx, column=1, value=row_data["Category"])
        ws_summary.cell(row=row_idx, column=2, value=row_data["Count"])

    # Add parameters not found section
    start_row = len(summary_df) + 5
    ws_summary.cell(row=start_row, column=1, value="Parameters Not Found in Data File").font = Font(bold=True)
    ws_summary.merge_cells(f'A{start_row}:B{start_row}')

    if missing_parameters_in_data:
        for i, param in enumerate(missing_parameters_in_data, start_row + 1):
            ws_summary.cell(row=i, column=1, value=param)
    else:
        ws_summary.cell(row=start_row + 1, column=1, value="All parameters were found in data file")

    # FIXED: Auto-adjust column widths for all sheets with proper handling for merged cells
    sheets_to_adjust = [ws_main, ws_summary]
    if len(wrong_parameters_df) > 0:
        sheets_to_adjust.append(ws_wrong)
    if len(missing_data_df) > 0:
        sheets_to_adjust.append(ws_missing)

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
            max_row = len(main_output_df) + 1
            max_col = len(main_headers)
        elif ws == ws_wrong and len(wrong_parameters_df) > 0:
            max_row = len(wrong_parameters_df) + 1
            max_col = len(wrong_headers)
        elif ws == ws_missing and len(missing_data_df) > 0:
            max_row = len(missing_data_df) + 1
            max_col = len(missing_headers)
        elif ws == ws_summary:
            max_row = start_row + len(missing_parameters_in_data) if missing_parameters_in_data else start_row + 1
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
    legend_row = len(main_output_df) + 3
    ws_main.cell(row=legend_row, column=1, value="LEGEND:").font = Font(bold=True)
    ws_main.cell(row=legend_row + 1, column=1, value="Green").fill = GREEN_FILL
    ws_main.cell(row=legend_row + 1, column=2, value="= Correct value")
    ws_main.cell(row=legend_row + 2, column=1, value="Yellow").fill = YELLOW_FILL
    ws_main.cell(row=legend_row + 2, column=2, value="= Incorrect value")
    ws_main.cell(row=legend_row + 3, column=1, value="Violet").fill = VIOLET_FILL
    ws_main.cell(row=legend_row + 3, column=2, value="= No data")
    ws_main.cell(row=legend_row + 4, column=1, value="Pink").fill = PatternFill(start_color="FFB6C1",
                                                                                end_color="FFB6C1", fill_type="solid")
    ws_main.cell(row=legend_row + 4, column=2, value="= Skipped (administrativeState and nRTAC)")

    # Add pattern explanations
    ws_main.cell(row=legend_row + 6, column=1, value="ENHANCED VALIDATION RULES:").font = Font(bold=True)
    ws_main.cell(row=legend_row + 7, column=1, value="• '14 = 14ms' → '14' is correct (extracts value before space/=)")
    ws_main.cell(row=legend_row + 8, column=1, value="• '-10 = -10 dBm' → '-10' is correct (handles negative numbers)")
    ws_main.cell(row=legend_row + 9, column=1,
                 value="• 'EnergyEfficiency=1' → 'vsDataEnergyEfficiency=1' (fuzzy key matching)")
    ws_main.cell(row=legend_row + 10, column=1,
                 value="• 'CgSwitch=Default' → 'SubNetwork=...,vsDataCgSwitch=Default' (path matching)")
    ws_main.cell(row=legend_row + 11, column=1, value="• '20 = 20slots en ZTD' → '20' for ZTD nodes (node-specific)")
    if SHEET_NAME == "NRCellCU":
        ws_main.cell(row=legend_row + 12, column=1, value="• 'Profile=0' for TDD cells, 'Profile=1' for FDD cells in TDD+FDD nodes")
    ws_main.cell(row=legend_row + 13, column=1, value="• N/A values in expected columns are treated as null")
    ws_main.cell(row=legend_row + 14, column=1,
                 value="• administrativeState and nRTAC parameter is skipped from validation")
    ws_main.cell(row=legend_row + 15, column=1, value="• French 'VRAI/FAUX' converted to English 'true/false'")

    wb.save(output_filename)

    print(f"🎉 SUCCESS!")
    print(f"📁 Output file created: {output_filename}")
    print(f"📊 Sheets created:")
    print(f"   - Parameter_Validation: Main validation results")
    if len(wrong_parameters_df) > 0:
        print(f"   - Wrong_Parameters: {len(wrong_parameters_df)} incorrect values")
    if len(missing_data_df) > 0:
        print(f"   - Missing_Data: {len(missing_data_df)} missing data records")
    print(f"   - Summary: Overall validation summary")
    print(f"📊 Final Validation Summary:")
    print(f"   🟢 {validation_stats['correct']} correct (exact match)")
    print(f"   🤖 {validation_stats['correct_fuzzy']} correct (fuzzy match)")
    print(f"   🔵 {validation_stats['correct_special']} correct (special rules)")
    print(f"   🔷 {validation_stats['correct_extracted']} correct (extracted values)")
    print(f"   🔢 {validation_stats['correct_numeric']} correct (numeric comparison)")
    print(f"   🟡 {validation_stats['incorrect']} incorrect")
    print(f"   🟣 {validation_stats['no_data']} missing data")
    print(f"   🎀 {validation_stats['skipped']} skipped (administrativeState and nRTAC)")

    # Print parameters not found
    if missing_parameters_in_data:
        print(f"\n❌ PARAMETERS NOT FOUND IN DATA FILE:")
        for param in missing_parameters_in_data:
            print(f"   - {param}")

except Exception as e:
    print(f"❌ Error saving output file: {e}")
    import traceback

    traceback.print_exc()

print(f"\n💡 Enhanced Validation Features:")
print(f"   • '14 = 14ms' → '14' is now correctly validated")
print(f"   • '-10 = -10 dBm' → '-10' is now correctly validated")
print(f"   • 'EnergyEfficiency=1' → 'vsDataEnergyEfficiency=1' (fuzzy key matching)")
print(f"   • 'CgSwitch=Default' → 'SubNetwork=NR_lte,...,vsDataCgSwitch=Default' (path matching)")
print(f"   • '20 = 20slots en ZTD' → '20' for ZTD nodes (node-specific validation)")
if SHEET_NAME == "NRCellCU":
    print(f"   • 'Profile=0' for TDD cells, 'Profile=1' for FDD cells in TDD+FDD nodes (NEW)")
print(f"   • N/A values in expected columns are treated as null")
print(f"   • administrativeState and nRTAC parameter is completely skipped")
print(f"   • Numeric comparison: 60 vs 60 matches even if different data types")
print(f"   • French 'VRAI/FAUX' converted to English 'true/false' in output")
print(f"   • Node type categorization now uses 'NeName' column")
print(f"   • SEPARATED NODE CATEGORIES: TDD+FDD, ZTD, CRZ, Ran4, SFR, BYT")
