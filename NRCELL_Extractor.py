import os

def extract_parameters(input_file_path, output_file_path, cell_type):
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
                    parameter_names.append((current_parent_parameter, sub_parameter))
            else:
                current_parent_parameter = stripped_line.split()[0].strip()
                parameter_names.append((None, current_parent_parameter))

    with open(output_file_path, 'w') as file:
        file.write('Struct;Parameter\n')
        for parent, sub in parameter_names:
            if parent:
                file.write(f"{parent};{sub}\n")
            else:
                file.write(f";{sub}\n")  

input_file_path = input("Please enter the path to the input file: ").strip()

cell_type = input("Please enter the cell type (NRCellDU or NRCellCU): ").strip()
output_file_name = f"Extracted_params_{cell_type}.csv"

script_directory = os.path.dirname(__file__)
output_file_path = os.path.join(script_directory, output_file_name)

extract_parameters(input_file_path, output_file_path, cell_type)
