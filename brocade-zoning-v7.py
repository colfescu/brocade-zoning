import openpyxl

def create_zoning_script(file_path, config_name):
    # Load the workbook and select the active sheet
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active

    # Extract SAN device names from specified cells
    san_device_name_1 = sheet['E2'].value
    san_device_name_2 = sheet['I2'].value
    san_device_name_3 = sheet['M2'].value
    san_device_name_4 = sheet['Q2'].value

    # Extract relevant columns from the sheet
    server_wwpns = [cell.value for cell in sheet['B'][1:]]
    server_aliases = [cell.value for cell in sheet['C'][1:]]

    san_wwpns_1 = [cell.value for cell in sheet['F'][1:]]
    san_aliases_1 = [cell.value for cell in sheet['G'][1:]]
    san_wwpns_2 = [cell.value for cell in sheet['J'][1:]]
    san_aliases_2 = [cell.value for cell in sheet['K'][1:]]
    san_wwpns_3 = [cell.value for cell in sheet['N'][1:]]
    san_aliases_3 = [cell.value for cell in sheet['O'][1:]]
    san_wwpns_4 = [cell.value for cell in sheet['R'][1:]]
    san_aliases_4 = [cell.value for cell in sheet['S'][1:]]

    aliases = set()
    zones = {}

    # Function to create aliases
    def create_alias(wwpn, alias):
        if wwpn and alias:
            aliases.add(f"alicreate \"{alias}\", \"{wwpn}\"")

    # Function to create zones
    def create_zone(server_alias, san_aliases, san_device_name):
        if server_alias and any(san_aliases):
            zone_name = f"Z_{server_alias}_{san_device_name}"
            zone_members = [server_alias] + [alias for alias in san_aliases if alias]
            zone_members_str = ";".join(zone_members)
            zones[zone_name] = f"zonecreate \"{zone_name}\", \"{zone_members_str}\""

    # Iterate through each row and create aliases for each SAN, skipping the header
    for i in range(len(server_wwpns)):
        create_alias(server_wwpns[i], server_aliases[i])
        create_alias(san_wwpns_1[i], san_aliases_1[i])
        create_alias(san_wwpns_2[i], san_aliases_2[i])
        create_alias(san_wwpns_3[i], san_aliases_3[i])
        create_alias(san_wwpns_4[i], san_aliases_4[i])

    # Group SAN aliases and WWPNs with their respective device names
    san_groups = [
        (san_device_name_1, san_aliases_1),
        (san_device_name_2, san_aliases_2),
        (san_device_name_3, san_aliases_3),
        (san_device_name_4, san_aliases_4),
    ]

    # Create zones for each server alias with all SAN aliases for each SAN device
    for i in range(len(server_wwpns)):
        for san_device_name, san_aliases in san_groups:
            create_zone(server_aliases[i], san_aliases, san_device_name)

    # Create the configuration command
    config_cmd = f"cfgcreate \"{config_name}\", \""
    config_cmd += ";".join(zones.keys())
    config_cmd += "\""

    # Add commands to save and enable the configuration
    cfgsave_cmd = "cfgsave"
    cfgenable_cmd = f"cfgenable \"{config_name}\""

    return aliases, zones.values(), config_cmd, cfgsave_cmd, cfgenable_cmd

def write_to_file(aliases, zones, config_cmd, cfgsave_cmd, cfgenable_cmd, output_file):
    with open(output_file, 'w') as f:
        for alias in aliases:
            f.write(alias + "\n")
        for zone in zones:
            f.write(zone + "\n")
        f.write(config_cmd + "\n")
        f.write(cfgsave_cmd + "\n")
        f.write(cfgenable_cmd + "\n")

# Define input and output paths
file_path = '/TUH/WWNs Zone Processing.xlsx'
config_name = 'BrocadeConfig_v7'
output_file = 'zoning_script_v7.txt'

# Create the zoning script
aliases, zones, config_cmd, cfgsave_cmd, cfgenable_cmd = create_zoning_script(file_path, config_name)
write_to_file(aliases, zones, config_cmd, cfgsave_cmd, cfgenable_cmd, output_file)

print(f"Zoning script created and saved to {output_file}")
