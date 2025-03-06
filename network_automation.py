import pandas as pd
from netmiko import ConnectHandler
import getpass
import re

# Constants

COMMANDS = [
    'terminal length 0', 'show ver', 'show running-config', 'show int des',
    'show ip int brief', 'show cdp neighbors', 'show lldp neighbors',
    'show vlan', 'show mac address-table', 'show switch', 'show ip arp',
    'show log', 'show clock'
]

# Functions

# Function to parse interface description after login to a switch
def parse_interface_description(output, hostname, ip_address):
    results = []
    lines = output.splitlines()
    for line in lines[1:]:
        line = line.strip()
        if not line:
            continue
        match = re.match(r"^(\S+)\s+((?:admin down|down|up))\s+(\S+)\s*(.*)$", line, re.IGNORECASE)
        if match:
            interface, status, protocol, description = match.groups()
            results.append({
                'hostname': hostname,
                'ip': ip_address,
                'Interface': interface,
                'Status': status.strip(),
                'Protocol': protocol.strip(),
                'Description': description.strip() if description else "N/A"
            })
        else:
            print(f"Error on line: {line}")
    return results

# Function to get the command outputs in a text file post login to the switch
def get_device_output(ip, username, password):
    device = {
        'device_type': 'cisco_ios',
        'host': ip,
        'username': username,
        'password': password,
    }
    try:
        connection = ConnectHandler(**device)
        hostname = connection.send_command('show run | include hostname').strip().split()[-1]
        with open(f"{hostname}.txt", 'w') as file:
            for command in COMMANDS:
                output = connection.send_command(command)
                file.write(f"\n\n{command}\n{'='*len(command)}\n{output}")
        connection.disconnect()
        print(f"Output saved to {hostname}.txt")
    except Exception as e:
        print(f"Failed to connect to {ip}: {e}")

# Function to handle switchport status
def handle_switchport_status(df_switches, username, password):
    results = []
    for idx, row in df_switches.iterrows():
        hostname = row['hostname']
        ip_address = row['ip']
        print(f"\nConnecting to {hostname} ({ip_address})...")
        device = {
            'device_type': 'cisco_ios',
            'ip': ip_address,
            'username': username,
            'password': password,
        }
        try:
            net_connect = ConnectHandler(**device)
            output = net_connect.send_command("show interface description")
            parsed_data = parse_interface_description(output, hostname, ip_address)
            results.extend(parsed_data)
            net_connect.disconnect()
            print(f"Successfully retrieved data from {hostname}.")
        except Exception as e:
            print(f"Failed to connect to {hostname} ({ip_address}): {e}")
            results.append({
                'hostname': hostname,
                'ip': ip_address,
                'Interface': 'N/A',
                'Status': 'N/A',
                'Protocol': 'N/A',
                'Description': str(e)
            })
    if results:
        output_file = "switch_interface_descriptions.xlsx"
        try:
            df_results = pd.DataFrame(results)
            df_results.to_excel(output_file, index=False)
            print(f"\nOutput saved successfully to {output_file}.")
        except Exception as e:
            print(f"Error writing results to Excel: {e}")
    else:
        print("No results to write.")

# Main function
def main():
    input_file = "switches.xlsx"
    try:
        df_switches = pd.read_excel(input_file)
    except Exception as e:
        print(f"Error reading Excel file {input_file}: {e}")
        return

    select_program = input("Enter number of the program you want to run:\n1 - Switchport status\n2 - Device backup")
    username = input("Enter the username: ")
    password = getpass.getpass("Enter the password: ")

    if select_program == '1':
        handle_switchport_status(df_switches, username, password)
    elif select_program == '2':
        for ip in df_switches['ip']:
            get_device_output(ip, username, password)
    else:
        print("Unknown program")

if __name__ == '__main__':
    main()