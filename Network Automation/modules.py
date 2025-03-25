import pandas as pd
from netmiko import ConnectHandler
import re

COMMANDS = [
    'terminal length 0', 'show ver', 'show run', 'show int des',
    'show ip int brief', 'show cdp neighbors', 'show lldp neighbors',
    'show vlan', 'show mac address-table', 'show switch', 'show ip arp',
    'show clock', 'show log'
]

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

def get_device_output(ip, hostname, username, password, secret_password, commands):
    device = {
        'device_type': 'cisco_ios',
        'host': ip,
        'username': username,
        'password': password,
        'secret': secret_password
    }
    try:
        connection = ConnectHandler(**device)
        connection.enable()
        with open(f"{hostname}.txt", 'w') as file:
            for command in commands:
                output = connection.send_command(command)
                file.write(f"\n\n{command}\n{'='*len(command)}\n{output}")
        connection.disconnect()
        print(f"Output saved to {hostname}.txt")
    except Exception as e:
        print(f"Failed to connect to {ip}: {e}")

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

def update_interface_description_and_shutdown(df_interfaces, username, password):
    for idx, row in df_interfaces.iterrows():
        hostname = row['hostname']
        ip_address = row['ip']
        interface = row['interface']
        description = row['description']
        print(f"\nConnecting to {hostname} ({ip_address}) to update {interface}...")

        device = {
            'device_type': 'cisco_ios',
            'ip': ip_address,
            'username': username,
            'password': password,
        }
        try:
            net_connect = ConnectHandler(**device)
            net_connect.enable()
            config_commands = [
                f"interface {interface}",
                f"description {description}",
                "shutdown"
            ]
            net_connect.send_config_set(config_commands)
            net_connect.disconnect()
            print(f"Successfully updated {interface} on {hostname}.")
        except Exception as e:
            print(f"Failed to connect to {hostname} ({ip_address}): {e}")