import pandas as pd
import getpass
from modules import get_device_output, handle_switchport_status, COMMANDS

def main():
    input_file = "devices.xlsx"
    try:
        df_switches = pd.read_excel(input_file)
    except Exception as e:
        print(f"Error reading Excel file {input_file}: {e}")
        return

    select_program = input("Enter number of the program you want to run:\n1 - Switchport status\n2 - Device backup\n")
    username = input("Enter the username: ")
    password = getpass.getpass("Enter the password: ")
    secret_password = getpass.getpass("Enter the secret password")

    if select_program == '1':
        handle_switchport_status(df_switches, username, password)
    elif select_program == '2':
        use_default_commands = input("Do you want to use the default commands? (yes/no): ").strip().lower()
        if use_default_commands == 'yes':
            commands = COMMANDS
        else:
            commands = input("Enter your commands separated by commas: ").split(',')
            commands = [cmd.strip() for cmd in commands]
        
        for idx, row in df_switches.iterrows():
            hostname = row['hostname']
            ip_address = row['ip']
            get_device_output(ip_address, hostname, username, password, secret_password, commands)
    else:
        print("Unknown program")

if __name__ == '__main__':
    main()