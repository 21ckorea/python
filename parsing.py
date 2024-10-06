import os
import re
import pandas as pd

def find_ips_in_file(file_path):
    """Find all IP addresses in a given file."""
    ip_pattern = re.compile(r'\b(?:[0-9]{1,3}\.){3}[0-9]{1,3}\b')
    ips = []
    with open(file_path, 'r', errors='ignore') as file:
        content = file.read()
        ips = ip_pattern.findall(content)
    return ips

def search_folder_for_ips(folder_path):
    """Search through all files in the specified folder for IP addresses."""
    results = []
    for root, _, files in os.walk(folder_path):
        for file in files:
            file_path = os.path.join(root, file)
            ips = find_ips_in_file(file_path)
            if ips:
                for ip in ips:
                    results.append({'Source Path': root, 'File Name': file, 'IP': ip})
                    # Print the result to the console
                    print(f"Source Path: {root}, File Name: {file}, IP: {ip}")
    return results

def write_results_to_excel(results, output_file):
    """Write the search results to an Excel file."""
    df = pd.DataFrame(results)
    df.to_excel(output_file, index=False)

# Example usage
folder_path = 'path/to/your/folder'
output_file = 'output.xlsx'
results = search_folder_for_ips(folder_path)
write_results_to_excel(results, output_file)