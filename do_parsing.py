import os
import re
import pandas as pd

def search_files_for_pattern(directory, extensions, pattern):
    results = []
    
    for root, dirs, files in os.walk(directory):
        for file in files:
            if any(file.endswith(ext) for ext in extensions):
                file_path = os.path.join(root, file)
                with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                    lines = f.readlines()
                    for i, line in enumerate(lines):
                        if pattern in line:
                            # Extracting the .do filename from the content
                            found_filename = re.search(r'[\w/]+\.do', line)
                            if found_filename:
                                found_filename = found_filename.group().split('/')[-1]
                            else:
                                found_filename = ''
                            result = {
                                'Filename': file,
                                'Line Number': i + 1,
                                'Found Content': line.strip(),
                                'Found Content Filename': found_filename
                            }
                            results.append(result)
                            # Print to console
                            print(f"Found in file: {file}, Line: {i + 1}")
                            print(f"Content: {line.strip()}")
                            print(f"Extracted Filename: {found_filename}")
    
    return results

def save_to_excel(data, output_file):
    df = pd.DataFrame(data)
    df.to_excel(output_file, index=False)

# Configuration
directory = 'C:/test'
extensions = ['.xfdl', '.jsp']
pattern = '.do'
output_file = 'results.xlsx'

# Execution
data = search_files_for_pattern(directory, extensions, pattern)
save_to_excel(data, output_file)

print(f'Results saved to {output_file}')