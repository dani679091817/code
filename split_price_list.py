import sys
import math
import pandas as pd

DEFAULT_INPUT_FILE = "liste prix vente article 2.xlsx"
CHUNK_SIZE = 2000
EXPECTED_COLUMNS = ["Référence", "Désignation", "Prix HT"]

input_file = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_INPUT_FILE

# Read the file, skipping the title row and using the second row as header
try:
    df = pd.read_excel(input_file, header=1)
except FileNotFoundError:
    print(f"Error: File '{input_file}' not found.")
    sys.exit(1)
except Exception as e:
    print(f"Error reading '{input_file}': {e}")
    sys.exit(1)

# Validate that the file has the expected three columns
if df.shape[1] != 3:
    print(f"Error: Expected 3 columns, found {df.shape[1]}.")
    sys.exit(1)

# Assign standardized column names
df.columns = EXPECTED_COLUMNS

total_rows = len(df)
num_files = math.ceil(total_rows / CHUNK_SIZE)

print(f"Total rows: {total_rows}")
print(f"Number of output files: {num_files}")

base_name = input_file.rsplit(".", 1)[0]

for i in range(num_files):
    start = i * CHUNK_SIZE
    end = min(start + CHUNK_SIZE, total_rows)
    chunk = df.iloc[start:end]

    output_filename = f"{base_name} - part {i + 1}.xlsx"
    try:
        chunk.to_excel(output_filename, index=False)
    except Exception as e:
        print(f"Error writing '{output_filename}': {e}")
        sys.exit(1)
    print(f"Created '{output_filename}' with {len(chunk)} rows")

print("Done.")
