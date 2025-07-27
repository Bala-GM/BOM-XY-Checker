import pandas as pd
import string

# Load Excel
file_path = r"D:/NX_BACKWORK/Feeder Setup_PROCESS/#Output/BOM/BOM.xlsx"
df = pd.read_excel(file_path)

# Clean and get unique Internal P/Ns
unique_parts = list(dict.fromkeys(df['Internal P/N'].dropna()))

# Generate Group codes like AA, AB, AC...
group_codes = [a + b for a in string.ascii_uppercase for b in string.ascii_uppercase]

# Map Internal P/N to Group
group_map = {part: group_codes[i] for i, part in enumerate(unique_parts)}

# Apply Group and Priority
df['Group'] = df['Internal P/N'].map(group_map)
df['Priority'] = df.groupby('Internal P/N').cumcount() + 1

# Save back to Excel
df.to_excel(file_path, index=False)
print("✅ Group and Priority columns updated successfully.")
