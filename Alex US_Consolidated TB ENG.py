import pandas as pd
import numpy as np
import os
import shutil

##########################################
# 1. Define file path and create new file name
##########################################
file_path = r"C:\Users\slee\OneDrive - SBP\Tax Returns\xSungkeun\Tax Return\2024\TB\Q4 Alex US Trial Balances.xlsx"
base, ext = os.path.splitext(file_path)
new_file_path = base + "Final" + ext  # e.g., "Q4 Alex US Trial BalancesFinal.xlsx"

# Copy the original file to the new file, preserving all original sheets
shutil.copy(file_path, new_file_path)

##########################################
# 2. Define sheet mapping (sheet name to desired column name)
##########################################
# In each sheet, column A is "Main" (account number) and column B is "Closing balance"
sheet_map = {
    "West": "West",
    "NE": "NE",
    "MW": "MW",
    "NSS": "NSS",
    "Direct LLC": "Direct LLC"
}

##########################################
# 3. Read each sheet and merge data based on "Main" (outer join)
##########################################
merged_df = None

for sheet_name, col_name in sheet_map.items():
    # Read columns A and B from each sheet
    df_temp = pd.read_excel(
        new_file_path,
        sheet_name=sheet_name,
        usecols=[0, 1],   # Column A and Column B
        header=0,         # Use the first row as header
        engine="openpyxl"
    )
    # Rename columns to "Main" and the corresponding sheet name
    df_temp.columns = ["Main", col_name]
    
    # Merge data based on "Main" using outer join
    if merged_df is None:
        merged_df = df_temp
    else:
        merged_df = pd.merge(merged_df, df_temp, on="Main", how="outer")

##########################################
# 4. Fill missing values with 0 in each numeric column
##########################################
for col in sheet_map.values():
    merged_df[col] = merged_df[col].fillna(0)

##########################################
# 5. Add a "Closing balance" column that sums the balances from all sheets
##########################################
merged_df["Closing balance"] = merged_df[list(sheet_map.values())].sum(axis=1)

##########################################
# 6. Reorder columns: "Main", sheet columns, then "Closing balance"
##########################################
final_cols = ["Main"] + list(sheet_map.values()) + ["Closing balance"]
merged_df = merged_df[final_cols]

##########################################
# 7. Append the new "Consolidated" sheet to the new file while preserving the original sheets
##########################################
with pd.ExcelWriter(new_file_path, mode="a", if_sheet_exists="replace", engine="openpyxl") as writer:
    merged_df.to_excel(writer, sheet_name="Consolidated", index=False)

print("Done! The 'Consolidated' sheet has been added to the new file:")
print(new_file_path)
