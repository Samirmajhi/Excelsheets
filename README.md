import pandas as pd

# Load Excel file and sheets
file_path = 'your_file.xlsx'
xls = pd.ExcelFile(file_path)

new_df = xls.parse('NewSheet')    # Replace with your actual sheet name
old_df = xls.parse('OldSheet')    # Replace with your actual sheet name

# Clean and normalize key columns
new_df['Code'] = new_df['Code'].astype(str).str.strip()
new_df['DR Scenerio'] = new_df['DR Scenerio'].astype(str).str.strip()
old_df['Code'] = old_df['Code'].astype(str).str.strip()
old_df['DR Scenerio'] = old_df['DR Scenerio'].astype(str).str.strip()

# Convert date columns
new_df['Recovery Plan'] = pd.to_datetime(new_df['Recovery Plan'], errors='coerce')
old_df['Recovery Plan'] = pd.to_datetime(old_df['Recovery Plan'], errors='coerce')

# New list to hold merged rows and set to track matched keys
merged_rows = []
matched_keys = set()

# Loop through new_df rows
for _, new_row in new_df.iterrows():
    code = new_row['Code']
    scenario = new_row['DR Scenerio']
    key = (code, scenario)

    match = old_df[(old_df['Code'] == code) & (old_df['DR Scenerio'] == scenario)]

    if not match.empty:
        old_row = match.iloc[0]

        new_date = new_row['Recovery Plan']
        old_date = old_row['Recovery Plan']

        if pd.isna(new_date) and pd.isna(old_date):
            merged_rte = new_row['RTE']
            merged_rpe = new_row['RPE']
            merged_date = pd.NaT
        elif pd.isna(new_date):
            merged_rte = old_row['RTE']
            merged_rpe = old_row['RPE']
            merged_date = old_date
        elif pd.isna(old_date):
            merged_rte = new_row['RTE']
            merged_rpe = new_row['RPE']
            merged_date = new_date
        else:
            if new_date > old_date:
                merged_rte = new_row['RTE']
                merged_rpe = new_row['RPE']
                merged_date = new_date
            else:
                merged_rte = old_row['RTE']
                merged_rpe = old_row['RPE']
                merged_date = old_date

        new_row['RTE'] = merged_rte
        new_row['RPE'] = merged_rpe
        new_row['Recovery Plan'] = merged_date

        matched_keys.add(key)
        merged_rows.append(new_row)
    else:
        # No match found — just add the row from new_df
        merged_rows.append(new_row)
        matched_keys.add(key)

# Add unmatched rows from old_df directly (no comparison)
for _, old_row in old_df.iterrows():
    key = (old_row['Code'], old_row['DR Scenerio'])
    if key not in matched_keys:
        merged_rows.append(old_row)

# Final merged DataFrame
final_df = pd.DataFrame(merged_rows)





import pandas as pd

# Load Excel file and specific sheets
file_path = 'your_excel_file.xlsx'
new_df = pd.read_excel(file_path, sheet_name='Plan')
old_df = pd.read_excel(file_path, sheet_name='Plan Old')

# Ensure Recovery Plan is datetime
new_df['Recovery Plan'] = pd.to_datetime(new_df['Recovery Plan'], errors='coerce')
old_df['Recovery Plan'] = pd.to_datetime(old_df['Recovery Plan'], errors='coerce')

# Fill NaN in RTE and RPE for safety
for col in ['RTE', 'RPE']:
    new_df[col] = new_df[col].fillna('')
    old_df[col] = old_df[col].fillna('')

# Merge logic
merged_rows = []
matched_keys = set()

for _, new_row in new_df.iterrows():
    code = new_row['Code']
    dr = new_row['DR Scenerio']
    key = (code, dr)

    old_match = old_df[(old_df['Code'] == code) & (old_df['DR Scenerio'] == dr)]

    if not old_match.empty:
        old_row = old_match.iloc[0]

        # Compare Recovery Plan
        new_date = new_row['Recovery Plan']
        old_date = old_row['Recovery Plan']

        # Decision logic
        if pd.isna(new_date) and pd.isna(old_date):
            final_row = new_row
        elif pd.isna(new_date):
            final_row = old_row.copy()
        elif pd.isna(old_date):
            final_row = new_row
        else:
            final_row = new_row if new_date > old_date else old_row

        matched_keys.add(key)
        merged_rows.append(final_row)
    else:
        # No match found, keep new row
        merged_rows.append(new_row)
        matched_keys.add(key)

# Add unmatched old_df rows
for _, old_row in old_df.iterrows():
    key = (old_row['Code'], old_row['DR Scenerio'])
    if key not in matched_keys:
        merged_rows.append(old_row)

# Final DataFrame
merged_df = pd.DataFrame(merged_rows)

# Save result
merged_df.to_excel('merged_output.xlsx', index=False)









import pandas as pd

# Load Excel file and sheets
file_path = 'your_file.xlsx'
xls = pd.ExcelFile(file_path)

new_df = xls.parse('NewSheet')    # Replace with your actual sheet name
old_df = xls.parse('OldSheet')    # Replace with your actual sheet name

# Clean and normalize key columns
new_df['Code'] = new_df['Code'].astype(str).str.strip()
new_df['DR Scenerio'] = new_df['DR Scenerio'].astype(str).str.strip()
old_df['Code'] = old_df['Code'].astype(str).str.strip()
old_df['DR Scenerio'] = old_df['DR Scenerio'].astype(str).str.strip()

# Convert date columns
new_df['Recovery Plan'] = pd.to_datetime(new_df['Recovery Plan'], errors='coerce')
old_df['Recovery Plan'] = pd.to_datetime(old_df['Recovery Plan'], errors='coerce')

# New list to hold merged rows
merged_rows = []

# Loop through each row in new_df
for _, new_row in new_df.iterrows():
    code = new_row['Code']
    scenario = new_row['DR Scenerio']

    # Try to find matching row in old_df
    match = old_df[(old_df['Code'] == code) & (old_df['DR Scenerio'] == scenario)]

    if not match.empty:
        old_row = match.iloc[0]

        # Apply your logic:
        new_date = new_row['Recovery Plan']
        old_date = old_row['Recovery Plan']

        if pd.isna(new_date) and pd.isna(old_date):
            merged_rte = new_row['RTE']
            merged_rpe = new_row['RPE']
            merged_date = pd.NaT

        elif pd.isna(new_date):
            merged_rte = old_row['RTE']
            merged_rpe = old_row['RPE']
            merged_date = old_date

        elif pd.isna(old_date):
            merged_rte = new_row['RTE']
            merged_rpe = new_row['RPE']
            merged_date = new_date

        else:
            if new_date > old_date:
                merged_rte = new_row['RTE']
                merged_rpe = new_row['RPE']
                merged_date = new_date
            else:
                merged_rte = old_row['RTE']
                merged_rpe = old_row['RPE']
                merged_date = old_date

        # Update row
        new_row['RTE'] = merged_rte
        new_row['RPE'] = merged_rpe
        new_row['Recovery Plan'] = merged_date

    # Append updated or original row
    merged_rows.append(new_row)

# Convert list back to DataFrame
final_df = pd.DataFrame(merged_rows)






import pandas as pd

# Load the sheets from the same Excel file
excel_path = "your_file.xlsx"
new_df = pd.read_excel(excel_path, sheet_name="Plan")
old_df = pd.read_excel(excel_path, sheet_name="Plan Old")

# Convert Recovery Plan columns to datetime
new_df["Recovery Plan"] = pd.to_datetime(new_df["Recovery Plan"], errors="coerce")
old_df["Recovery Plan"] = pd.to_datetime(old_df["Recovery Plan"], errors="coerce")

# Create a dictionary for quick lookup from old_df
old_dict = {
    (row["Code"], row["DR Scenerio"]): row
    for _, row in old_df.iterrows()
}

# Final output list
merged_rows = []

# Iterate over each row in new_df
for _, new_row in new_df.iterrows():
    code = new_row["Code"]
    scenario = new_row["DR Scenerio"]
    new_rp = new_row["Recovery Plan"]

    key = (code, scenario)
    old_row = old_dict.get(key)

    if old_row is None:
        # Case: App doesn't exist in old sheet → keep as-is
        merged_rows.append(new_row)
    else:
        old_rp = old_row["Recovery Plan"]

        # If both dates are NaT
        if pd.isna(new_rp) and pd.isna(old_rp):
            row = new_row.copy()
            row["RTE"] = new_row["RTE"]
            row["RPE"] = new_row["RPE"]
            row["Recovery Plan"] = pd.NaT
        elif pd.isna(new_rp) and not pd.isna(old_rp):
            row = new_row.copy()
            row["RTE"] = old_row["RTE"]
            row["RPE"] = old_row["RPE"]
            row["Recovery Plan"] = old_row["Recovery Plan"]
        elif not pd.isna(new_rp) and pd.isna(old_rp):
            row = new_row.copy()
            # Keep new row's values already
        else:
            # Both have dates → pick the latest
            row = new_row.copy()
            if new_rp > old_rp:
                pass  # keep new_row as-is
            else:
                row["RTE"] = old_row["RTE"]
                row["RPE"] = old_row["RPE"]
                row["Recovery Plan"] = old_row["Recovery Plan"]

        merged_rows.append(row)

# Now add rows from old_df that were not in new_df
new_keys = set((row["Code"], row["DR Scenerio"]) for _, row in new_df.iterrows())
for _, old_row in old_df.iterrows():
    key = (old_row["Code"], old_row["DR Scenerio"])
    if key not in new_keys:
        merged_rows.append(old_row)

# Convert the final list of rows to a DataFrame
final_df = pd.DataFrame(merged_rows)

# Save to Excel
final_df.to_excel("merged_result.xlsx", index=False)











import pandas as pd
import random

# Define some sample data
regions = ['North', 'South', 'East', 'West']
products = ['Laptop', 'Tablet', 'Smartphone', 'Monitor', 'Printer']
salespersons = ['Alice', 'Bob', 'Charlie', 'Diana', 'Eve']

# Create a list to hold generated records
data = []

# Generate 100 rows of fake sales data
for _ in range(100):
    region = random.choice(regions)
    product = random.choice(products)
    salesperson = random.choice(salespersons)
    units_sold = random.randint(1, 20)
    unit_price = random.randint(100, 2000)  # in dollars
    total_sales = units_sold * unit_price

    data.append({
        'Region': region,
        'Salesperson': salesperson,
        'Product': product,
        'Units Sold': units_sold,
        'Unit Price': unit_price,
        'Total Sales': total_sales
    })

# Create DataFrame
df = pd.DataFrame(data)

# Save to CSV
csv_filename = 'sales_data.csv'
df.to_csv(csv_filename, index=False)

print(f"File '{csv_filename}' created successfully!")





=IF(
    ISNA(MATCH(1, (exer_old!C:C=C37)*(exer_old!E:E=E37), 0)),
    F37,
    IF(AND(F37="", INDEX(exer_old!F:F, MATCH(1, (exer_old!C:C=C37)*(exer_old!E:E=E37), 0))=""),
        "",
        IF(F37="", INDEX(exer_old!F:F, MATCH(1, (exer_old!C:C=C37)*(exer_old!E:E=E37), 0)),
            IF(INDEX(exer_old!F:F, MATCH(1, (exer_old!C:C=C37)*(exer_old!E:E=E37), 0))="", F37,
                IF(INDEX(exer_old!F:F, MATCH(1, (exer_old!C:C=C37)*(exer_old!E:E=E37), 0)) > F37,
                    INDEX(exer_old!F:F, MATCH(1, (exer_old!C:C=C37)*(exer_old!E:E=E37), 0)),
                    F37
                )
            )
        )
    )
)



=IF(
  ISNA(MATCH(1, (PLAN_OLD!B:B=B2)*(PLAN_OLD!E:E=E2), 0)),
  I2,
  IF(
    AND(I2="", INDEX(PLAN_OLD!I:I, MATCH(1, (PLAN_OLD!B:B=B2)*(PLAN_OLD!E:E=E2), 0))=""),
    I2,
    IF(
      INDEX(PLAN_OLD!I:I, MATCH(1, (PLAN_OLD!B:B=B2)*(PLAN_OLD!E:E=E2), 0))="",
      I2,
      IF(
        I2="",
        INDEX(PLAN_OLD!I:I, MATCH(1, (PLAN_OLD!B:B=B2)*(PLAN_OLD!E:E=E2), 0)),
        IF(
          INDEX(PLAN_OLD!I:I, MATCH(1, (PLAN_OLD!B:B=B2)*(PLAN_OLD!E:E=E2), 0)) > I2,
          INDEX(PLAN_OLD!I:I, MATCH(1, (PLAN_OLD!B:B=B2)*(PLAN_OLD!E:E=E2), 0)),
          I2
        )
      )
    )
  )
)# Excelsheets
