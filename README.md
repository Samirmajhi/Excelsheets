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
