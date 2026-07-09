"""
generate_sample_data.py
------------------------
Creates a messy, realistic sample sales dataset (raw_data.csv) so you can
test the cleaning & reporting pipeline immediately without needing your
own file.

Run:
    python generate_sample_data.py
"""

import pandas as pd
import numpy as np
import random

random.seed(42)
np.random.seed(42)

n_rows = 300

regions = ["North", "South", "East", "West", "north", "SOUTH", " East ", None]
products = ["Laptop", "Mobile", "Tablet", "Headphones", "laptop", "MOBILE", None]
reps = ["Alice", "Bob", "Charlie", "David", "Eva", None]

rows = []
for i in range(n_rows):
    order_id = 1000 + i
    region = random.choice(regions)
    product = random.choice(products)
    rep = random.choice(reps)
    units = random.choice([1, 2, 3, 5, 10, None, -1])  # includes bad value (-1) and missing
    price = random.choice([250, 500, 750, 1000, None])
    date = random.choice(
        ["2024-01-05", "2024/01/06", "05-01-2024", "2024-02-15", "2024-02-30", None]
    )  # includes malformed/invalid date
    rows.append(
        {
            "OrderID": order_id,
            "Region": region,
            "Product": product,
            "SalesRep": rep,
            "UnitsSold": units,
            "UnitPrice": price,
            "OrderDate": date,
        }
    )

df = pd.DataFrame(rows)

# Inject duplicate rows on purpose
duplicates = df.sample(20, random_state=1)
df = pd.concat([df, duplicates], ignore_index=True)

# Shuffle
df = df.sample(frac=1, random_state=2).reset_index(drop=True)

df.to_csv("raw_data.csv", index=False)
print(f"Created raw_data.csv with {len(df)} rows (including duplicates & missing values).")
