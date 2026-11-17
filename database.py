import sqlite3
import pandas as pd

EXCEL_FILE = "IT_Assets.xlsx"
DATABASE_FILE = "assets.db"


def clean_value(value):
    if pd.isna(value):
        return "Not Assigned"

    value = str(value).strip()

    if value == "":
        return "Not Assigned"

    return value


# Read Excel
df = pd.read_excel(EXCEL_FILE, header=1)

# Connect to SQLite
conn = sqlite3.connect(DATABASE_FILE)

cursor = conn.cursor()

# Create assets table
cursor.execute("""
CREATE TABLE IF NOT EXISTS assets (
    asset_id TEXT PRIMARY KEY,
    asset_type TEXT,
    model TEXT,
    serial_no TEXT,
    status TEXT,
    assigned_to TEXT
)
""")

# Import Excel data
for _, row in df.iterrows():

    asset_id = clean_value(row["Asset No"])

    if asset_id == "Not Assigned":
        continue

    asset_type = clean_value(row["Asset Type"])
    model = clean_value(row["Model"])
    serial_no = clean_value(row["Serial No"])
    status = clean_value(row["Status"])
    assigned_to = clean_value(row["Employee"])

    cursor.execute("""
    INSERT OR REPLACE INTO assets
    (asset_id, asset_type, model, serial_no, status, assigned_to)
    VALUES (?, ?, ?, ?, ?, ?)
    """, (
        asset_id,
        asset_type,
        model,
        serial_no,
        status,
        assigned_to
    ))

conn.commit()
conn.close()

print("✅ Asset database created successfully!")