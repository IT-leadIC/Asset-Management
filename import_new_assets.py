
import sqlite3
import pandas as pd

EXCEL_FILE = "IT_Assets.xlsx"
DATABASE_FILE = "assets.db"


def clean_value(value):
    if pd.isna(value) or str(value).strip() == "":
        return "Not Assigned"
    return str(value).strip()


# Read the latest Excel data
df = pd.read_excel(EXCEL_FILE, header=1)

conn = sqlite3.connect(DATABASE_FILE)
cursor = conn.cursor()

# Confirm that the existing database is ready
cursor.execute("""
    SELECT name FROM sqlite_master
    WHERE type = 'table' AND name = 'assets'
""")

if cursor.fetchone() is None:
    conn.close()
    raise SystemExit(
        "Assets table not found. Check that assets.db exists."
    )

added = 0
existing = 0

for _, row in df.iterrows():
    asset_id = clean_value(row["Asset No"])

    if asset_id == "Not Assigned":
        continue

    cursor.execute("""
        INSERT OR IGNORE INTO assets
        (asset_id, asset_type, model, serial_no, status, assigned_to)
        VALUES (?, ?, ?, ?, ?, ?)
    """, (
        asset_id,
        clean_value(row["Asset Type"]),
        clean_value(row["Model"]),
        clean_value(row["Serial No"]),
        clean_value(row["Status"]),
        clean_value(row["Employee"])
    ))

    if cursor.rowcount == 1:
        added += 1
    else:
        existing += 1

conn.commit()
conn.close()

print(f"New assets added: {added}")
print(f"Existing assets left unchanged: {existing}")
print("Asset database sync completed.")
