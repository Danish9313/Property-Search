import pandas as pd
from main import CustomerDataStore, detect_property_columns, build_analysis_prompt
import config

df = pd.read_excel(config.EXCEL_FILE_PATH, dtype=str)
df = df.loc[:, ~df.columns.str.startswith("Unnamed")]
row = df.iloc[0]

json_store = CustomerDataStore(config.JSON_FOLDER_PATH)

cust_id = str(row.get("customer_id") or row.get("CUST_NUMBER") or "?").strip()
status_col = next((c for c in row.index if c.lower() == "property status"), "")
status = row.get(status_col, "")

print("=== ROW SELECTED ===")
print(f"  customer_id    : {cust_id}")
print(f"  Property Status: {status}")
print()

txt_files = json_store.find(cust_id)
print(f"=== PROPERTY TXT FILES FOR customer_id '{cust_id}' ===")
property_data = []
for i, item in enumerate(txt_files, 1):
    label = f"Property {i}"
    print(f"  [{label}] -> {item['filename']}")
    property_data.append((label, item["filename"], item["content"]))
print()

prompt = build_analysis_prompt(row, property_data)
print("=== PROMPT BUILT ===")
print(f"  Total characters : {len(prompt)}")
print(f"  Properties found : {len(property_data)}")
print()
print("--- PROMPT PREVIEW (first 600 chars) ---")
print(prompt[:600])
print("...")
