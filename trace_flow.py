import pandas as pd, re
from pathlib import Path

df = pd.read_excel("Updated Task - 206 Accounts- Pawnee Leasing Corporation - RG-2026-110.xlsx", dtype=str)
row = df.iloc[0]
lease_num = str(row["Lease #"]).strip()
lessee = str(row["Lessee"]).strip()
prop_details = str(row.get("Property Details", "")).strip()

print("=" * 60)
print("STEP 1 — Pick row from Excel")
print("=" * 60)
print(f"  Lease #          : {lease_num}")
print(f"  Lessee           : {lessee}")
print(f"  Property Details : {prop_details}")
print(f"  Filter passes?   : {'YES' if 'property found' in prop_details.lower() and not prop_details.lower().startswith('no') else 'NO — SKIPPED'}")

print()
print("=" * 60)
print("STEP 2 — Match Lease # to folder (numeric prefix match)")
print("=" * 60)
base = Path("206 Accounts - Pawnee Leasing Corporation - $6.04M - EFAs")
matched_folder = None
for d in base.iterdir():
    if not d.is_dir():
        continue
    m = re.match(r"^(\d+)", d.name)
    if m and m.group(1) == lease_num:
        matched_folder = d
        break

if matched_folder:
    print(f"  Lease # searched : {lease_num}")
    print(f"  Folder matched   : {matched_folder.name}")
else:
    print(f"  NO FOLDER FOUND for Lease # {lease_num}")

print()
print("=" * 60)
print("STEP 3 — Files loaded from matched folder")
print("=" * 60)
for f in sorted(matched_folder.glob("*.*")):
    size_kb = round(f.stat().st_size / 1024)
    print(f"  {f.suffix.upper():5s}  {size_kb:>6} KB  {f.name}")

print()
print("=" * 60)
print("STEP 4 — Send to OpenRouter API")
print("=" * 60)
print("  All files above get extracted (TXT read, PDF parsed)")
print("  Combined with Excel row data into one big prompt")
print("  Prompt sent to google/gemini-2.5-pro via OpenRouter")
print("  AI returns JSON with 15 analysis fields")
print("  Results written back to pawnee_EFA_output.xlsx")
