import pandas as pd
import os

db_path = 'master_procurement_database.csv'

if os.path.exists(db_path):
    print("Loading existing master database...")
    df = pd.read_csv(db_path)
    
    print(f"Original row count: {len(df)}")
    
    # 1. Filter out non-material rows (Commercial Parameters, stray metadata, etc.)
    junk_materials = ['Commercial Parameter', 'Basic Value', 'Sofs', 'Set', 'Sofa', 'Cat', 'nan', 'NaN', '']
    df = df[~df['Material_Description'].isin(junk_materials) & df['Material_Description'].notna()].copy()
    
    # 2. Filter out dummy test vendors if desired (optional, you can comment this out if you want to keep them)
    dummy_vendors = ['Captain America', 'Batman', 'Dog Infra', 'Cat Infra', 'Dog Cat Infra', 'Maira karira Tech']
    df = df[~df['Vendor_Name'].isin(dummy_vendors)].copy()
    
    # 3. Ensure proper numeric conversion and fix inflated Basic Amounts
    df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce').fillna(0.0)
    df['Rate_R0'] = pd.to_numeric(df['Rate_R0'], errors='coerce').fillna(0.0)
    df['Rate_R1'] = pd.to_numeric(df['Rate_R1'], errors='coerce').fillna(df['Rate_R0'])
    
    # Recalculate Basic Amount mathematically to eliminate manual entry errors
    df['Basic_Amount'] = df['Quantity'] * df['Rate_R1']
    
    # Save cleaned database
    df.to_csv(db_path, index=False)
    print(f"Database successfully cleaned! Remaining valid rows: {len(df)}")
else:
    print("Error: master_procurement_database.csv not found in this directory.")