import pandas as pd
import os

def display_excel_structure(excel_path):
    """Display the structure of the created Excel file."""
    
    if not os.path.exists(excel_path):
        print(f"File not found: {excel_path}")
        return
    
    try:
        # Read all sheets
        excel_file = pd.ExcelFile(excel_path)
        print(f"Excel File: {excel_path}")
        print(f"Sheets: {excel_file.sheet_names}")
        print("=" * 60)
        
        for sheet_name in excel_file.sheet_names:
            print(f"\n SHEET: {sheet_name}")
            print("-" * 40)
            
            df = pd.read_excel(excel_path, sheet_name=sheet_name)
            print(f"Rows: {len(df)}")
            print(f"Columns: {list(df.columns)}")
            
            if sheet_name == 'Registry_Staff':
                print("\n SAMPLE DATA:")
                
                # Show header counts
                registry_cols = ['TheraEX', 'Intuitive', 'Vitawerks', 'Vitawerks Cont']
                for col in registry_cols:
                    if col in df.columns:
                        non_empty = len(df[df[col].notna() & (df[col] != '')])
                        print(f"  {col}: {non_empty} staff members")
                
                print(f"\n First 10 rows:")
                pd.set_option('display.max_columns', None)
                pd.set_option('display.width', None)
                print(df.head(10).to_string(index=False))
                
            elif sheet_name == 'Summary':
                print(f"\nSUMMARY DATA:")
                print(df.to_string(index=False))
            
            print()
        
    except Exception as e:
        print(f"Error reading Excel file: {e}")

if __name__ == "__main__":
    import sys
    excel_path = sys.argv[1] if len(sys.argv) > 1 else "Document250616132824_v3.xlsx"
    display_excel_structure(excel_path)
