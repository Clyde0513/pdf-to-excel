import pandas as pd

def show_complete_data():
    excel_path = "Document250616132824_v3.xlsx"
    
    try:
        df = pd.read_excel(excel_path, sheet_name='Registry_Staff')
        
        print("=== COMPLETE REGISTRY STAFF DATA ===")
        print(f"Total rows: {len(df)}")
        print()
        
        # Set pandas to show all rows and columns
        pd.set_option('display.max_rows', None)
        pd.set_option('display.max_columns', None)
        pd.set_option('display.width', None)
        pd.set_option('display.max_colwidth', None)
        
        # Show only the relevant columns (skip Row_Number and Page)
        staff_cols = ['TheraEX', 'Intuitive', 'Vitawerks', 'Vitawerks Cont']
        
        print("All staff by category:")
        print("-" * 80)
        
        for col in staff_cols:
            if col in df.columns:
                # Get non-empty values
                staff_list = df[col].dropna()
                staff_list = staff_list[staff_list != '']
                
                print(f"\n{col}: ({len(staff_list)} people)")
                for i, name in enumerate(staff_list, 1):
                    print(f"  {i:2d}. {name}")
        
        print(f"\n=== VERIFICATION ===")
        print("Let's check the last few rows to see if data is complete:")
        print()
        
        # Show last 10 rows
        last_rows = df[['TheraEX', 'Intuitive', 'Vitawerks', 'Vitawerks Cont']].tail(10)
        print("Last 10 rows:")
        for idx, row in last_rows.iterrows():
            print(f"Row {idx + 1}: ", end="")
            for col in staff_cols:
                val = row[col] if pd.notna(row[col]) and row[col] != '' else '(empty)'
                print(f"{col}: {val[:20]:<20} ", end="")
            print()
        
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    show_complete_data()
