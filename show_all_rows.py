import pandas as pd

def show_all_rows():
    excel_path = "Document250616132824_v3.xlsx"
    
    df = pd.read_excel(excel_path, sheet_name='Registry_Staff')
    
    print("=== ALL STAFF DATA ===")
    print(f"Total rows: {len(df)}")
    print()
    
    # Show all data with row numbers
    staff_cols = ['TheraEX', 'Intuitive', 'Vitawerks', 'Vitawerks Cont']
    
    print("Row | TheraEX | Intuitive | Vitawerks | Vitawerks Cont")
    print("-" * 100)
    
    for idx, row in df.iterrows():
        row_num = idx + 1
        theraex = row['TheraEX'] if pd.notna(row['TheraEX']) and row['TheraEX'] != '' else ''
        intuitive = row['Intuitive'] if pd.notna(row['Intuitive']) and row['Intuitive'] != '' else ''
        vitawerks = row['Vitawerks'] if pd.notna(row['Vitawerks']) and row['Vitawerks'] != '' else ''
        vitawerks_cont = row['Vitawerks Cont'] if pd.notna(row['Vitawerks Cont']) and row['Vitawerks Cont'] != '' else ''
        
        print(f"{row_num:2d}  | {theraex:20s} | {intuitive:20s} | {vitawerks:25s} | {vitawerks_cont}")

if __name__ == "__main__":
    show_all_rows()
