import pandas as pd

def show_excel_output(file_path="automation_output.xlsx"):
    """
    Reads and displays the contents of both sheets from the specified Excel file.
    """
    try:
        # Read both sheets from the Excel file
        part_a_df = pd.read_excel(file_path, sheet_name="PartA_Company_Revenue")
        part_b_df = pd.read_excel(file_path, sheet_name="PartB_Contact_Enrichment")

        print("\n" + "="*50)
        print("          Part A: Company Revenue & Tier          ")
        print("="*50)
        
        if not part_a_df.empty:
            print(part_a_df.to_string())
        else:
            print("No data found in PartA_Company_Revenue sheet.")

        print("\n" + "="*50)
        print("        Part B: Contact Enrichment Results        ")
        print("="*50)

        if not part_b_df.empty:
            # To make it more readable, let's select and reorder key columns
            display_cols = [
                "Full Name", 
                "Current Company", 
                "Current Designation", 
                "LinkedIn URL", 
                "Work Email"
            ]
            # Filter for columns that actually exist to avoid errors
            existing_cols = [col for col in display_cols if col in part_b_df.columns]
            print(part_b_df[existing_cols].to_string())
        else:
            print("No data found in PartB_Contact_Enrichment sheet.")
            
        print("\n" + "="*50)
        print(f"Data read successfully from '{file_path}'")
        print("="*50)

    except FileNotFoundError:
        print(f"\n!!! Error: Output file '{file_path}' not found.")
        print("Please run 'agent_executor.py' first to generate the output file.")
    except Exception as e:
        print(f"\n!!! An error occurred while reading the Excel file: {e}")

if __name__ == "__main__":
    show_excel_output()
