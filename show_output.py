import pandas as pd

# Set pandas options to display all columns and prevent truncation
pd.set_option('display.max_columns', None)
pd.set_option('display.width', 200)

try:
    # Read both sheets from the Excel file
    df_part_a = pd.read_excel("automation_output.xlsx", sheet_name="PartA_Company_Revenue")
    df_part_b = pd.read_excel("automation_output.xlsx", sheet_name="PartB_Contact_Enrichment")

    print("--- Contents of PartA_Company_Revenue ---")
    print(df_part_a.to_string())
    print("\n" + "="*50 + "\n")
    print("--- Contents of PartB_Contact_Enrichment ---")
    print(df_part_b.to_string())

except FileNotFoundError:
    print("The file 'automation_output.xlsx' was not found.")
except Exception as e:
    print(f"An error occurred while reading the file: {e}")
