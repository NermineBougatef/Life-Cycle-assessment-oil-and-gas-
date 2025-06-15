import pandas as pd

def extract_tables_with_columns(file_path, sheet_name='Summary'):
    # Read the entire sheet without assuming headers
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)

    tables = []
    # Look for all header rows containing "Stages"
    for idx, row in df.iterrows():
        if str(row[0]).strip() == "Stages":
            # Use this row as header
            header = df.iloc[idx].tolist()
            # Collect table data until next blank or next header
            table_data = df.iloc[idx+1:]
            table_data = table_data[table_data[0].notna()]
            table_data.columns = header
            tables.append(table_data)

    # Extract "Stages" and "Total kg CO2 eq" columns from each table
    extracted = pd.concat([t[["Stages", "Total kg CO2 eq"]] for t in tables], ignore_index=True)

    return extracted

if __name__ == "__main__":
    input_file = "file8 (1).xlsx"  # Replace with your actual file name
    output_file = "Summary_of_CO2_emissions.xlsx"
    extracted_data = extract_tables_with_columns(input_file)
    extracted_data.to_excel(output_file, index=False)
    print(f"Extracted data saved to {output_file}")
