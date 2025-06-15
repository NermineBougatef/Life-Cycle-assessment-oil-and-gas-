import pandas as pd

# Load cleaned dataset from Excel file
df = pd.read_excel("extracted_component_co2.xlsx")

# Create a simple prototype calculator function
def co2_emission_calculator(component_name: str, user_quantity: float, user_unit: str):
    """
    Prototype function to calculate CO2 emissions based on user input
    """
    # Filter dataset for matching component
    match = df[df['Component'].str.contains(component_name, case=False, na=False)]

    if match.empty:
        return f"Component '{component_name}' not found in dataset."

    match = match.iloc[0]  # Take first matching row

    input_unit = match['Input_Unit']
    conversion_note = match['Notes']

    try:
        # Calculate scaling factor
        base_quantity = match['Input_Quantity']
        base_emission = match['CO2_eq_kg']

        scaling_factor = user_quantity / base_quantity
        estimated_emission = scaling_factor * base_emission

        # Build result
        result = {
            "Component": match['Component'],
            "User Quantity": f"{user_quantity} {user_unit}",
            "Emission Estimate (kg CO2 eq)": round(estimated_emission, 2),
            "Base Unit": f"{base_quantity} {input_unit}",
            "Emission Factor": f"{round(base_emission / base_quantity, 2)} kg CO2/{input_unit}",
            "Notes": conversion_note
        }
        return result

    except Exception as e:
        return f"Error calculating emissions: {str(e)}"

# Interactive user input
if __name__ == "__main__":
    component = input("Enter the component name: ")
    quantity = float(input("Enter the quantity: "))
    unit = input("Enter the unit (e.g., m3, kg): ")

    result = co2_emission_calculator(component, quantity, unit)
    print("\n--- CO2 Emission Estimate ---")
    if isinstance(result, dict):
        for key, value in result.items():
            print(f"{key}: {value}")
    else:
        print(result)
