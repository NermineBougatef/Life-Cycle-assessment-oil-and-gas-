# 📊 Excel CO₂ Emissions Data Cleaning & Formatting

This repository contains Python scripts and Jupyter Notebooks designed for **extracting**, **cleaning**, and **formatting CO₂ emissions data** from an Excel file (`file8.xlsx`) related to environmental impact and Life Cycle Assessment (LCA) studies.  
**The dataset used is specific to the United States (US).**

---

## 📁 Project Structure
├── extract_inventory_columns.py      # Extracts 'Component' and 'Total kg CO2 eq' from "Inventory" sheet
├── format_impacts_data.py           # Cleans and formats 'Impacts' sheet with merged headers
├── extract_summary_stages.py        # Extracts 'Stages' and 'Total kg CO2 eq' from multiple tables in "Summary" sheet
├── extracted_component_co2.xlsx     # Output: cleaned and styled Excel file from "Impacts"
├── columns_A_D_CO2_output.xlsx      # Output: simple cleaned file from "Inventory"
├── Summary_of_CO2_emissions.xlsx    # Output: consolidated CO2 data from "Summary" sheet
└── README.md                        # This file
