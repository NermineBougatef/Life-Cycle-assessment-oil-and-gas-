import pandas as pd

# 1. Définir le chemin du fichier
file_path = "file8 (1).xlsx"

# 2. Charger le fichier Excel
xls = pd.ExcelFile(file_path)
df = xls.parse("Impacts")

# 3. Utiliser la 2e ligne comme en-tête et supprimer les lignes inutiles
df.columns = df.iloc[1]
df = df.drop([0, 1]).reset_index(drop=True)
df = df.dropna(axis=1, how='all')  # Supprimer les colonnes vides

# 4. Renommer les colonnes pour plus de clarté
df = df.rename(columns={
    'Design Input': 'Input_Quantity',
    'Unit': 'Input_Unit',
    'Functional Input': 'Mass_kg',
    'kg CO2 eq': 'CO2_eq_kg',
    'Notes:': 'Notes'
})

# 5. Garder uniquement les colonnes nécessaires
df = df[[ 'Input_Quantity', 'Input_Unit', 'Mass_kg', 'CO2_eq_kg']]

# 6. Convertir les colonnes numériques
for col in ['Input_Quantity', 'Mass_kg', 'CO2_eq_kg']:
    df[col] = pd.to_numeric(df[col], errors='coerce')

# 7. Supprimer les lignes incomplètes
df_clean = df.dropna(subset=['Mass_kg', 'CO2_eq_kg'])

# 8. Exporter proprement

## Option 1 : Fichier Excel structuré
df_clean.to_excel("cleaned_co2_dataset.xlsx", index=False)
print("✅ Le fichier structuré a été sauvegardé sous 'cleaned_co2_dataset.xlsx'")

## Option 2 : Fichier CSV structuré avec séparateur ;
df_clean.to_csv("cleaned_co2_dataset.csv", index=False, sep=';')
print("✅ Le fichier structuré a aussi été sauvegardé sous 'cleaned_co2_dataset.csv'")

