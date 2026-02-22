import pandas as pd

file_path = "_h_batch_process_data.xlsx"
prod_file = "_h_batch_production_data.xlsx"

production_df = pd.read_excel(
    prod_file,
    sheet_name="BatchData"
)

print("Production DF Shape:", production_df.shape)

# Collect all batch sheets (ignore "Summary")
batch_sheets = [s for s in xls.sheet_names if s.startswith("Batch_")]

all_batches = []

for sheet in batch_sheets:
    df = pd.read_excel(file_path, sheet_name=sheet)

    # Extract Batch_ID from sheet name
    df["Batch_ID"] = sheet.replace("Batch_", "")

    all_batches.append(df)

# Combine into one dataframe
process_df = pd.concat(all_batches, ignore_index=True)

print(process_df.shape)
print(process_df["Batch_ID"].nunique())
energy_features = process_df.groupby("Batch_ID").agg({
    "Power_Consumption_kW": ["mean", "max", "std"],
    "Vibration_mm_s": ["mean", "max"],
    "Temperature_C": "std",
    "Pressure_Bar": "std"
})

energy_features.columns = [
    "avg_power",
    "max_power",
    "std_power",
    "avg_vibration",
    "max_vibration",
    "temp_variation",
    "pressure_variation"
]

energy_features = energy_features.reset_index()

print(energy_features.shape)
full_data = production_df.merge(
    energy_features,
    on="Batch_ID",
    how="left"
)

print(full_data.shape)