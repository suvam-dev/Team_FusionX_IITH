import pandas as pd

# ===============================
# STEP 1 — LOAD PROCESS DATA (MULTI-SHEET)
# ===============================

file_path = "_h_batch_process_data.xlsx"

# ⭐ YOU MUST CREATE xls FIRST
xls = pd.ExcelFile(file_path)

# Get all batch sheets
batch_sheets = [s for s in xls.sheet_names if s.startswith("Batch_")]

all_batches = []

for sheet in batch_sheets:
    df = pd.read_excel(file_path, sheet_name=sheet)

    # Extract Batch_ID from sheet name
    df["Batch_ID"] = sheet.replace("Batch_", "")

    all_batches.append(df)

# Combine all sheets
process_df = pd.concat(all_batches, ignore_index=True)

print("Process DF:", process_df.shape)

# ===============================
# STEP 2 — CREATE ENERGY FEATURES
# ===============================

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

print("Energy Features:", energy_features.shape)

# ===============================
# STEP 3 — LOAD PRODUCTION DATA
# ===============================

production_df = pd.read_excel(
    "_h_batch_production_data.xlsx",
    sheet_name="BatchData"
)

print("Production DF Shape:", production_df.shape)

# ===============================
# STEP 4 — MERGE
# ===============================

full_data = production_df.merge(
    energy_features,
    on="Batch_ID",
    how="left"
)

print("Merged Data:", full_data.shape)
from sklearn.preprocessing import MinMaxScaler

# Targets where HIGHER is better
high_good = [
    "Hardness",
    "Dissolution_Rate",
    "Content_Uniformity"
]

# Targets where LOWER is better
low_good = [
    "Friability",
    "Disintegration_Time",
    "avg_power"
]

scaler = MinMaxScaler()

full_data[high_good + low_good] = scaler.fit_transform(
    full_data[high_good + low_good]
)

# Composite industrial score
full_data["score"] = (
    full_data["Hardness"] +
    full_data["Dissolution_Rate"] +
    full_data["Content_Uniformity"] +
    (1 - full_data["Friability"]) +
    (1 - full_data["Disintegration_Time"]) +
    (1 - full_data["avg_power"])
) / 6

print(full_data[["Batch_ID", "score"]].head())
# Select top-performing batches
golden_batches = full_data.sort_values(
    by="score",
    ascending=False
).head(10)

print("\nTop Golden Batches:")
print(golden_batches[["Batch_ID", "score"]])
golden_params = golden_batches[[
    "Granulation_Time",
    "Binder_Amount",
    "Drying_Temp",
    "Drying_Time",
    "Machine_Speed",
    "Lubricant_Conc",
    "Moisture_Content"
]]

print("\nGolden Signature Parameters:")
print(golden_params)
# Define golden threshold (top 15% batches)
golden_threshold = full_data["score"].quantile(0.85)

print("\nGolden Score Threshold:", golden_threshold)

def recommend_adjustments(batch_id):

    current = full_data[full_data["Batch_ID"] == batch_id]

    if current.empty:
        print("Batch not found")
        return

    golden_mean = golden_params.mean()

    print(f"\nAdaptive Recommendations for Batch {batch_id}:\n")

    for col in golden_mean.index:
        current_val = current.iloc[0][col]
        target_val = golden_mean[col]

        diff = target_val - current_val
        percent_change = (diff / current_val) * 100 if current_val != 0 else 0

        if abs(percent_change) > 5:
            direction = "Increase" if diff > 0 else "Decrease"
            print(f"{direction} {col} by {round(abs(percent_change),1)}%")
def adaptive_golden_update():

    global golden_params

    # Find batches exceeding threshold
    new_gold = full_data[full_data["score"] >= golden_threshold]

    # Extract parameter columns
    updated_params = new_gold[[
        "Granulation_Time",
        "Binder_Amount",
        "Drying_Temp",
        "Drying_Time",
        "Machine_Speed",
        "Lubricant_Conc",
        "Moisture_Content"
    ]]

    golden_params = updated_params

    print("\nGolden Signature Database Updated!")
    print("Total Golden Batches:", len(golden_params))

adaptive_golden_update()
recommend_adjustments("T001")

# ==================================
# PARETO OPTIMIZATION ENGINE
# ==================================

def get_pareto_front(data):

    pareto = []

    for i, row in data.iterrows():
        dominated = False

        for j, other in data.iterrows():

            # Objectives:
            # maximize score
            # maximize Hardness
            # minimize avg_power

            if (
                other["score"] >= row["score"] and
                other["Hardness"] >= row["Hardness"] and
                other["avg_power"] <= row["avg_power"] and
                (
                    other["score"] > row["score"] or
                    other["Hardness"] > row["Hardness"] or
                    other["avg_power"] < row["avg_power"]
                )
            ):
                dominated = True
                break

        if not dominated:
            pareto.append(row)

    return pd.DataFrame(pareto)


pareto_front = get_pareto_front(full_data)

print("\nPareto Optimal Batches:")
print(pareto_front[["Batch_ID", "score", "Hardness","avg_power"]])
# ==================================
# USER INPUT INTERFACE
# ==================================
