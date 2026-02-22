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
            ):1
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

def add_new_batch():
    """Add a new batch with manual data entry"""
    global full_data, golden_batches, pareto_front
    
    print("\n" + "="*50)
    print("ADD NEW BATCH DATA")
    print("="*50)
    
    batch_id = input("Enter Batch ID (e.g., T999): ").strip().upper()
    
    # Check if batch already exists
    if batch_id in full_data["Batch_ID"].values:
        print(f"\n❌ Batch {batch_id} already exists!")
        return
    
    try:
        # Production parameters
        print("\n--- Production Parameters ---")
        granulation_time = float(input("Granulation Time (minutes): "))
        binder_amount = float(input("Binder Amount (g): "))
        drying_temp = float(input("Drying Temperature (°C): "))
        drying_time = float(input("Drying Time (minutes): "))
        machine_speed = float(input("Machine Speed (RPM): "))
        lubricant_conc = float(input("Lubricant Concentration (%): "))
        moisture_content = float(input("Moisture Content (%): "))
        
        # Quality metrics
        print("\n--- Quality Metrics ---")
        hardness = float(input("Hardness (N): "))
        dissolution_rate = float(input("Dissolution Rate (%): "))
        content_uniformity = float(input("Content Uniformity (%): "))
        friability = float(input("Friability (%): "))
        disintegration_time = float(input("Disintegration Time (minutes): "))
        
        # Energy metrics
        print("\n--- Energy Metrics ---")
        avg_power = float(input("Average Power Consumption (kW): "))
        max_power = float(input("Max Power Consumption (kW): "))
        std_power = float(input("Std Power Consumption (kW): "))
        avg_vibration = float(input("Average Vibration (mm/s): "))
        max_vibration = float(input("Max Vibration (mm/s): "))
        temp_variation = float(input("Temperature Variation (°C): "))
        pressure_variation = float(input("Pressure Variation (Bar): "))
        
        # Create new batch dictionary
        new_batch = {
            "Batch_ID": batch_id,
            "Granulation_Time": granulation_time,
            "Binder_Amount": binder_amount,
            "Drying_Temp": drying_temp,
            "Drying_Time": drying_time,
            "Machine_Speed": machine_speed,
            "Lubricant_Conc": lubricant_conc,
            "Moisture_Content": moisture_content,
            "Hardness": hardness,
            "Dissolution_Rate": dissolution_rate,
            "Content_Uniformity": content_uniformity,
            "Friability": friability,
            "Disintegration_Time": disintegration_time,
            "avg_power": avg_power,
            "max_power": max_power,
            "std_power": std_power,
            "avg_vibration": avg_vibration,
            "max_vibration": max_vibration,
            "temp_variation": temp_variation,
            "pressure_variation": pressure_variation
        }
        
        # Normalize the new batch using existing scaler
        high_good = ["Hardness", "Dissolution_Rate", "Content_Uniformity"]
        low_good = ["Friability", "Disintegration_Time", "avg_power"]
        
        normalized_values = scaler.transform(
            pd.DataFrame([new_batch])[high_good + low_good]
        )[0]
        
        for i, col in enumerate(high_good + low_good):
            new_batch[col] = normalized_values[i]
        
        # Calculate score
        new_batch["score"] = (
            new_batch["Hardness"] +
            new_batch["Dissolution_Rate"] +
            new_batch["Content_Uniformity"] +
            (1 - new_batch["Friability"]) +
            (1 - new_batch["Disintegration_Time"]) +
            (1 - new_batch["avg_power"])
        ) / 6
        
        # Add to full_data
        full_data = pd.concat([full_data, pd.DataFrame([new_batch])], ignore_index=True)
        
        print(f"\n✅ Batch {batch_id} added successfully!")
        print(f"Score: {new_batch['score']:.4f}")
        
        # Update golden batches and pareto front
        golden_batches = full_data.sort_values(by="score", ascending=False).head(10)
        pareto_front = get_pareto_front(full_data)
        
    except ValueError:
        print("\n❌ Invalid input. Please enter valid numbers.")

while True:
    print("\n" + "="*50)
    print("BATCH OPTIMIZATION MENU")
    print("="*50)
    print("1. Get recommendations for a specific batch")
    print("2. View Pareto optimal batches")
    print("3. View golden batches")
    print("4. View all batch scores")
    print("5. Add new batch data")
    print("6. Exit")
    print("="*50)
    
    choice = input("\nSelect an option (1-6): ").strip()
    
    if choice == "1":
        batch_id = input("Enter Batch ID (e.g., T001): ").strip().upper()
        recommend_adjustments(batch_id)
    
    elif choice == "2":
        print("\nPareto Optimal Batches:")
        print(pareto_front[["Batch_ID", "score", "Hardness", "avg_power"]])
    
    elif choice == "3":
        print("\nTop Golden Batches:")
        print(golden_batches[["Batch_ID", "score"]])
    
    elif choice == "4":
        print("\nAll Batch Scores:")
        print(full_data[["Batch_ID", "score"]].sort_values(by="score", ascending=False))
    
    elif choice == "5":
        add_new_batch()
    
    elif choice == "6":
        print("\nExiting... Goodbye!")
        break
    
    else:
        print("\nInvalid option. Please select 1-6.")
