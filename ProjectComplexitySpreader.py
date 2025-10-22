#%%
# Build the month-by-month breakout as requested.

import pandas as pd
import numpy as np
import math
#from caas_jupyter_tools import display_dataframe_to_user
#%%
# Load data
df = pd.read_excel(r"C:\Users\thebner\OneDrive - Q2e\PlatformBaselineByPhase_HardCode.xlsx", sheet_name="Baselines")
#%%
# Ensure expected columns exist
expected_cols = {'Complexity', 'Role', 'Region', 'Phase', 'PhaseDurDays', 'Hours per Day'}
missing = expected_cols - set(df.columns)
if missing:
    raise ValueError(f"Missing expected columns: {missing}")
#%%
# Define the ordered phase sequence for timeline placement
phase_order = [
    'Build',
    '1st Third of Config',
    '2nd Third of Config',
    '3rd Third of Config',
    'LPT',
    'Go-Live',
    'Transition to Support'
]

# Sanity: restrict to known phases in our order to avoid surprises
df = df[df['Phase'].isin(phase_order)].copy()
#%%
# Compute per-complexity phase window length as the MAX duration across roles/regions (for timeline offsets)
phase_windows = (
    df.groupby(['Complexity', 'Phase'])['PhaseDurDays']
    .max()
    .reset_index()
)
#%%
# Build a mapping for each complexity: phase -> (start_day, window_days)
complexity_phase_timeline = {}
for comp, comp_block in phase_windows.groupby('Complexity'):
    start = 1
    mapping = {}
    # iterate in the defined order; if a phase is missing for a complexity, treat as 0-day window
    for ph in phase_order:
        if ph in set(comp_block['Phase']):
            window_days = int(comp_block.loc[comp_block['Phase'] == ph, 'PhaseDurDays'].iloc[0])
        else:
            window_days = 0
        mapping[ph] = (start, window_days)
        start += window_days
    complexity_phase_timeline[comp] = mapping
#%%
# Determine global project length in days and months across all complexities
total_days_by_comp = {
    comp: sum(window for (_, window) in mapping.values())
    for comp, mapping in complexity_phase_timeline.items()
}
global_total_days = max(total_days_by_comp.values()) if total_days_by_comp else 0
global_total_months = int(math.ceil(global_total_days / 30.0)) if global_total_days > 0 else 0
#%%
# Function to expand a single row into month allocations
def allocate_row_to_months(row):
    comp = row['Complexity']
    phase = row['Phase']
    hours_per_day = float(row['Hours per Day'])
    dur_days = int(row['PhaseDurDays'])

    start_day, _window_days = complexity_phase_timeline[comp][phase]
    # This role/region may have fewer days than the phase window; use its own duration
    # Allocate day-by-day into 30-day months
    month_alloc = {}
    for day_offset in range(dur_days):
        day = start_day + day_offset
        month_idx = (day - 1) // 30 + 1  # Month 1 = days 1-30, etc.
        month_alloc[month_idx] = month_alloc.get(month_idx, 0.0) + hours_per_day
    return month_alloc
#%%
# Compute allocations per Complexity + Role + Region
group_cols = ['Complexity', 'Role', 'Region']
allocations = {}

for idx, row in df.iterrows():
    key = (row['Complexity'], row['Role'], row['Region'])
    month_alloc = allocate_row_to_months(row)
    if key not in allocations:
        allocations[key] = {}
    for m, hrs in month_alloc.items():
        allocations[key][m] = allocations[key].get(m, 0.0) + hrs
#%%
# Build the final table with Month 1...Month N
month_cols = [f"Month {i}" for i in range(1, global_total_months + 1)]
rows_out = []
for (comp, role, region), month_map in allocations.items():
    row_dict = {'Complexity': comp, 'Role': role, 'Region': region}
    for i in range(1, global_total_months + 1):
        row_dict[f"Month {i}"] = round(month_map.get(i, 0.0), 4)
    rows_out.append(row_dict)

out_df = pd.DataFrame(rows_out, columns=['Complexity', 'Role', 'Region'] + month_cols)
#%%
# Sort for readability
out_df = out_df.sort_values(['Complexity', 'Role', 'Region']).reset_index(drop=True)
#%%
# Save to files for download
csv_path = r"C:\Users\thebner\OneDrive - Q2e\Monthly_Breakout_By_Complexity_Role_Region.csv"
xlsx_path = r"C:\Users\thebner\OneDrive - Q2e\Monthly_Breakout_By_Complexity_Role_Region.xlsx"
out_df.to_csv(csv_path, index=False)
with pd.ExcelWriter(xlsx_path, engine='xlsxwriter') as writer:
    out_df.to_excel(writer, sheet_name="Monthly Breakout", index=False)

# Show a preview table to the user
print("Monthly Breakout by Complexity, Role, Region", out_df.head(50))

(csv_path, xlsx_path, out_df.shape, global_total_days, global_total_months)
