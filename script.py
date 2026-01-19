import pandas as pd
import re
from datetime import datetime

# ---------------------------
# File paths
# ---------------------------
file_a = "File_A.csv"
file_b = "File_B.csv"

# ---------------------------
# STEP 1: Read File A
# ---------------------------
df_a = pd.read_csv(file_a, usecols=[0])
df_a = df_a.iloc[1:]  # Skip header
df_a.columns = ["Device_Name"]

unique_a = set(df_a["Device_Name"].dropna().unique())

# ---------------------------
# STEP 2: Read File B (from row 11)
# ---------------------------
df_b = pd.read_csv(
    file_b,
    usecols=[1, 4],
    skiprows=10,
    header=None
)

df_b.columns = ["Device_Name", "Attendance"]

df_b["Attendance"] = df_b["Attendance"].replace(
    "முன்பே சப்மிட் செய்துவிட்டேன்", 0
)

df_b["Attendance"] = pd.to_numeric(df_b["Attendance"], errors="coerce").fillna(0)

df_c = (
    df_b.groupby("Device_Name", as_index=False)["Attendance"]
    .max()
)

# ---------------------------
# STEP 3: Compare A vs B
# ---------------------------
existing_devices = set(df_c["Device_Name"])
missing_devices = unique_a - existing_devices

df_c["Category"] = "Polls"

missing_df = pd.DataFrame({
    "Device_Name": list(missing_devices),
    "Attendance": 0,
    "Category": "Not_Entered"
})

final_df = pd.concat([df_c, missing_df], ignore_index=True)

# ---------------------------
# STEP 4: Zoom total
# ---------------------------
zoom_total = final_df["Attendance"].sum()

# ---------------------------
# STEP 5: Extract meeting date from line 3
# ---------------------------
df_temp = pd.read_csv(
    file_b,
    usecols=[3],
    skiprows=10,
    header=None
)

meeting_ts = df_temp.iloc[0, 0]
meeting_date = pd.to_datetime(meeting_ts).date()
report_title = f"Report for the Meeting held on {meeting_date}"

# ---------------------------
# STEP 6: User input
# ---------------------------
in_person = int(input("Enter In-Person attendance count: "))
total_attendance = zoom_total + in_person

# ---------------------------
# STEP 7: Write Excel Report
# ---------------------------

file_c = f"Output_{meeting_date}.xlsx"

with pd.ExcelWriter(file_c, engine="openpyxl") as writer:
    sheet = "Report"

    # Title
    pd.DataFrame([report_title]).to_excel(
        writer,
        sheet_name=sheet,
        index=False,
        header=False,
        startrow=0
    )

    # Main table
    final_df.to_excel(
        writer,
        sheet_name=sheet,
        index=False,
        startrow=2
    )

    # Summary
    summary_start = 2 + len(final_df) + 2

    summary_data = [
        ["Meeting Attendance - Summary", ""],
        ["On Zoom", zoom_total],
        ["In Person", in_person],
        ["Total", total_attendance]
    ]

    pd.DataFrame(summary_data).to_excel(
        writer,
        sheet_name=sheet,
        index=False,
        header=False,
        startrow=summary_start
    )

print(f"{file_c} generated successfully.")
