import streamlit as st
import pandas as pd
from datetime import datetime
import io

st.set_page_config(page_title="Meeting Attendance", layout="centered")

st.title("📊 Meeting Attendance Generation Dashboard")

# ---------------------------
# Helper function to load CSV safely
# ---------------------------
def load_csv(file, usecols=None, skiprows=None, header='infer', sep=","):
    if file is None:
        st.warning("No file uploaded yet!")
        return None
    try:
        file.seek(0)  # Reset pointer before reading
        df = pd.read_csv(file, usecols=usecols, skiprows=skiprows, header=header, sep=sep)
        return df
    except pd.errors.EmptyDataError:
        st.error("The uploaded CSV is empty!")
    except pd.errors.ParserError:
        st.error("CSV parsing error! Maybe wrong separator or malformed CSV.")
    except UnicodeDecodeError:
        try:
            file.seek(0)
            df = pd.read_csv(file, usecols=usecols, skiprows=skiprows, header=header, sep=sep, encoding="utf-8-sig")
            return df
        except Exception as e:
            st.error(f"Failed to read CSV with utf-8-sig: {e}")
    return None

# ---------------------------
# File upload widgets
# ---------------------------
file_a = st.file_uploader("Upload File A (CSV)", type=["csv"])
file_b = st.file_uploader("Upload File B (CSV)", type=["csv"])

in_person = st.number_input(
    "Enter In-Person Attendance Count",
    min_value=0,
    step=1
)

if file_a and file_b and st.button("Generate Report"):
    # ---------------------------
    # STEP 1: Load File A
    # ---------------------------
    df_a = load_csv(file_a, usecols=[0])
    if df_a is None:
        st.stop()
    df_a = df_a.iloc[1:]  # Skip header row if needed
    df_a.columns = ["Device_Name"]
    unique_a = set(df_a["Device_Name"].dropna().unique())
    
    # st.write("File A Preview:")
    # st.dataframe(df_a)

    # ---------------------------
    # STEP 2: Load File B (from row 11)
    # ---------------------------
    df_b = load_csv(file_b, usecols=[1, 4], skiprows=10, header=None)
    if df_b is None:
        st.stop()
    df_b.columns = ["Device_Name", "Attendance"]
    
    # Convert Attendance to numeric
    df_b["Attendance"] = df_b["Attendance"].replace(
        "முன்பே சப்மிட் செய்துவிட்டேன்", 0
    )
    df_b["Attendance"] = pd.to_numeric(df_b["Attendance"], errors="coerce").fillna(0)

    # st.write("File B Preview (from row 11):")
    #st.dataframe(df_b)

    # ---------------------------
    # STEP 3: Process and compare
    # ---------------------------
    df_c = df_b.groupby("Device_Name", as_index=False)["Attendance"].max()
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
    # STEP 5: Extract meeting date from File B (row 11, col D)
    # ---------------------------
    df_temp = load_csv(file_b, usecols=[3], skiprows=10, header=None)
    if df_temp is None or df_temp.empty:
        st.error("Could not extract meeting date from File B!")
        st.stop()
    meeting_ts = df_temp.iloc[0, 0]
    meeting_date = pd.to_datetime(meeting_ts).date()
    report_title = f"Report for the Meeting held on {meeting_date}"

    # ---------------------------
    # STEP 6: Total attendance
    # ---------------------------
    total_attendance = zoom_total + in_person

    # ---------------------------
    # STEP 7: Display final report
    # ---------------------------
    st.write(f"### {report_title}")
    st.write("Final Device Attendance:")
    st.dataframe(final_df)

    st.write("Meeting Attendance - Summary")
    summary_df = pd.DataFrame({
        "Category": ["On Zoom", "In Person", "Total"],
        "Count": [zoom_total, in_person, total_attendance]
    })
    st.dataframe(summary_df)


    # ---------------------------
    # STEP 8: Create Excel in memory
    # ---------------------------
    output = io.BytesIO()
    file_name = f"Output_{meeting_date}.xlsx"

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        sheet = "Report"

        pd.DataFrame([report_title]).to_excel(
            writer,
            sheet_name=sheet,
            index=False,
            header=False,
            startrow=0
        )

        final_df.to_excel(
            writer,
            sheet_name=sheet,
            index=False,
            startrow=2
        )

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

    st.success("Report generated successfully 🎉")

    st.download_button(
        label="📥 Download Excel Report",
        data=output.getvalue(),
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
