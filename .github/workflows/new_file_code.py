import pandas as pd
import os

# ---------------- PATH ----------------
base_path = os.getcwd()

# ---------------- INPUT FILE ----------------
input_file = os.path.join(base_path, "HLR_HSS_input.xlsx")
if not os.path.exists(input_file):
    raise FileNotFoundError(f"❌ Input file not found: {input_file}")

output_file = os.path.join(base_path, "KPI_report_HLR_HSS.xlsx")
print("📥 Using input file:", input_file)

# ---------------- READ ALL SHEETS ----------------
excel_file = pd.ExcelFile(input_file)
df = pd.concat(
    [pd.read_excel(input_file, sheet_name=s) for s in excel_file.sheet_names],
    ignore_index=True
)

# ---------------- CLEAN DATA ----------------
df['Period start time'] = pd.to_datetime(df['Period start time'], errors='coerce')
df = df.dropna(subset=['Period start time'])
second_col = df.columns[1]

# ---------------- KPI DEFINITIONS ----------------
if second_col == "HSSFE name":
    print("Detected HSS Report")
    key_col = "HSSFE name"
    interfaces = {
        "S6a_Interface": {
            "S6ULR": ("S6a ULR FR", "FR"),
            "S6AIR": ("S6 AIR FR", "FR")
        },
        "Cx_Interface": {
            "CxUAR": ("Cx UAR FR", "FR"),
            "CxSAR": ("Cx SAR FR", "FR")
        },
        "Sh_Interface": {
            "ShUDR": ("Sh UDR FR", "FR"),
            "ShPUR": ("Sh PUR FR", "FR")
        }
    }
elif second_col == "NTHLRFE name":
    print("Detected HLR Report")
    key_col = "NTHLRFE name"
    interfaces = {
        "HLR_Interface": {
            "VLR": ("VLR LUP FR", "FR"),
            "SRI": ("SRI FR", "FR"),
            "VLR_SUCC": ("HLR_VLRLU_success", "SR"),
            "SRI_FR": ("HLR_SRI_failure_ratio", "FR")
        }
    }
else:
    raise Exception("Unknown report format")

# ---------------- HELPER FUNCTIONS ----------------
def format_label(ts):
    return ts.strftime('%d%b_%H:%M') if pd.notna(ts) else "NA"

def remark_logic(today, prev, metric_type):
    if pd.isna(today) or pd.isna(prev):
        return "Data Missing"
    diff = today - prev
    if metric_type == "FR":
        return "In degraded" if diff > 1 else "In Trend"
    if metric_type == "SR":
        return "In degraded" if (today < 99 or diff < 0) else "In Trend"

# ---------------- DATE DETECTION ----------------
df["date"] = df["Period start time"].dt.date
df["time"] = df["Period start time"].dt.time

dates = sorted(df["date"].unique())
today_date = dates[-1]
yesterday_date = dates[-2] if len(dates) >= 2 else None

# Latest timestamp today
today_ts = df[df["date"] == today_date]["Period start time"].max()
latest_time = today_ts.time()

# Same timestamp yesterday
yesterday_ts = None
if yesterday_date:
    same_time_rows = df[
        (df["date"] == yesterday_date) &
        (df["time"] == latest_time)
    ]
    if not same_time_rows.empty:
        yesterday_ts = same_time_rows["Period start time"].max()

today_label = format_label(today_ts)
y_label = format_label(yesterday_ts)

df_today = df[df["Period start time"] == today_ts]
df_y = df[df["Period start time"] == yesterday_ts] if yesterday_ts else pd.DataFrame()

# ---------------- BUILD INTERFACE DASHBOARDS ----------------
results = {}

for interface, metrics in interfaces.items():
    result = pd.DataFrame()

    for short, (col, metric_type) in metrics.items():

        if col not in df.columns:
            continue

        # ✅ FIX: GROUP BY to avoid duplicate blades
        today = (
            df_today.groupby(key_col)[col]
            .mean()  # change to .max() if needed
            .reset_index()
            .rename(columns={col: f"{short}_{today_label}"})
        )

        yest = (
            df_y.groupby(key_col)[col]
            .mean()
            .reset_index()
            .rename(columns={col: f"{short}_{y_label}"})
        ) if not df_y.empty else pd.DataFrame()

        # Merge
        temp = pd.merge(yest, today, on=key_col, how="outer")

        # Difference
        temp[f"{short}_Difference_N-1"] = (
            temp[f"{short}_{today_label}"] - temp[f"{short}_{y_label}"]
        )

        # Remarks
        temp[f"{short}_Remarks_N-1"] = temp.apply(
            lambda x: remark_logic(
                x.get(f"{short}_{today_label}"),
                x.get(f"{short}_{y_label}"),
                metric_type
            ),
            axis=1
        )

        if result.empty:
            result = temp
        else:
            result = pd.merge(result, temp, on=key_col, how="outer")

    # ✅ Optional: sort blades
    if not result.empty:
        result = result.sort_values(by=key_col)

    results[interface] = result

# ---------------- EXPORT TO EXCEL ----------------
with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
    workbook = writer.book

    header = workbook.add_format({
        "bold": True, "bg_color": "#4F81BD",
        "font_color": "white", "align": "center"
    })

    green = workbook.add_format({"bg_color": "#C6EFCE"})
    red = workbook.add_format({"bg_color": "#FFC7CE"})
    yellow = workbook.add_format({"bg_color": "#FFF2CC"})

    for sheet, data in results.items():

        if data.empty:
            continue

        data.to_excel(writer, sheet_name=sheet, index=False)
        worksheet = writer.sheets[sheet]

        rows = len(data)
        cols = len(data.columns)

        # Header format
        for c in range(cols):
            worksheet.write(0, c, data.columns[c], header)

        key_col_idx = data.columns.get_loc(key_col)

        for c, name in enumerate(data.columns):

            if "Remarks" in name:

                worksheet.conditional_format(
                    1, c, rows, c,
                    {'type': 'text', 'criteria': 'containing', 'value': 'In Trend', 'format': green}
                )

                worksheet.conditional_format(
                    1, c, rows, c,
                    {'type': 'text', 'criteria': 'containing', 'value': 'Data Missing', 'format': yellow}
                )

                worksheet.conditional_format(
                    1, c, rows, c,
                    {'type': 'text', 'criteria': 'containing', 'value': 'In degraded', 'format': red}
                )

                # Highlight blade if degraded
                col_letter = chr(65 + c)
                worksheet.conditional_format(
                    1, key_col_idx, rows, key_col_idx,
                    {
                        'type': 'formula',
                        'criteria': f'=${col_letter}2="In degraded"',
                        'format': red
                    }
                )

print("✅ Automation Completed Successfully")
print("📄 Output file:", output_file)