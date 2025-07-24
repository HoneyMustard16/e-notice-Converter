import pandas as pd


def convert_excel_to_text(excel_file, output_file):
    # Read the Excel file with specified data types for t_long and t_lat
    head_df = pd.read_excel(excel_file, sheet_name="HEAD")
    notice_df = pd.read_excel(excel_file, sheet_name="NOTICE", dtype=str)
    antenna_df = pd.read_excel(excel_file, sheet_name="ANTENNA", dtype=str)
    rx_station_df = pd.read_excel(excel_file, sheet_name="RX STATION", dtype=str)
    tail_df = pd.read_excel(excel_file, sheet_name="TAIL")

    # Start building the output text
    output_lines = []

    # Process HEAD section
    output_lines.append("<HEAD>")
    for index, row in head_df.iterrows():
        for col in head_df.columns:
            if "date" in col.lower() or "sent" in col.lower():
                date_value = pd.to_datetime(row[col]).date()
                output_lines.append(f"{col}={str(date_value)}")
            else:
                output_lines.append(f"{col}={str(row[col])}")  # Convert to string
    output_lines.append("</HEAD>")

    for index, notice_row in notice_df.iterrows():
        output_lines.append("<NOTICE>")
        for col in notice_df.columns:
            if "t_long" in col.lower():
                # Use the value directly, preserving sign and leading zeros
                output_lines.append(
                    f"{col}={notice_row[col].zfill(7)}"
                )  # Ensure it has leading zeros
            elif "t_lat" in col.lower():
                # Use the value directly, preserving sign and leading zeros
                output_lines.append(
                    f"{col}={notice_row[col].zfill(6)}"
                )  # Ensure it has leading zeros
            else:
                output_lines.append(
                    f"{col}={str(notice_row[col])}"
                )  # Convert to string

        # Link ANTENNA and RX STATION using t_adm_ref_id
        adm_ref_id = notice_row["t_adm_ref_id"]
        antenna_row = antenna_df[antenna_df["t_adm_ref_id"] == adm_ref_id]

        if not antenna_row.empty:
            output_lines.append("<ANTENNA>")
            for col in antenna_row.columns:
                if col != "t_adm_ref_id":
                    output_lines.append(f"{col}={str(antenna_row.iloc[0][col])}")

            # Now include RX STATION inside ANTENNA
            rx_station_row = rx_station_df[rx_station_df["t_adm_ref_id"] == adm_ref_id]
            if not rx_station_row.empty:
                output_lines.append("<RX_STATION>")
                for col in rx_station_row.columns:
                    if col != "t_adm_ref_id":
                        output_lines.append(f"{col}={str(rx_station_row.iloc[0][col])}")
                output_lines.append("</RX_STATION>")
            output_lines.append("</ANTENNA>")
        output_lines.append("</NOTICE>")

    # Process TAIL section
    output_lines.append("<TAIL>")
    for index, row in tail_df.iterrows():
        for col in tail_df.columns:
            if "date" in col.lower() or "sent" in col.lower():
                date_value = pd.to_datetime(row[col]).date()
                output_lines.append(f"{col}={str(date_value)}")
            else:
                output_lines.append(f"{col}={str(row[col])}")
    output_lines.append("</TAIL>")

    # Write to output file
    with open(output_file, "w") as f:
        for line in output_lines:
            f.write(f"{line}\n")


# Example usage
convert_excel_to_text(
    "T12 DataFace 401 - 600.xlsx", "T12 e-notice 401 - 600 HF Pemerintah.txt"
)
