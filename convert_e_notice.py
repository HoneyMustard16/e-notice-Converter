import pandas as pd


def convert_excel_to_text(excel_file, output_file):
    """
    Converts an Excel file containing e-notice data into a text format.
    Processes multiple sheets: HEAD, NOTICE, ANTENNA, RX STATION, TAIL.
    Outputs a structured text file with XML-like tags.
    """
    # Read the Excel file sheets into separate DataFrames
    # HEAD and TAIL sheets are read with default types
    # NOTICE, ANTENNA, RX STATION sheets are read as strings to preserve formatting
    head_df = pd.read_excel(excel_file, sheet_name="HEAD")
    notice_df = pd.read_excel(excel_file, sheet_name="NOTICE", dtype=str)
    antenna_df = pd.read_excel(excel_file, sheet_name="ANTENNA", dtype=str)
    rx_station_df = pd.read_excel(excel_file, sheet_name="RX STATION", dtype=str)
    tail_df = pd.read_excel(excel_file, sheet_name="TAIL")

    # List to accumulate lines of the output text file
    output_lines = []

    # Process HEAD section: write opening tag
    output_lines.append("<HEAD>")
    # Iterate over each row in HEAD sheet
    for index, row in head_df.iterrows():
        # For each column in the row
        for col in head_df.columns:
            # If the column name contains 'date' or 'sent', convert to date string
            if "date" in col.lower() or "sent" in col.lower():
                if not pd.isna(row[col]):
                    date_value = pd.to_datetime(row[col]).date()
                    output_lines.append(f"{col}={str(date_value)}")
            else:
                # Otherwise, convert the value to string and append
                if not pd.isna(row[col]):
                    val = str(row[col])
                    if val != "":
                        output_lines.append(f"{col}={val}")
    # Write closing tag for HEAD section
    output_lines.append("</HEAD>")

    # Process each row in NOTICE sheet
    for index, notice_row in notice_df.iterrows():
        output_lines.append("<NOTICE>")
        # Iterate over each column in NOTICE row
        for col in notice_df.columns:
            # For longitude columns, pad with leading zeros to length 7
            if "t_long" in col.lower():
                if pd.notna(notice_row[col]):
                    value = str(notice_row[col]).zfill(7)
                    output_lines.append(f"{col}={value}")
            # For latitude columns, pad with leading zeros to length 6
            elif "t_lat" in col.lower():
                if pd.notna(notice_row[col]):
                    value = str(notice_row[col]).zfill(6)
                    output_lines.append(f"{col}={value}")
            else:
                # For other columns, convert to string and append
                value = str(notice_row[col])
                if value == "nan":
                    value = ""
                if value != "":
                    output_lines.append(f"{col}={value}")

        # Link ANTENNA and RX STATION data using t_adm_ref_id from NOTICE row
        adm_ref_id = notice_row["t_adm_ref_id"]
        antenna_row = antenna_df[antenna_df["t_adm_ref_id"] == adm_ref_id]

        # If matching ANTENNA data found
        if not antenna_row.empty:
            output_lines.append("<ANTENNA>")
            # Append all ANTENNA columns except the reference id
            for col in antenna_row.columns:
                if col != "t_adm_ref_id":
                    value = str(antenna_row.iloc[0][col])
                    if value == "nan":
                        value = ""
                    if value != "":
                        output_lines.append(f"{col}={value}")

            # Find matching RX STATION rows for the same reference id
            rx_station_rows = rx_station_df[rx_station_df["t_adm_ref_id"] == adm_ref_id]

            if not rx_station_rows.empty:
                # Check if any RX STATION row has t_geo_type == "MULTIPOINT"
                if (rx_station_rows["t_geo_type"].str.upper() == "MULTIPOINT").any():
                    output_lines.append("<RX_STATION>")
                    output_lines.append("t_geo_type=MULTIPOINT")
                    # For each MULTIPOINT row, output a POINT block with latitude and longitude
                    multipoint_rows = rx_station_rows[
                        rx_station_rows["t_geo_type"].str.upper() == "MULTIPOINT"
                    ]
                    for _, point_row in multipoint_rows.iterrows():
                        output_lines.append("<POINT>")
                        t_lat = (
                            str(point_row["t_lat"])
                            if pd.notna(point_row["t_lat"])
                            else ""
                        )
                        t_long = (
                            str(point_row["t_long"])
                            if pd.notna(point_row["t_long"])
                            else ""
                        )
                        output_lines.append(f"t_lat={t_lat}")
                        output_lines.append(f"t_long={t_long}")
                        output_lines.append("</POINT>")
                    output_lines.append("</RX_STATION>")
                else:
                    # For other geo types, output a single RX_STATION block
                    output_lines.append("<RX_STATION>")
                    rx_row = rx_station_rows.iloc[0]
                    t_geo_type = str(rx_row.get("t_geo_type", "")).upper()

                    # Append columns selectively based on geo type
                    for col in rx_station_rows.columns:
                        if col == "t_adm_ref_id":
                            continue
                        if t_geo_type == "ZONE" and col in [
                            "t_lat",
                            "t_long",
                            "t_radius",
                        ]:
                            continue
                        if t_geo_type == "CIRCLE" and col == "t_zone_id":
                            continue
                        if t_geo_type == "POINT" and col in [
                            "t_zone_id",
                            "t_radius",
                        ]:
                            continue
                        value = str(rx_row[col])
                        if value == "nan":
                            value = ""
                        if value != "":
                            output_lines.append(f"{col}={value}")
                    output_lines.append("</RX_STATION>")
            output_lines.append("</ANTENNA>")
        output_lines.append("</NOTICE>")

    # Process TAIL section: write opening tag
    output_lines.append("<TAIL>")
    # Iterate over each row in TAIL sheet
    for index, row in tail_df.iterrows():
        # For each column in the row
        for col in tail_df.columns:
            # Convert date or sent columns to date string
            if "date" in col.lower() or "sent" in col.lower():
                if not pd.isna(row[col]):
                    date_value = pd.to_datetime(row[col]).date()
                    output_lines.append(f"{col}={str(date_value)}")
            else:
                # Otherwise convert to string
                if not pd.isna(row[col]):
                    val = str(row[col])
                    if val != "":
                        output_lines.append(f"{col}={val}")
    # Write closing tag for TAIL section
    output_lines.append("</TAIL>")

    # Write all accumulated lines to the output text file
    with open(output_file, "w") as f:
        for line in output_lines:
            f.write(f"{line}\n")
