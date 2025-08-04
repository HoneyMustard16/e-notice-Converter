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
                # Convert to string and handle NaN/None values
                value = str(notice_row[col]) if pd.notna(notice_row[col]) else ""
                output_lines.append(
                    f"{col}={value.zfill(7)}"
                )  # Ensure it has leading zeros
            elif "t_lat" in col.lower():
                # Convert to string and handle NaN/None values
                value = str(notice_row[col]) if pd.notna(notice_row[col]) else ""
                output_lines.append(
                    f"{col}={value.zfill(6)}"
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
            rx_station_rows = rx_station_df[rx_station_df["t_adm_ref_id"] == adm_ref_id]

            if not rx_station_rows.empty:
                # Check if any row has t_geo_type == "MULTIPOINT"
                if (rx_station_rows["t_geo_type"].str.upper() == "MULTIPOINT").any():
                    output_lines.append("<RX_STATION>")
                    output_lines.append("t_geo_type=MULTIPOINT")
                    # For each point (row), output a POINT block with t_lat and t_long
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
                    output_lines.append("<RX_STATION>")
                    # Get the first row since we output only one RX_STATION block
                    rx_row = rx_station_rows.iloc[0]
                    t_geo_type = str(rx_row.get("t_geo_type", "")).upper()

                    for col in rx_station_rows.columns:
                        if col == "t_adm_ref_id":
                            continue
                        # Skip columns based on t_geo_type
                        if t_geo_type == "ZONE" and col in [
                            "t_lat",
                            "t_long",
                            "t_radius",
                        ]:
                            continue
                        if t_geo_type == "CIRCLE" and col == "t_zone_id":
                            continue
                        if t_geo_type == "POINT" and col in [
                            "t_ctry",
                            "t_site_name",
                            "t_noise_temp",
                        ]:
                            continue
                        output_lines.append(f"{col}={str(rx_row[col])}")
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
