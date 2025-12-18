import pandas as pd

def create_dict_blacklisted(excel_file_path):
    # Load Excel
    df = pd.read_excel(excel_file_path)

    # Normalize column names
    df.columns = df.columns.str.strip()

    # Forward-fill operator names (handles blank rows)
    df["PayTV IPTV Operator"] = df["PayTV IPTV Operator"].ffill()

    # Clean partnerStationdId values
    df["partnerStationdId"] = (
        df["partnerStationdId"]
        .astype(str)
        .str.replace("\u00a0", "", regex=True)
        .str.strip()
    )

    # Add prefix to each ID
    prefix = "epgProvider:st."
    df["partnerStationdId"] = prefix + df["partnerStationdId"]

    # Build dictionary
    operator_station_map = (
        df.groupby("PayTV IPTV Operator")["partnerStationdId"]
        .apply(list)
        .to_dict()
    )

    return operator_station_map

if __name__ == "__main__":
    ## Blacklisted stations
    blacklist_file = "Stingray Blacklisted Channels - PayTV.xlsx"
    blacklist_location = f"TivoPlus_source_files/{blacklist_file}"

    values = create_dict_blacklisted(blacklist_location)
    print(values)
