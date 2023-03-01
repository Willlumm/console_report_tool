# Script to load GSD and GFK console market data from Excel and text files, process, combine, and load into an Excel report.
# Files in the input folder are manually replaced weekly with the latest files exported from GSD and GFK.
# Files in the past folder are used to avoid exporting multiple years of data from GSD/GFK.
# Files in the extrap folder are used to extrapolate values and append additional data.
# William Lumme

import os
import pandas as pd
import re
import xlwings as xw

# Hardcoded classes, territories, and HD sizes.
classes = {
    "4 PRO":            "PRO",
    "4 SLIM":           "SLIM",
    "DIGITAL EDITION":  "DIGITAL EDITION",
    "SWITCH LITE":      "LITE",
    "OLED":             "OLED",
    "SWITCH 64 GB":     "OLED",
    "XBOX ONE S":       "ONE S",
    "XBOX ONE X":       "ONE X",
    "XBOX SERIES S":    "SERIES S",
    "XBOX SERIES X":    "SERIES X",
}

territories = {
    "GSA":      "SWITZERLAND",
    "BENE":     "BENELUX",
    "OCEANIA":  "ANZ",
    "ASIA":     "JAPAN",
}

hdsizes = {
    "32 GB":                "32 GB",
    "64 GB":                "64 GB",
    "OLED":                 "64 GB",
    "250 GB":               "250 GB",
    "500 GB":               "500 GB",
    "500GB":                "500 GB",
    "SONY PLAYSTATION 4":   "500 GB",
    "512":                  "512 GB",
    "825":                  "825 GB",
    "1 TB":                 "1 TB",
    "1TB":                  "1 TB",    
    "2 TB":                 "2 TB",
    "2TB":                  "2 TB"
}

# Load new GSD data from Excel file. Throws error if no file matching the name scheme is found.
def load_gsd():
    for filename in os.listdir("input"):
        if re.fullmatch(r"\w{8}-(\w{4}-){3}\w{12}\.xlsx", filename):
            print(f"Loading new GSD data from {filename}...")
            gsd_input = filename
            break
    else:
        raise Exception("GSD file not found")
    return pd.read_excel(f"input/{gsd_input}")

# Load new GFK data from text file. 
def load_gfk():
    print("Loading new GFK data...")
    return pd.read_csv("input/NEW_HW_DATA.txt", sep="\t")

# Load past GSD data from CSV files.
def load_past_gsd():
    dfs = []
    for filename in os.listdir("past"):
        if re.fullmatch(r"gsd.*\.csv", filename):
            print(f"Loading past GSD data from {filename}...")
            df = pd.read_csv(f"past/{filename}")
            dfs.append(df)
    return pd.concat(dfs)

# Load past GFK data from text files.
def load_past_gfk():
    dfs = []
    for filename in os.listdir("past"):
        if re.fullmatch(r"gfk.*\.txt", filename):
            print(f"Loading past GFK data from {filename}...")
            df = pd.read_csv(f"past/{filename}", sep="\t")
            dfs.append(df)
    df = pd.concat(dfs)
    return df

# Process GSD data.
def process_gsd(gsd, dates):
    print("Processing GSD data...")

    # Filter by country and platform.
    gsd = gsd[gsd["Country"] != "UNITED KINGDOM"]
    gsd = gsd[gsd["Platform"].isin(["PS4", "PS5", "SWITCH", "XBOX ONE", "XBOX SERIES"])]

    # Add default fields.
    gsd["Source"] = "GSD"
    gsd["CLASS"] = "ORIGINAL"

    # Update class and territory.
    for string, tag in classes.items():
        gsd.loc[gsd["SKU"].str.contains(string), "CLASS"] = tag
    gsd["Territory"] = gsd["Territory"].replace(territories)

    # Add fiscal date.
    gsd = gsd.merge(dates, how="left", left_on=["Year", "Week"], right_on=["YEAR", "WEEK"])

    # Extrapolate values.
    gsd_extrap = pd.read_csv("extrap/EXTRAPOLATION HW GSD.csv")
    gsd_extrap["Territory"] = gsd_extrap["Territory"].str.upper()
    gsd = gsd.merge(gsd_extrap, how="left", left_on=["Country", "FY", "Week", "Platform"], right_on=["Territory", "FY", "Week", "Format"])
    gsd["Units 100%"] = gsd["Units"] / gsd["Extrapolation"]
    gsd["Value Euro 100%"] = gsd["Values"] / gsd["Extrapolation"]
    gsd["Value Local 100%"] = ""
    
    # Rename and return data.
    gsd = gsd.rename(columns={"HD Size": "HDSize", "Territory_x": "Territory", "Units": "Panel Units", "Values": "Panel Value EURO"})
    gsd = gsd[["Source", "SKU", "Platform", "Bundle", "HDSize", "CLASS", "Country", "Territory", "FY", "Year", "MONTH NEW", "Week", "Panel Units", "Panel Value EURO", "Extrapolation", "Units 100%", "Value Euro 100%", "Value Local 100%"]]
    return gsd

# Process GFK data.
def process_gfk(gfk, dates):
    print("Processing GFK data...")

    # Add default fields.
    gfk["Source"] = "GFK"
    gfk["CLASS"] = "ORIGINAL"
    gfk["HDSize"] = "UNKNOWN"

    # Add fiscal date.
    gfk = gfk.merge(dates, how="left", left_on=["Year (W)", "Week (W)"], right_on=["YEAR", "WEEK"])

    # Extrapolate values.
    gfk_extrap = pd.read_csv("extrap/EXTRAPOLATION HW GFK.csv")
    gfk = gfk.merge(gfk_extrap, how="left", left_on=["Country", "FY", "Main Platform"], right_on=["Territory", "FY", "Format"])
    gfk[["Units Panel (W)", "Value Panel (W)"]] = gfk[["Units Panel (W)", "Value Panel (W)"]].replace(",", "", regex=True)
    gfk["Units 100%"] = gfk["Units Panel (W)"].astype(float) / gfk["Extrapolation"]
    gfk["Value Euro 100%"] = gfk["Value Panel (W)"].astype(float) / gfk["Extrapolation"]
    gfk["Value Local 100%"] = ""

    # Update bundle, platform, territory, class, and HD size
    gfk["Bundle"] = gfk["Bundle"].replace({0: "STANDALONE", 1: "BUNDLE"})
    gfk["Main Platform"] = gfk["Main Platform"].replace("NINTENDO SWITCH", "SWITCH")
    gfk["Territory"] = gfk["Territory"].str.upper()
    for string, tag in classes.items():
        gfk.loc[gfk["Article Name"].str.contains(string, na=False), "CLASS"] = tag
    for string, tag in hdsizes.items():
        gfk.loc[gfk["Article Name"].str.contains(string, na=False), "HDSize"] = tag
    
    # Rename and return data.
    gfk = gfk.rename(columns={"Article Name": "SKU", "Main Platform": "Platform", "Year (W)": "Year", "Week (W)": "Week", "Units Panel (W)": "Panel Units", "Value Panel (W)": "Panel Value EURO"})
    gfk = gfk[gfk["Platform"].isin(["PS4", "PS5", "SWITCH", "XBOX ONE", "XBOX SERIES"])]
    gfk = gfk[["Source", "SKU", "Platform", "Bundle", "HDSize", "CLASS", "Country", "Territory", "FY", "Year", "MONTH NEW", "Week", "Panel Units", "Panel Value EURO", "Extrapolation", "Units 100%", "Value Euro 100%", "Value Local 100%"]]
    return gfk

def main():
    # Date table for joining financial year
    dates = pd.read_csv("extrap/DATES.csv")

    gsd = load_gsd()
    gsd_past = load_past_gsd()
    gsd = pd.concat([gsd, gsd_past])
    gsd = process_gsd(gsd, dates)
    
    gfk = load_gfk()
    gfk_past = load_past_gfk()
    gfk = pd.concat([gfk, gfk_past])
    gfk = process_gfk(gfk, dates)

    # Load substitute data for missing values.
    extrap = pd.read_csv("extrap/Germany Extrap.csv")

    # Combine and save in report.
    output = pd.concat([gsd, gfk, extrap])
    output = output[output.FY >= 2021]
    wb = xw.Book("EUA Weekly Console HW Report CY23 WK07.xlsx")
    sheet = wb.sheets("weekly")
    sheet.range("A:R").clear_contents()
    print("Saving in report...")
    sheet["A1"].options(index=False, header=True).value = output

if __name__=="__main__":
    main()