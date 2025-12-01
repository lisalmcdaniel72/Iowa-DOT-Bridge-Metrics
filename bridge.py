#!/usr/bin/env python
# coding: utf-8

# In[52]:


import pandas as pd
import numpy as np
import datetime as dt
import openpyxl
import re
import streamlit as st
import io # Added for BytesIO

# -------------------------------
# Helper Functions
# -------------------------------

def raw_file(input_df):
    """
    Process the raw input DataFrame and return the cleaned RAW DataFrame.
    """
    processed_df = input_df.copy()
    processed_df = processed_df.drop_duplicates()
    return processed_df

def act8_fil(input_df):
    """
    Process the ACT8 DataFrame and return the cleaned version.
    """
    processed_df = input_df.copy()
    processed_df = processed_df.drop_duplicates()
    return processed_df

# -------------------------------
# Action Item Functions (From Part 1)
# -------------------------------

def action7(RAW):
    """
    Process Action 7:
    - Remove certain districts
    - Filter by Year Built, Load Rating, Main Structure Type, Open/Posted status
    - Exclude rows with specific critical locations and comment patterns
    """
    ACT7 = RAW.copy()

    # Remove State Bridges by district
    districts = [f"State Bridges > District {i}" for i in range(1, 7)]
    ACT7 = ACT7[~ACT7["Parent Asset"].isin(districts)]

    # Remove old bridges (Year Built < 1994)
    ACT7 = ACT7[~((ACT7["NBI 027 Year Built"] <= 1994) |
                  (ACT7["NBI 027 Year Built"].isna() & (ACT7["B.W.01: Year Built"] <= 1994)))]

    # Keep only Operating Rating = 2 (or NaN)
    ACT7["NBI 063 Method Used Operating Rating"] = pd.to_numeric(
        ACT7["NBI 063 Method Used Operating Rating"], errors='coerce'
    )
    ACT7 = ACT7[(ACT7["NBI 063 Method Used Operating Rating"] == 2) |
                (ACT7["NBI 063 Method Used Operating Rating"].isna())]

    # Load Rating Method == ASR
    ACT7 = ACT7[ACT7["B.LR.04: Load Rating Method"] == "ASR"]

    # Filter out certain Main Structure Types or Span Materials
    ACT7["NBI 043 Main Structure Type"] = pd.to_numeric(
        ACT7["NBI 043 Main Structure Type"], errors='coerce'
    )
    MAIN_STRUC = [701, 702, 300, 400, 301, 401]
    SPAN_MAT = ["M01", "M02", "SX", "T01", "T02", "T03", "T04", "TX", "X"]
    ACT7 = ACT7[~(ACT7["NBI 043 Main Structure Type"].isin(MAIN_STRUC) |
                  ACT7["NBI 043 Main Structure Type"].isna() |
#                  ACT7["NBI 043 Main Structure Type"] == "") |
                  ACT7["B.SP.04: Span Material - Main"].isin(SPAN_MAT))]

    # Filter out closed or posted bridges
    ACT7 = ACT7[~((ACT7["NBI 041 Open, Posted Or Closed"] == "K") |
                  (ACT7["NBI 041 Open, Posted Or Closed"].isna() &
                   (ACT7["B.PS.01: Load Posting Status"] == "C")))]

    # Remove rows based on critical location keywords
    CRI_LOC = ["timber", "plank", "long", "trans", "pile", "piling", "standard", "std"]
    CRI_LOC_PAT = "|".join(CRI_LOC)
    ACT7 = ACT7[~ACT7["critical location"].str.contains(CRI_LOC_PAT, case=False, na=False)]

    # Remove rows based on comment keywords
    COM = ["30 ksi", "flatcar", "testing", "salvage", "standard"]
    COM_PAT = "|".join(COM)
    ACT7["Comments"] = ACT7["Comments"].astype(str)
    ACT7["Comment Inv Rating"] = ACT7["Comment Inv Rating"].astype(str)
    ACT7 = ACT7[~ACT7["Comments"].str.contains(COM_PAT, case=False, na=False)]
    ACT7 = ACT7[~ACT7["Comment Inv Rating"].str.contains(COM_PAT, case=False, na=False)]

    return ACT7


def action9(RAW):
    """
    Process Action 9:
    - Remove certain districts
    - Filter bridges based on Operating Rating, Load Rating Method, Design Load
    - Filter Open/Posted status and Year Built
    - Classify bridges as Standard or Non-Standard
    """
    ACT9 = RAW.copy()

    districts = [f"State Bridges > District {i}" for i in range(1, 7)]
    ACT9 = ACT9[~ACT9["Parent Asset"].isin(districts)]

    # Operating Rating == 1 or NaN
    ACT9["NBI 063 Method Used Operating Rating"] = pd.to_numeric(
        ACT9["NBI 063 Method Used Operating Rating"], errors='coerce'
    )
    ACT9 = ACT9[(ACT9["NBI 063 Method Used Operating Rating"] == 1) |
                (ACT9["NBI 063 Method Used Operating Rating"].isna())]

    # Load Rating Method == LFR
    ACT9 = ACT9[ACT9["B.LR.04: Load Rating Method"] == "LFR"]

    # Design Load
    ACT9 = ACT9[(ACT9["NBI 031 Design Load"] == "A") |
                (ACT9["NBI 031 Design Load"].isna() &
                 (ACT9["B.LR.01: Design Load"] == "HL93"))]

    # Filter Open/Posted status
    ACT9 = ACT9[~((ACT9["NBI 041 Open, Posted Or Closed"] == "K") |
                  (ACT9["NBI 041 Open, Posted Or Closed"].isna() &
                   (ACT9["B.PS.01: Load Posting Status"] == "C")))]

    # Filter Year Built
    ACT9 = ACT9[~((ACT9["NBI 027 Year Built"] < 2010) |
                  (ACT9["NBI 027 Year Built"].isna() &
                   (ACT9["B.W.01: Year Built"] < 2010)))]

    # Classify Standard vs Non-Standard
    COM = ["standard", "std", "STANDARD"]
    COM_PAT = "|".join(COM)
    ACT9["Comments"] = ACT9["Comments"].astype(str).fillna("")
    ACT9["Comment Inv Rating"] = ACT9["Comment Inv Rating"].astype(str).fillna("")

    ACT9_S = ACT9[(ACT9["Comments"].str.contains(COM_PAT, case=False, na=False)) |
                  (ACT9["Comment Inv Rating"].str.contains(COM_PAT, case=False, na=False))].copy()
    ACT9_S["Standard/Non-Standard"] = "Standard"

    ACT9_NS = ACT9[~((ACT9["Comments"].str.contains(COM_PAT, case=False, na=False)) |
                     (ACT9["Comment Inv Rating"].str.contains(COM_PAT, case=False, na=False)))].copy()
    ACT9_NS["Standard/Non-Standard"] = "Non-Standard"

    ACT9_F = pd.concat([ACT9_S, ACT9_NS], ignore_index=True)
    return ACT9_F


def action15(RAW):
    """
    Process Action 15:
    - Remove certain districts
    - Keep bridges with Load Rating Method == AR
    - Split and filter bridges built before 1972 and before 1992 based on structure type and span type
    - Filter based on Year Built and Year Reconstructed
    """
    ACT15 = RAW.copy()

    # Remove Parent Asset = State Bridges > District #
    districts = [f"State Bridges > District {i}" for i in range(1, 7)]
    ACT15 = ACT15[~ACT15["Parent Asset"].isin(districts)]

    # Keep Load Rating Method == AR
    ACT15 = ACT15[ACT15["B.LR.04: Load Rating Method"] == "AR"]

    # Before 1972 subset
    SPAN_MAT = ["F01", "F02", "F03", "F04", "P01", "P02"]
    ACT15_72 = ACT15[
        ACT15["NBI 043 Main Structure Type"].astype(str).str.endswith("19") |
        (ACT15["NBI 043 Main Structure Type"].isna() & ACT15["B.SP.06: Span Type - Main"].isin(SPAN_MAT))
    ].copy()

    ACT15_72 = ACT15_72[
        ((ACT15_72["NBI 027 Year Built"] <= 1972) |
         (ACT15_72["NBI 027 Year Built"].isna() & (ACT15_72["B.W.01: Year Built"] < 1972))) &
        ((ACT15_72["NBI 106 Year Reconst"].isna()) | (ACT15_72["NBI 106 Year Reconst"] <= 1972))
    ]

    # Before 1992 subset
    ACT15_92 = ACT15[
        (~ACT15["NBI 043 Main Structure Type"].astype(str).str.endswith("19")) |
        (ACT15["NBI 043 Main Structure Type"].isna() & ~ACT15["B.SP.06: Span Type - Main"].isin(SPAN_MAT))
    ].copy()

    ACT15_92 = ACT15_92[
        ((ACT15_92["NBI 027 Year Built"] <= 1992) |
         (ACT15_92["NBI 027 Year Built"].isna() & (ACT15_92["B.W.01: Year Built"] < 1992))) &
        ((ACT15_92["NBI 106 Year Reconst"].isna()) | (ACT15_92["NBI 106 Year Reconst"] <= 1992))
    ]

    # Combine subsets
    ACT15_F = pd.concat([ACT15_72, ACT15_92], ignore_index=True)

    return ACT15_F


def action16(RAW):
    """
    Process Action 16:
    - Remove certain districts
    - Keep bridges with Load Rating Method == AR
    - Split and filter bridges built after 1972 and after 1992 based on structure type and span type
    - Remove bridges with standard/design-related comments
    """
    ACT16 = RAW.copy()

    # Remove Parent Asset = State Bridges > District #
    districts = [f"State Bridges > District {i}" for i in range(1, 7)]
    ACT16 = ACT16[~ACT16["Parent Asset"].isin(districts)]

    # Keep Load Rating Method == AR
    ACT16 = ACT16[ACT16["B.LR.04: Load Rating Method"] == "AR"]

    # After 1972 subset
    SPAN_MAT = ["F01", "F02", "F03", "F04", "P01", "P02"]
    ACT16_72 = ACT16[
        ACT16["NBI 043 Main Structure Type"].astype(str).str.endswith("19") |
        (ACT16["NBI 043 Main Structure Type"].isna() & ACT16["B.SP.06: Span Type - Main"].isin(SPAN_MAT))
    ].copy()

    ACT16_72 = ACT16_72[
        ((ACT16_72["NBI 027 Year Built"] > 1972) |
         (ACT16_72["NBI 027 Year Built"].isna() & (ACT16_72["B.W.01: Year Built"] > 1972))) &
        ((ACT16_72["NBI 106 Year Reconst"].isna()) |
         (ACT16_72["NBI 106 Year Reconst"] == 0) |
         (ACT16_72["NBI 106 Year Reconst"] > 1972))
    ]

    # After 1992 subset
    ACT16_92 = ACT16[
        (~ACT16["NBI 043 Main Structure Type"].astype(str).str.endswith("19")) |
        (ACT16["NBI 043 Main Structure Type"].isna() & ~ACT16["B.SP.06: Span Type - Main"].isin(SPAN_MAT))
    ].copy()

    ACT16_92 = ACT16_92[
        ((ACT16_92["NBI 027 Year Built"] > 1992) |
         (ACT16_92["NBI 027 Year Built"].isna() & (ACT16_92["B.W.01: Year Built"] > 1992))) &
        ((ACT16_92["NBI 106 Year Reconst"].isna()) |
         (ACT16_92["NBI 106 Year Reconst"] == 0) |
         (ACT16_92["NBI 106 Year Reconst"] > 1992))
    ]

    # Combine subsets
    ACT16_F = pd.concat([ACT16_72, ACT16_92], ignore_index=True)

    # Remove bridges with standard/design-related comments
    COM = [
        "standard", "std", "design load per certified", "based on field measurements", "HL-93",
        "exterior wall reinforcing is inadequate", "bridge plan was HS20", "shop drawing not available",
        "exterior wall under reinforced", "shop drawings not available", "bridge plans was HS20",
        "Per field measurements", "assignment", "design load", "HS20 design", "HS-20 live",
        "HS 20 design", "high fill depth", "Unable to Provide"
    ]
    COM_PAT = "|".join(COM)
    ACT16_F["Comments"] = ACT16_F["Comments"].astype(str)
    ACT16_F["Comment Inv Rating"] = ACT16_F["Comment Inv Rating"].astype(str)

    ACT16_F = ACT16_F[~(
        ACT16_F["Comments"].str.contains(COM_PAT, case=False, na=False) |
        ACT16_F["Comment Inv Rating"].str.contains(COM_PAT, case=False, na=False)
    )]

    return ACT16_F


def action17(RAW):
    """
    Process Action 17:
    - Remove County and City Bridges
    - Keep bridges with Load Rating Method == AR
    - Split and filter bridges built before 1972 and before 1992 based on structure type and span type
    """
    ACT17 = RAW.copy()

    # Remove County and City Bridges
    bridge = ["County Bridges", "City Bridges"]
    bridge_pat = "|".join(bridge)
    ACT17 = ACT17[~ACT17["Parent Asset"].str.contains(bridge_pat, case=False, na=False)]

    # Keep Load Rating Method == AR
    ACT17 = ACT17[ACT17["B.LR.04: Load Rating Method"] == "AR"]

    # Ensure numeric fields
    ACT17["NBI 027 Year Built"] = pd.to_numeric(ACT17["NBI 027 Year Built"], errors='coerce')
    ACT17["NBI 106 Year Reconst"] = pd.to_numeric(ACT17["NBI 106 Year Reconst"], errors='coerce')

    # Before 1972 subset
    SPAN_MAT = ["F01", "F02", "F03", "F04", "P01", "P02"]
    ACT17_72 = ACT17[
        ACT17["NBI 043 Main Structure Type"].astype(str).str.endswith("19") |
        (ACT17["NBI 043 Main Structure Type"].isna() & ACT17["B.SP.06: Span Type - Main"].isin(SPAN_MAT))
    ].copy()

    ACT17_72 = ACT17_72[
        ((ACT17_72["NBI 027 Year Built"] <= 1972) |
         (ACT17_72["NBI 027 Year Built"].isna() & (ACT17_72["B.W.01: Year Built"] < 1972))) &
        (ACT17_72["NBI 106 Year Reconst"].isna() |
         (ACT17_72["NBI 106 Year Reconst"] <= 1972) |
         (ACT17_72["NBI 106 Year Reconst"] == 0))
    ]

    # Before 1992 subset
    ACT17_92 = ACT17[
        (~ACT17["NBI 043 Main Structure Type"].astype(str).str.endswith("19")) |
        (ACT17["NBI 043 Main Structure Type"].isna() & ~ACT17["B.SP.06: Span Type - Main"].isin(SPAN_MAT))
    ].copy()

    ACT17_92 = ACT17_92[
        ((ACT17_92["NBI 027 Year Built"] <= 1992) |
         (ACT17_92["NBI 027 Year Built"].isna() & (ACT17_92["B.W.01: Year Built"] < 1992))) &
        (ACT17_92["NBI 106 Year Reconst"].isna() |
         (ACT17_92["NBI 106 Year Reconst"] <= 1992) |
         (ACT17_92["NBI 106 Year Reconst"] == 0))
    ]

    # Combine subsets
    ACT17_F = pd.concat([ACT17_72, ACT17_92], ignore_index=True)

    return ACT17_F


def action18(RAW):
    """
    Process Action 18:
    - Remove County and City Bridges
    - Keep bridges with Load Rating Method == AR
    - Split and filter bridges built after 1972 and after 1992 based on structure type and span type
    - Remove bridges with specific comments
    """
    ACT18 = RAW.copy()

    # Remove County and City Bridges
    bridge = ["County Bridges", "City Bridges"]
    bridge_pat = "|".join(bridge)
    ACT18 = ACT18[~ACT18["Parent Asset"].str.contains(bridge_pat, case=False, na=False)]

    # Keep Load Rating Method == AR
    ACT18 = ACT18[ACT18["B.LR.04: Load Rating Method"] == "AR"]

    # After 1972 subset
    SPAN_MAT = ["F01", "F02", "F03", "F04", "P01", "P02"]
    ACT18_72 = ACT18[
        ACT18["NBI 043 Main Structure Type"].astype(str).str.endswith("19") |
        (ACT18["NBI 043 Main Structure Type"].isna() & ACT18["B.SP.06: Span Type - Main"].isin(SPAN_MAT))
    ].copy()

    ACT18_72 = ACT18_72[
        ((ACT18_72["NBI 027 Year Built"] > 1972) |
         (ACT18_72["NBI 027 Year Built"].isna() & (ACT18_72["B.W.01: Year Built"] > 1972))) &
        (ACT18_72["NBI 106 Year Reconst"].isna() |
         (ACT18_72["NBI 106 Year Reconst"] > 1972) |
         (ACT18_72["NBI 106 Year Reconst"] == 0))
    ]

    # After 1992 subset
    ACT18_92 = ACT18[
        (~ACT18["NBI 043 Main Structure Type"].astype(str).str.endswith("19")) |
        (ACT18["NBI 043 Main Structure Type"].isna() &
         ~ACT18["B.SP.06: Span Type - Main"].isin(SPAN_MAT))
    ].copy()

    ACT18_92 = ACT18_92[
        ((ACT18_92["NBI 027 Year Built"] > 1992) |
         (ACT18_92["NBI 027 Year Built"].isna() & (ACT18_92["B.W.01: Year Built"] > 1992))) &
        (ACT18_92["NBI 106 Year Reconst"].isna() |
         (ACT18_92["NBI 106 Year Reconst"] > 1992) |
         (ACT18_92["NBI 106 Year Reconst"] == 0))
    ]

    # Combine subsets
    ACT18_F = pd.concat([ACT18_72, ACT18_92], ignore_index=True)

    # Remove bridges with specific comments
    COM = ["standard", "std", "parametric", "LFR", "NBI 64", "NBI 66"]
    COM_PAT = "|".join(COM)
    ACT18_F["Comments"] = ACT18_F["Comments"].astype(str).fillna("")
    ACT18_F["Comment Inv Rating"] = ACT18_F["Comment Inv Rating"].astype(str).fillna("")

    ACT18_F = ACT18_F[~(
        ACT18_F["Comments"].str.contains(COM_PAT, case=False, na=False) |
        ACT18_F["Comment Inv Rating"].str.contains(COM_PAT, case=False, na=False)
    )]

    return ACT18_F


def action19(RAW):
    """
    Process Action 19:
    - Remove State and Border Bridges
    - Keep Load Rating Method == EJ
    - Remove bridges with specific structure types, span types, or critical locations
    - Categorize remaining bridges into sub-categories
    """
    ACT19 = RAW.copy()

    # Remove State and Border Bridges
    bridge = ["State Bridges", "Border Bridges"]
    bridge_pat = "|".join(bridge)
    ACT19 = ACT19[~ACT19["Parent Asset"].str.contains(bridge_pat, case=False, na=False)]

    # Keep Load Rating Method == EJ
    ACT19 = ACT19[ACT19["B.LR.04: Load Rating Method"] == "EJ"]

    # Keep only open bridges
    ACT19 = ACT19[ACT19["NBI 041 Open, Posted Or Closed"] != "K"]

    # Convert numeric columns
    def try_numeric(val):
        try:
            f = float(val)
            return int(f) if f.is_integer() else f
        except:
            return val

    ACT19["NBI 043 Main Structure Type"] = ACT19["NBI 043 Main Structure Type"].apply(try_numeric)
    ACT19["NBI 043 Main Structure Type Str"] = ACT19["NBI 043 Main Structure Type"].apply(lambda x: str(x).strip())

    # Removal conditions
    cond_319 = ACT19["NBI 043 Main Structure Type"] == 319
    cond_startswith = ACT19["NBI 043 Main Structure Type Str"].str.startswith(("1", "2", "5", "6"))
    cond_nan_nbi043_span_ty = ACT19["NBI 043 Main Structure Type"].isna() & ACT19["B.SP.06: Span Type - Main"].isin(["P01", "P02"])
    cond_nan_span_ty_span_mat = ACT19["B.SP.06: Span Type - Main"].isna() & ACT19["B.SP.04: Span Material - Main"].isin(["C01","C02","C03","C04","C05","CX"])
    cond_critloc = (ACT19["critical location"].str.contains("timber|plank|pile", case=False, na=False) |
                    ACT19["critical location.1"].str.contains("timber|plank|pile", case=False, na=False))

    # Apply removal mask
    remove_mask = cond_319 | cond_startswith | cond_nan_nbi043_span_ty | cond_nan_span_ty_span_mat | cond_critloc
    ACT19 = ACT19[~remove_mask].reset_index(drop=True)
    ACT19 = ACT19.drop(columns=["NBI 043 Main Structure Type Str"])

    # Sub-categories
    ACT19_2 = ACT19[ACT19["Comment Inv Rating"].astype(str).str.contains("std|standard", case=False, na=False)].copy()
    ACT19_2["Action 19 Sub-Category"] = "Standard Bridge"

    ACT19_3 = ACT19.merge(ACT19_2, how="left", indicator=True)
    ACT19_3 = ACT19_3[ACT19_3["_merge"] == "left_only"].drop(columns=["_merge"])
    ACT19_3 = ACT19_3[ACT19_3["Comment Inv Rating"].astype(str).str.contains("test", case=False, na=False)]
    ACT19_3["Action 19 Sub-Category"] = "Bridge was load tested."

    ACT19_1 = ACT19.merge(pd.concat([ACT19_2, ACT19_3]).drop_duplicates(), how="left", indicator=True)
    ACT19_1 = ACT19_1[ACT19_1["_merge"] == "left_only"].drop(columns=["_merge"])
    ACT19_1 = ACT19_1[ACT19_1["Comment Inv Rating"].astype(str).str.contains("poor|deteriorat|post|decay|damage|clos", case=False, na=False)]
    ACT19_1["Action 19 Sub-Category"] = "Severe Deterioration"

    ACT19_4 = ACT19.merge(pd.concat([ACT19_1, ACT19_2, ACT19_3]).drop_duplicates(), how="left", indicator=True)
    ACT19_4 = ACT19_4[ACT19_4["_merge"] == "left_only"].drop(columns=["_merge"])
    ACT19_4["Action 19 Sub-Category"] = "Not Permitted"

    # Final combined dataset
    ACT19_F = pd.concat([ACT19_1, ACT19_2, ACT19_3, ACT19_4], ignore_index=True)
    ACT19_F["Comments"] = ACT19_F["Comments"].fillna("")
    ACT19_F["Comment Inv Rating"] = ACT19_F["Comment Inv Rating"].fillna("")

    return ACT19_F


def action20(RAW):
    """
    Process Action 20:
    - Remove State and Border Bridges
    - Keep Load Rating Method == EJ
    - Filter based on open/posted status, structure type, span material, critical location, and comments
    """
    ACT20 = RAW.copy()

    # Remove State and Border Bridges
    bridge = ["State Bridges", "Border Bridges"]
    bridge_pat = "|".join(bridge)
    ACT20 = ACT20[~ACT20["Parent Asset"].str.contains(bridge_pat, case=False, na=False)]

    # Keep Load Rating Method == EJ
    ACT20 = ACT20[ACT20["B.LR.04: Load Rating Method"] == "EJ"]

    # Remove closed or posted bridges
    ACT20 = ACT20[~((ACT20["NBI 041 Open, Posted Or Closed"] == "K") |
                    (ACT20["NBI 041 Open, Posted Or Closed"].isna() & (ACT20["B.PS.01: Load Posting Status"] == "C")))]

    # Keep concrete bridges
    ACT20["NBI 043 Main Structure Type"] = pd.to_numeric(ACT20["NBI 043 Main Structure Type"], errors='coerce')
    concrete = [101,102,104,105,106,119,121,122,100,201,202,204,205,206,219,221,222,200]
    SPAN_MAT = ["C01","C02","C03","C04","C05"]
    ACT20 = ACT20[(ACT20["NBI 043 Main Structure Type"].isin(concrete)) |
                  (ACT20["NBI 043 Main Structure Type"].isna() & ACT20["B.SP.04: Span Material - Main"].isin(SPAN_MAT))]

    # Remove bridges with critical locations
    CRI_LOC = ["timber","plank","long","trans","pil"]
    CRI_LOC_PAT = "|".join(CRI_LOC)
    ACT20 = ACT20[~(ACT20["critical location"].str.contains(CRI_LOC_PAT, case=False, na=False) |
                    ACT20["critical location.1"].str.contains(CRI_LOC_PAT, case=False, na=False))]

    # Filter based on comments
    COM = ["based on a parametric","based on the parametric","no signs of distress","sufficient",
           "available","no plans","unable to provide","agreed with FHWA","software",
           "deterioration","MBE 6A.5.11","standard","std"]
    COM_PAT = "|".join(COM)
    ACT20 = ACT20[(ACT20["Comments"].isna() | (ACT20["Comments"] == "")) &
                  ~(ACT20["Comment Inv Rating"].str.contains(COM_PAT, case=False, na=False))]

    return ACT20


# Using the second, more detailed definition of Action 21
def action21(RAW: pd.DataFrame) -> pd.DataFrame:
    """
    Process Action 21:
    - Remove County and City Bridges
    - Keep Load Rating Method == EJ
    - Filter based on open/posted status, structure type, span material, critical location, and comments
    """
    ACT21 = RAW.copy()

    # Remove County and City Bridges
    bridges = ["County Bridges", "City Bridges"]
    bridge_pat = "|".join(bridges)
    ACT21 = ACT21[~ACT21["Parent Asset"].str.contains(bridge_pat, case=False, na=False)]

    # Keep Load Rating Method == EJ
    ACT21 = ACT21[ACT21["B.LR.04: Load Rating Method"] == "EJ"]

    # Remove closed or posted bridges
    ACT21 = ACT21[~((ACT21["NBI 041 Open, Posted Or Closed"] == "K") |
                    (ACT21["NBI 041 Open, Posted Or Closed"].isna() &
                     (ACT21["B.PS.01: Load Posting Status"] == "C")))]

    # Keep concrete bridges
    ACT21["NBI 043 Main Structure Type"] = pd.to_numeric(ACT21["NBI 043 Main Structure Type"], errors='coerce')
    concrete_types = [101, 102, 104, 105, 106, 119, 121, 122, 100,
                      211, 212, 214, 215, 216, 219, 221, 222, 210]
    span_materials = ["C01", "C02", "C03", "C04", "C05", "CX"]
    ACT21 = ACT21[(ACT21["NBI 043 Main Structure Type"].isin(concrete_types)) |
                  (ACT21["NBI 043 Main Structure Type"].isna() &
                   ACT21["B.SP.04: Span Material - Main"].isin(span_materials))]

    # Remove bridges with critical locations
    critical_locs = ["timber", "plank", "long", "trans", "pil"]
    critical_pat = "|".join(critical_locs)
    ACT21 = ACT21[~(ACT21["critical location"].str.contains(critical_pat, case=False, na=False) |
                    ACT21["critical location.1"].str.contains(critical_pat, case=False, na=False))]

    # Filter based on comments
    comments = ["parametric", "illegible", "missing", "no plans", "per section 6.1.4",
                "The following bridge has been inspected", "standard", "std"]
    comment_pat = "|".join(comments)
    ACT21["Comments"] = ACT21["Comments"].fillna("")
    ACT21["Comment Inv Rating"] = ACT21["Comment Inv Rating"].fillna("")
    ACT21 = ACT21[~(ACT21["Comments"].str.contains(comment_pat, case=False, na=False) |
                    ACT21["Comment Inv Rating"].str.contains(comment_pat, case=False, na=False))]

    return ACT21


def action22(RAW: pd.DataFrame) -> pd.DataFrame:
    """
    Process Action 22:
    - Remove State and Border Bridges
    - Keep Load Rating Method == EJ
    - Filter by open/posted status, structure type, span type, and comments
    """
    ACT22 = RAW.copy()

    # Remove State and Border Bridges
    bridges = ["State Bridges", "Border Bridges"]
    bridge_pat = "|".join(bridges)
    ACT22 = ACT22[~ACT22["Parent Asset"].str.contains(bridge_pat, case=False, na=False)]

    # Keep Load Rating Method == EJ
    ACT22 = ACT22[ACT22["B.LR.04: Load Rating Method"] == "EJ"]

    # Remove closed or posted bridges
    ACT22 = ACT22[~((ACT22["NBI 041 Open, Posted Or Closed"] == "K") |
                    (ACT22["NBI 041 Open, Posted Or Closed"].isna() &
                     (ACT22["B.PS.01: Load Posting Status"] == "C")))]

    # Convert NBI 043 Main Structure Type to numeric
    ACT22["NBI 043 Main Structure Type"] = pd.to_numeric(ACT22["NBI 043 Main Structure Type"], errors='coerce')

    # Keep NBI 043 = 319 or if NaN, Span Type = P02
    ACT22_1 = ACT22[(ACT22["NBI 043 Main Structure Type"] == 319) |
                    (ACT22["NBI 043 Main Structure Type"].isna() &
                     (ACT22["B.SP.06: Span Type - Main"] == "P02"))].copy()

    # Remaining bridges with culvert-related comments
    ACT22_1c = ACT22_1.reset_index(drop=True)
    ACT22_2 = ACT22.merge(ACT22_1c, how="left", indicator=True)
    ACT22_2 = ACT22_2[ACT22_2["_merge"] == "left_only"].drop(columns=["_merge"])

    culvert_comments = ["CMP", "corrugated", "metal culvert"]
    culvert_pat = "|".join(culvert_comments)
    ACT22_2["Comments"] = ACT22_2["Comments"].fillna("")
    ACT22_2["Comment Inv Rating"] = ACT22_2["Comment Inv Rating"].fillna("")
    ACT22_2 = ACT22_2[(ACT22_2["Comments"].str.contains(culvert_pat, case=False, na=False) |
                       ACT22_2["Comment Inv Rating"].str.contains(culvert_pat, case=False, na=False))]

    # Combine the two filtered sets
    ACT22_F = pd.concat([ACT22_1, ACT22_2], ignore_index=True)

    return ACT22_F

# -------------------------------
# RAW2, RAW3, and Action Functions (From Part 2)
# -------------------------------

def make_RAW2(
    RAW,
    ACT7, ACT8, ACT9_F, ACT15_F, ACT16_F, ACT17_F, ACT18_F, ACT19_F,
    ACT20, ACT21, ACT22_F
):
    """
    Build RAW2 from RAW and Action 7–22 datasets.
    Includes:
        - Concatenation of all action outputs
        - Deduplication
        - Normalization of first 42 columns
        - Manual overrides
        - Anti-join RAW - ACT7_22
    """

    #Concatenate all action results.
    ACTS = [
        ACT7, ACT8, ACT9_F, ACT15_F, ACT16_F, ACT17_F, ACT18_F,
        ACT19_F, ACT20, ACT21, ACT22_F
    ]

    ACT7_22 = pd.concat(ACTS, ignore_index=True)

    #Drop duplicates on first 42 columns
    first42 = RAW.columns[:42].tolist()
    ACT7_22 = ACT7_22.drop_duplicates(subset=first42).reset_index(drop=True)

    #Manual override
    if "Bridge ID" in ACT7_22.columns:
        ACT7_22.loc[
            ACT7_22["Bridge ID"] == "180TH ST.",
            "NBI 063 Method Used Operating Rating"
        ] = "F"

    #Drop duplicates in RAW.
    RAW = RAW.drop_duplicates(subset=first42)

    #Normalization function.
    def normalize_mixed(val):
        if pd.isna(val):
            return ""
        try:
            f = float(val)
            return str(int(f)) if f.is_integer() else str(f)
        except:
            return str(val).strip().lower()

    #Apply normalization to RAW and ACT7_22.
    for col in first42:
        if col in RAW.columns:
            RAW[col] = RAW[col].apply(normalize_mixed)
        if col in ACT7_22.columns:
            ACT7_22[col] = ACT7_22[col].apply(normalize_mixed)

        #Replace "nan" and fill blanks.
        if col in RAW.columns:
            RAW[col] = RAW[col].replace("nan", "").fillna("")
            RAW[col] = RAW[col].astype(str).str.strip()
        if col in ACT7_22.columns:
            ACT7_22[col] = ACT7_22[col].replace("nan", "").fillna("")
            ACT7_22[col] = ACT7_22[col].astype(str).str.strip()

    #Drop duplicates again after normalization.
    RAW = RAW.drop_duplicates(subset=first42)
    ACT7_22 = ACT7_22.drop_duplicates(subset=first42)

    #Anti-join RAW - ACT7_22 → RAW2.
    RAW2 = RAW.merge(
        ACT7_22[first42], on=first42, how="left", indicator=True
    )
    RAW2 = RAW2[RAW2["_merge"] == "left_only"] \
           .drop(columns=["_merge"]) \
           .reset_index(drop=True)

    #Final
    RAW2 = RAW2.drop_duplicates(subset=first42)

    return RAW2, ACT7_22 # Return ACT7_22 for counts


def action2(RAW2):
    """
    Action 2:
    - Remove State and Border Bridges
    - Keep bridges with only SU7 traffic (multi-lane or one-lane)
    - Remove based on Load Rating Method rules
    - Remove based on Open/Posted/Closed status
    - Remove rows with comments indicating standards or parametrics
    - Drop unneeded columns
    """
    ACT2 = RAW2.copy()

    #Remove State & Border Bridges

    bridge = ["State Bridges", "Border Bridges"]
    bridge_pat = "|".join(bridge)
    ACT2 = ACT2[~ACT2["Parent Asset"].str.contains(bridge_pat, case=False, na=False)]

    #Traffic Tons Columns

    tons = [
        "Multi Lane Traffic: Type SU4 Tons", "Multi Lane Traffic: Type SU5 Tons", 
        "Multi Lane Traffic: Type SU6 Tons", "Multi Lane Traffic: Type SU7 Tons",    
        "One Lane Traffic: Type SU4 Tons", "One Lane Traffic: Type SU5 Tons", 
        "One Lane Traffic: Type SU6 Tons", "One Lane Traffic: Type SU7 Tons"
    ]

    #Convert to numeric, fill NaN with 0.
    for col in tons:
        if col in ACT2.columns:
            ACT2[col] = pd.to_numeric(ACT2[col], errors="coerce").fillna(0)
        else:
            ACT2[col] = 0 # Add column if missing

    #Keep only where SU7 > 0 and all others = 0.
    ACT2 = ACT2[
        ((ACT2["Multi Lane Traffic: Type SU4 Tons"] == 0) &
         (ACT2["Multi Lane Traffic: Type SU5 Tons"] == 0) &
         (ACT2["Multi Lane Traffic: Type SU6 Tons"] == 0) &
         (ACT2["Multi Lane Traffic: Type SU7 Tons"] > 0)) |
        ((ACT2["One Lane Traffic: Type SU4 Tons"] == 0) &
         (ACT2["One Lane Traffic: Type SU5 Tons"] == 0) &
         (ACT2["One Lane Traffic: Type SU6 Tons"] == 0) &
         (ACT2["One Lane Traffic: Type SU7 Tons"] > 0))
    ]

    #Load Rating Method Rules
    ACT2["NBI 063 Method Used Operating Rating"] = pd.to_numeric(
        ACT2["NBI 063 Method Used Operating Rating"], errors="coerce"
    )

    ACT2["B.LR.04: Load Rating Method"] = (
        ACT2["B.LR.04: Load Rating Method"].astype(str).str.strip().str.upper()
    )

    ACT2 = ACT2[~(
        (ACT2["B.LR.04: Load Rating Method"] == "EJ") |
        (ACT2["B.LR.04: Load Rating Method"].isna() &
         (ACT2["NBI 063 Method Used Operating Rating"] == 0))
    )]

    #Open/Posted/Closed rules
    OPC = ["P", "R", "K"]
    LPS = ["C", "PP", "PR"]

    ACT2["NBI 041 Open, Posted Or Closed"] = (
        ACT2["NBI 041 Open, Posted Or Closed"].astype(str).str.strip().str.upper()
    )
    ACT2["B.PS.01: Load Posting Status"] = (
        ACT2["B.PS.01: Load Posting Status"].astype(str).str.strip().str.upper()
    )

    ACT2 = ACT2[~(
        ACT2["NBI 041 Open, Posted Or Closed"].isin(OPC) |
        (ACT2["NBI 041 Open, Posted Or Closed"].isna() &
         ACT2["B.PS.01: Load Posting Status"].isin(LPS))
    )]

    #Remove based on comments
    COM = [
        "SU4", "close bridge", "closed",
        "based on a parametric", "based on the parametric",
        "J7", "J24", "Standards"
    ]

    COM_PAT = "|".join([re.escape(w) for w in COM])
    STAN_PAT = r"(?<!non[-\s])standard|(?<!non[-\s])std"
    FINAL_PAT = f"{COM_PAT}|{STAN_PAT}"

    ACT2["Comments"] = ACT2["Comments"].astype(str)
    ACT2["Comment Inv Rating"] = ACT2["Comment Inv Rating"].astype(str)

    ACT2 = ACT2[
        ((ACT2["Comments"].isna() | (ACT2["Comments"].str.strip() == "")) &
         (ACT2["Comment Inv Rating"].isna() | (ACT2["Comment Inv Rating"].str.strip() == ""))) |
        (~(
            ACT2["Comments"].str.contains(FINAL_PAT, case=False, na=False, regex=True) |
            ACT2["Comment Inv Rating"].str.contains(FINAL_PAT, case=False, na=False, regex=True)
        ))
    ]

    #Drop unnecessary columns.
    ACT2 = ACT2.drop(
        columns=["Standard/Non-Standard", "Action 19 Sub-Category"],
        errors='ignore'
    )

    return ACT2


def action3(RAW2):

    ACT3 = RAW2.copy()

    #Exclude State and Border Bridges
    bridge = ["State Bridge", "Border Bridge"]
    bridge_pat = "|".join(bridge)
    ACT3 = ACT3[~(ACT3["Parent Asset"].str.contains(bridge_pat, case=False, na=False))]

    #Convert traffic columns to numeric.
    tons = [
        "Multi Lane Traffic: Type SU4 Tons", "Multi Lane Traffic: Type SU5 Tons",
        "Multi Lane Traffic: Type SU6 Tons", "Multi Lane Traffic: Type SU7 Tons",
        "One Lane Traffic: Type SU4 Tons", "One Lane Traffic: Type SU5 Tons",
        "One Lane Traffic: Type SU6 Tons", "One Lane Traffic: Type SU7 Tons"
    ]
    for col in tons:
        if col in ACT3.columns:
            ACT3[col] = pd.to_numeric(ACT3[col], errors="coerce").fillna(0)
        else:
            ACT3[col] = 0

    #Keep rows where ALL traffic = 0
    ACT3 = ACT3[
        (ACT3["Multi Lane Traffic: Type SU4 Tons"] == 0) &
        (ACT3["Multi Lane Traffic: Type SU5 Tons"] == 0) &
        (ACT3["Multi Lane Traffic: Type SU6 Tons"] == 0) &
        (ACT3["Multi Lane Traffic: Type SU7 Tons"] == 0) &
        (ACT3["One Lane Traffic: Type SU4 Tons"] == 0) &
        (ACT3["One Lane Traffic: Type SU5 Tons"] == 0) &
        (ACT3["One Lane Traffic: Type SU6 Tons"] == 0) &
        (ACT3["One Lane Traffic: Type SU7 Tons"] == 0)
    ]

    #Exclude Load Rating Methods EJ / AR
    LRM = ["EJ", "AR"]
    ACT3["B.LR.04: Load Rating Method"] = (
        ACT3["B.LR.04: Load Rating Method"].astype(str).str.strip().str.upper()
    )
    ACT3 = ACT3[~ACT3["B.LR.04: Load Rating Method"].isin(LRM)]

    #Remove posted / restricted / closed bridges
    OPC = ["K", "P", "D"]
    LPS = ["C", "PP", "PR", "TP", "TR"]

    ACT3["NBI 041 Open, Posted Or Closed"] = (
        ACT3["NBI 041 Open, Posted Or Closed"].astype(str).str.strip().str.upper()
    )
    ACT3["B.PS.01: Load Posting Status"] = (
        ACT3["B.PS.01: Load Posting Status"].astype(str).str.strip().str.upper()
    )

    ACT3 = ACT3[~(
        ACT3["NBI 041 Open, Posted Or Closed"].isin(OPC) |
        ((ACT3["NBI 041 Open, Posted Or Closed"] == "") &
         ACT3["B.PS.01: Load Posting Status"].isin(LPS))
    )]

    #Remove parametric / standard comments.
    COM = ["SU4", "close bridge", "closed",
           "based on a parametric", "based on the parametric",
           "J7", "J24", "Standards"]

    COM_PAT = "|".join([re.escape(word) for word in COM])
    STAN_PAT = r"(?<!non[-\s])standard|(?<!non[-\s])std"
    FINAL_PAT = f"{COM_PAT}|{STAN_PAT}"

    ACT3["Comments"] = ACT3["Comments"].astype(str)
    ACT3["Comment Inv Rating"] = ACT3["Comment Inv Rating"].astype(str)

    ACT3 = ACT3[
        ((ACT3["Comments"].isna()) | (ACT3["Comments"].str.strip() == "")) &
        ((ACT3["Comment Inv Rating"].isna()) | (ACT3["Comment Inv Rating"].str.strip() == "")) |
        (~(
            ACT3["Comments"].str.contains(FINAL_PAT, case=False, na=False, regex=True) |
            ACT3["Comment Inv Rating"].str.contains(FINAL_PAT, case=False, na=False, regex=True)
        ))
    ]

    #Convert relevant columns to numeric.
    ACT3["NBI 063 Method Used Operating Rating"] = (
        ACT3["NBI 063 Method Used Operating Rating"].astype(str).str.upper().str.strip()
    )
    ACT3["NBI 064 Operating Rating"] = pd.to_numeric(
        ACT3["NBI 064 Operating Rating"], errors='coerce'
    )
    ACT3["B.LR.06: Operating Load Rating Factor"] = pd.to_numeric(
        ACT3["B.LR.06: Operating Load Rating Factor"], errors='coerce'
    )

    #Keep rows with valid Year Built
    ACT3 = ACT3[
        ACT3["B.W.01: Year Built"].notna() &
        (ACT3["B.W.01: Year Built"].astype(str).str.strip() != "")
    ]

    #Remove G1 low OR bridges.
    G1 = ["1", "2", "3", "4", "5", "A", "C"]
    ACT3 = ACT3[~(
        (ACT3["NBI 063 Method Used Operating Rating"].isin(G1)) &
        (ACT3["NBI 064 Operating Rating"] < 45)
    )]

    #Remove G2 low OR bridges.
    G2 = ["6", "7", "8", "F", "D"]
    ACT3 = ACT3[~(
        (ACT3["NBI 063 Method Used Operating Rating"].isin(G2)) &
        ((ACT3["NBI 064 Operating Rating"] < 1.26) |
         (ACT3["NBI 064 Operating Rating"].isna() &
          (ACT3["B.LR.06: Operating Load Rating Factor"] < 1.26)))
    )]

    #Drop unneeded columns.
    ACT3 = ACT3.drop(["Standard/Non-Standard", "Action 19 Sub-Category"],
                     axis=1, errors='ignore')

    return ACT3


def run_action8m_and_raw3(RAW_in, ACT8_in):

    RAW = RAW_in.copy()
    ACT8M = ACT8_in.copy()

    first42_cols = RAW.columns[:42].tolist()

    ACT8M = ACT8M.drop_duplicates(subset=first42_cols).reset_index(drop=True)

    #Fix specific override for Bridge ID
    if "Bridge ID" in ACT8M.columns:
        ACT8M["NBI 063 Method Used Operating Rating"] = ACT8M[
            "NBI 063 Method Used Operating Rating"
        ].astype(str)

        ACT8M.loc[
            ACT8M["Bridge ID"] == "180TH ST.",
            "NBI 063 Method Used Operating Rating"
        ] = "F"

    def normalize_mixed(val):
        if pd.isna(val) or str(val).strip().lower() in ["", "nan", "none"]:
            return "0"
        try:
            f = float(val)
            if f == 0:
                return "0"
            elif f.is_integer():
                return str(int(f))
            else:
                return str(f)
        except:
            return str(val).strip()

    #Normalize RAW and ACT8M for first 42 columns
    for col in first42_cols:
        if col in RAW.columns:
            RAW[col] = RAW[col].apply(normalize_mixed)
        if col in ACT8M.columns:
            ACT8M[col] = ACT8M[col].apply(normalize_mixed)

    #Merge and keep rows in RAW that are NOT in ACT8M (RAW3)
    RAW3 = RAW.merge(
        ACT8M[first42_cols],
        on=first42_cols,
        how="left",
        indicator=True
    )

    RAW3 = (
        RAW3[RAW3["_merge"] == "left_only"]
        .drop(columns=["_merge"])
        .reset_index(drop=True)
    )

    return ACT8M, RAW3


def action5(RAW3):
    ACT5 = RAW3.copy()

    #Remove County and City Bridges
    bridge = ["County Bridges", "City Bridges"]
    bridge_pat = "|".join(bridge)
    ACT5 = ACT5[~(ACT5["Parent Asset"].str.contains(bridge_pat, case=False, na=False))]

    #Traffic tonnage columns
    tons = [
        "Multi Lane Traffic: Type SU4 Tons", "Multi Lane Traffic: Type SU5 Tons",
        "Multi Lane Traffic: Type SU6 Tons", "Multi Lane Traffic: Type SU7 Tons",
        "One Lane Traffic: Type SU4 Tons", "One Lane Traffic: Type SU5 Tons",
        "One Lane Traffic: Type SU6 Tons", "One Lane Traffic: Type SU7 Tons"
    ]

    #Convert tons columns to numeric and fill NaN with 0
    for col in tons: 
        if col in ACT5.columns:
            ACT5[col] = pd.to_numeric(ACT5[col], errors="coerce") 
            ACT5[col] = ACT5[col].fillna(0).astype(int) #V2 Addition
        else:
            ACT5[col] = 0

    # Keep only rows where only SU7 has traffic
    ACT5 = ACT5[
        ((ACT5["Multi Lane Traffic: Type SU4 Tons"] == 0) &
         (ACT5["Multi Lane Traffic: Type SU5 Tons"] == 0) &
         (ACT5["Multi Lane Traffic: Type SU6 Tons"] == 0) &
         (ACT5["Multi Lane Traffic: Type SU7 Tons"] > 0))
        |
        ((ACT5["One Lane Traffic: Type SU4 Tons"] == 0) &
         (ACT5["One Lane Traffic: Type SU5 Tons"] == 0) &
         (ACT5["One Lane Traffic: Type SU6 Tons"] == 0) &
         (ACT5["One Lane Traffic: Type SU7 Tons"] > 0))
    ]

    # Filter NBI 063 != 0
    ACT5["NBI 063 Method Used Operating Rating"] = pd.to_numeric(
        ACT5["NBI 063 Method Used Operating Rating"], errors="coerce"
    )
    ACT5 = ACT5[ACT5["NBI 063 Method Used Operating Rating"] != 0]

    # Remove NBI 041 = K/P/R or if NA and Posting Status is C/PP/PR
    OPC = ["P", "R", "K"]
    LPS = ["C", "PP", "PR"]

    ACT5 = ACT5[
        ~(
            (ACT5["NBI 041 Open, Posted Or Closed"].isin(OPC))
            |
            (ACT5["NBI 041 Open, Posted Or Closed"].isna()
             & ACT5["B.PS.01: Load Posting Status"].isin(LPS))
        )
    ]

    #Remove based on comments
    COM = ["SU4", "close bridge", "closed", "based on a parametric",
           "based on the parametric", "Standards"]

    COM_PAT = "|".join([re.escape(word) for word in COM])
    STAN_PAT = r"(?<!non[-\s])standard|(?<!non[-\s])std"
    FINAL_PAT = f"{COM_PAT}|{STAN_PAT}"

    ACT5["Comments"] = ACT5["Comments"].astype(str)
    ACT5["Comment Inv Rating"] = ACT5["Comment Inv Rating"].astype(str)

    ACT5 = ACT5[
        ~(
            ACT5["Comments"].str.contains(FINAL_PAT, case=False, na=False, regex=True)
            |
            ACT5["Comment Inv Rating"].str.contains(FINAL_PAT, case=False, na=False, regex=True)
        )
    ]

    return ACT5


def action6(RAW3):
    ACT6 = RAW3.copy()

    # -----------------------------
    # Remove City and County Bridges
    # -----------------------------
    bridge = ["City Bridges", "County Bridges"]
    bridge_pat = "|".join(bridge)
    ACT6 = ACT6[~ACT6["Parent Asset"].str.contains(bridge_pat, case=False, na=False)]

    # -----------------------------
    # Traffic tonnage columns
    # -----------------------------
    tons = [
        "Multi Lane Traffic: Type SU4 Tons", "Multi Lane Traffic: Type SU5 Tons",
        "Multi Lane Traffic: Type SU6 Tons", "Multi Lane Traffic: Type SU7 Tons",
        "One Lane Traffic: Type SU4 Tons", "One Lane Traffic: Type SU5 Tons",
        "One Lane Traffic: Type SU6 Tons", "One Lane Traffic: Type SU7 Tons"
    ]

    # Convert tons columns to numeric and fill NaN with 0
    for col in tons:
        if col in ACT6.columns:
            ACT6[col] = pd.to_numeric(ACT6[col], errors="coerce")
            ACT6[col] = ACT6[col].fillna(0).astype(int)
        else:
            ACT6[col] = 0

    # Keep only rows where ALL traffic columns are zero
    ACT6 = ACT6[
        ((ACT6["Multi Lane Traffic: Type SU4 Tons"] == 0) &
         (ACT6["Multi Lane Traffic: Type SU5 Tons"] == 0) &
         (ACT6["Multi Lane Traffic: Type SU6 Tons"] == 0) &
         (ACT6["Multi Lane Traffic: Type SU7 Tons"] == 0)) &
        ((ACT6["One Lane Traffic: Type SU4 Tons"] == 0) &
         (ACT6["One Lane Traffic: Type SU5 Tons"] == 0) &
         (ACT6["One Lane Traffic: Type SU6 Tons"] == 0) &
         (ACT6["One Lane Traffic: Type SU7 Tons"] == 0))
    ]

    # -----------------------------
    # Load Rating Method filters
    # -----------------------------
    LRM = ["EJ", "AR"]
    ACT6["B.LR.04: Load Rating Method"] = ACT6["B.LR.04: Load Rating Method"].astype(str).str.strip().str.upper()
    ACT6 = ACT6[~ACT6["B.LR.04: Load Rating Method"].isin(LRM)]

    # -----------------------------
    # Remove based on NBI 041 / Load Posting Status
    # -----------------------------
    OPC = ["R", "P", "K"]
    LPS = ["C", "PP", "PR"]
    ACT6["NBI 041 Open, Posted Or Closed"] = ACT6["NBI 041 Open, Posted Or Closed"].astype(str).str.strip().str.upper()
    ACT6["B.PS.01: Load Posting Status"] = ACT6["B.PS.01: Load Posting Status"].astype(str).str.strip().str.upper()
    ACT6 = ACT6[
        ~(
            ACT6["NBI 041 Open, Posted Or Closed"].isin(OPC) |
            (ACT6["NBI 041 Open, Posted Or Closed"].isna() & ACT6["B.PS.01: Load Posting Status"].isin(LPS))
        )
    ]

    # -----------------------------
    # Remove based on Comments
    # -----------------------------
    COM = ["SU4", "close bridge", "closed", "based on a parametric", "based on the parametric", "Standards"]
    COM_PAT = "|".join([re.escape(word) for word in COM])
    STAN_PAT = r"(?<!non[-\s])standard|(?<!non[-\s])std"
    FINAL_PAT = f"{COM_PAT}|{STAN_PAT}"

    ACT6["Comments"] = ACT6["Comments"].astype(str)
    ACT6["Comment Inv Rating"] = ACT6["Comment Inv Rating"].astype(str)

    ACT6 = ACT6[
        ((ACT6["Comments"].isna() | (ACT6["Comments"].str.strip() == "")) &
         (ACT6["Comment Inv Rating"].isna() | (ACT6["Comment Inv Rating"].str.strip() == ""))) |
        ~(ACT6["Comments"].str.contains(FINAL_PAT, case=False, na=False, regex=True) |
          ACT6["Comment Inv Rating"].str.contains(FINAL_PAT, case=False, na=False, regex=True))
    ]

    # -----------------------------
    # Convert for numeric filtering
    # -----------------------------
    ACT6["NBI 063 Method Used Operating Rating"] = ACT6["NBI 063 Method Used Operating Rating"].astype(str).str.strip().str.upper()
    ACT6["NBI 064 Operating Rating"] = pd.to_numeric(ACT6["NBI 064 Operating Rating"], errors="coerce")
    ACT6["B.LR.06: Operating Load Rating Factor"] = pd.to_numeric(ACT6["B.LR.06: Operating Load Rating Factor"], errors="coerce")

    # -----------------------------
    # Remove G1 low rating bridges
    # -----------------------------
    G1 = ["1", "2", "3", "4", "5", "A", "C"]
    ACT6 = ACT6[~((ACT6["NBI 063 Method Used Operating Rating"].isin(G1)) &
                  (ACT6["NBI 064 Operating Rating"] < 45))]

    # -----------------------------
    # Remove Year Built = 0
    # -----------------------------
    ACT6 = ACT6[~(ACT6["B.W.01: Year Built"].astype(str) == "0")]

    # -----------------------------
    # Remove G2 / low rating bridges
    # -----------------------------
    G2 = ["6", "7", "8", "F", "D", "f", "d"]
    ACT6 = ACT6[~(
        (ACT6["NBI 063 Method Used Operating Rating"].isin(G2) & (ACT6["NBI 064 Operating Rating"] < 1.26)) |
        (ACT6["NBI 063 Method Used Operating Rating"].isna() & (ACT6["B.LR.06: Operating Load Rating Factor"] < 1.26))
    )]

    # -----------------------------
    # Remove blank or NA Method Used
    # -----------------------------
    ACT6 = ACT6[~(ACT6["NBI 063 Method Used Operating Rating"].isna() |
                  (ACT6["NBI 063 Method Used Operating Rating"] == ""))]

    return ACT6


# -------------------------------
# Excel Generation Function
# -------------------------------

def generate_bridge_excel(
    RAW, RAW2, RAW3,
    ACT2, ACT3, ACT5, ACT6,
    ACT7, ACT8, ACT9_F, ACT9_CT_S, ACT9_CT_NS,
    ACT15_F, ACT16_F, ACT17_F, ACT18_F,
    ACT19_F, ACT19_CT_SB, ACT19_CT_SD, ACT19_CT_NP, ACT19_CT_BLT,
    ACT20, ACT21, ACT22_F
):
    """
    Generates a fully formatted Excel workbook in-memory for the bridge metrics.
    Returns a BytesIO object suitable for Streamlit download_button.
    """
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer: 

        RAW.to_excel(writer, sheet_name="RAW", float_format="%.3f", startrow=4, startcol=0, index=False)
        RAW2.to_excel(writer, sheet_name="RAW2", float_format="%.3f", startrow=4, startcol=0, index=False)
        RAW3.to_excel(writer, sheet_name="RAW3", float_format="%.3f", startrow=4, startcol=0, index=False)
        ACT7.to_excel(writer, sheet_name="ACTION7", float_format="%.3f", startrow=4, startcol=0, index=False)
        ACT8.to_excel(writer, sheet_name="ACTION8", float_format="%.3f", startrow=4, startcol=0, index=False)
        ACT9_F.to_excel(writer, sheet_name="ACTION9", float_format="%.3f", startrow=5, startcol=0, index=False)
        ACT15_F.to_excel(writer, sheet_name="ACTION15", float_format="%.3f", startrow=4, startcol=0, index=False)
        ACT16_F.to_excel(writer, sheet_name="ACTION16", float_format="%.3f", startrow=4, startcol=0, index=False)
        ACT17_F.to_excel(writer, sheet_name="ACTION17", float_format="%.3f", startrow=4, startcol=0, index=False)
        ACT18_F.to_excel(writer, sheet_name="ACTION18", float_format="%.3f", startrow=4, startcol=0, index=False)
        ACT19_F.to_excel(writer, sheet_name="ACTION19", float_format="%.3f", startrow=7, startcol=0, index=False)
        ACT20.to_excel(writer, sheet_name="ACTION20", float_format="%.3f", startrow=4, startcol=0, index=False)
        ACT21.to_excel(writer, sheet_name="ACTION21", float_format="%.3f", startrow=4, startcol=0, index=False)
        ACT22_F.to_excel(writer, sheet_name="ACTION22", float_format="%.3f", startrow=4, startcol=0, index=False)
        ACT2.to_excel(writer, sheet_name="ACTION2", float_format="%.3f", startrow=4, startcol=0, index=False)
        ACT3.to_excel(writer, sheet_name="ACTION3", float_format="%.3f", startrow=4, startcol=0, index=False)
        ACT5.to_excel(writer, sheet_name="ACTION5", float_format="%.3f", startrow=4, startcol=0, index=False)
        ACT6.to_excel(writer, sheet_name="ACTION6", float_format="%.3f", startrow=4, startcol=0, index=False)

        # --- Access workbook and sheets for formatting ---
        workbook = writer.book
        sheet1 = writer.sheets["ACTION7"]
        sheet2 = writer.sheets["ACTION8"]
        sheet3 = writer.sheets["ACTION9"]
        sheet4 = writer.sheets["ACTION15"]
        sheet5 = writer.sheets["ACTION16"]
        sheet6 = writer.sheets["ACTION17"]
        sheet7 = writer.sheets["ACTION18"]
        sheet8 = writer.sheets["ACTION19"]
        sheet9 = writer.sheets["ACTION20"]
        sheet10 = writer.sheets["ACTION21"]
        sheet11 = writer.sheets["ACTION22"]
        sheet12 = writer.sheets["ACTION2"]
        sheet13 = writer.sheets["ACTION3"]
        sheet14 = writer.sheets["ACTION5"]
        sheet15 = writer.sheets["ACTION6"]
        sheet16 = writer.sheets["RAW"]
        sheet17 = writer.sheets["RAW2"]
        sheet18 = writer.sheets["RAW3"]

            # --- Define formats ---
        Aleft = workbook.add_format({"bold": True, "align": "left"})
        bold = workbook.add_format({"bold": True})

            # --- Write Headers and Counts ---
        sheet1.write("A1", "Action Item 7 (formerly Action Item 3 in Metric 13): LPA bridges built after 1994 with ASR load ratings. These bridges must be updated to LFR.", Aleft)
        sheet1.write("A3", "Total Bridges:")
        sheet1.write("B3", len(ACT7), bold)

        sheet2.write("A1", "Action Item 8: LPA bridges reconstructed after 1994 with ASR load ratings. These bridges must be updated to LFR. ", Aleft)
        sheet2.write("A3", "Total Bridges:")
        sheet2.write("B3", len(ACT8), bold)

        sheet3.write("A1", "Action Item 9 (formerly Action Item 5 in Metric 13):  LPA bridges designed LRFD after October 1st, 2010 but are rated LFR. These bridges must be updated to LRFR.", Aleft)
        sheet3.write("A2", "Total Bridges:")
        sheet3.write("B2", len(ACT9_F), bold)
        sheet3.write("A3","Standard:")
        sheet3.write("B3", ACT9_CT_S, bold)
        sheet3.write("A4","Non-Standard:")
        sheet3.write("B4", ACT9_CT_NS, bold)

        sheet4.write("A1", "Action Item 15 (formerly Action Item 11 in Metric 13): LPA bridges with Assigned load ratings where ratings are not appropriate. Load Rating calculations are needed.", Aleft)
        sheet4.write("A3", "Total Bridges:")
        sheet4.write("B3", len(ACT15_F), bold)

        sheet5.write("A1", "Action Item 16 (formerly Action Item 5 in Metric 15): LPA bridges have Assigned load ratings. Documentation is needed to state the criteria for how the Assigned ratings are appropriate.", Aleft)
        sheet5.write("A3", "Total Bridges:")
        sheet5.write("B3", len(ACT16_F), bold)

        sheet6.write("A1", "Action Item 17 (formerly Action Item 12 in Metric 13): DOT owned bridges have Assigned load ratings where ratings are not appropriate. Load Rating calculations are needed.", Aleft)
        sheet6.write("A3", "Total Bridges:")
        sheet6.write("B3", len(ACT17_F), bold)

        sheet7.write("A1", "Action Item 18 (formerly Action Item 6 in Metric 15): DOT bridges have Assigned load ratings. Documentation is needed to state the criteria for how the Assigned ratings are appropriate.", Aleft)
        sheet7.write("A3", "Total Bridges:")
        sheet7.write("B3", len(ACT18_F), bold)

        sheet8.write("A1", "Action Item 19 (formerly Action Item 13 in Metric 13): LPA bridges rated with engineering judgement needing load rating calculations or documentation explaining why EJ was used.", Aleft)
        sheet8.write("A2", "Total Bridges:")
        sheet8.write("B2", len(ACT19_F), bold) 
        sheet8.write("A3","Standard Bridge:")
        sheet8.write("B3", ACT19_CT_SB, bold)
        sheet8.write("A4","Severe Deterioration:")
        sheet8.write("B4", ACT19_CT_SD, bold)
        sheet8.write("A5","Not Permitted:")
        sheet8.write("B5", ACT19_CT_NP, bold)
        sheet8.write("A6","Bridge was Load Tested:")
        sheet8.write("B6", ACT19_CT_BLT, bold)

        sheet9.write("A1", "Action Item 20 (formerly Action Item 7 in Metric 15): LPA bridges rated with engineering judgement valid documentation.", Aleft)
        sheet9.write("A3", "Total Bridges:")
        sheet9.write("B3", len(ACT20), bold) 

        sheet10.write("A1", "Action Item 21 (formerly Action Item 8 in Metric 15): DOT bridges rated with engineering judgement valid documentation.", Aleft)
        sheet10.write("A3", "Total Bridges:")
        sheet10.write("B3", len(ACT21), bold) 

        sheet11.write("A1", "Action Item 22 Query excludes bridges built in 2022 and after.", Aleft)
        sheet11.write("A3", "Total Bridges:")
        sheet11.write("B3", len(ACT22_F), bold) 

        sheet12.write("A1", "Action Item 2 (formerly Action Item 1 in Metric 15): LPA bridges that only have the controlling Specialized Hauling Vehicle rating entered. They must have all Specialized Hauling Vehicles ratings included in the Load Rating tables.", Aleft)
        sheet12.write("A3", "Total Bridges:")
        sheet12.write("B3", len(ACT2), bold) 

        sheet13.write("A1", "Action Item 3 (formerly Action Item 2 in Metric 15):  LPA bridges where the DOT parametric study was used to determine the load ratings for SHV’s but a note needs to be included in the comment field of the Load Rating Report.", Aleft)
        sheet13.write("A3", "Total Bridges:")
        sheet13.write("B3", len(ACT3), bold)

        sheet14.write("A1", "Action Item 5 (formerly Action Item 3 in Metric 15): DOT bridges that have only entered the controlling Specialized Hauling Vehicle rating.", Aleft)
        sheet14.write("A3", "Total Bridges:")
        sheet14.write("B3", len(ACT5), bold) 

        sheet15.write("A1", "Action Item 6 (formerly Action Item 4 in Metric 15): DOT bridges where the DOT parametric study was used to determine the load ratings for SHV’s but a note needs to be included in the comment field of the Load Rating Report.", Aleft)
        sheet15.write("A3", "Total Bridges:")
        sheet15.write("B3", len(ACT6), bold) 

        sheet16.write("A1", "Raw Data From Original Query", Aleft)
        sheet17.write("A1", "Raw Data Minus Bridges in Action 7, Action 8, Action 9, Action 15, Action 16, Action 17, Action 18, Action 19, Action 20, Action 21 and Action 22", Aleft)
        sheet17.write("A2", "Used to Calculate Action 2 and Action 3", Aleft)
        sheet18.write("A1", "Raw Data Minus Bridges in Action 8", Aleft)
        sheet18.write("A2", "Used to Calculate Action 5 and Action 6", Aleft)

    # writer.save() # This is not needed, 'with' statement handles saving/closing
    output.seek(0)

    return output


# In[ ]:




