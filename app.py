#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import streamlit as st
import io
import pandas as pd
from bridge import (
    raw_file,
    act8_fil,
    generate_bridge_excel,
    action2, action3, action5, action6,
    action7, action9, action15, action16,
    action17, action18, action19, action20, action21, action22,
    make_RAW2, run_action8m_and_raw3
)

st.set_page_config(page_title="Iowa DOT Bridge Metrics", layout="wide")
st.title("Iowa DOT Bridge Metrics")
st.text("Upload the RAW and Action 8 Excel files.\nClick Run.")

# Upload files
raw_file_uploader = st.file_uploader("Upload RAW Excel File", type=["xlsx", "xls"])
act8_file_uploader = st.file_uploader("Upload Action 8 Excel File", type=["xlsx", "xls"])

# Initialize session state
if "results_ready" not in st.session_state:
    st.session_state.results_ready = False


# ----------------- PROCESSING BUTTON -------------------
if st.button("Run"):
    if raw_file_uploader is None or act8_file_uploader is None:
        st.error("Please upload both files before running.")
    else:
        with st.spinner("Processing..."):
            RAW_loaded = pd.read_excel(raw_file_uploader)
            ACT8_loaded = pd.read_excel(act8_file_uploader)

            RAW = raw_file(RAW_loaded)
            ACT8 = act8_fil(ACT8_loaded)

            ACT7 = action7(RAW)
            ACT9_F = action9(RAW)
            ACT9_CT_S = len(ACT9_F[ACT9_F["Standard/Non-Standard"] == "Standard"])
            ACT9_CT_NS = len(ACT9_F[ACT9_F["Standard/Non-Standard"] == "Non-Standard"])
            ACT15_F = action15(RAW)
            ACT16_F = action16(RAW)
            ACT17_F = action17(RAW)
            ACT18_F = action18(RAW)
            ACT19_F = action19(RAW)
            ACT19_CT_SD = len(ACT19_F[ACT19_F["Action 19 Sub-Category"] == "Severe Deterioration"])
            ACT19_CT_SB = len(ACT19_F[ACT19_F["Action 19 Sub-Category"] == "Standard Bridge"])
            ACT19_CT_BLT = len(ACT19_F[ACT19_F["Action 19 Sub-Category"] == "Bridge was load tested."])
            ACT19_CT_NP = len(ACT19_F[ACT19_F["Action 19 Sub-Category"] == "Not Permitted"])
            ACT20 = action20(RAW)
            ACT21 = action21(RAW)
            ACT22_F = action22(RAW)

            RAW2, _ = make_RAW2(
                RAW.copy(), ACT7, ACT8, ACT9_F, ACT15_F, ACT16_F,
                ACT17_F, ACT18_F, ACT19_F, ACT20, ACT21, ACT22_F
            )

            ACT2 = action2(RAW2)
            ACT3 = action3(RAW2)

            ACT8M, RAW3 = run_action8m_and_raw3(RAW.copy(), ACT8.copy())

            ACT5 = action5(RAW3)
            ACT6 = action6(RAW3)

            excel_file = generate_bridge_excel(
                RAW, RAW2, RAW3,
                ACT2, ACT3, ACT5, ACT6,
                ACT7, ACT8, ACT9_F, ACT9_CT_S, ACT9_CT_NS,
                ACT15_F, ACT16_F, ACT17_F, ACT18_F,
                ACT19_F, ACT19_CT_SB, ACT19_CT_SD, ACT19_CT_NP, ACT19_CT_BLT,
                ACT20, ACT21, ACT22_F
            )

            # Save everything in session state
            st.session_state.excel_file = excel_file
            st.session_state.counts = {
                "Action 2": len(ACT2),
                "Action 3": len(ACT3),
                "Action 5": len(ACT5),
                "Action 6": len(ACT6),
                "Action 7": len(ACT7),
                "Action 8 (Uploaded)": len(ACT8),
                "Action 9": len(ACT9_F),
                "Action 15": len(ACT15_F),
                "Action 16": len(ACT16_F),
                "Action 17": len(ACT17_F),
                "Action 18": len(ACT18_F),
                "Action 19": len(ACT19_F),
                "Action 20": len(ACT20),
                "Action 21": len(ACT21),
                "Action 22": len(ACT22_F),
                "RAW (Original)": len(RAW),
                "RAW2 (RAW - Actions 7-22)": len(RAW2),
                "RAW3 (RAW - Action 8)": len(RAW3)
            }
            st.session_state.results_ready = True

        st.success("Processing complete!")


# ----------------- DISPLAY RESULTS (PERSISTENT) -------------------
if st.session_state.results_ready:

    st.subheader("Summary Counts")
    st.table(pd.DataFrame(
        list(st.session_state.counts.items()),
        columns=["Action", "Number of Bridges"]
    ))

    st.download_button(
        label="Download Bridge Metrics Excel",
        data=st.session_state.excel_file,
        file_name="Bridge_Metrics_Output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

