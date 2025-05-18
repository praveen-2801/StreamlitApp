import streamlit as st
import pandas as pd
from io import BytesIO

# Sidebar navigation
st.sidebar.title("Main Menu")
option = st.sidebar.radio("Choose option : ", ["Home", "Pricing Analysis V2"])

if option == "Home":
    st.title("Welcome , this app allows you to merge multiple excel files into one file")

if option == "Pricing Analysis V2":
    st.title("Pricing Analysis V2")

    # File uploader
    uploaded_files = st.file_uploader("Upload Excel files", type=["xlsx"], accept_multiple_files=True)

    if uploaded_files:
        st.success(f"{len(uploaded_files)} file(s) uploaded.")

        if st.button("Merge Files"):
            try:
                # Read and concatenate Excel files
                merged_df = pd.concat([pd.read_excel(file) for file in uploaded_files], ignore_index=True)

                st.success("Files successfully merged!")
                # st.dataframe(merged_df)

                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    merged_df.to_excel(writer, index=False, sheet_name="Merged")
                    # writer.save()
                output.seek(0)
                # Download button
                st.download_button(
                    label="Download Merged Excel",
                    data=output,
                    file_name="merged_pricing_data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"An error occurred during merging: {e}")
