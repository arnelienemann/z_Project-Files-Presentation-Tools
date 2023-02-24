import streamlit as st
import pandas as pd

def app():

    st.header("How to: Merge Erika and Rogator Data")

    # Allow the user to upload the first Excel file
    uploaded_file1 = st.file_uploader("Choose the Rogator Excel file (.xls):")

    # Read the first Excel file into a Pandas DataFrame
    if uploaded_file1 is not None:
        df1 = pd.read_excel(uploaded_file1)

    # Allow the user to upload the second Excel file
    uploaded_file2 = st.file_uploader("Choose the Erika file (.csv):", type="csv")

    # Read the second Excel file into a Pandas DataFrame
    if uploaded_file2 is not None:
        df2 = pd.read_csv(uploaded_file2, sep=";")
        df2 = df2.rename(columns={'participationId': 'UserID'})

        # Display the first and second DataFrame
        #st.write("First DataFrame:", df1)
        #st.write("Second DataFrame:", df2)


    if uploaded_file1 and uploaded_file2:
        #common_column = st.selectbox("Select the common column:", df1.columns)
        # Combine the two DataFrames using a common column
        result = pd.merge(df1, df2, on="UserID")

        # Display the resulting DataFrame
        st.write("Combined dataset:", result)

        # Download the resulting DataFrame as an Excel file
        excel_file = result.to_excel("data-vlookup/Combined Dataset.xlsx")
        with open("data-vlookup/Combined Dataset.xlsx", "rb") as download_file:
            st.download_button(label = 'Save combined dataset', data = download_file, file_name = 'Combined Dataset.xlsx') 
