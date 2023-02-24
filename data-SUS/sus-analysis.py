import streamlit as st
import xlrd
import numpy
import pandas as pd
#import plotly
#import Python-pptx

#file = st.file_uploader("Upload:") #use copy from clipboard instead

SUS = {}
file = pd.DataFrame()

st.header("SUS Score")
st.write("Step 1: copy the SUS columns into your clipboard and click the button below.")

if st.button("Read Clipboard"):
    file = pd.read_clipboard()


if file is not None:
        
    df = file
    st.dataframe(df)

    for column in df.columns:   

        if "_1" in column and not "_10" in column:
            SUS["Usage frequency"] = df[column].mean() -1
        if "_2" in column:
            SUS["Complexity"] = 6 - df[column].mean()   
        if "_3" in column:
            SUS["Easy to use"] = df[column].mean() -1
        if "_4" in column:
            SUS["Assistance needed"] = 6 - df[column].mean()  
        if "_5" in column:
            SUS["Features well integrated"] = df[column].mean() -1
        if "_6" in column:
            SUS["Inconsistency"] = 6 - df[column].mean()  
        if "_7" in column:
            SUS["Quick to learn"] = df[column].mean() -1
        if "_8" in column:
            SUS["Cumbersome to use"] = 6 - df[column].mean()  
        if "_9" in column:
            SUS["Confident use"] = df[column].mean() -1
        if "_10" in column:
            SUS["A lot learning required"] = 6 - df[column].mean()  

    st.write(SUS)
    


#python -m streamlit run sus-analysis.py