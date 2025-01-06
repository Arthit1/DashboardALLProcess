import pandas as pd
import streamlit as st
import os
import matplotlib.pyplot as plt
from matplotlib import rcParams

# Function to ensure UTF-8 encoding for DataFrame
def ensure_utf8_encoding(df):
    return df.apply(lambda x: x.str.encode('utf-8').str.decode('utf-8') if x.dtype == 'object' else x)

# Set custom CSS to use a Thai-compatible font
st.markdown(
    """
    <style>
        body {
            font-family: 'TH SarabunPSK', sans-serif;
        }
    </style>
    """,
    unsafe_allow_html=True
)

# Set Thai-compatible font for matplotlib
rcParams['font.family'] = 'TH SarabunPSK'

# Title of the dashboard
st.title("Tara-Silom Data Dashboard")

# Sidebar section for data input
st.sidebar.header("Upload Data")
data_mode = st.sidebar.radio(
    "How would you like to provide the data?",
    ("Upload File", "Specify Path")
)

# Load data based on user input
data = None
if data_mode == "Upload File":
    uploaded_file = st.sidebar.file_uploader(
        "Upload your Excel file", type=["xlsx", "xls"]
    )
    if uploaded_file:
        data = pd.read_excel(uploaded_file, sheet_name="Tara-Silom")  # Load only the "Tara-Silom" sheet
        data = ensure_utf8_encoding(data)  # Ensure encoding
elif data_mode == "Specify Path":
    file_path = st.sidebar.text_input(
        "Enter the path to your Excel file",
        value=r"D:\Code\ALLProcess\filtered_data.xlsx"  # Example default path
    )
    if st.sidebar.button("Load Data"):
        if os.path.exists(file_path):
            try:
                data = pd.read_excel(file_path, sheet_name="Tara-Silom")  # Load only the "Tara-Silom" sheet
                data = ensure_utf8_encoding(data)  # Ensure encoding
                st.sidebar.success("Data loaded successfully!")
            except Exception as e:
                st.sidebar.error(f"Error loading file: {e}")
        else:
            st.sidebar.error("The specified file path does not exist.")

# Display data and chart if loaded
if data is not None:
    st.subheader("Data from 'Tara-Silom' Sheet")
    st.write(data)

    # Check if the column 'สถานะของเอกสาร' exists in the data
    status_column = 'สถานะของเอกสาร'
    if status_column in data.columns:
        # Convert statuses to strings to ensure proper rendering
        unique_statuses = data[status_column].astype(str).unique()
        selected_statuses = st.multiselect(
            "Select the statuses to include in the chart",
            options=unique_statuses,
            default=[]  # Default to clear filter (no statuses selected)
        )

        if selected_statuses:
            # Filter data based on selected statuses
            filtered_data = data[data[status_column].isin(selected_statuses)]
            status_counts = filtered_data[status_column].value_counts()

            # Display the pie chart with counts inside the chart area
            st.subheader("Pie Chart of Selected สถานะของเอกสาร ALL Process")
            fig, ax = plt.subplots()
            wedges, texts, autotexts = ax.pie(
                status_counts,
                labels=status_counts.index,
                autopct='%1.1f%%',
                startangle=140,
                colors=plt.cm.tab20.colors[:len(status_counts)]  # Adjust colors based on the number of statuses
            )

            # Customize font for autopct
            for text in autotexts:
                text.set_fontsize(10)  # Set font size for percentage labels
            for text, count in zip(texts, status_counts):
                text.set_text(f"{text.get_text()} ({count} รายการ) ")  # Add count to the label

            ax.set_title('Pie Chart of Selected สถานะของเอกสาร')
            st.pyplot(fig)

        else:
            st.info("Please select at least one status to display the chart.")
    else:
        st.error(f"Column '{status_column}' not found in the data.")
else:
    st.info("Please upload a file or specify a path to load data.")
