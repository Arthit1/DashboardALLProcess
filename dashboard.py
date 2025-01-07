import pandas as pd
import streamlit as st
import os
import matplotlib.pyplot as plt
from io import BytesIO
import requests
from matplotlib import font_manager as fm

# Function to download font from Google Fonts and register it with Matplotlib
def register_google_font(font_name):
    url = f"https://fonts.googleapis.com/css2?family={font_name.replace(' ', '+')}:wght@400;500;600;700&display=swap"
    response = requests.get(url)
    if response.status_code == 200:
        ttf_url = response.text.split("url(")[1].split(")")[0].replace('"', '')
        font_data = requests.get(ttf_url).content
        font_file = BytesIO(font_data)
        font_path = fm.FontProperties(fname=font_file).get_name()
        fm.fontManager.addfont(font_path)
        plt.rcParams["font.family"] = font_path
    else:
        st.error(f"Failed to download font '{font_name}' from Google Fonts.")

# Register the Sarabun font dynamically
register_google_font("Sarabun")

# Apply Google Fonts for Streamlit UI
st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@400;500;600;700&display=swap');
        body {
            font-family: 'Sarabun', sans-serif;
        }
        .streamlit-expanderHeader {
            font-family: 'Sarabun', sans-serif;
        }
    </style>
    """, unsafe_allow_html=True)

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
    # File uploader
    uploaded_file = st.sidebar.file_uploader(
        "Upload your Excel file", type=["xlsx", "xls"]
    )
    if uploaded_file:
        data = pd.read_excel(uploaded_file, sheet_name="Tara-Silom")  # Load only the "Tara-Silom" sheet
elif data_mode == "Specify Path":
    # Text input for file path
    file_path = st.sidebar.text_input(
        "Enter the path to your Excel file",
        value=r"D:\Code\ALLProcess\filtered_data.xlsx"  # Example default path
    )
    if st.sidebar.button("Load Data"):
        if os.path.exists(file_path):
            try:
                data = pd.read_excel(file_path, sheet_name="Tara-Silom")  # Load only the "Tara-Silom" sheet
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

            # Customize font for the pie chart
            for text in autotexts:
                text.set_fontsize(10)  # Set font size for percentage labels
            for text, count in zip(texts, status_counts):
                text.set_text(f"{text.get_text()} ({count} รายการ)")  # Add count to the label

            ax.set_title('Pie Chart of Selected สถานะของเอกสาร', fontsize=14)
            st.pyplot(fig)

        else:
            st.info("Please select at least one status to display the chart.")
    else:
        st.error(f"Column '{status_column}' not found in the data.")
else:
    st.info("Please upload a file or specify a path to load data.")
