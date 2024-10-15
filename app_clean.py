import numpy as np
import pandas as pd
from collections import Counter
from datetime import datetime
from copy import copy
from IPython.display import display, HTML
import streamlit as st # type: ignore
from streamlit_echarts import st_echarts # type: ignore
import json
import requests
pd.options.display.float_format = "{:,.2f}".format
divider_style = """
    <hr style="border: none; 
    height: 5px; 
    background-color: white; 
    border-radius: 10px; 
    margin: 20px 0;
    opacity: 0.8">
"""
def create_pie_chart(miss_data, corr_data):
    options = {
                "tooltip": {"trigger": "item"},
                "legend": {
                    "top": "5%", 
                    "left": "center",
                    "textStyle": {"color": "#fff"}  # Legend text color to white
                },
                "series": [
                    {
                        "name": "Mismatch Ratio",
                        "type": "pie",
                        "radius": ["40%", "70%"],
                        "avoidLabelOverlap": False,
                        "itemStyle": {
                            "borderRadius": 0,
                            "borderColor": "#fff",
                            "borderWidth": 0,
                        },
                        "label": {
                            "show": False,
                            "position": "center",
                            "color": "#fff",  # Text color white
                            "fontSize": 16,
                            "fontWeight": "bold"
                        },
                        "labelLine": {"show": False},
                        "data": [
                            {"value": miss_data, "name": "Mismatch", "itemStyle": {"color": "#ff6961"}},  # Soft red for mismatch
                            {"value":  corr_data, "name": "Correct", "itemStyle": {"color": "#90ee90"}},  # Soft green for correct
                        ],
                    }
                ],
            }
        
    st_echarts(
        options=options, height="300px",
    )
# Set Streamlit to use the wider layout mode
st.set_page_config(layout="wide", page_title="Kelompok Data Viewer")

def main():
    # Custom CSS to center the title
    st.markdown("""
        <style>
        .centered-title {
            text-align: center;
            text-decoration: underline;
        }
        </style>
        """, unsafe_allow_html=True)

    # Centered title using custom class
    st.markdown("<h1 class='centered-title'>SSKI QUALITY ASSURANCE</h1>", unsafe_allow_html=True)
    st.markdown(divider_style, unsafe_allow_html=True)

    # Centralized styling for the DataFrames
    dataframe_style = {
        'text-align': 'center',
        'background-color': '#E8F6F3'
    }

    def highlight_rows(row):
        # Index will be used to check the first and last rows
        if row.name == 0:
            return ['background-color: #A9DFBF; color: black'] * len(row)  # Green background for the first row
        elif row.name == len(df_clean) - 1:
            return ['background-color: #F5B7B1; color: black'] * len(row)  # Yellow background for the last row
        else:
            return ['background-color: #F9E79F; color: black'] * len(row)  # Red background for all other rows

    # Function to display the DataFrame
    def display_dataframe(input_df):
        st.dataframe(input_df.style.set_properties(**{'text-align': 'center'}).set_table_styles(
            [{'selector': 'th', 'props': [('text-align', 'center'), ('background-color', '#E8F6F3')]}]
        ).format(precision=2))

    
    

    file_path = "https://raw.githubusercontent.com/YudisthiraPutra/streamlit/c695f97e81e9a82ecd007c7438a73ec042a26cb7/data_test.json"

    # Load the JSON file
    response = requests.get(file_path)
    data = response.json()

    raw_data = data['raw_data']
    raw_keys_list = list(raw_data.keys())

    clean_data = data['clean_data']
    clean_keys_list = list(clean_data.keys())

    summary_data = data['summary_data']
    sum_keys_list = list(summary_data.keys())

    horizontal_clean_data = data['horizontal_clean_data']
    hor_clean_keys_list = list(horizontal_clean_data.keys())

    horizontal_raw_data = data['horizontal_raw_data']
    hor_raw_keys_list = list(horizontal_raw_data.keys())

    error_counts = {}
    ver_error_count = 0
    ver_total_count = 0
    for i in range(len(clean_data)):
        df_clean = pd.DataFrame(clean_data[clean_keys_list[i]])
        df_raw = pd.DataFrame(raw_data[raw_keys_list[i]])

        sski_path = df_clean['Path'][0]  # Assuming 'Path' column exists
        sski_number = sski_path.split('.')[1]  # Extract the number after 'SSKI'


        column_count_error = len(df_clean.columns) - 2
        ver_error_count += column_count_error

        column_count_correct = len(df_raw.columns) -2
        ver_total_count += column_count_correct
        error_counts[sski_number] = error_counts.get(sski_number, 0) + column_count_error


    total_error_rows = 0
    total_correct_rows = 0
    total_rows = 0
    hor_error = {}
    for item in hor_raw_keys_list:
        final = pd.DataFrame(horizontal_raw_data[item])
        if final is not None and not final.empty:
            if item in horizontal_clean_data:
                clean = pd.DataFrame(horizontal_clean_data[item])
                total_rows_in_table = len(final)  # Total rows in the table
                error_rows = len(clean)  # Error rows come from `clean`
                correct_rows = total_rows_in_table - error_rows  # Correct rows are the difference
                hor_error[item] = error_rows
                # Update counts
                total_error_rows += error_rows
            
            total_correct_rows += correct_rows
            total_rows += total_rows_in_table
    



    # Define layout with two columns
    col1_g, col2_g, col3_g = st.columns((2, 2,4))
    with col1_g:
        st.markdown("<p style='text-align: center;'><span style='text-align: center;font-weight: bold; text-decoration: underline;'>Vertical Check Mismatch Ratio</span></p>", unsafe_allow_html=True)
        create_pie_chart(ver_error_count,ver_total_count)
        st.markdown(f"<p style='text-align: center;'><span style='font-weight: bold; text-decoration: underline;'>{(ver_error_count/(ver_error_count + ver_total_count)) * 100:.2f}%</span> of mismatches.</p>", unsafe_allow_html=True)

    with col2_g:
        st.markdown("<p style='text-align: center;'><span style='text-align: center;font-weight: bold; text-decoration: underline;'>Horizontal Check Mismatch Ratio</span></p>", unsafe_allow_html=True)
        create_pie_chart(total_error_rows,total_rows)
        st.markdown(f"<p style='text-align: center;'><span style='font-weight: bold; text-decoration: underline;'>{(total_error_rows/total_rows) * 100:.2f}%</span> of mismatches.</p>", unsafe_allow_html=True)


    with col3_g:
        total_ver = sum(error_counts.values())
        st.markdown("<h1 style='text-align: center;'>General Overview</h1>", unsafe_allow_html=True)
        st.markdown(divider_style, unsafe_allow_html=True)
        st.subheader("Mismatch Counts for Vertical Checks")
        
        # Expander for vertical mismatch counts
        with st.expander("Show Vertical Mismatches Count"):
            # Create two columns for better layout
            col1, col2 = st.columns(2)
            with col1:
                for i, (sski_number, count) in enumerate(error_counts.items()):
                    if i % 2 == 0:
                        st.markdown(f"SSKI - {sski_number}: {count} mismatch(es)")
            
            with col2:
                for i, (sski_number, count) in enumerate(error_counts.items()):
                    if i % 2 != 0:
                        st.markdown(f"SSKI - {sski_number}: {count} mismatch(es)")
        
        st.subheader("Mismatch Counts for Horizontal Checks")
        
        # Expander for horizontal mismatch counts
        with st.expander("Show Horizontal Mismatches Count"):
            col1, col2 = st.columns(2)
            with col1:
                for i, (sski_number, count) in enumerate(hor_error.items()):
                    if i % 2 == 0:
                        st.markdown(f"SSKI - {sski_number}: {count} mismatch(es)")
            
            with col2:
                for i, (sski_number, count) in enumerate(hor_error.items()):
                    if i % 2 != 0:
                        st.markdown(f"SSKI - {sski_number}: {count} mismatch(es)")
    st.markdown(divider_style, unsafe_allow_html=True)
    # Define layout with two columns
    col1, col2 = st.columns((1, 4))
    with col1:
        table_list = ["1", "2", "3", "4", "5a", "5b", "5c", "5d", "5d.1", "5.d.2", "6", 
                    "7", "8", "9", "10", "11a", "12", "13", "14", "15", "16a", 
                    "17", "18", "19", "20"]
        
        st.text("Ingin melakukan apa?")
        # Dynamically create buttons for each item in the table_list
        for table in table_list:
            if st.button(f"Lihat SSKI Tabel {table}"):
                st.session_state.selected_table = table
            

    with col2:
        if 'selected_table' in st.session_state:
            selected_number = st.session_state.selected_table
            print(selected_number)
            filtered_keys = [key for key in clean_keys_list if key.split('-')[0] == str(selected_number)]
            count = 0
            for i in filtered_keys:
                df_clean = pd.DataFrame(clean_data[i])
                if df_clean is not None and not (len(df_clean.columns) == 2 and 'Path' in df_clean.columns):
                    df_summary = pd.DataFrame(summary_data[i])

                    if count == 0:
                        st.markdown("<h1 class='centered-title'>VERTICAL CHECK</h1>", unsafe_allow_html=True)

                    st.markdown(divider_style, unsafe_allow_html=True)
                    st.subheader(f"SSKI {i}")

                    display_dataframe(df_summary)

                    with st.expander("See Detail?"):
                        st.write("""
                        **Penjelasan Warna:**
                        - 🟩 : Aggregat
                        - 🟨 : Calculated
                        - 🟥 : Selisih
                        """)
                        st.dataframe(df_clean.style.apply(highlight_rows, axis=1).set_properties(
                        **{'text-align': 'center'}).set_table_styles(
                        [{'selector': 'th', 'props': [('text-align', 'center'), ('background-color', '#E8F6F3')]}]
                        ).format(precision=2))
                    count +=1
            

            if selected_number in horizontal_clean_data:
                st.markdown("<h1 class='centered-title'>HORIZONTAL CHECK</h1>", unsafe_allow_html=True)
                df_clean = pd.DataFrame(horizontal_clean_data[selected_number])
                st.subheader(f"SSKI - {selected_number}")
                st.dataframe(
                    df_clean.style.set_properties(**{'text-align': 'center'})
                    .set_table_styles([{'selector': 'th', 
                                        'props': [('text-align', 'center'), 
                                                ('background-color', '#E8F6F3')]}])
                )
                count +=1

            if count == 0:
                st.text("Tabel ini sudah konsisten")

        else:
            st.markdown("<h1 class='centered-title'>VERTICAL CHECK</h1>", unsafe_allow_html=True)
            for i in range(len(clean_data)):
                
                df_clean = pd.DataFrame(clean_data[clean_keys_list[i]])
                if df_clean is not None and not (len(df_clean.columns) == 2 and 'Path' in df_clean.columns):
                    df_summary = pd.DataFrame(summary_data[sum_keys_list[i]])

                    st.markdown(divider_style, unsafe_allow_html=True)
                    st.subheader(f"SSKI {clean_keys_list[i]}")

                    display_dataframe(df_summary)

                    with st.expander("See Detail?"):
                        st.write("""
                        **Penjelasan Warna:**
                        - 🟩 : Aggregat
                        - 🟨 : Calculated
                        - 🟥 : Selisih
                        """)
                        st.dataframe(df_clean.style.apply(highlight_rows, axis=1).set_properties(
                        **{'text-align': 'center'}).set_table_styles(
                        [{'selector': 'th', 'props': [('text-align', 'center'), ('background-color', '#E8F6F3')]}]
                        ).format(precision=2))

            st.markdown("<h1 class='centered-title'>HORIZONTAL CHECK</h1>", unsafe_allow_html=True)

            table_list = ["1", "2", "3", "4", "5a", "5b", "5c", "5d", "5d.1", "5.d.2", "6", 
                                "7", "8", "9", "10", "11a", "12", "13", "14", "15", "16a", 
                                "17", "18", "19", "20"]
            for item in hor_clean_keys_list:
                df_clean = pd.DataFrame(horizontal_clean_data[item])
                st.subheader(f"SSKI - {item}")
                st.dataframe(df_clean.style.set_properties(**{'text-align': 'center'}).set_table_styles(
                    [{'selector': 'th', 'props': [('text-align', 'center'), ('background-color', '#E8F6F3')]}]
                ).format(precision=2))
            
        

            

# Example usage of the main function
list_tahun = ['2022']  # Define list_tahun as needed
if __name__ == "__main__":
    main()
