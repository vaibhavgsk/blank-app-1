import streamlit as st
import pandas as pd
import io

# buffer to use for Excel writer
buffer = io.BytesIO()

# Title of the app
st.title("Upload a Raw Excel File")

# File uploader widget
uploaded_file = st.file_uploader("Choose a file", type=["xlsx"])

# Check if a file has been uploaded
if uploaded_file is not None:

    if uploaded_file.type == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
        # Read the Excel file into a DataFrame
        df = pd.read_excel(uploaded_file)
        st.write("Uploaded DataFrame:", df)  # Check uploaded data

    else:
        st.error("Unsupported file type")

if uploaded_file is not None and st.button("Run Transformation"):
    if 'df' in locals():  # Ensure DataFrame exists before proceeding
        last_header_0 = None
        # Check if 'Header' column exists in the uploaded file
        if 'Header' not in df.columns:
            st.error("'Header' column is missing from the uploaded file.")
        else:
            for i, row in df.iterrows():
                if row['Header'] == 0:
                    # Store the current Header=0 row's values in temp variable named last_header_0
                    last_header_0 = row
                elif last_header_0 is not None:
                    # For each Header=1 row, fill values from the last Header=0 row (i.e it will not overwrite existing non-empty values)
                    # Replace only the empty strings with values from Header=0 row
                    for col in df.columns:
                        if row[col] == '' or pd.isna(row[col]):
                            df.at[i, col] = last_header_0[col]

            # Filter out rows where 'Header' is not 1 (just in case)
            Report = pd.DataFrame({
                'SoldTo': df[df['Header'] == 1]['SoldTo'],
                'ShipTo': df[df['Header'] == 1]['ShipTo'],
                'Order Number': df[df['Header'] == 1]['PONumberBSTNK'],
                'Material': df[df['Header'] == 1]['Material'],
                'Qty': df[df['Header'] == 1]['Quant'],
                'Customer Name': df[df['Header'] == 1]['ShipToName1'],
                'Customer Name 2': df[df['Header'] == 1]['ShipToName2'],
                'Street Address': df[df['Header'] == 1]['ShipToStreet1'],
                'City': df[df['Header'] == 1]['ShipToCity'],
                'District Suburb': df[df['Header'] == 1]['ShipToRegion'],
                'Postcode': df[df['Header'] == 1]['ShiptoPostCode'],
            })

            # Fill and format columns
            Report['SoldTo'] = Report['SoldTo'].astype('int64').astype('str')
            Report['ShipTo'] = Report['ShipTo'].astype('str')
            Report['Order Number'] = Report['Order Number'].astype('str')
            Report['Material'] = Report['Material'].astype('int64').astype('str')
            Report['Qty'] = Report['Qty'].fillna(0).astype('int')
            Report['Customer Name'] = Report['Customer Name'].astype('str')
            Report['Customer Name 2'] = Report['Customer Name 2'].astype('str')
            Report['Street Address'] = Report['Street Address'].astype('str')
            Report['City'] = Report['City'].astype('str')
            Report['District Suburb'] = Report['District Suburb'].astype('str')
            Report['Postcode'] = Report['Postcode'].astype('int64').astype('str')

            Report.reset_index(drop=True, inplace=True)

            st.write("Transformed DataFrame:", Report)  # Check transformed data

            # Write the transformed DataFrame to the Excel buffer
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                Report.to_excel(writer, index=False, sheet_name="Report")

            # Prepare the download button
            buffer.seek(0)  # Ensure the buffer is at the beginning

            download2 = st.download_button(
                label="Download data as Excel",
                data=buffer,
                file_name='Report.xls',
                mime='application/vnd.ms-excel'
            )

            st.success("Transformation executed successfully!")

    else:
        st.error("Please upload a file and ensure data is available before transforming.")
