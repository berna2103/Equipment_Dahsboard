import streamlit as st
import pandas as pd
import os
import plotly.express as px
from geopy.geocoders import Nominatim
import plotly.graph_objects as go
from datetime import datetime

st.set_page_config(layout="wide")

excel_file_with_coordinates = "./data/output_excel.xlsx"
REPORT_URL = 'https://elekta.my.salesforce.com/00O6g000006RqPX'
df = pd.DataFrame()
st.title('Region Overview')

###################### LOAD DATA ##################################################################################
st.sidebar.title('Load Data')

uploaded_file = st.sidebar.file_uploader("Choose SCV File", type=['csv'])


if uploaded_file:
    if os.path.exists(excel_file_with_coordinates):
        df = pd.read_excel("./data/output_excel.xlsx")
   
    else:
        df = pd.read_csv(uploaded_file)
    
    df.rename(columns= {
            'Account': 'account',
            'Location': 'location',
            'IP Street': 'street',
            'IP City': 'city',
            'IP State': 'state',
            'IP Zip/Postal Code': 'zipcode',
            'Primary Technician: Member Name': 'primary_fse',
            'Secondary Technician Name': 'secondary_fse',
            'EoL Date IP': 'eol',
            'Device Age': 'device_age',
            'Customer/Device Acceptance Date': 'cat',
            'EoGS Date IP': 'eogs',
            'Installed Product: Installed Product': 'ip'

        }, inplace=True
    )

    # Drop rows where column 'B' has NaN
    df_clean = df.dropna(subset=['location'])

    def combine_address(row):
        zipcode = row['zipcode'][:5]
        return f"{row['street']}, {row['city']}, {row['state']}, {zipcode}"

    df_clean['address'] = df_clean.apply(combine_address, axis=1)

    # Initialize geocoder
    geolocator = Nominatim(user_agent="LocatingMachines")

    # Function to get latitude and longitude from an address
    def geocode_address(address):
        print(f"Geocoding: {address}")
        try:
            location = geolocator.geocode(address, timeout=10)

            if location:
                print(f"Found: {location.latitude}, {location.longitude}")  # Print the found coordinates
                return pd.Series([location.latitude, location.longitude])
            else:
                print("Location not found!")
                return pd.Series([None, None])

        except Exception as e:
            print(f"Error: {e}")
            return pd.Series([None, None])



    # Apply geocoding to get latitude and longitude

    if os.path.exists(excel_file_with_coordinates):
        print("File Already Exists!")
    else:
        df_clean[['latitude', 'longitude']] = df_clean['address'].apply(geocode_address)
        df_clean.to_excel('output_excel_coordinates_missing.xlsx', sheet_name='data')
        df_clean_no_coordinates = df_clean.dropna(subset=['latitude', 'longitude'])
        df_clean_no_coordinates.to_excel('./data/output_excel.xlsx', sheet_name='data')
        st.write(df_clean_no_coordinates)

    # Region Metrics
    head_count = len(df_clean['primary_fse'].unique())
    average_equipment_age = df_clean['device_age'].mean()
    median_equipment_age = df_clean['device_age'].median()
    df_clean['ip_device_type'] = df['ip'].str.split('/').str[0] 
    df_clean['serial_number'] = df['ip'].str.split('/').str[0]  + df_clean['ip'].apply(lambda x: str(x.split('/')[2]))

    device_type_frequency = df_clean['ip_device_type'].value_counts()
    
    # Convert unique values and counts to lists
    unique_values = device_type_frequency.index.tolist()
    counts = device_type_frequency.tolist()


    container = st.container(border=True)

    with container:
        col1, col2, col3 = st.columns(3)
        col4,col5, col6 = st.columns(3)
        with col1:
            st.metric(label= 'Head Count', value= head_count)
        with col2:
            st.metric(label='Average Device Age', value= f'{round(average_equipment_age, 2)} yrs')
        with col3:
            st.metric(label='Total IPs', value=(len(df_clean['device_age'])))

        # Step 4: Display metrics in rows of 3 columns
        for i in range(0, len(unique_values), 3):
            # Create 3 columns for the current row
            cols = st.columns(3)
            # Display up to 3 metrics per row
            for j in range(3):
                if i + j < len(unique_values):
                    # Assign value and count to the appropriate column
                    cols[j].metric(label=unique_values[i + j], value=counts[i + j])

    st.divider()

    ############### PLOT DEVICE AGE #################################################
    st.header('Device Age')
    sorted_data = df_clean.sort_values(by='device_age')

    # Create a list of colors based on device_age condition
    colors = ['red' if age > 12 else 'blue' for age in sorted_data['device_age']]

    # Create the bar chart with custom bar width
    device_age_fig = go.Figure(go.Bar(
        y=sorted_data['device_age'],
        x=sorted_data['serial_number'],
        marker=dict(color=colors),
        textposition='auto',
        
    ))
    # Add a vertical line for today's date from top to bottom of the chart
    device_age_fig.add_hline(
        y=12,
        line_width=3,
        line_dash="dash",
        line_color="gray",
        annotation_text="Due for replacement age > 12",  # Optional: Adds text near the line
        annotation_position="top left"
    )

    st.plotly_chart(device_age_fig)

    ############END OF PLOT DEVICE AGE ##################################################

    ############ Plot points on a MAP using Plotly Express #############################
    st.header("Region's Device Locations")
    fig = px.scatter_mapbox(df_clean, 
                            hover_data=['account','address','ip'],
                            lat="latitude", 
                            lon="longitude", 
                            hover_name="location",

                            zoom=5, 
                            height=600)

    # Set map style and layout
    fig.update_layout(mapbox_style="open-street-map")

    st.plotly_chart(fig)
    #############END of MAP ##############################################################


    ############### PLOT EOL and EOGL    #################################################

    st.title("EOL and EOGS")
    # Convert the EOL and EOGL columns to datetime format
    df_clean['eol'] = pd.to_datetime(df_clean['eol'], errors='coerce')
    df_clean['eogs'] = pd.to_datetime(df_clean['eogs'], errors='coerce')


    # Filter out rows where either EOL or EOGL is missing
    df_filtered = df_clean.dropna(subset=['eol', 'eogs'])
    df_filtered['serial_number'] = df_clean['ip'].apply(lambda x: x.split('/')[2])
    df_filtered['ip_type'] = df_clean['ip'].apply(lambda x: x.split('/')[0])


    container = st.container(border=True)
    with container:
        st.header('Installed product EOL timeline')
        # Get today's date as a pandas datetime object
        today = pd.to_datetime(datetime.today())

        # Filter out rows where either EOL or EOGL is missing
        df_filtered = df_clean.dropna(subset=['eol', 'eogs'])

        # Ensure that the 'eogs' column (End of Guaranteed Support) is a datetime object
        df_filtered['eogs'] = pd.to_datetime(df_filtered['eogs'], errors='coerce')

        # Create the DataFrame for the timeline, ensuring all datetime fields are in pandas datetime format
        df_timeline = pd.DataFrame({
            'Account': df_filtered['account'],
            'Product': df_filtered['ip'],
            'Start': [today] * len(df_filtered),  # Create a Start column with today's date for each row
            'End': df_filtered['eol']  # EOGL as the end of guaranteed support
        })


        # Ensure there are no missing datetime values in 'End' after conversion
        df_timeline = df_timeline.dropna(subset=['End'])

        # Add a new column to indicate whether the End date is in the past or future
        df_timeline['Status'] = df_timeline['End'].apply(lambda x: 'Past' if x < today else 'Future')


        # Create the timeline chart using Plotly Express
        fig = px.timeline(df_timeline, 
                          x_start='Start', 
                          x_end='End', 
                          y='Product', 
                          height=1000, 
                          hover_data=['Account'],
                          color='Status',
                          color_discrete_map={'Past': 'red', 'Future': 'green'} ,
                          title="EOL Product Timeline")

        # # Add today's date as a vertical line using the add_vline method
        # fig.add_vline(x=today, line=dict(color="green", dash="dash"), annotation_text="Today", annotation_position="top")

        # Update layout to show date formatting on the x-axis
        fig.update_layout(
            xaxis_title="Date",
            yaxis_title="Products",
            xaxis=dict(tickformat="%m-%d-%Y"),
            showlegend=False,

        )

        # Display the figure in Streamlit
        st.plotly_chart(fig)

    ################### EOGS ###########################################################################################3
    eogs_container = st.container(border=True)
    with eogs_container: 

        st.header('Installed product EOGS timeline')
        # Get today's date as a pandas datetime object

        # Create the DataFrame for the timeline, ensuring all datetime fields are in pandas datetime format
        df_timeline = pd.DataFrame({
            'Account': df_filtered['account'],
            'Product': df_filtered['ip'],
            'Start': [today] * len(df_filtered),  # Create a Start column with today's date for each row
            'End': df_filtered['eogs']  # EOGL as the end of guaranteed support
        })

        # Add a new column to indicate whether the End date is in the past or future
        df_timeline['Status'] = df_timeline['End'].apply(lambda x: 'Past' if x < today else 'Future')


        # Create the timeline chart using Plotly Express
        fig = px.timeline(df_timeline, 
                          x_start='Start', 
                          x_end='End', 
                          y='Product', 
                          height=1000, 
                          hover_data=['Account'],
                          color='Status',
                          color_discrete_map={'Past': 'red', 'Future': 'green'} ,
                          title="EOL Product Timeline")

        # # Add today's date as a vertical line using the add_vline method
        # fig.add_vline(x=today, line=dict(color="green", dash="dash"), annotation_text="Today", annotation_position="top")

        # Update layout to show date formatting on the x-axis
        fig.update_layout(
            xaxis_title="Date",
            yaxis_title="Products",
            xaxis=dict(tickformat="%m-%d-%Y"),
            showlegend=False,

        )

        # Display the figure in Streamlit
        st.plotly_chart(fig)

else:
    st.header('Upload Report')
    st.write('1. CLM Primary Sites Filter by Manger')
    st.markdown("Report on CLM(%s): " % REPORT_URL)
    st.write('2. Customized and enter RSM name.')
    st.write('3. Run report')
    st.write('4. Export to csv file.')