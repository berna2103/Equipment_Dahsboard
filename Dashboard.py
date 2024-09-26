import streamlit as st
import pandas as pd
import os
import plotly.express as px
from geopy.geocoders import Nominatim
import plotly.graph_objects as go
from datetime import datetime

st.set_page_config(layout="wide")

excel_file_with_coordinates = "./data/output_excel.xlsx"
df = pd.DataFrame()
st.title('Region Overview')


if os.path.exists(excel_file_with_coordinates):
    df = pd.read_excel("./data/output_excel.xlsx")
else:
    df = pd.read_csv("./data/Equipment_Bernardo.csv")

df.rename(
    columns= {
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
        location = geolocator.geocode(address, timeout=5)
        
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

# Region Metrics
head_count = len(df_clean['secondary_fse'].unique())
average_equipment_age = df_clean['device_age'].mean()
median_equipment_age = df_clean['device_age'].median()

st.dataframe(df_clean)

container = st.container(border=True)

with container:
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric(label= 'Head Count', value= head_count)
    with col2:
        st.metric(label='Average Device Age', value= f'{round(average_equipment_age, 2)} yrs')
    with col3:
        st.metric(label='Total IPs', value=(len(df_clean['device_age'])))

############### PLOT DEVICE AGE #################################################
sorted_data = df_clean.sort_values(by='device_age')
# Create the bar chart with custom bar width
device_age_fig = go.Figure(go.Bar(
    y=sorted_data['device_age'],
    x=sorted_data['ip'],
    # width=[0.9]*len(sorted_data),  # Set a fixed width for all bars
      # horizontal bar chart (optional)
))

st.plotly_chart(device_age_fig)
st.write(sorted_data)

############END OF PLOT DEVICE AGE ##################################################

############ Plot points on a MAP using Plotly Express #############################
fig = px.scatter_mapbox(df_clean, 
                        lat="latitude", 
                        lon="longitude", 
                        hover_name="location",
                        
                        zoom=5, 
                        height=800)

# Set map style and layout
fig.update_layout(mapbox_style="open-street-map")

st.plotly_chart(fig)
#############END of MAP ##############################################################


############### PLOT EOL and EOGL    #################################################
# Convert the EOL and EOGL columns to datetime format
df_clean['eol'] = pd.to_datetime(df_clean['eol'], errors='coerce')
df_clean['eogs'] = pd.to_datetime(df_clean['eogs'], errors='coerce')


# Filter out rows where either EOL or EOGL is missing
df_filtered = df_clean.dropna(subset=['eol', 'eogs'])
df_filtered['serial_number'] = df_clean['ip'].apply(lambda x: x.split('/')[2])
df_filtered['ip_type'] = df_clean['ip'].apply(lambda x: x.split('/')[0])

# Create a Plotly figure
fig = go.Figure()

# Add scatter plots for EOL and EOGL
fig.add_trace(go.Scatter(
    x=df_filtered['ip'],
    y=df_filtered['eol'],
    mode='markers',
    name='End of Life (EOL)',
    marker=dict(color='red')
))

fig.add_trace(go.Scatter(
    x=df_filtered['ip'],
    y=df_filtered['eogs'],
    mode='markers',
    name='End of Guaranteed Support (EOGS)',
    marker=dict(color='blue')
))

# Add a vertical line for today's date
today = datetime.today()
# Add an annotation for today's date instead of using a line
fig.add_trace(go.Scatter(
    x=[df_filtered['ip'].iloc[-1]],  # Position annotation near the last product
    y=[today],  # Today's date on the y-axis
    mode='lines+text',
    text=['Today'],
    showlegend=False,
    textposition="bottom right"
))


# Update layout
fig.update_layout(
    title="Installed Products with EOL and EOGL",
    xaxis_title="Products",
    yaxis_title="Date",
    xaxis=dict(tickvals=df_filtered['ip']),
)

# Display the figure in Streamlit
st.plotly_chart(fig)

############### PLOT EOL and EOGL    #################################################

# Get today's date as a pandas datetime object
today = pd.to_datetime(datetime.today())

# Filter out rows where either EOL or EOGL is missing
df_filtered = df_clean.dropna(subset=['eol', 'eogs'])

# Ensure that the 'eogs' column (End of Guaranteed Support) is a datetime object
df_filtered['eogs'] = pd.to_datetime(df_filtered['eogs'], errors='coerce')

# Create the DataFrame for the timeline, ensuring all datetime fields are in pandas datetime format
df_timeline = pd.DataFrame({
    'Product': df_filtered['ip'],
    'Start': [today] * len(df_filtered),  # Create a Start column with today's date for each row
    'End': df_filtered['eogs']  # EOGL as the end of guaranteed support
})


# Ensure there are no missing datetime values in 'End' after conversion
df_timeline = df_timeline.dropna(subset=['End'])

# Create the timeline chart using Plotly Express
fig = px.timeline(df_timeline, x_start='Start', x_end='End', y='Product', title="Product Support Timeline")

# # Add today's date as a vertical line using the add_vline method
# fig.add_vline(x=today, line=dict(color="green", dash="dash"), annotation_text="Today", annotation_position="top")

# Update layout to show date formatting on the x-axis
fig.update_layout(
    xaxis_title="Date",
    yaxis_title="Products",
    xaxis=dict(tickformat="%Y-%m-%d"),
    showlegend=False
)

# Display the figure in Streamlit
st.plotly_chart(fig)