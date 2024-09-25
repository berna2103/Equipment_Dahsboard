import streamlit as st
import pandas as pd
import plotly.express as px
from geopy.geocoders import Nominatim
import nbformat



st.title('Region Overview')

df = pd.read_csv('./Equipment.csv')

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
df_clean[['latitude', 'longitude']] = df_clean['address'].apply(geocode_address)
df_clean.to_excel('output_excel_coordinates_missing.xlsx', sheet_name='data')

df_clean_no_coordinates = df_clean.dropna(subset=['latitude', 'longitude'])

df_clean_no_coordinates.to_excel('output_excel.xlsx', sheet_name='data')

# Plot points on a map using Plotly Express
fig = px.scatter_mapbox(df_clean_no_coordinates, 
                        lat="latitude", 
                        lon="longitude", 
                        hover_name="address",
                        
                        zoom=5, 
                        height=800)

# Set map style and layout
fig.update_layout(mapbox_style="open-street-map")

st.plotly_chart(fig)