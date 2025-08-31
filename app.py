# freakyluffy/client_flask/client_flask-d6698893f32401e64d38b73f4ecd0c16c7652afe/app.py
from flask import Flask, render_template, request, redirect, url_for
import pandas as pd
import numpy as np
import os

app = Flask(__name__)

# --- In-memory data storage ---
# This will hold the combined data from all uploaded files.
df_store = None

# --- !!! IMPORTANT CONFIGURATION !!! ---
# You MUST update this dictionary with the total number of available rooms for each hotel.
# The key is the 'Property Name' from your Excel file, and the value is its room capacity.
HOTEL_CAPACITY = {
    "The Hyde Dubai (HB8Y1)": 350,
    # Add other hotels here, for example:
    # "Hotel 2": 275,
    # "Hotel 3": 410,
}

def calculate_metrics(df, capacity_df):
    """Helper function to calculate metrics for any given subset of data."""
    if df.empty:
        return {'occ_percent': 0, 'occ_vs_ly': 0, 'forecast_rev': 0}

    # Merge with capacity to get total available rooms for the period
    merged_df = df.merge(capacity_df, on='Property Name', how='left')
    total_available_rooms = merged_df['capacity'].sum()

    # Calculate sums
    occ_ty = merged_df['Occupancy On Books This Year'].sum()
    occ_ly = merged_df['Occupancy On Books STLY'].sum()
    
    # Calculate metrics
    metrics = {
        'occ_percent': (occ_ty / total_available_rooms) * 100 if total_available_rooms else 0,
        'occ_vs_ly': ((occ_ty - occ_ly) / occ_ly) * 100 if occ_ly else 0,
        'forecast_rev': merged_df['Forecasted Room Revenue This Year'].sum()
    }
    return metrics

@app.route('/')
def index():
    """Renders the upload page."""
    return render_template('index.html', dashboards_exist=(df_store is not None))

@app.route('/process', methods=['POST'])
def process_file():
    """Processes uploaded files and combines them into one dataset."""
    global df_store
    
    uploaded_files = request.files.getlist("file")
    if not uploaded_files or all(f.filename == '' for f in uploaded_files):
        return "Error: No files were selected.", 400

    data_frames = []
    for file in uploaded_files:
        if file and (file.filename.lower().endswith('.xlsx') or file.filename.lower().endswith('.xls')):
            try:
                df = pd.read_excel(file, engine='openpyxl')
                data_frames.append(df)
            except Exception as e:
                return f"An error occurred while processing '{file.filename}': {e}", 500
    
    if data_frames:
        df_store = pd.concat(data_frames, ignore_index=True)
    
    return redirect(url_for('main_dashboard'))

@app.route('/dashboard')
def main_dashboard():
    """Builds the hierarchical data structure and renders the main dashboard."""
    if df_store is None:
        return redirect(url_for('index'))

    try:
        df = df_store.copy()
        
        # --- 1. Data Cleaning and Preparation ---
        df.columns = df.columns.str.strip()
        df['Occupancy Date'] = pd.to_datetime(df['Occupancy Date'])
        
        numeric_cols = [
            'Occupancy On Books This Year', 'Occupancy On Books STLY', 
            'Forecasted Room Revenue This Year'
        ]
        for col in numeric_cols:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            
        df['Month'] = df['Occupancy Date'].dt.strftime('%B')
        df['Day'] = df['Occupancy Date'].dt.strftime('%b %d')

        # Create a DataFrame for hotel capacities
        capacity_df = pd.DataFrame(list(HOTEL_CAPACITY.items()), columns=['Property Name', 'capacity_per_day'])
        
        # Expand capacity to cover all dates in the data
        date_range = pd.date_range(df['Occupancy Date'].min(), df['Occupancy Date'].max(), name='Occupancy Date')
        capacity_by_date = pd.MultiIndex.from_product([capacity_df['Property Name'], date_range], names=['Property Name', 'Occupancy Date']).to_frame(index=False)
        capacity_by_date = capacity_by_date.merge(capacity_df, on='Property Name').rename(columns={'capacity_per_day': 'capacity'})
        
        # --- 2. Build the Hierarchical Structure ---
        portfolio_data = {'name': 'All Hotels (Portfolio Total)', 'children': []}
        
        # Calculate Portfolio Totals
        portfolio_data['metrics'] = calculate_metrics(df, capacity_by_date)
        
        # Level 1: Hotels
        for hotel_name, hotel_group in df.groupby('Property Name'):
            hotel_capacity_df = capacity_by_date[capacity_by_date['Property Name'] == hotel_name]
            hotel_data = {
                'name': hotel_name,
                'metrics': calculate_metrics(hotel_group, hotel_capacity_df),
                'children': []
            }
            
            # Level 2: Months
            month_order = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
            hotel_group['Month'] = pd.Categorical(hotel_group['Month'], categories=month_order, ordered=True)
            
            for month_name, month_group in hotel_group.sort_values('Month').groupby('Month', observed=False):
                month_capacity_df = hotel_capacity_df[hotel_capacity_df['Occupancy Date'].dt.strftime('%B') == month_name]
                month_data = {
                    'name': month_name,
                    'metrics': calculate_metrics(month_group, month_capacity_df),
                    'children': []
                }
                
                # Level 3: Business Views (e.g., Groups, Contracts)
                for view_name, view_group in month_group.groupby('Business View'):
                    view_capacity_df = month_capacity_df # Capacity doesn't change by view
                    view_data = {
                        'name': view_name,
                        'metrics': calculate_metrics(view_group, view_capacity_df),
                        'children': []
                    }
                    
                    # Level 4: Days
                    for day_name, day_group in view_group.sort_values('Occupancy Date').groupby('Day'):
                        day_capacity_df = view_capacity_df[view_capacity_df['Occupancy Date'].dt.strftime('%b %d') == day_name]
                        day_data = {
                            'name': day_name,
                            'metrics': calculate_metrics(day_group, day_capacity_df),
                            'children': [] # Days are the lowest level
                        }
                        view_data['children'].append(day_data)
                    month_data['children'].append(view_data)
                hotel_data['children'].append(month_data)
            portfolio_data['children'].append(hotel_data)
            
        return render_template('dashboard.html', data=portfolio_data)

    except Exception as e:
        # Provide a more detailed error message for debugging
        import traceback
        return f"An error occurred while building the dashboard: {e}\n<pre>{traceback.format_exc()}</pre>", 500

@app.route('/reset')
def reset_data():
    """Clears all stored data."""
    global df_store
    df_store = None
    return redirect(url_for('index'))

# if __name__ == '__main__':
#     # For local development
#     app.run(debug=True)

if __name__ == '__main__':
    # For Render deployment
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
