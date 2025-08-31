from flask import Flask, render_template, request, redirect, url_for
import pandas as pd
import os
import traceback
from pymongo import MongoClient
from dotenv import load_dotenv

app = Flask(__name__)
load_dotenv() # Loads environment variables from a .env file for local development

# --- Connect to MongoDB ---
MONGO_URI = os.environ.get('MONGO_URI')
if not MONGO_URI:
    raise RuntimeError("MONGO_URI environment variable is not set.")
client = MongoClient(MONGO_URI)
db = client['hotel_dashboard_db']
collection = db['reports']

# --- !!! IMPORTANT CONFIGURATION !!! ---
HOTEL_CAPACITY = {
    "The Hyde Dubai (HB8Y1)": 350,
    # Add other hotels here, for example:
    # "Another Hotel": 275,
}

def calculate_metrics(df, capacity_df):
    """Helper function to calculate metrics for any given subset of data."""
    if df.empty:
        return {'occ_percent': 0, 'occ_vs_ly': 0, 'forecast_rev': 0}

    # Merge with capacity data to get total available rooms for the period
    merged_df = df.merge(capacity_df, on=['Property Name', 'Occupancy Date'], how='left')
    total_available_rooms = merged_df['capacity'].sum()

    occ_ty = merged_df['Occupancy On Books This Year'].sum()
    occ_ly = merged_df['Occupancy On Books STLY'].sum()
    
    return {
        'occ_percent': (occ_ty / total_available_rooms) * 100 if total_available_rooms else 0,
        'occ_vs_ly': ((occ_ty - occ_ly) / occ_ly) * 100 if occ_ly else 0,
        'forecast_rev': merged_df['Forecasted Room Revenue This Year'].sum()
    }

@app.route('/')
def index():
    """Renders the upload page."""
    dashboards_exist = collection.count_documents({}) > 0
    return render_template('index.html', dashboards_exist=dashboards_exist)

@app.route('/process', methods=['POST'])
def process_file():
    """Reads uploaded files and saves their data to MongoDB."""
    uploaded_files = request.files.getlist("file")
    if not uploaded_files or all(f.filename == '' for f in uploaded_files):
        return "Error: No files were selected.", 400

    for file in uploaded_files:
        if file and (file.filename.lower().endswith('.xlsx') or file.filename.lower().endswith('.xls')):
            try:
                df = pd.read_excel(file, engine='openpyxl')
                df.columns = df.columns.str.strip()
                records = df.to_dict('records')
                if records:
                    collection.insert_many(records)
            except Exception as e:
                return f"An error occurred while processing '{file.filename}': {e}", 500
    
    # Redirect to the main unified dashboard after upload
    return redirect(url_for('main_dashboard'))

@app.route('/dashboard')
def main_dashboard():
    """Builds the single, unified hierarchical dashboard from all data in MongoDB."""
    cursor = collection.find({})
    df = pd.DataFrame(list(cursor))
    if df.empty:
        return redirect(url_for('index'))

    try:
        # --- 1. Data Cleaning and Preparation ---
        df['Occupancy Date'] = pd.to_datetime(df['Occupancy Date'])
        numeric_cols = ['Occupancy On Books This Year', 'Occupancy On Books STLY', 'Forecasted Room Revenue This Year']
        for col in numeric_cols:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            
        df['Month'] = df['Occupancy Date'].dt.strftime('%B')
        df['Day'] = df['Occupancy Date'].dt.strftime('%b %d')

        # --- 2. Build Capacity DataFrame ---
        capacity_df = pd.DataFrame(list(HOTEL_CAPACITY.items()), columns=['Property Name', 'capacity_per_day'])
        # Create a full date range covering all data
        date_range = pd.date_range(df['Occupancy Date'].min(), df['Occupancy Date'].max(), name='Occupancy Date')
        # Create a row for each hotel for each day in the date range
        capacity_by_date = pd.MultiIndex.from_product([capacity_df['Property Name'], date_range], names=['Property Name', 'Occupancy Date']).to_frame(index=False)
        capacity_by_date = capacity_by_date.merge(capacity_df, on='Property Name').rename(columns={'capacity_per_day': 'capacity'})
        
        # --- 3. Build the Hierarchical Structure ---
        portfolio_data = {'name': 'All Hotels (Portfolio Total)', 'children': []}
        portfolio_data['metrics'] = calculate_metrics(df, capacity_by_date)
        
        # Level 1: Hotels
        for hotel_name, hotel_group in df.groupby('Property Name'):
            hotel_data = {'name': hotel_name, 'metrics': calculate_metrics(hotel_group, capacity_by_date), 'children': []}
            
            # Level 2: Months
            month_order = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
            hotel_group['Month'] = pd.Categorical(hotel_group['Month'], categories=month_order, ordered=True)
            
            for month_name, month_group in hotel_group.sort_values('Month').groupby('Month', observed=False):
                month_data = {'name': month_name, 'metrics': calculate_metrics(month_group, capacity_by_date), 'children': []}
                
                # Level 3: Business Views
                for view_name, view_group in month_group.groupby('Business View'):
                    view_data = {'name': view_name, 'metrics': calculate_metrics(view_group, capacity_by_date), 'children': []}
                    
                    # Level 4: Days
                    for day_name, day_group in view_group.sort_values('Occupancy Date').groupby('Day'):
                        day_data = {'name': day_name, 'metrics': calculate_metrics(day_group, capacity_by_date), 'children': []}
                        view_data['children'].append(day_data)
                    month_data['children'].append(view_data)
                hotel_data['children'].append(month_data)
            portfolio_data['children'].append(hotel_data)
            
        return render_template('dashboard.html', data=portfolio_data)

    except Exception as e:
        return f"An error occurred while building the dashboard: {e}\n<pre>{traceback.format_exc()}</pre>", 500

@app.route('/reset')
def reset_data():
    """Clears all data from the MongoDB collection."""
    collection.delete_many({})
    return redirect(url_for('index'))

# if __name__ == '__main__':
#     # For local development
#     app.run(debug=True)

# if __name__ == '__main__':
#     # For Render deployment
#     port = int(os.environ.get('PORT', 5000))
#     app.run(host='0.0.0.0', port=port, debug=False)
# if __name__ == '__main__':
#     # For local development
#     app.run(debug=True)

if __name__ == '__main__':
    # For Render deployment
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
