# freakyluffy/client_flask/client_flask-d6698893f32401e64d38b73f4ecd0c16c7652afe/app.py
from flask import Flask, render_template, request, redirect, url_for
import pandas as pd
from urllib.parse import unquote
import os
import traceback
from pymongo import MongoClient

app = Flask(__name__)

# --- Connect to MongoDB ---
# Get the connection string from Render's environment variables
MONGO_URI = os.environ.get('MONGO_URI')
if not MONGO_URI:
    raise RuntimeError("MONGO_URI environment variable is not set.")

client = MongoClient(MONGO_URI)
db = client['hotel_dashboard_db'] # Database name
collection = db['reports']        # Collection name

# --- !!! IMPORTANT CONFIGURATION !!! ---
HOTEL_CAPACITY = {
    "The Hyde Dubai (HB8Y1)": 350,
}

def calculate_metrics(df, capacity_df):
    """Helper function to calculate metrics for any given subset of data."""
    if df.empty:
        return {'occ_percent': 0, 'occ_vs_ly': 0, 'forecast_rev': 0}
    merged_df = df.merge(capacity_df, on='Property Name', how='left')
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
    # Check if any documents exist in the collection
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
                # Add a column for the source filename to identify the data
                df['filename'] = file.filename
                # Convert DataFrame to a list of dictionaries to insert into MongoDB
                records = df.to_dict('records')
                if records:
                    collection.insert_many(records)
            except Exception as e:
                return f"An error occurred while processing '{file.filename}': {e}", 500
    
    return redirect(url_for('dashboard_index'))

@app.route('/dashboards')
def dashboard_index():
    """Shows a list of all dashboards by getting distinct filenames from MongoDB."""
    dashboard_names = collection.distinct('filename')
    return render_template('dashboards.html', dashboard_names=dashboard_names)

def load_df_from_mongo(filename):
    """Helper function to fetch data for a specific file and convert it to a DataFrame."""
    cursor = collection.find({'filename': filename})
    df = pd.DataFrame(list(cursor))
    if not df.empty:
        df['Occupancy Date'] = pd.to_datetime(df['Occupancy Date'])
    return df

@app.route('/summary/<path:filename>')
def summary(filename):
    """Displays the summary view for a file's data stored in MongoDB."""
    filename = unquote(filename)
    df = load_df_from_mongo(filename)
    if df.empty:
        return "Error: No data found for this dashboard.", 404
    
    # ... The rest of the logic is the same, as it operates on the DataFrame
    try:
        summary_cols = ['Occupancy On Books This Year', 'Booked Room Revenue This Year']
        for col in summary_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            else:
                return f"Error: Required column '{col}' not found in '{filename}'.", 400

        summary_data = df.groupby('Property Name').agg({
            'Occupancy On Books This Year': 'sum',
            'Booked Room Revenue This Year': 'sum'
        }).reset_index()

        summary_data['ADR On Books This Year'] = summary_data.apply(
            lambda row: (row['Booked Room Revenue This Year'] / row['Occupancy On Books This Year']) if row['Occupancy On Books This Year'] != 0 else 0,
            axis=1
        )
        return render_template('summary.html', summary_data=summary_data.to_dict(orient='records'), filename=filename)
    except Exception as e:
        return f"An error occurred creating the summary for '{filename}': {e}", 500

@app.route('/hotel/<path:filename>/<hotel_name>')
def hotel_detail(filename, hotel_name):
    """Displays the detailed view for a hotel from a specific file in MongoDB."""
    filename = unquote(filename)
    hotel_name = unquote(hotel_name)
    
    # Query MongoDB for the specific hotel within the specific file
    cursor = collection.find({'filename': filename, 'Property Name': hotel_name})
    hotel_df = pd.DataFrame(list(cursor))
    if hotel_df.empty: return "Error: Hotel data not found.", 404

    try:
        hotel_df['Occupancy Date'] = pd.to_datetime(hotel_df['Occupancy Date'])
        # (The rest of the data processing logic remains the same)
        numeric_cols = ['Occupancy On Books This Year', 'Occupancy On Books STLY', 'Forecasted Room Revenue This Year', 'Booked Room Revenue ST2Y']
        for col in numeric_cols:
            if col in hotel_df.columns:
                hotel_df[col] = pd.to_numeric(hotel_df[col], errors='coerce').fillna(0)
            else:
                hotel_df[col] = 0

        hotel_df['Month'] = hotel_df['Occupancy Date'].dt.strftime('%B')
        
        capacity_df = pd.DataFrame(list(HOTEL_CAPACITY.items()), columns=['Property Name', 'capacity_per_day'])
        date_range = pd.date_range(hotel_df['Occupancy Date'].min(), hotel_df['Occupancy Date'].max(), name='Occupancy Date')
        capacity_by_date = pd.MultiIndex.from_product([capacity_df['Property Name'], date_range], names=['Property Name', 'Occupancy Date']).to_frame(index=False)
        capacity_by_date = capacity_by_date.merge(capacity_df, on='Property Name').rename(columns={'capacity_per_day': 'capacity'})
        
        overall_totals = calculate_metrics(hotel_df, capacity_by_date)
        
        month_order = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
        hotel_df['Month'] = pd.Categorical(hotel_df['Month'], categories=month_order, ordered=True)
        hotel_df = hotel_df.sort_values(by=['Month', 'Occupancy Date'])

        monthly_data = {}
        for name, group in hotel_df.groupby('Month', observed=False):
            month_capacity_df = capacity_by_date[capacity_by_date['Occupancy Date'].dt.strftime('%B') == name]
            monthly_data[name] = {'totals': calculate_metrics(group, month_capacity_df), 'records': group.to_dict(orient='records')}
        
        detail_data = {'name': hotel_name, 'totals': overall_totals, 'monthly_data': monthly_data, 'columns': hotel_df.columns.tolist()}
        return render_template('dashboard.html', data=detail_data, filename=filename)

    except Exception as e:
        return f"An error occurred generating the detail view: {e}\n<pre>{traceback.format_exc()}</pre>", 500

@app.route('/reset')
def reset_data():
    """Clears all dashboards by deleting all documents in the collection."""
    collection.delete_many({})
    return redirect(url_for('index'))
# if __name__ == '__main__':
#     # For local development
#     app.run(debug=True)

if __name__ == '__main__':
    # For Render deployment
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
