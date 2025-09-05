from flask import Flask, render_template, request, redirect, url_for, send_file
import pandas as pd
import os
import traceback
from pymongo import MongoClient
from dotenv import load_dotenv
import io

app = Flask(__name__)
load_dotenv()

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
}

from flask import Flask, render_template, request, redirect, url_for, send_file
import pandas as pd
import os
import traceback
from pymongo import MongoClient
from dotenv import load_dotenv
import io

app = Flask(__name__)
load_dotenv()

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
}

def calculate_metrics(df, capacity_df):
    """Helper function to calculate metrics for any given subset of data."""
    if df.empty:
        return {
            'occ_ty': 0, 'occ_ly': 0, 'occ_diff': 0, 'occ_st2y': 0, 'occ_forecast': 0,
            'booked_rev_ty': 0, 'booked_rev_ly': 0, 'booked_rev_diff': 0, 'booked_rev_st2y': 0,
            'forecast_rev_ty': 0
        }
    
    # Sum all the required metrics
    occ_ty = df['Occupancy On Books This Year'].sum()
    occ_ly = df['Occupancy On Books STLY'].sum()
    occ_st2y = df['Occupancy On Books ST2Y'].sum()
    occ_forecast = df['Occupancy Forecast This Year'].sum()
    booked_rev_ty = df['Booked Room Revenue This Year'].sum()
    booked_rev_ly = df['Booked Room Revenue STLY'].sum()
    booked_rev_st2y = df['Booked Room Revenue ST2Y'].sum()
    forecast_rev_ty = df['Forecasted Room Revenue This Year'].sum()
    
    return {
        'occ_ty': occ_ty,
        'occ_ly': occ_ly,
        'occ_diff': occ_ty - occ_ly,
        'occ_st2y': occ_st2y,
        'occ_forecast': occ_forecast,
        'booked_rev_ty': booked_rev_ty,
        'booked_rev_ly': booked_rev_ly,
        'booked_rev_diff': booked_rev_ty - booked_rev_ly,
        'booked_rev_st2y': booked_rev_st2y,
        'forecast_rev_ty': forecast_rev_ty
    }

@app.route('/')
def index():
    """Renders the upload page."""
    data_exists = collection.count_documents({}) > 0
    return render_template('index.html', data_exists=data_exists)

@app.route('/process', methods=['POST'])
def process_file():
    """Clears old data, reads new files, saves to MongoDB, and redirects to the success page."""
    collection.delete_many({})
    # --- ONLY CHANGE IS HERE ---
    # Use .getlist() to handle one or more uploaded files
    uploaded_files = request.files.getlist("file")
    
    if not uploaded_files or all(f.filename == '' for f in uploaded_files):
        return "Error: No files were selected.", 400
        
    # The rest of the function works as is, because it already loops through the files
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
    return redirect(url_for('success'))

@app.route('/success')
def success():
    """Shows a success page with a button to download the report."""
    return render_template('success.html')

def flatten_hierarchy_for_excel(item, level=0, output_list=None):
    """Recursively flattens data and includes the hierarchy level for each row with 12 columns."""
    if output_list is None:
        output_list = []
    
    metrics = item['metrics']
    output_list.append((
        {
            '': item['name'],  # Column A
            'Occupancy On Books This Year': metrics['occ_ty'],  # Column B
            'Occupancy On Books STLY': metrics['occ_ly'],  # Column C
            'Difference (TY-LY)': metrics['occ_diff'],  # Column D
            'Occupancy On Books ST2Y': metrics['occ_st2y'],  # Column E
            'Occupancy Forecast This Year': metrics['occ_forecast'],  # Column F
            'Empty Column 1': '',  # Column G (empty)
            'Booked Room Revenue This Year': metrics['booked_rev_ty'],  # Column H
            'Booked Room Revenue STLY': metrics['booked_rev_ly'],  # Column I
            'Revenue Difference (TY-LY)': metrics['booked_rev_diff'],  # Column J
            'Booked Room Revenue ST2Y': metrics['booked_rev_st2y'],  # Column K
            'Forecasted Room Revenue This Year': metrics['forecast_rev_ty'],  # Column L
            'Empty Column 2': ''  # Column M (empty)
        },
        level
    ))
    if 'children' in item:
        for child in item['children']:
            flatten_hierarchy_for_excel(child, level + 1, output_list)
    return output_list

@app.route('/download-report')
def download_report():
    """Generates the hierarchical data and serves it as a downloadable Excel file with collapsible rows."""
    cursor = collection.find({})
    df = pd.DataFrame(list(cursor))
    if df.empty:
        return "No data found to generate a report.", 404

    try:
        # --- Data processing logic ---
        df['Occupancy Date'] = pd.to_datetime(df['Occupancy Date'])
        
        # Define all numeric columns that need to be processed
        numeric_cols = [
            'Occupancy On Books This Year', 'Occupancy On Books STLY', 'Occupancy On Books ST2Y',
            'Booked Room Revenue This Year', 'Booked Room Revenue STLY', 'Booked Room Revenue ST2Y',
            'Forecasted Room Revenue This Year', 'Occupancy Forecast This Year'
        ]
        
        for col in numeric_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            else:
                df[col] = 0  # Add missing columns with zero values
        
        df['Month'] = df['Occupancy Date'].dt.strftime('%B')
        df['Day'] = df['Occupancy Date'].dt.strftime('%b %d')
        
        capacity_df = pd.DataFrame(list(HOTEL_CAPACITY.items()), columns=['Property Name', 'capacity_per_day'])
        date_range = pd.date_range(df['Occupancy Date'].min(), df['Occupancy Date'].max(), name='Occupancy Date')
        capacity_by_date = pd.MultiIndex.from_product([capacity_df['Property Name'], date_range], names=['Property Name', 'Occupancy Date']).to_frame(index=False)
        capacity_by_date = capacity_by_date.merge(capacity_df, on='Property Name').rename(columns={'capacity_per_day': 'capacity'})
        
        portfolio_data = {'name': 'All Hotels (Portfolio Total)', 'children': []}
        portfolio_data['metrics'] = calculate_metrics(df, capacity_by_date)
        
        for hotel_name, hotel_group in df.groupby('Property Name'):
            hotel_data = {'name': hotel_name, 'metrics': calculate_metrics(hotel_group, capacity_by_date), 'children': []}
            month_order = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
            hotel_group['Month'] = pd.Categorical(hotel_group['Month'], categories=month_order, ordered=True)
            for month_name, month_group in hotel_group.sort_values('Month').groupby('Month', observed=False):
                month_data = {'name': month_name, 'metrics': calculate_metrics(month_group, capacity_by_date), 'children': []}
                for view_name, view_group in month_group.groupby('Business View'):
                    view_data = {'name': view_name, 'metrics': calculate_metrics(view_group, capacity_by_date), 'children': []}
                    for day_name, day_group in view_group.sort_values('Occupancy Date').groupby('Day'):
                        day_data = {'name': day_name, 'metrics': calculate_metrics(day_group, capacity_by_date), 'children': []}
                        view_data['children'].append(day_data)
                    month_data['children'].append(view_data)
                hotel_data['children'].append(month_data)
            portfolio_data['children'].append(hotel_data)

        # --- Generate Excel with Collapsible Rows ---
        flat_data_with_levels = flatten_hierarchy_for_excel(portfolio_data)
        report_data = [item[0] for item in flat_data_with_levels]
        report_df = pd.DataFrame(report_data)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            report_df.to_excel(writer, index=False, sheet_name='Portfolio_Report')
            worksheet = writer.sheets['Portfolio_Report']

            # Set row grouping levels
            for i, (_, level) in enumerate(flat_data_with_levels):
                if level > 0:
                    worksheet.row_dimensions[i + 2].outline_level = level
            
            # This ensures the summary (parent) row is shown above its children
            worksheet.sheet_properties.outlinePr.summaryBelow = False

            # Auto-adjust column widths
            for column_cells in worksheet.columns:
                length = max(len(str(cell.value)) for cell in column_cells)
                worksheet.column_dimensions[column_cells[0].column_letter].width = max(length + 2, 15)
        
        output.seek(0)
        
        return send_file(output, as_attachment=True, download_name='Portfolio_Report_Grouped.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    except Exception as e:
        return f"An error occurred: {e}\n<pre>{traceback.format_exc()}</pre>", 500

@app.route('/reset')
def reset_data():
    """Clears all data from the MongoDB collection."""
    collection.delete_many({})
    return redirect(url_for('index'))


if __name__ == '__main__':
    # For Render deployment
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
