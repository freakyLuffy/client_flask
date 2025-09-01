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
        return {'occ_percent': 0, 'occ_vs_ly': 0, 'forecast_rev': 0}
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
    data_exists = collection.count_documents({}) > 0
    return render_template('index.html', data_exists=data_exists)

@app.route('/process', methods=['POST'])
def process_file():
    """Clears old data, reads new files, saves to MongoDB, and redirects to the success page."""
    collection.delete_many({})
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
    return redirect(url_for('success'))

@app.route('/success')
def success():
    """Shows a success page with a button to download the report."""
    return render_template('success.html')

def flatten_hierarchy_for_excel(item, level=0, output_list=None):
    """Recursively flattens data and includes the hierarchy level for each row."""
    if output_list is None:
        output_list = []
    
    output_list.append((
        {
            '': item['name'],
            'Occupancy on Books': f"{item['metrics']['occ_percent']:.1f}%",
            'Occupancy vs. Last Year': f"{item['metrics']['occ_vs_ly']:+.1f}%",
            'Total Forecasted Room Rev': f"${item['metrics']['forecast_rev']:,.0f}"
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
        # --- (Data processing logic remains the same) ---
        df['Occupancy Date'] = pd.to_datetime(df['Occupancy Date'])
        numeric_cols = ['Occupancy On Books This Year', 'Occupancy On Books STLY', 'Forecasted Room Revenue This Year']
        for col in numeric_cols:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
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
            
            # *** NEW: Tell Excel to open with the groups collapsed ***
            # This ensures the summary (parent) row is shown above its children
            worksheet.sheet_properties.outlinePr.summaryBelow = False

            # Auto-adjust column widths
            for column_cells in worksheet.columns:
                length = max(len(str(cell.value)) for cell in column_cells)
                worksheet.column_dimensions[column_cells[0].column_letter].width = length + 2
        
        output.seek(0)
        
        return send_file(output, as_attachment=True, download_name='Portfolio_Report_Grouped.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    except Exception as e:
        return f"An error occurred: {e}\n<pre>{traceback.format_exc()}</pre>", 500

@app.route('/reset')
def reset_data():
    """Clears all data from the MongoDB collection."""
    collection.delete_many({})
    return redirect(url_for('index'))

# if __name__ == '__main__':
#     app.run(debug=True)
# if __name__ == '__main__':
#     # For Render deployment
#     port = int(os.environ.get('PORT', 5001))
#     app.run(host='0.0.0.0', port=port, debug=False)

# if __name__ == '__main__':
#     # For Render deployment
#     port = int(os.environ.get('PORT', 5000))
#     app.run(host='0.0.0.0', port=port, debug=False)

# if __name__ == '__main__':
#     app.run(debug=True)
# if __name__ == '__main__':
#     # For Render deployment
#     port = int(os.environ.get('PORT', 5001))
#     app.run(host='0.0.0.0', port=port, debug=False)

# if __name__ == '__main__':
#     # For Render deployment
#     port = int(os.environ.get('PORT', 5000))
#     app.run(host='0.0.0.0', port=port, debug=False)

if __name__ == '__main__':
    # For Render deployment
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
