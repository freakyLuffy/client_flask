# freakyluffy/client_flask/client_flask-d6698893f32401e64d38b73f4ecd0c16c7652afe/app.py
from flask import Flask, render_template, request, redirect, url_for
import pandas as pd
from urllib.parse import unquote

app = Flask(__name__)

# --- In-memory data storage ---
dashboards_store = {}

@app.route('/')
def index():
    """Renders the upload page."""
    return render_template('index.html', dashboards_exist=bool(dashboards_store))

@app.route('/process', methods=['POST'])
def process_file():
    """Processes uploaded files and adds them as new, separate dashboards."""
    global dashboards_store
    
    uploaded_files = request.files.getlist("file")
    if not uploaded_files or all(f.filename == '' for f in uploaded_files):
        return "Error: No files were selected.", 400

    for file in uploaded_files:
        if file and (file.filename.lower().endswith('.xlsx') or file.filename.lower().endswith('.xls')):
            try:
                df = pd.read_excel(file, engine='openpyxl')
                df.columns = df.columns.str.strip()
                df['Occupancy Date'] = pd.to_datetime(df['Occupancy Date'])
                dashboards_store[file.filename] = df
            except Exception as e:
                return f"An error occurred while processing '{file.filename}': {e}", 500
    
    return redirect(url_for('dashboard_index'))

@app.route('/dashboards')
def dashboard_index():
    """Shows a list of all uploaded dashboards."""
    return render_template('dashboards.html', dashboard_names=list(dashboards_store.keys()))

@app.route('/summary/<path:filename>')
def summary(filename):
    """Displays the summary view for a SINGLE dashboard."""
    filename = unquote(filename)
    df = dashboards_store.get(filename)
    if df is None: return "Error: Dashboard not found.", 404
    
    try:
        df_copy = df.copy()
        summary_cols = ['Occupancy On Books This Year', 'Booked Room Revenue This Year']
        for col in summary_cols:
            if col in df_copy.columns:
                df_copy[col] = pd.to_numeric(df_copy[col], errors='coerce').fillna(0)
            else:
                return f"Error: Required column '{col}' not found in '{filename}'.", 400

        summary_data = df_copy.groupby('Property Name').agg({
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
    """Displays the detailed view for a specific hotel from a specific dashboard."""
    filename = unquote(filename)
    hotel_name = unquote(hotel_name)
    df = dashboards_store.get(filename)
    if df is None: return "Error: Dashboard not found.", 404

    try:
        hotel_df = df[df['Property Name'] == hotel_name].copy()
        if hotel_df.empty: return "Error: Hotel not found.", 404

        sum_cols = ['Booked Room Revenue ST2Y', 'Forecasted Room Revenue This Year', 'Occupancy Forecast This Year']
        for col in sum_cols:
            if col in hotel_df.columns:
                hotel_df[col] = pd.to_numeric(hotel_df[col], errors='coerce').fillna(0)
            else:
                hotel_df[col] = 0

        hotel_df['Month'] = hotel_df['Occupancy Date'].dt.strftime('%B')
        overall_totals = {
            'total_revenue_st2y': hotel_df['Booked Room Revenue ST2Y'].sum(),
            'total_forecast_revenue': hotel_df['Forecasted Room Revenue This Year'].sum(),
            'total_occupancy_forecast': hotel_df['Occupancy Forecast This Year'].sum()
        }
        
        # Ensures months are always sorted chronologically
        month_order = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
        hotel_df['Month'] = pd.Categorical(hotel_df['Month'], categories=month_order, ordered=True)
        hotel_df = hotel_df.sort_values(by=['Month', 'Occupancy Date'])

        monthly_data = {}
        grouped = hotel_df.groupby('Month', observed=False)
        for name, group in grouped:
            monthly_totals = {
                'month_revenue_st2y': group['Booked Room Revenue ST2Y'].sum(),
                'month_forecast_revenue': group['Forecasted Room Revenue This Year'].sum(),
                'month_occupancy_forecast': group['Occupancy Forecast This Year'].sum()
            }
            monthly_data[name] = {'totals': monthly_totals, 'records': group.to_dict(orient='records')}
        
        detail_data = {
            'name': hotel_name,
            'totals': overall_totals,
            'monthly_data': monthly_data,
            'columns': hotel_df.columns.tolist()
        }
        return render_template('results.html', hotel_data=detail_data, filename=filename)
    except Exception as e:
        return f"An error occurred generating the detail view: {e}", 500

@app.route('/reset')
def reset_data():
    """Clears all stored dashboards."""
    global dashboards_store
    dashboards_store = {}
    return redirect(url_for('index'))

if __name__ == '__main__':
    # For Render deployment
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
