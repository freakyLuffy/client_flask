# freakyluffy/client_flask/client_flask-d6698893f32401e64d38b73f4ecd0c16c7652afe/app.py
from flask import Flask, render_template, request, redirect, url_for
import pandas as pd

app = Flask(__name__)

# --- In-memory data storage ---
df_store = None

@app.route('/')
def index():
    """Renders the main page with the file upload form."""
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process_file():
    """
    Handles the uploaded file, loads it into the in-memory store,
    and redirects to the summary view.
    """
    global df_store
    
    if 'file' not in request.files or not request.files['file'].filename:
        return "Error: No file was selected.", 400

    file = request.files['file']
    if not (file.filename.lower().endswith('.xlsx') or file.filename.lower().endswith('.xls')):
        return "Error: Please upload a valid Excel file (.xlsx or .xls).", 400

    try:
        df_store = pd.read_excel(file, engine='openpyxl')
        df_store.columns = df_store.columns.str.strip()
        df_store['Occupancy Date'] = pd.to_datetime(df_store['Occupancy Date'])
        return redirect(url_for('summary'))
    except Exception as e:
        return f"An error occurred while processing the file: {e}", 500

@app.route('/summary')
def summary():
    """Displays the high-level summary table of all hotels."""
    if df_store is None:
        return redirect(url_for('index'))
    
    try:
        summary_cols = ['Occupancy On Books This Year', 'Booked Room Revenue This Year']
        for col in summary_cols:
            if col in df_store.columns:
                df_store[col] = pd.to_numeric(df_store[col], errors='coerce').fillna(0)
            else:
                return f"Error: Required column '{col}' not found in the file.", 400

        summary_data = df_store.groupby('Property Name').agg({
            'Occupancy On Books This Year': 'sum',
            'Booked Room Revenue This Year': 'sum'
        }).reset_index()

        summary_data['ADR On Books This Year'] = summary_data.apply(
            lambda row: (row['Booked Room Revenue This Year'] / row['Occupancy On Books This Year']) 
                        if row['Occupancy On Books This Year'] != 0 else 0,
            axis=1
        )
        return render_template('summary.html', summary_data=summary_data.to_dict(orient='records'))
    except Exception as e:
        return f"An error occurred while creating the summary: {e}", 500

@app.route('/hotel/<hotel_name>')
def hotel_detail(hotel_name):
    """Displays the detailed, monthly breakdown for a single, selected hotel."""
    if df_store is None:
        return redirect(url_for('index'))

    try:
        hotel_df = df_store[df_store['Property Name'] == hotel_name].copy()
        if hotel_df.empty:
            return "Error: Hotel not found.", 404

        sum_cols = ['Booked Room Revenue ST2Y', 'Forecasted Room Revenue This Year', 'Occupancy Forecast This Year']
        for col in sum_cols:
            if col in hotel_df.columns:
                hotel_df[col] = pd.to_numeric(hotel_df[col], errors='coerce').fillna(0)
            else:
                hotel_df[col] = 0

        hotel_df['Month'] = hotel_df['Occupancy Date'].dt.strftime('%B')
        
        # Calculate overall totals for the hotel
        overall_totals = {
            'total_revenue_st2y': hotel_df['Booked Room Revenue ST2Y'].sum(),
            'total_forecast_revenue': hotel_df['Forecasted Room Revenue This Year'].sum(),
            'total_occupancy_forecast': hotel_df['Occupancy Forecast This Year'].sum()
        }

        month_order = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
        hotel_df['Month'] = pd.Categorical(hotel_df['Month'], categories=month_order, ordered=True)
        hotel_df = hotel_df.sort_values(by=['Month', 'Occupancy Date'])

        monthly_data = {}
        grouped = hotel_df.groupby('Month', observed=False)
        for name, group in grouped:
            # *** NEW: Calculate totals for each month ***
            monthly_totals = {
                'month_revenue_st2y': group['Booked Room Revenue ST2Y'].sum(),
                'month_forecast_revenue': group['Forecasted Room Revenue This Year'].sum(),
                'month_occupancy_forecast': group['Occupancy Forecast This Year'].sum()
            }
            # *** Store both monthly totals and daily records ***
            monthly_data[name] = {
                'totals': monthly_totals,
                'records': group.to_dict(orient='records')
            }
        
        detail_data = {
            'name': hotel_name,
            'totals': overall_totals,
            'monthly_data': monthly_data,
            'columns': hotel_df.columns.tolist()
        }
        
        return render_template('results.html', hotel_data=detail_data)

    except Exception as e:
        return f"An error occurred while generating the detail view: {e}", 500

if __name__ == '__main__':
    app.run(debug=True)