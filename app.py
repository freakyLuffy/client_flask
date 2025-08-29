from flask import Flask, render_template, request, redirect, url_for, flash, jsonify
import pandas as pd
import os
from werkzeug.utils import secure_filename
from datetime import datetime
import json
from collections import defaultdict

app = Flask(__name__)
app.secret_key = 'your-secret-key-here'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# Ensure upload folder exists
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def parse_date(date_str):
    """Parse date string and return month-year"""
    try:
        if isinstance(date_str, str):
            date_obj = datetime.strptime(date_str, '%d-%b-%Y')
        else:
            date_obj = date_str
        return date_obj.strftime('%B %Y')
    except:
        return 'Unknown'

def process_excel_data(file_path):
    """Process Excel file and return hierarchical data structure"""
    try:
        df = pd.read_excel(file_path)
        
        # Clean column names
        df.columns = df.columns.str.strip()
        
        # Create hierarchical structure
        hierarchy = defaultdict(lambda: defaultdict(lambda: defaultdict(list)))
        
        for _, row in df.iterrows():
            property_name = row['Property Name']
            occupancy_date = row['Occupancy Date']
            business_view = row['Business View']
            
            # Parse month from date
            month_year = parse_date(occupancy_date)
            
            # Convert row to dictionary for easy access
            row_data = {
                'day_of_week': row['Day of Week'],
                'occupancy_date': str(occupancy_date),
                'occupancy_on_books': row['Occupancy On Books This Year'],
                'arrivals': row['Arrivals This Year'],
                'departures': row['Departures This Year'],
                'cancelled': row['Cancelled This Year'],
                'no_show': row['No Show This Year'],
                'booked_room_revenue': row['Booked Room Revenue This Year'],
                'adr_on_books': row['ADR On Books This Year'],
                'occupancy_forecast': row.get('Occupancy Forecast This Year', 0)
            }
            
            hierarchy[property_name][month_year][business_view].append(row_data)
        
        return dict(hierarchy)
    
    except Exception as e:
        print(f"Error processing Excel file: {str(e)}")
        return None

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        flash('No file selected')
        return redirect(request.url)
    
    file = request.files['file']
    if file.filename == '':
        flash('No file selected')
        return redirect(request.url)
    
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)
        
        # Process the Excel file
        hierarchy_data = process_excel_data(file_path)
        
        if hierarchy_data:
            return render_template('results.html', data=hierarchy_data)
        else:
            flash('Error processing the Excel file. Please check the format.')
            return redirect(url_for('index'))
    else:
        flash('Invalid file type. Please upload an Excel file (.xlsx or .xls)')
        return redirect(url_for('index'))

@app.template_filter('format_currency')
def format_currency(value):
    """Format number as currency"""
    try:
        return f"${float(value):,.2f}"
    except:
        return "$0.00"

@app.template_filter('format_number')
def format_number(value):
    """Format number with commas"""
    try:
        return f"{float(value):,.0f}"
    except:
        return "0"

if __name__ == '__main__':
    # For Render deployment
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
