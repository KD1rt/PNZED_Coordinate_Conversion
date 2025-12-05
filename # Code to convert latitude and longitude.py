"""
PNZED Coordinate Converter Flask App
Converts latitude/longitude to Northing/Easting (NAD83 NC State Plane)
"""
import pandas as pd
import geopandas as gpd
from shapely.geometry import Point
import os
from flask import Flask, request, render_template, send_file, redirect, url_for
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'outputs'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# Create necessary folders
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

ALLOWED_EXTENSIONS = {'csv', 'xlsx', 'xls'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def home():
    """Display the upload form"""
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert():
    """Handle file upload and coordinate conversion"""
    try:
        # Get project name from form
        project_name = request.form.get('project_name', '').strip()
        
        if not project_name:
            return "Error: Please provide a project name", 400
        
        # Check if file was uploaded
        if 'file' not in request.files:
            return "Error: No file uploaded", 400
        
        file = request.files['file']
        
        if file.filename == '':
            return "Error: No file selected", 400
        
        if not allowed_file(file.filename):
            return "Error: Invalid file type. Please upload CSV or Excel file", 400
        
        # Save uploaded file
        filename = secure_filename(file.filename)
        input_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(input_path)
        
        # Read the file based on extension
        if filename.lower().endswith('.csv'):
            df = pd.read_csv(input_path)
        else:
            df = pd.read_excel(input_path)
        
        # Validate required columns exist
        if 'x' not in df.columns or 'y' not in df.columns:
            os.remove(input_path)
            return "Error: File must contain 'x' and 'y' columns for longitude and latitude", 400
        
        # Check for valid coordinates
        if df['x'].isnull().any() or df['y'].isnull().any():
            os.remove(input_path)
            return "Error: File contains missing coordinate values", 400
        
        # Create points from x (longitude), y (latitude) coordinates
        geometry = gpd.points_from_xy(df['x'], df['y'])
        
        # Create GeoDataFrame with WGS84 (EPSG:4326) - FIXED: removed space in EPSG code
        gdf = gpd.GeoDataFrame(df, geometry=geometry, crs="EPSG:4326")
        
        # Reproject to NAD83 NC State Plane Feet 2011 (EPSG:6543)
        gdf_projected = gdf.to_crs(epsg=6543)
        
        # Extract Easting and Northing coordinates
        gdf_projected['Easting'] = gdf_projected.geometry.x
        gdf_projected['Northing'] = gdf_projected.geometry.y
        
        # Create output filename
        output_filename = f'{secure_filename(project_name)}_converted_Northing_Easting.xlsx'
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)
        
        # Export to Excel (index=False removes row numbers)
        gdf_projected.to_excel(output_path, index=False)
        
        # Clean up input file
        os.remove(input_path)
        
        # Get file size for display
        file_size = os.path.getsize(output_path)
        file_size_kb = file_size / 1024
        
        # Store output info in session-like manner (using query params for simplicity)
        # In production, consider using Flask sessions
        return render_template('results.html', 
                             filename=output_filename,
                             filesize=f"{file_size_kb:.1f} KB")
        
    except Exception as e:
        # Clean up any created files on error
        if 'input_path' in locals() and os.path.exists(input_path):
            os.remove(input_path)
        return f"Error processing file: {str(e)}", 500

@app.route('/download/<filename>')
def download(filename):
    """Download the converted file"""
    try:
        file_path = os.path.join(app.config['OUTPUT_FOLDER'], filename)
        if os.path.exists(file_path):
            return send_file(
                file_path,
                as_attachment=True,
                download_name=filename,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        else:
            return "File not found", 404
    except Exception as e:
        return f"Error downloading file: {str(e)}", 500

if __name__ == '__main__':
    print("=" * 50)
    print("PNZED Coordinate Converter")
    print("=" * 50)
    print("Starting Flask server...")
    print("Access the application at: http://127.0.0.1:5000")
    print("=" * 50)
    app.run(host='0.0.0.0', port=5000, debug=True)