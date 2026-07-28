# PNZED Coordinate Converter

A Flask web application that converts geographic coordinates from WGS84 latitude/longitude into North Carolina State Plane coordinates (NAD83, US Survey Feet).

The tool was designed to streamline the conversion of ArcGIS Online coordinate exports into a format commonly required for engineering, survey, and Civil 3D workflows.

---

## Key Features

- Upload CSV or Excel files
- Batch coordinate conversion
- Converts WGS84 coordinates to NC State Plane
- Outputs Easting and Northing values
- Download-ready Excel output
- Web-based interface

---

## Technology Stack

- Python
- Flask
- Pandas
- GeoPandas
- Shapely

---

## Coordinate Systems

### Input

```text
EPSG:4326
WGS84 Latitude / Longitude
```

### Output

```text
EPSG:6543
NAD83 North Carolina State Plane (US Feet)
```

---

## Workflow

1. Upload CSV or Excel file.
2. Validate coordinate fields.
3. Create point geometries from latitude and longitude.
4. Reproject points to NC State Plane.
5. Extract Easting and Northing coordinates.
6. Export results to Excel.

---

## Input Requirements

Required fields:

```text
x = Longitude
y = Latitude
```

Example:

```csv
x,y
-78.6382,35.7796
-78.8503,35.9132
```

---

## Installation

```bash
git clone https://github.com/yourusername/pnzed-coordinate-converter.git

cd pnzed-coordinate-converter

pip install flask pandas geopandas shapely openpyxl
```

Run the application:

```bash
python app.py
```

---

## Skills Demonstrated

- Coordinate reference systems (CRS)
- Datum transformations
- GIS automation
- GeoPandas workflows
- Flask web development
- Engineering and survey data preparation

---

## Use Cases

- ArcGIS Online exports
- Civil 3D data preparation
- Survey coordination
- Utility and infrastructure projects
- Construction workflows
