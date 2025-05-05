# TwinTabler

TwinTabler is a web application that processes Excel files by splitting rows with odd and even IDs into separate datasets. Odd-ID rows are displayed starting from Column A with original headers, while even-ID rows are displayed starting from Column G with "_2" appended to each header.

## Features

- Upload Excel (.xlsx) files for processing
- Automatic separation of odd and even ID rows
- Formatted output with clear column labeling
- Instant file download after processing

## Requirements

- Python 3.11+
- Flask
- pandas
- openpyxl
- gunicorn (for production deployment)

## Installation

1. Clone the repository:
   ```
   git clone https://github.com/saverio46/TwinTabler.git
   cd TwinTabler
   ```

2. Install dependencies:
   ```
   pip install email-validator flask flask-sqlalchemy gunicorn openpyxl pandas psycopg2-binary
   ```

3. Run the application:
   ```
   python main.py
   ```
   
   Or for production:
   ```
   gunicorn --bind 0.0.0.0:5000 main:app
   ```

## Usage

1. Access the web interface in your browser
2. Upload an Excel file with customer data (must include ID column)
3. The application processes the file according to the specified rules
4. Download the processed file automatically

## How It Works

The application separates rows with odd and even IDs:
- Odd-ID rows are placed starting from Column A with original headers
- Even-ID rows are placed starting from Column G with "_2" appended to headers# TwinTabler
# TwinTabler
