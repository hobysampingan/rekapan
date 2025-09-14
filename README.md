# Excel Order Parser - SKU Quantity Summarizer

This Streamlit application allows users to upload an Excel file containing order data and generates a summarized table of quantities grouped by Seller SKU. The app provides an easy-to-print markdown table with proper alignment.

## Features

- File upload interface for Excel files (.xlsx or .xls)
- Automatic data parsing and processing
- Grouping and summing quantities by Seller SKU
- Sorted output in alphabetical order
- Printable table with professional styling
- Error handling for invalid files or missing columns

## Requirements

- Python 3.x
- streamlit
- pandas

## Installation

1. Ensure Python 3.x is installed on your system.
2. Clone or download this repository.
3. Install the required packages using pip:
   ```
   pip install -r requirements.txt
   ```

## Usage

1. Run the Streamlit app:
   ```
   streamlit run app.py
   ```
2. Open the provided URL in your web browser (usually http://localhost:8501).
3. Upload your Excel file using the file uploader.
4. The app will process the data and display the summarized table.
5. Click the "Print Table" button to open the table in a new tab and print it.

## Excel File Format

The uploaded Excel file should follow this structure:
- Row 0: Header row (e.g., "Order ID", etc.)
- Row 1: Description row (e.g., "Platform unique order ID.")
- Row 2 and below: Data rows

The app automatically maps to the correct columns for Seller SKU and Quantity.

## Error Handling

- Displays an error message if the file format is invalid or required columns are missing.
- Handles non-numeric quantity values by converting them to 0.
- Shows a message if no file is uploaded.

## License

This project is open-source. Feel free to modify and distribute.