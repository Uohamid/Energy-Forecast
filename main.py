import pandas as pd

def extract_excel_data(file_path):
    """
    Extract data from multiple worksheets in an Excel file.
    
    Args:
        file_path (str): Path to the Excel file
        
    Returns:
        dict: Dictionary containing DataFrames for each worksheet
    """
    try:
        # Initialize dictionary to store all data
        extracted_data = {}
        
        # Read Excel file
        xl = pd.ExcelFile(file_path)
        
        # Extract Customer Details
        if 'Customer Details' in xl.sheet_names:
            customer_details = pd.read_excel(xl, 'Customer Details')
            # Clean column names
            customer_details.columns = [col.strip() for col in customer_details.columns]
            extracted_data['customer_details'] = customer_details
        
        # Extract Export kWh
        if 'Export kWh' in xl.sheet_names:
            export_kwh = pd.read_excel(xl, 'Export kWh')
            export_kwh.columns = [col.strip() for col in export_kwh.columns]
            # Convert Date to datetime
            export_kwh['Date'] = pd.to_datetime(export_kwh['Date'])
            extracted_data['export_kwh'] = export_kwh
        
        # Extract Meteo Forecast kW
        if 'Meteo Forecast kW' in xl.sheet_names:
            meteo_forecast = pd.read_excel(xl, 'Meteo Forecast kW')
            meteo_forecast.columns = [col.strip() for col in meteo_forecast.columns]
            # Convert datetime columns if they exist
            for col in meteo_forecast.columns:
                if 'Time' in col or 'Date' in col:
                    meteo_forecast[col] = pd.to_datetime(meteo_forecast[col])
            extracted_data['meteo_forecast'] = meteo_forecast
        
        # Extract Time Lookup
        if 'Time Lookup' in xl.sheet_names:
            time_lookup = pd.read_excel(xl, 'Time Lookup')
            time_lookup.columns = [col.strip() for col in time_lookup.columns]
            extracted_data['time_lookup'] = time_lookup
        
        return extracted_data
    
    except FileNotFoundError:
        print(f"Error: File '{file_path}' not found")
        return None
    except Exception as e:
        print(f"Error occurred while reading Excel file: {str(e)}")
        return None

def print_data_summary(data_dict):
    """
    Print a summary of the extracted data.
    
    Args:
        data_dict (dict): Dictionary containing DataFrames
    """
    if not data_dict:
        print("No data to summarize")
        return
        
    for sheet_name, df in data_dict.items():
        print(f"\n{sheet_name.replace('_', ' ').title()} Summary:")
        print(f"Number of rows: {len(df)}")
        print(f"Columns: {', '.join(df.columns)}")
        print(f"First few rows:\n{df.head().to_string()}\n")

if __name__ == "__main__":
    # Example usage
    file_path = "Data.xlsx"  # Replace with actual file path
    extracted_data = extract_excel_data(file_path)
    
    if extracted_data:
        print_data_summary(extracted_data)
        
        # Example: Accessing specific data
        # Get customer details
        if 'customer_details' in extracted_data:
            print("\nCustomer Names:")
            print(extracted_data['customer_details']['Customer Name'].tolist())
        
        # Get export kWh for a specific date
        if 'export_kwh' in extracted_data:
            sample_date = pd.to_datetime('2023-04-01')
            daily_export = extracted_data['export_kwh'][
                extracted_data['export_kwh']['Date'] == sample_date
            ]
            print(f"\nExport kWh on {sample_date.date()}:")
            print(daily_export[['Period', 'kWh']].to_string())