import pandas as pd
from datetime import datetime

class DataProcessor:
    @staticmethod
    def load_file(file_path):
        """Load CSV or Excel file."""
        if file_path.lower().endswith('.csv'):
            encodings = ['utf-8', 'latin1', 'cp1252']
            for encoding in encodings:
                try:
                    df = pd.read_csv(file_path, encoding=encoding, header=None)
                    return df
                except UnicodeDecodeError:
                    continue
            raise ValueError("Unable to read the CSV file with supported encodings")
        elif file_path.lower().endswith('.xlsx'):
            try:
                # First try with default engine
                try:
                    df = pd.read_excel(file_path, header=None)
                    return df
                except PermissionError:
                    # If permission error, try with 'openpyxl' engine
                    df = pd.read_excel(file_path, header=None, engine='openpyxl')
                    return df
            except PermissionError:
                raise ValueError("The Excel file appears to be open in another program. Please close it and try again.")
            except Exception as e:
                raise ValueError(f"Unable to read the Excel file: {str(e)}")
        else:
            raise ValueError("Unsupported file format. Please use .csv or .xlsx files")

    @staticmethod
    def get_wells(df):
        """Extract unique wells from the dataframe."""
        if df is not None and len(df) > 0:
            # Get well names from columns starting from column D (index 3)
            well_names = df.iloc[0, 3:].tolist()  # Get from first row
            return [str(name) for name in well_names if pd.notna(name)]
        return []

    @staticmethod
    def get_well_data(df, well):
        """Get data for a specific well."""
        if df is not None and well:
            # Find the column index for the selected well
            well_col = None
            for col in range(3, df.shape[1]):  # Start from column D (index 3)
                if str(df.iloc[0, col]) == str(well):
                    well_col = col
                    break
            
            if well_col is None:
                return None

            # Get the date from row 4 (index 3)
            sample_date = str(df.iloc[3, well_col]).strip()
            try:
                date_obj = pd.to_datetime(sample_date)
                formatted_date = date_obj.strftime('%b %Y')
            except:
                formatted_date = datetime.now().strftime('%b %Y')

            # Get analytes (rows starting from row 7 until 'Notes:')
            analytes = []
            for idx, value in enumerate(df.iloc[:, 0]):
                if pd.isna(value) or str(value).strip() == 'Notes:':
                    break
                if idx >= 6:  # Start from row 7
                    analyte_name = str(value).strip()
                    analyte_value = str(df.iloc[idx, well_col]).strip()
                    
                    # Get the AWQS value from column B (index 1)
                    try:
                        awqs = float(str(df.iloc[idx, 1]).strip())
                    except (ValueError, TypeError):
                        awqs = None
                    
                    # Check if value contains 'U' and replace with 'ND'
                    if 'U' in analyte_value:
                        analyte_value = 'ND'
                    
                    # Check exceedance
                    exceeds = False
                    if analyte_value != 'ND':
                        try:
                            numeric_value = float(analyte_value)
                            if awqs is not None and numeric_value > awqs:
                                exceeds = True
                        except (ValueError, TypeError):
                            pass
                    
                    analytes.append({
                        'name': analyte_name,
                        'value': analyte_value,
                        'exceeds': exceeds
                    })
            
            return {
                'date': formatted_date,
                'analytes': analytes
            }
        return None

    @staticmethod
    def is_historical_file(df):
        """Determine if the file is a historical data file (multiple columns per well, sampling dates in row 4)."""
        if df is not None and df.shape[0] > 4:
            # Get well names from first row, starting from column D (index 3)
            well_names = df.iloc[0, 3:].tolist()
            # If there are duplicate well names, it's likely a historical file
            return len(well_names) != len(set(well_names))
        return False

    @staticmethod
    def get_well_sampling_dates(df):
        """For historical files, return a dict mapping well name to list of (col_idx, sampling_date) tuples."""
        if df is not None and df.shape[0] > 4:
            well_dates = {}
            for col in range(3, df.shape[1]):
                well = str(df.iloc[0, col]).strip()
                date = str(df.iloc[3, col]).strip()
                if well and date:
                    if well not in well_dates:
                        well_dates[well] = []
                    well_dates[well].append((col, date))
            return well_dates
        return {}