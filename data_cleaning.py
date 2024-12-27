import pandas as pd

def clean_data(data: pd.DataFrame, city_column: str) -> pd.DataFrame:
    """
    Clean the 'City' column by removing unwanted characters and using lstrip() and rstrip().
    """
    if city_column:
        data[city_column] = data[city_column].str.lstrip('_')
        data[city_column] = data[city_column].str.rstrip('_')
    return data


def clean_dates(data, date_column):
    """
    Cleans and standardizes the date column with a standardized date column in the format YYYY-MM-DD.
    """
    try:
        if date_column not in data.columns:
            raise ValueError(f"Column '{date_column}' not found in the dataset.")
        
        # Convert the date column to datetime format and standardize it
        data[date_column] = pd.to_datetime(data[date_column], errors='coerce')

        # Warn about invalid dates
        invalid_dates = data[date_column].isna().sum()
        if invalid_dates > 0:
            print(f"Warning: {invalid_dates} invalid dates were found and set to NaT.")
        
        data[date_column] = data[date_column].dt.strftime('%Y-%m-%d')
        return data

    except Exception as e:
        print(f"Error in clean_dates: {e}")
        return data


def remove_duplicate_rows(data):
    data_cleaned = data.drop_duplicates()

    # Print the cleaned data or check how many rows were removed
    print(f"Removed {len(data) - len(data_cleaned)} duplicate rows.")
    return data_cleaned
